"""Read a workbook in memory and plan blank-only edits for a native PPT table."""
import base64
import io
import json
import re
import sys
import unicodedata
from decimal import Decimal, ROUND_HALF_UP

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


def key(value):
    return re.sub(r"[\s‘’“”'\"【】]", "", unicodedata.normalize("NFKC", str(value or ""))).casefold()


ALIASES = {"hair": ["护发", "洗护发", "头发护理"], "skin": ["护肤", "护肤品"],
           "baby": ["婴儿护理", "母婴"], "fem": ["女性护理"]}


def unique(items, label):
    if len(items) != 1:
        raise ValueError(f"{label}应有一个匹配，实际有 {len(items)} 个")
    return items[0]


def formatted(cell, is_base):
    value = cell.value
    if value is None or cell.data_type == "e":
        raise ValueError(f"{cell.coordinate} 缺少可用的源数据")
    if isinstance(value, (float, int)):
        number = Decimal(str(value))
        if is_base:
            if number != number.to_integral_value():
                raise ValueError(f"{cell.coordinate} 的 Base 不是整数")
            return str(int(number))
        if "%" in cell.number_format:
            number *= 100
        return str(number.quantize(Decimal(1), rounding=ROUND_HALF_UP)) + "%"
    return str(value).strip()


def plan(ws, prefix, grid, aliases):
    if not re.fullmatch(r"[A-Za-z0-9_]+", prefix):
        raise ValueError("题号只能包含字母、数字和下划线")
    cols = [key(x) for x in grid[0][1:]]
    rows = [key(r[0]) for r in grid[1:]]
    if any(not keys or "" in keys or len(keys) != len(set(keys)) for keys in (cols, rows)):
        raise ValueError("目标表的品类或行标签为空或重复")
    terms = {c: {c, *(key(x) for x in ALIASES.get(c, []))} for c in cols}
    for label, values in aliases.items():
        if key(label) not in terms or not isinstance(values, list) or not all(isinstance(x, str) and key(x) for x in values):
            raise ValueError("别名须按目标列名填写，例如 Hair=洗发护发")
        terms[key(label)].update(map(key, values))
    pattern = re.compile(r"(?<![A-Za-z0-9_])" + re.escape(prefix) + r"_\d+(?!\w)", re.I | re.ASCII)
    data = list(ws.iter_rows())
    markers = [(i, cell) for i, line in enumerate(data) for cell in line[:1]
               if pattern.search(str(cell.value or ""))]
    if len(markers) != len(cols):
        raise ValueError(f"找到 {len(markers)} 个题块，目标有 {len(cols)} 个品类")
    filled, conflicts, seen, same = [], [], set(), 0
    previous = -1
    seen_ids = set()
    for idx, (start, title) in enumerate(markers):
        block_id = pattern.search(str(title.value)).group().lower()
        if block_id in seen_ids:
            raise ValueError(f"重复题号 {block_id}")
        seen_ids.add(block_id)
        name = key(title.value)
        category = unique([c for c in cols if any(t in name for t in terms[c])], f"{title.coordinate} 品类")
        if category in seen:
            raise ValueError(f"品类重复 {category}")
        seen.add(category)
        # The closest preceding Base belongs to this question. Its own column header
        # must occur between the previous question and this Base.
        bases = [i for i in range(previous + 1, start) if key(data[i][0].value) == "base"]
        if not bases:
            raise ValueError(f"{title.coordinate} 上方找不到 Base")
        base = bases[-1]
        headers = [(i, j) for i in range(previous + 1, start + 1) for j, cell in enumerate(data[i])
                   if key(cell.value) == "total"]
        if not headers:
            raise ValueError(f"{title.coordinate} 找不到 Total")
        last_header = max(i for i, _ in headers)
        column = unique([j for i, j in headers if i == last_header], "Total 列")
        end = markers[idx + 1][0] if idx + 1 < len(markers) else len(data)
        source_rows = {"base": base}
        for i in range(start + 1, end):
            k = key(data[i][0].value)
            if k == "base" or any(key(c.value) == "total" for c in data[i][1:]):
                break
            if k in rows:
                if k in source_rows:
                    raise ValueError(f"源表行为重复 {data[i][0].value}")
                source_rows[k] = i
        dest_col = cols.index(category) + 1
        for dest_row, label in enumerate(rows, 1):
            if label not in source_rows:
                raise ValueError(f"{title.coordinate} 找不到行 {grid[dest_row][0]}")
            cell = data[source_rows[label]][column]
            new = formatted(cell, label == "base")
            old = grid[dest_row][dest_col].strip()
            entry = dict(row=dest_row, col=dest_col, value=new, old=old,
                         source=f"{ws.title}!{get_column_letter(column+1)}{source_rows[label]+1}")
            if not old:
                filled.append(entry)
            elif old == new or (label == "base" and old.replace(",", "") == new):
                same += 1
            else:
                conflicts.append(entry)
        previous = start
    return dict(edits=filled, conflicts=conflicts, same=same)


def main():
    arg = json.load(sys.stdin)
    book = load_workbook(io.BytesIO(base64.b64decode(arg["file"], validate=True)), data_only=True, read_only=True)
    try:
        if arg.get("transform"):
            from transform import transform
            return print(json.dumps(transform(book[arg["sheet"]], arg["transform"])))
        result = plan(book[arg["sheet"]], arg["prefix"], arg["grid"], arg.get("aliases", {}))
        print(json.dumps(result))
    finally:
        book.close()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(json.dumps({"error": str(exc)}))
        sys.exit(1)
