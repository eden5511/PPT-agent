"""Transform a named Excel range using the supported Chinese command grammar."""
import re
from decimal import Decimal, InvalidOperation


def transform(ws, command):
    grammar = (r"区域([A-Za-z]+[1-9]\d*:[A-Za-z]+[1-9]\d*)"
               r"(?:，筛选([^，=]+)=([^，]+))?"
               r"(?:，按([^，]+?)分组对([^，]+?)求和)?"
               r"(?:，按([^，]+?)(升序|降序))?"
               r"(?:，选取([^，]+))?")
    match = re.fullmatch(grammar, command.strip())
    if not match:
        raise ValueError("区域、筛选、分组、排序或选列语法无法识别")
    area, field, value, group, metric, sort, direction, select = match.groups()
    cells = list(ws[area.upper()])
    header = [str(c.value or "").strip() for c in cells[0]]
    if "" in header or len(header) != len(set(header)):
        raise ValueError("区域首行需要唯一且非空的列名")
    records = [dict(zip(header, [c.value for c in row])) for row in cells[1:] if any(c.value is not None for c in row)]
    for column in (field, group, metric):
        if column and column not in header:
            raise ValueError(f"未找到列 {column}")
    if field:
        records = [r for r in records if str(r[field]).strip() == value]
    if group:
        sums = {}
        for r in records:
            if r[group] is None:
                raise ValueError("存在空分组")
            try:
                sums[r[group]] = sums.get(r[group], Decimal(0)) + Decimal(str(r[metric]))
            except InvalidOperation as exc:
                raise ValueError(f"{metric} 含非数值") from exc
        header = [group, metric]
        records = [{group: k, metric: v} for k, v in sums.items()]
    if sort:
        if sort not in header:
            raise ValueError(f"排序列不存在：{sort}")
        records.sort(key=lambda r: Decimal(str(r[sort])), reverse=direction == "降序")
    columns = select.split("、") if select else header
    if not all(c in header for c in columns):
        raise ValueError("选取的列不存在")
    if not records:
        raise ValueError("区域内没有符合条件的数据")
    return {"matrix": [columns] + [["" if r[c] is None else str(r[c]) for c in columns] for r in records]}
