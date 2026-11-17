"""
PPT 执行器 - 根据解析后的意图，调用 python-pptx 执行具体操作
这就是"工具调用"的核心：AI 决定用什么工具、传什么参数
"""
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
import openpyxl


class PPTExecutor:
    """执行 PPT 操作的工具集合"""

    def execute(self, intent: dict, context) -> str:
        """
        主执行入口
        intent: 意图解析器输出的结构化指令
        context: 上下文管理器，包含当前 PPT 路径等信息
        """
        ppt_path = context.ppt_path
        prs = Presentation(ppt_path)

        from pathlib import Path
        from tempfile import NamedTemporaryFile
        from os import unlink
        if intent["type"] == "unknown":
            return intent.get("msg", "无法识别")
        with NamedTemporaryFile(prefix=Path(ppt_path).stem + "_修改版_", suffix=".pptx",
                                dir=Path(ppt_path).resolve().parent, delete=False) as temp:
            output = temp.name
        try:
            result = self._dispatch(prs, intent, output)
            if Path(output).stat().st_size:
                context.set_ppt(output)
                return result + "\n已保存 " + output
        except BaseException:
            unlink(output)
            raise
        unlink(output)
        return result

    def _dispatch(self, prs, intent, ppt_path):
        if intent["type"] == "cross":
            from pathlib import Path
            from addin.workbook import plan
            page = intent["page"]
            if not 1 <= page <= len(prs.slides):
                raise ValueError("页码超出范围")
            tables = [s.table for s in prs.slides[page-1].shapes if s.has_table]
            if len(tables) != 1:
                raise ValueError("目标页需要有一张可唯一确定的表格")
            table = tables[0]
            grid = [[cell.text for cell in row.cells] for row in table.rows]
            source = Path(intent["source"])
            if not source.is_absolute():
                source = Path(ppt_path).parent / source
            book = openpyxl.load_workbook(source, read_only=True, data_only=True)
            try:
                result = plan(book[intent["sheet"]], intent["prefix"], grid, {})
            finally:
                book.close()
            for edit in result["edits"]:
                cell = table.cell(edit["row"], edit["col"])
                cell.text = edit["value"]
            if result["edits"]:
                prs.save(ppt_path)
            report = f'一致 {result["same"]} 格，补空 {len(result["edits"])} 格，已有差异 {len(result["conflicts"])} 格'
            for diff in result["conflicts"]:
                report += f'\n{diff["source"]}：Excel {diff["value"]}，PPT {diff["old"]}'
            return report
        if intent["type"] == "title_text":
            page = intent["page"]
            if not 1 <= page <= len(prs.slides):
                raise ValueError("页码超出范围")
            title = prs.slides[page - 1].shapes.title
            if title is None:
                raise ValueError("该页没有标题占位符")
            title.text = intent["text"]
            prs.save(ppt_path)
            return f"已修改第 {page} 页标题"
        if intent["type"] == "set_style":
            return self._set_style(prs, intent, ppt_path)
        elif intent["type"] == "replace_text":
            return self._replace_text(prs, intent, ppt_path)
        elif intent["type"] == "fill_data":
            return self._fill_data(prs, intent, ppt_path)
        elif intent["type"] == "unknown":
            return f"❓ {intent.get('msg', '未知指令')}"
        else:
            return f"❌ 暂不支持的指令类型: {intent['type']}"

    def _hex_to_rgb(self, hex_color: str):
        """把十六进制颜色转成 RGBColor"""
        hex_color = hex_color.lstrip("#")
        return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

    def _set_style(self, prs, intent, path):
        """修改字体样式：颜色、字号"""
        target = intent["target"]
        element = intent["element"]
        color = intent.get("color")
        size = intent.get("size")

        # 确定要处理的幻灯片
        if target == "all":
            slides = list(prs.slides)
        else:
            if target > len(prs.slides):
                return f"❌ 页码 {target} 超出范围，共 {len(prs.slides)} 页"
            slides = [prs.slides[target - 1]]

        modified_count = 0

        for slide in slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue

                # 判断是否要修改这个形状
                should_modify = False
                if element == "all_text":
                    should_modify = True
                elif element == "title" and shape.is_placeholder:
                    # 标题占位符类型：1=TITLE, 3=CENTER_TITLE
                    if shape.placeholder_format.type in [1, 3]:
                        should_modify = True
                elif element == "body":
                    if shape.is_placeholder and shape.placeholder_format.type == 2:
                        should_modify = True
                else:
                    should_modify = False

                if should_modify:
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if intent.get("font"):
                                from pptx.oxml.xmlchemy import OxmlElement
                                from pptx.oxml.ns import qn
                                run.font.name = intent["font"]
                                props = run._r.get_or_add_rPr()
                                east_asian = props.find(qn("a:ea"))
                                if east_asian is None:
                                    east_asian = OxmlElement("a:ea")
                                    props.append(east_asian)
                                east_asian.set("typeface", intent["font"])
                            if size:
                                run.font.size = Pt(size)
                            if color:
                                run.font.color.rgb = self._hex_to_rgb(color)
                    modified_count += 1

        prs.save(path)
        target_desc = f"第 {target} 页" if target != "all" else "全部页面"
        return f"已修改 {target_desc} 的 {modified_count} 个文本框，字号={size}Pt，颜色=#{color}"

    def _replace_text(self, prs, intent, path):
        """批量替换文字"""
        old = intent["old_text"]
        new = intent["new_text"]

        if not old:
            return "❌ 未能识别要替换的文字"

        count = 0
        for slide in prs.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if old in run.text:
                            run.text = run.text.replace(old, new)
                            count += 1

        prs.save(path)
        return f'已替换 {count} 处 "{old}" → "{new}"' 

    def _fill_data(self, prs, intent, path):
        """从 Excel 读取数据，填充到 PPT 表格"""
        source = intent["source"]
        page = intent["page"]

        if page > len(prs.slides):
            return f"❌ 页码 {page} 超出范围"

        # 读取 Excel
        try:
            wb = openpyxl.load_workbook(source)
            ws = wb.active
            data = []
            for row in ws.iter_rows(values_only=True):
                data.append(row)
        except Exception as e:
            return f"❌ 读取 Excel 失败: {e}"

        # 找到目标页的第一个表格并填充
        slide = prs.slides[page - 1]
        table_found = False

        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                table = shape.table
                table_found = True
                filled = 0
                for i, row_data in enumerate(data):
                    if i >= len(table.rows):
                        break
                    for j, val in enumerate(row_data):
                        if j >= len(table.columns):
                            break
                        if val is not None:
                            table.cell(i, j).text = str(val)
                            filled += 1
                prs.save(path)
                return f"✅ 已从 {source} 填充 {filled} 个单元格到第 {page} 页表格"

        if not table_found:
            return f"⚠️ 第 {page} 页没有找到表格，请先在 PPT 里插入表格"
