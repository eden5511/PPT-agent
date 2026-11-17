import unittest
from openpyxl import Workbook
from workbook import discover, plan


class SourceDiscoveryTest(unittest.TestCase):
    def test_rbs_named_base_alias_and_total(self):
        wb = Workbook()
        wb.active.title = "说明"
        ws = wb.create_sheet("当期")
        for row in [["", "Total"], ["Base - 洗护发产品", "100*"],
                    ["购物行为 - 洗护发产品 [B6B8E57_01]"], ["咨询客服", 8]]:
            ws.append(row)
        grid = [["", "Hair"], ["Base", "100"], ["咨询客服", ""]]
        result = discover(wb, grid)["candidates"]
        self.assertEqual([(x["sheet"], x["prefix"]) for x in result], [("当期", "B6B8E57")])
        matched = plan(ws, result[0]["prefix"], grid, {})
        self.assertEqual(matched["same"], 1)
        self.assertEqual(matched["edits"][0]["value"], "8%")
        self.assertEqual(matched["edits"][0]["source"], "当期!B4")
