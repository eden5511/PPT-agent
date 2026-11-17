import unittest
from openpyxl import Workbook
from workbook import plan


class AlignmentTest(unittest.TestCase):
    def setUp(self):
        self.sheet = Workbook().active
        for row in [
            ["", "Total", "其他"], ["Base", 100, 90], ["Q_01 护发"],
            ["行为二", "-", 1], ["行为一", .08, 5],
            ["", "其他", "Total"], ["Base", 3, 200], ["Q_02 护肤"],
            ["行为一", 9, .16], ["行为二", 8, 0]
        ]:
            self.sheet.append(row)
        for addr in ["B5","C9","C10"]:
            self.sheet[addr].number_format = "0%"
        self.grid = [["", "Skin", "Hair"], ["Base", "200", "100"],
                     ["行为一", "旧值", ""], ["行为二", "0%", "-"]]

    def test_fill_and_conflict(self):
        result = plan(self.sheet, "Q", self.grid, {})
        self.assertEqual(result["same"], 4)
        self.assertEqual([(e["row"], e["col"], e["value"], e["source"]) for e in result["edits"]],
                         [(2, 2, "8%", "Sheet!B5")])
        self.assertEqual(result["conflicts"][0]["value"], "16%")

    def test_duplicate_target(self):
        self.grid[0][1] = "Hair"
        with self.assertRaises(ValueError):
            plan(self.sheet, "Q", self.grid, {})

class TransformTest(unittest.TestCase):
    def test_filter_group_sort_project(self):
        from transform import transform
        ws = Workbook().active
        for row in [["城市","品牌","销量"],["上海","A",2],["北京","B",100],["上海","B",4],["上海","A",3]]:
            ws.append(row)
        result = transform(ws, "区域A1:C5，筛选城市=上海，按品牌分组对销量求和，按销量降序，选取品牌、销量")
        self.assertEqual(result["matrix"], [["品牌","销量"],["A","5"],["B","4"]])
