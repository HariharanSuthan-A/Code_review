"""
Unit tests for WorkbookManager. These don't touch the LLM/agent layer at
all — they verify the deterministic Excel operations the agent relies on
actually behave correctly, since a tool bug is invisible to the LLM and
will just produce confidently wrong spreadsheets.

Run with:  pytest tests/test_workbook_manager.py -v
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "backend"))

import pytest

from app.excel.workbook_manager import WorkbookManager, WorkbookError


@pytest.fixture
def wm(tmp_path):
    manager = WorkbookManager(workbooks_dir=str(tmp_path))
    manager.new_workbook("test.xlsx")
    return manager


def test_new_workbook_has_default_sheet(wm):
    summary = wm.get_summary()
    assert len(summary["sheets"]) == 1


def test_add_row_and_read_back(wm):
    wm.add_row(["Name", "Age", "Salary"])
    wm.add_row(["Arun", 24, 35000])
    preview = wm.sheet_preview()
    assert preview["preview"][0] == ["Name", "Age", "Salary"]
    assert preview["preview"][1] == ["Arun", 24, 35000]


def test_find_row_by_column_value(wm):
    wm.add_row(["Name", "Age", "Salary"])
    wm.add_row(["Arun", 24, 35000])
    wm.add_row(["Meena", 27, 42000])
    row = wm.find_row_by_value("Name", "Meena")
    assert row == 3


def test_update_row_by_header(wm):
    wm.add_row(["Name", "Age", "Salary"])
    wm.add_row(["Arun", 24, 35000])
    wm.update_row(2, {"Salary": 40000})
    preview = wm.sheet_preview()
    assert preview["preview"][1][2] == 40000


def test_delete_row(wm):
    wm.add_row(["Name"])
    wm.add_row(["Arun"])
    wm.add_row(["Meena"])
    wm.delete_row(2)
    preview = wm.sheet_preview()
    assert preview["preview"] == [["Name"], ["Meena"]]


def test_add_column(wm):
    wm.add_row(["Name"])
    wm.add_row(["Arun"])
    wm.add_column("Department", ["AI"])
    headers = wm.get_headers()
    assert headers == ["Name", "Department"]


def test_create_rename_delete_sheet(wm):
    wm.create_sheet("Summary")
    assert "Summary" in wm.list_sheets()
    wm.rename_sheet("Summary", "Totals")
    assert "Totals" in wm.list_sheets()
    assert "Summary" not in wm.list_sheets()
    wm.delete_sheet("Totals")
    assert "Totals" not in wm.list_sheets()


def test_cannot_delete_last_sheet(wm):
    only_sheet = wm.list_sheets()[0]
    with pytest.raises(WorkbookError):
        wm.delete_sheet(only_sheet)


def test_apply_formula(wm):
    wm.add_row(["Name", "Salary"])
    wm.add_row(["Arun", 35000])
    wm.add_row(["Meena", 42000])
    result = wm.apply_simple_formula("B4", "SUM", "B2:B3")
    assert result["formula"] == "=SUM(B2:B3)"


def test_apply_general_formula(wm):
    wm.add_row(["Name", "Salary"])
    wm.add_row(["Arun", 35000])
    wm.add_row(["Meena", 42000])
    result = wm.apply_formula("C2", '=IF(B2>30000,"High","Low")')
    # Backend-computed: the ANSWER is stored, never the raw '=...' text.
    assert wm.read_range("C2") == [["High"]]
    assert wm._ws(None)["C2"].value == "High"
    assert result["formula"] == '=IF(B2>30000,"High","Low")'
    assert result["value"] == "High"


def test_formula_preview_shows_answer_not_formula_text(wm):
    wm.add_row(["Rating"])
    for v in ["Excellent", "Good", "Excellent"]:
        wm.add_row([v])
    formula = '=CONCATENATE("Excellent: ",COUNTIF(A2:A4,"Excellent"),", Good: ",COUNTIF(A2:A4,"Good"))'
    result = wm.apply_formula("B2", formula)
    assert result["value"] == "Excellent: 2, Good: 1"
    assert wm.read_range("B2") == [["Excellent: 2, Good: 1"]]
    assert wm.sheet_preview(max_rows=5)[ "preview"][1][1] == "Excellent: 2, Good: 1"
    # backend-computed answer stored (no raw formula pasted into the cell)
    assert wm._ws(None)["B2"].value == "Excellent: 2, Good: 1"


def test_compute_formula_answers_in_chat_without_writing(wm):
    wm.add_row(["Name", "Salary"])
    wm.add_row(["Arun", 35000])
    wm.add_row(["Meena", 42000])
    result = wm.compute_formula("=SUM(B2:B3)")
    assert result["value"] == 77000
    # nothing was written to the sheet
    assert wm.read_range("C2") == [[None]]


def test_apply_formula_rejects_bad_syntax(wm):
    wm.add_row(["Name"])
    with pytest.raises(WorkbookError):
        wm.apply_formula("A2", "SUM(B1:B2)")  # missing leading '='
    with pytest.raises(WorkbookError):
        wm.apply_formula("A2", "=SUM(B1:B2")  # unbalanced parens


def test_sort_range(wm):
    wm.add_row(["Name", "Salary"])
    wm.add_row(["Arun", 35000])
    wm.add_row(["Meena", 42000])
    wm.add_row(["Kavin", 28000])
    wm.sort_range("Salary", ascending=False)
    preview = wm.sheet_preview()
    salaries = [row[1] for row in preview["preview"][1:]]
    assert salaries == [42000, 35000, 28000]


def test_filter_rows(wm):
    wm.add_row(["Name", "Salary"])
    wm.add_row(["Arun", 35000])
    wm.add_row(["Meena", 42000])
    wm.add_row(["Kavin", 28000])
    results = wm.filter_rows("Salary", "gt", 30000)
    names = sorted(r["values"][0] for r in results)
    assert names == ["Arun", "Meena"]


def test_undo_reverts_last_change(wm):
    wm.add_row(["Name"])
    wm.add_row(["Arun"])
    before = wm.sheet_preview()["preview"]
    wm.delete_row(2)
    wm.undo()
    after = wm.sheet_preview()["preview"]
    assert before == after


def test_format_cells_does_not_raise(wm):
    wm.add_row(["Name", "Salary"])
    wm.add_row(["Arun", 35000])
    result = wm.format_cells("B2:B2", bold=True, fill_color="D9F2E3", number_format='"$"#,##0.00')
    assert result["range"] == "B2:B2"


def test_save_and_reopen(wm, tmp_path):
    wm.add_row(["Name"])
    wm.add_row(["Arun"])
    saved = wm.save_workbook(save_as="roundtrip.xlsx")
    path = saved["saved_path"]
    assert os.path.exists(path)

    wm2 = WorkbookManager(workbooks_dir=str(tmp_path))
    wm2.open_workbook(path)
    preview = wm2.sheet_preview()
    assert preview["preview"] == [["Name"], ["Arun"]]


def test_insert_column_before_and_after(wm):
    wm.add_row(["Name", "Salary"])
    wm.add_row(["Arun", 35000])
    wm.insert_column_before("Salary", "Dept", ["AI"])
    assert wm.get_headers() == ["Name", "Dept", "Salary"]
    assert wm.read_range("B2") == [["AI"]]
    wm.insert_column_after("Salary", "Bonus", [5000])
    assert wm.get_headers() == ["Name", "Dept", "Salary", "Bonus"]
    assert wm.read_range("D2") == [[5000]]


def test_insert_row_before_and_after(wm):
    wm.add_row(["Name"])
    wm.add_row(["Arun"])
    wm.add_row(["Meena"])
    wm.insert_row_before(2, ["Inserted"])
    assert wm.sheet_preview()["preview"] == [["Name"], ["Inserted"], ["Arun"], ["Meena"]]
    wm.insert_row_after(2, ["After"])
    assert wm.sheet_preview()["preview"][2] == ["After"]


def test_move_column_and_row(wm):
    wm.add_row(["A", "B", "C"])
    wm.add_row([1, 2, 3])
    wm.move_column("A", 3)
    assert wm.get_headers() == ["B", "C", "A"]
    wm.move_row(2, 2)  # no-op
    assert wm.sheet_preview()["preview"][1] == [2, 3, 1]


def test_copy_range_and_clear_range(wm):
    wm.add_row(["Name", "Salary"])
    wm.add_row(["Arun", 35000])
    wm.copy_range("A1:B2", "D1")
    assert wm.read_range("D1:E2") == [["Name", "Salary"], ["Arun", 35000]]
    wm.clear_range("D1:E2")
    assert wm.read_range("D1:E2") == [[None, None], [None, None]]
    # original untouched
    assert wm.read_range("A1") == [["Name"]]


def test_merge_and_unmerge(wm):
    wm.add_row(["a", "b", "c"])
    wm.merge_cells("A1:C1")
    ws = wm._ws(None)
    assert "A1:C1" in [str(r) for r in ws.merged_cells.ranges]
    wm.unmerge_cells("A1:C1")
    assert len(list(ws.merged_cells.ranges)) == 0


def test_freeze_panes_and_dimensions(wm):
    wm.add_row(["a", "b"])
    wm.freeze_panes("B2")
    assert wm._ws(None).freeze_panes == "B2"
    wm.freeze_panes("A1")
    assert wm._ws(None).freeze_panes is None
    wm.set_column_width("A", 25)
    assert wm._ws(None).column_dimensions["A"].width == 25
    wm.set_row_height(1, 30)
    assert wm._ws(None).row_dimensions[1].height == 30
    result = wm.autofit_columns()
    assert "A" in result["widths"]


def test_duplicate_sheet_and_comment(wm):
    wm.add_row(["Name"])
    wm.duplicate_sheet(wm.list_sheets()[0], "Copy")
    assert "Copy" in wm.list_sheets()
    wm.set_cell_comment("A1", "hello")
    assert wm._ws(None)["A1"].comment.text == "hello"


def test_conditional_format_and_validation(wm):
    wm.add_row(["Score"])
    wm.add_row([10])
    wm.add_row([200])
    assert wm.conditional_format("A2:A3", "greater_than", threshold=100)["rule_type"] == "greater_than"
    assert wm.conditional_format("A2:A3", "between", minimum=0, maximum=300)["rule_type"] == "between"
    assert wm.conditional_format("A2:A3", "color_scale")["rule_type"] == "color_scale"
    assert wm.conditional_format("A2:A3", "data_bars")["rule_type"] == "data_bars"
    assert wm.conditional_format("A2:A3", "duplicate_values")["rule_type"] == "duplicate_values"
    assert wm.data_validation("A2:A3", "list", options=["a", "b"])["validation_type"] == "list"
    assert wm.data_validation("A2:A3", "whole_number", minimum=1, maximum=10)["validation_type"] == "whole_number"
    assert wm.data_validation("A2:A3", "decimal", minimum=0.5, maximum=9.5)["validation_type"] == "decimal"


def test_suggest_formula_returns_context(wm):
    wm.add_row(["Region", "Sales"])
    wm.add_row(["West", 100])
    ctx = wm.suggest_formula("total sales where region is West")
    assert ctx["headers"] == ["Region", "Sales"]
    assert "SUMIFS" in ctx["guidance"]


def test_structural_ops_support_undo(wm):
    wm.add_row(["A", "B"])
    before = wm.sheet_preview()["preview"]
    wm.insert_column_before("B", "X")
    wm.undo()
    assert wm.sheet_preview()["preview"] == before


def test_delete_row_out_of_range_raises(wm):
    wm.add_row(["Name"])
    wm.add_row(["Arun"])
    with pytest.raises(WorkbookError):
        wm.delete_row(99)
    with pytest.raises(WorkbookError):
        wm.delete_row(0)
    # failed deletes must not pollute the undo stack
    assert wm.sheet_preview()["preview"] == [["Name"], ["Arun"]]


def test_update_row_out_of_range_raises(wm):
    wm.add_row(["Name"])
    wm.add_row(["Arun"])
    with pytest.raises(WorkbookError):
        wm.update_row(99, {"Name": "X"})
    with pytest.raises(WorkbookError):
        wm.update_row(0, {"Name": "X"})


def test_delete_column_out_of_range_raises(wm):
    wm.add_row(["A", "B"])
    with pytest.raises(WorkbookError):
        wm.delete_column("ZZ")
    with pytest.raises(WorkbookError):
        wm.delete_column("Nope")


def test_move_row_missing_source_raises(wm):
    wm.add_row(["A"])
    wm.add_row(["x"])
    with pytest.raises(WorkbookError):
        wm.move_row(99, 1)


def test_set_row_height_rejects_bad_row(wm):
    wm.add_row(["A"])
    with pytest.raises(WorkbookError):
        wm.set_row_height(0, 20)


def test_write_cell_rejects_bad_coordinate(wm):
    with pytest.raises(WorkbookError):
        wm.write_cell("not-a-cell", "x")
    with pytest.raises(WorkbookError):
        wm.write_range("bogus", [[1]])


def test_new_workbook_sanitizes_traversal(wm, tmp_path):
    summary = wm.new_workbook("../../tmp/pwned.xlsx")
    assert summary["path"] == os.path.join(str(tmp_path), "pwned.xlsx")
    assert os.path.abspath(summary["path"]).startswith(os.path.abspath(str(tmp_path)))
    with pytest.raises(WorkbookError):
        wm.new_workbook("..")


def test_save_as_cannot_escape_workbooks_dir(wm, tmp_path):
    wm.add_row(["A"])
    saved = wm.save_workbook(save_as="ok.xlsx")
    assert os.path.abspath(saved["saved_path"]).startswith(os.path.abspath(str(tmp_path)))


def test_analyze_normalizes_blank_headers(wm):
    wm.add_row(["Name", "Salary"])
    wm.add_row(["Arun", 35000])
    wm.write_cell("C2", "extra")  # data with no header above it -> blank header
    summary = wm.analyze()
    assert "Column_3" in summary["columns"]
    assert summary["numeric_summary"]["Salary"]["count"] == 1.0
    # a stray numeric value in the header row must not crash analysis either
    wm.apply_simple_formula("D1", "SUM", "B2:B2")
    assert wm.analyze()["row_count"] == 1
