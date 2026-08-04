"""测试数据视图管理（扩展）"""

import pytest
import tempfile
import os
import json
import shutil
from pathlib import Path
from openpyxl import Workbook

from abacus.core.work.data_view import DataViewCapability


@pytest.fixture
def sample_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(["ID", "Name", "Region", "Sales", "Profit"])
    ws.append([1, "Product A", "East", 1000, 200])
    ws.append([2, "Product B", "West", 1500, 300])
    ws.append([3, "Product A", "East", 800, 160])
    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    wb.save(tmp.name)
    tmp.close()
    return tmp.name


@pytest.fixture
def work_file(sample_excel):
    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    tmp.close()
    shutil.copy(sample_excel, tmp.name)
    os.unlink(sample_excel)
    return tmp.name


def _cleanup(filepath):
    config_file = Path(filepath).parent / f"{Path(filepath).stem}_views.json"
    if Path(filepath).exists():
        os.unlink(filepath)
    if config_file.exists():
        os.unlink(config_file)


class TestDataViewExtended:
    def test_create_view_with_filters(self, work_file):
        cap = DataViewCapability()
        result = cap.execute(None, file=work_file, sheet="Sales", action="create",
                           view_name="EastView", columns=["ID", "Name", "Sales"],
                           filters={"Region": ["East", "West"]})
        assert result["success"] is True
        assert result["view_name"] == "EastView"
        assert result["filters"] == {"Region": ["East", "West"]}
        _cleanup(work_file)

    def test_get_view_with_filters(self, work_file):
        cap = DataViewCapability()
        cap.execute(None, file=work_file, sheet="Sales", action="create",
                   view_name="FilteredView", columns=["ID", "Name", "Sales"],
                   filters={"Region": ["East"]})
        get_result = cap.execute(None, file=work_file, sheet="Sales", action="get",
                               view_name="FilteredView")
        assert get_result["success"] is True
        assert get_result["rows"] >= 1
        assert "ID" in get_result["columns"]
        _cleanup(work_file)

    def test_list_multiple_views(self, work_file):
        cap = DataViewCapability()
        cap.execute(None, file=work_file, sheet="Sales", action="create",
                   view_name="View1", columns=["ID", "Name"])
        cap.execute(None, file=work_file, sheet="Sales", action="create",
                   view_name="View2", columns=["ID", "Sales"])
        cap.execute(None, file=work_file, sheet="Sales", action="create",
                   view_name="View3", columns=["ID", "Region"])
        list_result = cap.execute(None, file=work_file, sheet="Sales", action="list")
        assert list_result["success"] is True
        assert list_result["count"] == 3
        assert "View1" in list_result["views"]
        assert "View2" in list_result["views"]
        assert "View3" in list_result["views"]
        _cleanup(work_file)

    def test_delete_multiple_views(self, work_file):
        cap = DataViewCapability()
        cap.execute(None, file=work_file, sheet="Sales", action="create",
                   view_name="View1", columns=["ID", "Name"])
        cap.execute(None, file=work_file, sheet="Sales", action="create",
                   view_name="View2", columns=["ID", "Sales"])
        delete_result = cap.execute(None, file=work_file, sheet="Sales", action="delete",
                                   view_name="View1")
        assert delete_result["success"] is True
        list_result = cap.execute(None, file=work_file, sheet="Sales", action="list")
        assert list_result["count"] == 1
        assert "View2" in list_result["views"]
        _cleanup(work_file)

    def test_invalid_action(self, work_file):
        cap = DataViewCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=work_file, sheet="Sales", action="invalid")
        _cleanup(work_file)

    def test_create_without_view_name(self, work_file):
        cap = DataViewCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=work_file, sheet="Sales", action="create",
                       columns=["ID", "Name"])
        _cleanup(work_file)

    def test_create_without_columns(self, work_file):
        cap = DataViewCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=work_file, sheet="Sales", action="create",
                       view_name="TestView")
        _cleanup(work_file)

    def test_get_nonexistent_view(self, work_file):
        cap = DataViewCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=work_file, sheet="Sales", action="get",
                       view_name="NonexistentView")
        _cleanup(work_file)

    def test_delete_nonexistent_view(self, work_file):
        cap = DataViewCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=work_file, sheet="Sales", action="delete",
                       view_name="NonexistentView")
        _cleanup(work_file)

    def test_missing_file(self):
        cap = DataViewCapability()
        with pytest.raises(Exception):
            cap.execute(None, sheet="Sales", action="list")

    def test_missing_sheet(self, work_file):
        cap = DataViewCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=work_file, action="list")
        _cleanup(work_file)

    def test_missing_action(self, work_file):
        cap = DataViewCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=work_file, sheet="Sales")
        _cleanup(work_file)

    def test_file_not_found(self):
        cap = DataViewCapability()
        with pytest.raises(Exception):
            cap.execute(None, file="nonexistent.xlsx", sheet="Sales", action="list")

    def test_get_view_with_single_value_filter(self, work_file):
        cap = DataViewCapability()
        cap.execute(None, file=work_file, sheet="Sales", action="create",
                   view_name="EastOnly", columns=["ID", "Name", "Sales"],
                   filters={"Region": "East"})
        result = cap.execute(None, file=work_file, sheet="Sales", action="get",
                           view_name="EastOnly")
        assert result["success"] is True
        assert result["rows"] >= 1
        _cleanup(work_file)

    def test_create_view_override(self, work_file):
        cap = DataViewCapability()
        cap.execute(None, file=work_file, sheet="Sales", action="create",
                   view_name="View1", columns=["ID", "Name"])
        cap.execute(None, file=work_file, sheet="Sales", action="create",
                   view_name="View1", columns=["ID", "Sales"])
        result = cap.execute(None, file=work_file, sheet="Sales", action="get",
                           view_name="View1")
        assert "Sales" in result["columns"]
        _cleanup(work_file)

    def test_schema(self):
        cap = DataViewCapability()
        schema = cap.schema
        assert len(schema) > 0
        names = [s.name for s in schema]
        assert "file" in names
        assert "sheet" in names
        assert "action" in names
