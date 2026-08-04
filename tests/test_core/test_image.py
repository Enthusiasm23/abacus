"""测试插入图片到单元格"""

import pytest
import tempfile
import os
from pathlib import Path
from openpyxl import Workbook
from PIL import Image as PILImage

from abacus.core.work.image import InsertImageCapability


@pytest.fixture
def sample_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["ID", "Name"])
    ws.append([1, "Product A"])
    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    wb.save(tmp.name)
    tmp.close()
    return tmp.name


@pytest.fixture
def sample_image():
    """创建一个简单的测试图片"""
    img = PILImage.new('RGB', color='red', size=(100, 100))
    tmp = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
    img.save(tmp.name)
    tmp.close()
    return tmp.name


class TestInsertImage:
    def test_capability_properties(self):
        cap = InsertImageCapability()
        assert cap.name == "insert_excel_image"
        assert cap.chapter == "work"
        assert "图片" in cap.description

    def test_missing_file(self):
        cap = InsertImageCapability()
        with pytest.raises(Exception):
            cap.execute(None, sheet="Sheet1", cell="A1", image_path="test.png")

    def test_missing_sheet(self, sample_excel):
        cap = InsertImageCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=sample_excel, cell="A1", image_path="test.png")
        os.unlink(sample_excel)

    def test_missing_cell(self, sample_excel):
        cap = InsertImageCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=sample_excel, sheet="Sheet1", image_path="test.png")
        os.unlink(sample_excel)

    def test_missing_image_path(self, sample_excel):
        cap = InsertImageCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=sample_excel, sheet="Sheet1", cell="A1")
        os.unlink(sample_excel)

    def test_file_not_found(self, sample_image):
        cap = InsertImageCapability()
        with pytest.raises(Exception):
            cap.execute(None, file="nonexistent.xlsx", sheet="Sheet1",
                       cell="A1", image_path=sample_image)
        os.unlink(sample_image)

    def test_image_not_found(self, sample_excel):
        cap = InsertImageCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=sample_excel, sheet="Sheet1",
                       cell="A1", image_path="nonexistent.png")
        os.unlink(sample_excel)

    def test_sheet_not_found(self, sample_excel, sample_image):
        cap = InsertImageCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=sample_excel, sheet="Nonexistent",
                       cell="A1", image_path=sample_image)
        os.unlink(sample_excel)
        os.unlink(sample_image)

    def test_insert_image_success(self, sample_excel, sample_image):
        cap = InsertImageCapability()
        result = cap.execute(None, file=sample_excel, sheet="Sheet1",
                           cell="A3", image_path=sample_image)
        assert result["success"] is True
        assert result["cell"] == "A3"
        assert "width" in result
        assert "height" in result
        os.unlink(sample_excel)
        os.unlink(sample_image)

    def test_insert_image_with_dimensions(self, sample_excel, sample_image):
        cap = InsertImageCapability()
        result = cap.execute(None, file=sample_excel, sheet="Sheet1",
                           cell="B1", image_path=sample_image,
                           width=200, height=150)
        assert result["success"] is True
        assert result["width"] == 200
        assert result["height"] == 150
        os.unlink(sample_excel)
        os.unlink(sample_image)

    def test_schema(self):
        cap = InsertImageCapability()
        schema = cap.schema
        assert len(schema) == 6
        names = [s.name for s in schema]
        assert "file" in names
        assert "sheet" in names
        assert "cell" in names
        assert "image_path" in names
        assert "width" in names
        assert "height" in names
