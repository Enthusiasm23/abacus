"""数据清洗 - 去重、缺失值处理、格式化"""

import logging
from pathlib import Path
from typing import Any

import pandas as pd

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError, ValidationError

logger = logging.getLogger(__name__)


class DataCleaningCapability(Capability):
    """数据清洗 - 去重、缺失值处理、格式化"""

    @property
    def name(self) -> str:
        return "clean_data"

    @property
    def chapter(self) -> str:
        return "grain"

    @property
    def description(self) -> str:
        return "数据清洗（去重、缺失值处理、格式化）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(name="file", type="string", description="文件路径", required=True),
            CapabilitySchema(name="sheet", type="string", description="工作表名称", required=False),
            CapabilitySchema(
                name="output", type="string", description="输出文件路径", required=False
            ),
            CapabilitySchema(
                name="operations", type="array", description="清洗操作列表", required=False
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        file_path = params.get("file")
        sheet_name = params.get("sheet")
        output = params.get("output")
        operations = params.get("operations", ["remove_duplicates", "handle_missing"])

        if not file_path:
            raise ValidationError("执行失败: 缺少必要参数 file")

        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"文件操作失败: 文件不存在 {file_path}")

        suffix = path.suffix.lower()
        if suffix in [".xlsx", ".xls"]:
            if sheet_name:
                df = pd.read_excel(path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(path, sheet_name=0)
        elif suffix == ".csv":
            df = pd.read_csv(path)
        else:
            raise DataError(f"数据操作失败: 不支持的文件格式 {suffix}")

        original_rows = len(df)
        report = {"original_rows": original_rows}

        for op in operations:
            if op == "remove_duplicates":
                df = df.drop_duplicates()
                report["duplicates_removed"] = original_rows - len(df)

            elif op == "handle_missing":
                missing_before = df.isnull().sum().sum()
                df = df.dropna()
                report["missing_removed"] = missing_before

            elif op == "strip_whitespace":
                for col in df.select_dtypes(include=["object"]).columns:
                    df[col] = df[col].str.strip()
                report["whitespace_stripped"] = True

            elif op == "convert_types":
                df = df.convert_dtypes()
                report["types_converted"] = True

        report["final_rows"] = len(df)
        report["rows_affected"] = original_rows - len(df)

        if output:
            output_path = Path(output)
            if output_path.suffix.lower() == ".csv":
                df.to_csv(output_path, index=False)
            else:
                df.to_excel(output_path, index=False)
            report["output_file"] = str(output_path)

        return report
