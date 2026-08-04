"""CSV/Excel 合并 - 智能列匹配、去重、冲突解决"""

import logging
from pathlib import Path
from typing import Any

import pandas as pd

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError, ValidationError

logger = logging.getLogger(__name__)


class CSVMergeCapability(Capability):
    """CSV/Excel 合并 - 智能列匹配、去重、冲突解决"""

    @property
    def name(self) -> str:
        return "merge_files"

    @property
    def chapter(self) -> str:
        return "transport"

    @property
    def description(self) -> str:
        return "合并多个 CSV/Excel 文件（智能列匹配、去重、冲突解决）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(name="files", type="array", description="文件路径列表", required=True),
            CapabilitySchema(
                name="output", type="string", description="输出文件路径", required=True
            ),
            CapabilitySchema(
                name="merge_type",
                type="string",
                description="合并类型（concat/merge/join）",
                required=False,
            ),
            CapabilitySchema(
                name="on", type="string", description="合并键（merge/join 时）", required=False
            ),
            CapabilitySchema(name="dedup", type="boolean", description="是否去重", required=False),
            CapabilitySchema(
                name="dedup_columns", type="array", description="去重列", required=False
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        files = params.get("files", [])
        output = params.get("output")
        merge_type = params.get("merge_type", "concat")
        on = params.get("on")
        dedup = params.get("dedup", False)
        dedup_columns = params.get("dedup_columns")

        if not files:
            raise ValidationError("执行失败: 缺少必要参数 files")
        if not output:
            raise ValidationError("执行失败: 缺少必要参数 output")

        # 加载所有文件
        dataframes = []
        for file_path in files:
            path = Path(file_path)
            if not path.exists():
                raise FileNotFoundError(f"文件操作失败: 文件不存在 {file_path}")

            suffix = path.suffix.lower()
            if suffix == ".csv":
                df = pd.read_csv(path)
            elif suffix in [".xlsx", ".xls"]:
                df = pd.read_excel(path)
            else:
                raise DataError(f"数据操作失败: 不支持的文件格式 {suffix}")

            dataframes.append(df)

        # 执行合并
        if merge_type == "concat":
            result = pd.concat(dataframes, ignore_index=True)
        elif merge_type == "merge":
            if not on:
                raise ValidationError("执行失败: merge 合并需要 on 参数")
            result = dataframes[0]
            for df in dataframes[1:]:
                result = pd.merge(result, df, on=on, how="outer")
        elif merge_type == "join":
            if not on:
                raise ValidationError("执行失败: join 合并需要 on 参数")
            result = dataframes[0]
            for df in dataframes[1:]:
                result = result.join(df.set_index(on), on=on)
        else:
            raise DataError(f"数据操作失败: 不支持的合并类型 {merge_type}")

        # 去重
        if dedup:
            before_count = len(result)
            if dedup_columns:
                result = result.drop_duplicates(subset=dedup_columns)
            else:
                result = result.drop_duplicates()
            duplicates_removed = before_count - len(result)
        else:
            duplicates_removed = 0

        # 保存
        output_path = Path(output)
        if output_path.suffix.lower() == ".csv":
            result.to_csv(output_path, index=False)
        else:
            result.to_excel(output_path, index=False)

        return {
            "output": output,
            "files_merged": len(files),
            "total_rows": len(result),
            "total_columns": len(result.columns),
            "duplicates_removed": duplicates_removed,
            "merge_type": merge_type,
        }
