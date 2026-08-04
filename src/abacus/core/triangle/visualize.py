"""勾股章 - 数据可视化：用 matplotlib 生成图表"""

import logging
from pathlib import Path
from typing import Any

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError, ValidationError

logger = logging.getLogger(__name__)


class VisualizeCapability(Capability):
    """数据可视化：用 matplotlib 生成图表"""

    @property
    def name(self) -> str:
        return "visualize"

    @property
    def chapter(self) -> str:
        return "triangle"

    @property
    def description(self) -> str:
        return "数据可视化（生成 PNG/SVG/PDF 图表）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(name="file", type="string", description="数据文件路径", required=True),
            CapabilitySchema(name="sheet", type="string", description="工作表名称", required=False),
            CapabilitySchema(
                name="chart_type",
                type="string",
                description="图表类型（bar/line/pie/scatter/heatmap）",
                required=True,
            ),
            CapabilitySchema(name="x_column", type="string", description="X轴列名", required=False),
            CapabilitySchema(name="y_column", type="string", description="Y轴列名", required=False),
            CapabilitySchema(
                name="output", type="string", description="输出图片路径", required=True
            ),
            CapabilitySchema(name="title", type="string", description="图表标题", required=False),
            CapabilitySchema(
                name="width", type="number", description="图片宽度（英寸）", required=False
            ),
            CapabilitySchema(
                name="height", type="number", description="图片高度（英寸）", required=False
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        file_path = params.get("file")
        sheet_name = params.get("sheet")
        chart_type = params.get("chart_type")
        x_column = params.get("x_column")
        y_column = params.get("y_column")
        output = params.get("output")
        title = params.get("title")
        width = params.get("width", 10)
        height = params.get("height", 6)

        if not file_path:
            raise ValidationError("执行失败: 缺少必要参数 file")
        if not output:
            raise ValidationError("执行失败: 缺少必要参数 output")

        try:
            import matplotlib

            matplotlib.use("Agg")
            import matplotlib.pyplot as plt
            import pandas as pd
        except ImportError:
            raise DataError("数据操作失败: matplotlib 未安装，请运行: pip install matplotlib pandas")

        # 加载数据
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"文件操作失败: 文件不存在 {file_path}")

        suffix = path.suffix.lower()
        if suffix in [".xlsx", ".xls"]:
            df = pd.read_excel(path, sheet_name=sheet_name or 0)
        elif suffix == ".csv":
            df = pd.read_csv(path)
        else:
            raise DataError(f"数据操作失败: 不支持的文件格式 {suffix}")

        # 创建图表
        fig, ax = plt.subplots(figsize=(width, height))

        if chart_type == "bar":
            self._create_bar_chart(ax, df, x_column, y_column)
        elif chart_type == "line":
            self._create_line_chart(ax, df, x_column, y_column)
        elif chart_type == "pie":
            self._create_pie_chart(ax, df, x_column, y_column)
        elif chart_type == "scatter":
            self._create_scatter_chart(ax, df, x_column, y_column)
        elif chart_type == "heatmap":
            self._create_heatmap(fig, ax, df)
        else:
            raise DataError(f"数据操作失败: 不支持的图表类型 {chart_type}")

        if title:
            ax.set_title(title)

        plt.tight_layout()
        plt.savefig(output, dpi=150, bbox_inches="tight")
        plt.close()

        return {"file": file_path, "output": output, "chart_type": chart_type, "created": True}

    def _create_bar_chart(self, ax, df, x_col, y_col):
        """创建柱状图"""
        if x_col and y_col:
            ax.bar(df[x_col], df[y_col])
            ax.set_xlabel(x_col)
            ax.set_ylabel(y_col)
        else:
            df.plot(kind="bar", ax=ax)

    def _create_line_chart(self, ax, df, x_col, y_col):
        """创建折线图"""
        if x_col and y_col:
            ax.plot(df[x_col], df[y_col], marker="o")
            ax.set_xlabel(x_col)
            ax.set_ylabel(y_col)
        else:
            df.plot(kind="line", ax=ax, marker="o")

    def _create_pie_chart(self, ax, df, x_col, y_col):
        """创建饼图"""
        if x_col and y_col:
            ax.pie(df[y_col], labels=df[x_col], autopct="%1.1f%%")
        else:
            df.iloc[:, 0].value_counts().plot(kind="pie", ax=ax, autopct="%1.1f%%")

    def _create_scatter_chart(self, ax, df, x_col, y_col):
        """创建散点图"""
        if x_col and y_col:
            ax.scatter(df[x_col], df[y_col])
            ax.set_xlabel(x_col)
            ax.set_ylabel(y_col)
        else:
            df.plot(kind="scatter", x=df.columns[0], y=df.columns[1], ax=ax)

    def _create_heatmap(self, fig, ax, df):
        """创建热力图"""
        import numpy as np

        numeric_df = df.select_dtypes(include=[np.number])
        if numeric_df.empty:
            raise DataError("数据操作失败: 未找到数值列，无法创建热力图")
        im = ax.imshow(numeric_df.corr(), cmap="coolwarm", aspect="auto")
        ax.set_xticks(range(len(numeric_df.columns)))
        ax.set_yticks(range(len(numeric_df.columns)))
        ax.set_xticklabels(numeric_df.columns, rotation=45, ha="right")
        ax.set_yticklabels(numeric_df.columns)
        import matplotlib.pyplot as _plt

        _plt.colorbar(im, ax=ax)
