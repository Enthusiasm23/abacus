"""商功章 - 高级图表：用 xlsxwriter 实现高级图表"""

import logging
from typing import Any

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError

logger = logging.getLogger(__name__)


class AdvancedChartCapability(Capability):
    """高级图表：用 xlsxwriter 创建高级图表"""

    @property
    def name(self) -> str:
        return "create_advanced_chart"

    @property
    def chapter(self) -> str:
        return "work"

    @property
    def description(self) -> str:
        return "创建高级图表（组合图、双轴图、瀑布图、甘特图）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(name="file", type="string", description="输出文件路径", required=True),
            CapabilitySchema(name="data", type="object", description="图表数据", required=True),
            CapabilitySchema(
                name="chart_type",
                type="string",
                description="图表类型（combo/dual_axis/waterfall/gantt）",
                required=True,
            ),
            CapabilitySchema(name="title", type="string", description="图表标题", required=False),
            CapabilitySchema(name="x_axis", type="string", description="X轴标题", required=False),
            CapabilitySchema(name="y_axis", type="string", description="Y轴标题", required=False),
        ]

    def execute(self, context: Any, **params) -> Any:
        file_path = params.get("file")
        data = params.get("data", {})
        chart_type = params.get("chart_type")
        title = params.get("title")
        x_axis = params.get("x_axis")
        y_axis = params.get("y_axis")

        if not file_path:
            raise DataError("file parameter is required")

        try:
            import xlsxwriter

            self._xlsxwriter = xlsxwriter
        except ImportError:
            raise DataError("xlsxwriter not installed. Run: pip install xlsxwriter")

        return self._create_chart(file_path, data, chart_type, title, x_axis, y_axis)

    def _create_chart(
        self,
        filepath: str,
        data: dict,
        chart_type: str,
        title: str = None,
        x_axis: str = None,
        y_axis: str = None,
    ) -> dict:
        """创建高级图表"""
        try:
            wb = self._xlsxwriter.Workbook(filepath)

            # 创建数据工作表
            ws_data = wb.add_worksheet("Data")
            headers = data.get("headers", [])
            rows = data.get("rows", [])

            for col, header in enumerate(headers):
                ws_data.write(0, col, header)
            for row_idx, row in enumerate(rows, 1):
                for col_idx, value in enumerate(row):
                    ws_data.write(row_idx, col_idx, value)

            # 创建图表工作表
            ws_chart = wb.add_worksheet("Charts")

            if chart_type == "combo":
                self._create_combo_chart(wb, ws_chart, data, title)
            elif chart_type == "dual_axis":
                self._create_dual_axis_chart(wb, ws_chart, data, title)
            elif chart_type == "waterfall":
                self._create_waterfall_chart(wb, ws_chart, data, title)
            elif chart_type == "gantt":
                self._create_gantt_chart(wb, ws_chart, data, title)
            else:
                raise DataError(f"Unknown chart type: {chart_type}")

            wb.close()

            return {"file": filepath, "chart_type": chart_type, "title": title, "created": True}

        except Exception as e:
            logger.error(f"Failed to create chart: {e}")
            raise DataError(str(e))

    def _create_combo_chart(self, wb, ws, data: dict, title: str):
        """创建组合图"""
        chart1 = wb.add_chart({"type": "column"})
        chart2 = wb.add_chart({"type": "line"})

        headers = data.get("headers", [])
        rows = data.get("rows", [])

        if len(headers) >= 2:
            chart1.add_series(
                {
                    "name": headers[1],
                    "categories": f"=Data!$A$2:$A${len(rows) + 1}",
                    "values": f"=Data!$B$2:$B${len(rows) + 1}",
                }
            )

        if len(headers) >= 3:
            chart2.add_series(
                {
                    "name": headers[2],
                    "categories": f"=Data!$A$2:$A${len(rows) + 1}",
                    "values": f"=Data!$C$2:$C${len(rows) + 1}",
                }
            )
            chart1.combine(chart2)

        chart1.set_title({"name": title or "组合图"})
        chart1.set_x_axis({"name": headers[0] if headers else ""})
        chart1.set_y_axis({"name": "数值"})

        ws.insert_chart("A1", chart1)

    def _create_dual_axis_chart(self, wb, ws, data: dict, title: str):
        """创建双轴图"""
        chart = wb.add_chart({"type": "column"})
        line_chart = wb.add_chart({"type": "line"})

        headers = data.get("headers", [])
        rows = data.get("rows", [])

        if len(headers) >= 3:
            chart.add_series(
                {
                    "name": headers[1],
                    "categories": f"=Data!$A$2:$A${len(rows) + 1}",
                    "values": f"=Data!$B$2:$B${len(rows) + 1}",
                }
            )
            line_chart.add_series(
                {
                    "name": headers[2],
                    "categories": f"=Data!$A$2:$A${len(rows) + 1}",
                    "values": f"=Data!$C$2:$C${len(rows) + 1}",
                    "y2_axis": True,
                }
            )

        chart.combine(line_chart)
        chart.set_title({"name": title or "双轴图"})
        ws.insert_chart("A1", chart)

    def _create_waterfall_chart(self, wb, ws, data: dict, title: str):
        """创建瀑布图"""
        chart = wb.add_chart({"type": "column", "subtype": "stacked"})

        headers = data.get("headers", [])
        rows = data.get("rows", [])

        # 瀑布图需要特殊处理
        ws.write(0, 0, "项目")
        ws.write(0, 1, "基础")
        ws.write(0, 2, "变化")
        ws.write(0, 3, "总计")

        for i, row in enumerate(rows, 1):
            ws.write(i, 0, row[0])
            ws.write(i, 1, 0)  # 基础（隐藏）
            ws.write(i, 2, row[1])  # 变化
            ws.write(i, 3, 0)  # 总计（计算）

        chart.add_series(
            {
                "name": "基础",
                "categories": f"=Data!$A$2:$A${len(rows) + 1}",
                "values": f"=Data!$B$2:$B${len(rows) + 1}",
                "fill": {"none": True},
                "border": {"none": True},
            }
        )
        chart.add_series(
            {
                "name": "变化",
                "categories": f"=Data!$A$2:$A${len(rows) + 1}",
                "values": f"=Data!$C$2:$C${len(rows) + 1}",
            }
        )

        chart.set_title({"name": title or "瀑布图"})
        ws.insert_chart("A1", chart)

    def _create_gantt_chart(self, wb, ws, data: dict, title: str):
        """创建甘特图"""
        chart = wb.add_chart({"type": "bar", "subtype": "stacked"})

        headers = data.get("headers", [])
        rows = data.get("rows", [])

        # 甘特图需要开始时间和持续时间
        ws.write(0, 0, "任务")
        ws.write(0, 1, "开始")
        ws.write(0, 2, "持续")

        for i, row in enumerate(rows, 1):
            ws.write(i, 0, row[0])
            ws.write(i, 1, row[1])  # 开始时间
            ws.write(i, 2, row[2])  # 持续时间

        # 隐藏开始时间列
        chart.add_series(
            {
                "name": "开始",
                "categories": f"=Data!$A$2:$A${len(rows) + 1}",
                "values": f"=Data!$B$2:$B${len(rows) + 1}",
                "fill": {"none": True},
                "border": {"none": True},
            }
        )
        chart.add_series(
            {
                "name": "持续",
                "categories": f"=Data!$A$2:$A${len(rows) + 1}",
                "values": f"=Data!$C$2:$C${len(rows) + 1}",
            }
        )

        chart.set_title({"name": title or "甘特图"})
        chart.set_y_axis({"reverse": True})
        ws.insert_chart("A1", chart)
