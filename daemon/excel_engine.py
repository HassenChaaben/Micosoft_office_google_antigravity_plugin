"""
Microsoft Excel Autonomous Automation Engine
Uses win32com to control running Excel instances:
- Reading active worksheet data, cell matrices, and formulas
- Writing structured 2D ranges with auto-formatting and AutoFit
- Constructing professional financial models (Revenue, EBITDA, DCF, Margins)
- Generating native Excel charts (Clustered Column, Line, Pie)
"""

import sys
from typing import Dict, Any, List, Optional

class ExcelEngine:
    def __init__(self):
        pass

    def _get_excel_app(self):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

        try:
            import win32com.client
            app = win32com.client.GetObject(Class="Excel.Application")
            if app and app.Workbooks.Count > 0:
                return app
        except Exception:
            pass
        return None

    def read_sheet(self) -> Dict[str, Any]:
        """
        Reads active worksheet name, dimensions, used range values, formulas, and current selection.
        """
        app = self._get_excel_app()
        if not app:
            return {
                "type": "simulated",
                "sheet_name": "Sheet1",
                "active_cell": "A1",
                "used_range": "A1:D5",
                "headers": ["Quarter", "Revenue", "COGS", "Gross Profit"],
                "rows": [
                    ["Q1 2026", 1250000, 450000, 800000],
                    ["Q2 2026", 1420000, 510000, 910000],
                    ["Q3 2026", 1680000, 590000, 1090000],
                    ["Q4 2026", 1950000, 680000, 1270000]
                ]
            }

        try:
            try:
                wb = app.ActiveWorkbook
            except Exception:
                wb = app.Workbooks.Add()

            sheet = app.ActiveSheet
            used = sheet.UsedRange
            values = used.Value if used else []
            formulas = used.Formula if used else []
            sel = app.Selection

            return {
                "type": "live",
                "workbook": wb.Name,
                "sheet_name": sheet.Name,
                "used_range": used.Address if used else "A1",
                "selection": sel.Address if sel else "A1",
                "row_count": used.Rows.Count if used else 0,
                "col_count": used.Columns.Count if used else 0,
                "data": values
            }
        except Exception as e:
            return {"type": "error", "message": str(e)}

    def write_cell_range(self, start_cell: str, matrix: List[List[Any]], format_table: bool = True) -> Dict[str, Any]:
        """
        Writes a 2D matrix of values/formulas to active Excel sheet starting at start_cell (e.g., 'A1').
        Applies executive table styling (navy header, white bold text, zebra striping, auto-fit).
        """
        if not matrix or not matrix[0]:
            return {"success": True, "range": start_cell, "rows": 0, "cols": 0}

        app = self._get_excel_app()
        if not app:
            return {
                "success": True,
                "simulated": True,
                "start_cell": start_cell,
                "rows_written": len(matrix),
                "cols_written": len(matrix[0]) if matrix else 0
            }

        try:
            try:
                wb = app.ActiveWorkbook
            except Exception:
                wb = app.Workbooks.Add()

            sheet = app.ActiveSheet
            start_range = sheet.Range(start_cell)
            row_start = start_range.Row
            col_start = start_range.Column

            num_rows = len(matrix)
            num_cols = len(matrix[0]) if num_rows > 0 else 0

            # Write values cell by cell or block
            end_cell = sheet.Cells(row_start + num_rows - 1, col_start + num_cols - 1)
            target_range = sheet.Range(start_range, end_cell)

            for r_idx, row in enumerate(matrix):
                for c_idx, val in enumerate(row):
                    cell = sheet.Cells(row_start + r_idx, col_start + c_idx)
                    if isinstance(val, str) and val.startswith("="):
                        cell.Formula = val
                    else:
                        cell.Value = val

            # Apply professional table formatting
            if format_table and num_rows > 0:
                # Header row: Dark Navy fill (#1E3A8A = 0x8A3A1E in BGR), White bold text
                header_range = sheet.Range(start_range, sheet.Cells(row_start, col_start + num_cols - 1))
                header_range.Font.Bold = True
                header_range.Font.Color = 0xFFFFFF # White
                header_range.Interior.Color = 0x8A3A1E # #1E3A8A Navy
                header_range.HorizontalAlignment = -4108 # xlCenter

                # Data rows: Zebra striping and currency formatting
                for r_idx in range(1, num_rows):
                    row_range = sheet.Range(sheet.Cells(row_start + r_idx, col_start), sheet.Cells(row_start + r_idx, col_start + num_cols - 1))
                    if r_idx % 2 == 1:
                        row_range.Interior.Color = 0xFCFAF8 # #F8FAFC light zebra
                    else:
                        row_range.Interior.Color = 0xFFFFFF

                # AutoFit all columns in the table
                target_range.Columns.AutoFit()

            return {
                "success": True,
                "range": target_range.Address,
                "rows": num_rows,
                "cols": num_cols
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def apply_financial_model(self, model_type: str = "saas_revenue", currency: str = "$") -> Dict[str, Any]:
        """
        Creates a professional Wall Street-standard financial projection model with
        Revenue, COGS, Gross Profit, Operating Expenses, EBITDA, and Margin percentages.
        """
        app = self._get_excel_app()
        matrix = [
            ["Metric", "FY2024", "FY2025", "FY2026E", "FY2027E", "CAGR"],
            ["Subscription Revenue", 5200000, 8100000, 12800000, 19500000, "=(E2/B2)^(1/3)-1"],
            ["Professional Services", 800000, 1100000, 1400000, 1800000, "=(E3/B3)^(1/3)-1"],
            ["Total Revenue", "=SUM(B2:B3)", "=SUM(C2:C3)", "=SUM(D2:D3)", "=SUM(E2:E3)", "=(E4/B4)^(1/3)-1"],
            ["Cost of Goods Sold (COGS)", 1200000, 1800000, 2600000, 3700000, "=(E5/B5)^(1/3)-1"],
            ["Gross Profit", "=B4-B5", "=C4-C5", "=D4-D5", "=E4-E5", "=(E6/B6)^(1/3)-1"],
            ["Gross Margin %", "=B6/B4", "=C6/C4", "=D6/D4", "=E6/E4", "-"],
            ["R&D Expenses", 1800000, 2600000, 3900000, 5400000, "=(E8/B8)^(1/3)-1"],
            ["Sales & Marketing", 2100000, 3100000, 4600000, 6200000, "=(E9/B9)^(1/3)-1"],
            ["General & Administrative", 700000, 950000, 1300000, 1700000, "=(E10/B10)^(1/3)-1"],
            ["Total Operating Expenses", "=SUM(B8:B10)", "=SUM(C8:C10)", "=SUM(D8:D10)", "=SUM(E8:E10)", "=(E11/B11)^(1/3)-1"],
            ["EBITDA", "=B6-B11", "=C6-C11", "=D6-D11", "=E6-E11", "-"],
            ["EBITDA Margin %", "=B12/B4", "=C12/C4", "=D12/D4", "=E12/E4", "-"]
        ]

        if not app:
            return {"success": True, "simulated": True, "model": model_type, "rows": len(matrix)}

        res = self.write_cell_range("A1", matrix, format_table=True)
        try:
            sheet = app.ActiveSheet
            # Format Currency rows
            for row_idx in [2, 3, 4, 5, 6, 8, 9, 10, 11, 12]:
                sheet.Range(f"B{row_idx}:E{row_idx}").NumberFormat = f"{currency}#,##0"
            # Format Percentage rows
            for row_idx in [7, 13]:
                sheet.Range(f"B{row_idx}:E{row_idx}").NumberFormat = "0.0%"
            # Format CAGR column
            sheet.Range("F2:F11").NumberFormat = "0.0%"

            # Total Revenue & EBITDA double bottom border (xlDouble = 9)
            for row_idx in [4, 6, 11, 12]:
                border_range = sheet.Range(f"A{row_idx}:F{row_idx}")
                border_range.Font.Bold = True
                border_range.Borders(9).LineStyle = 1 # xlContinuous top
                border_range.Borders(9).Weight = 2
                border_range.Borders(4).LineStyle = 9 # xlDouble bottom

            return {"success": True, "model": model_type, "range": "A1:F13"}
        except Exception as e:
            return {"success": True, "model": model_type, "formatting_note": str(e)}

    def create_chart(self, chart_type: str = "column", data_range: str = "A1:E4", title: str = "Revenue Breakdown") -> Dict[str, Any]:
        """
        Inserts native Excel chart (column = xlColumnClustered, line = xlLine, pie = xlPie).
        """
        app = self._get_excel_app()
        if not app:
            return {"success": True, "simulated": True, "chart": chart_type, "range": data_range}

        try:
            sheet = app.ActiveSheet
            source = sheet.Range(data_range)

            # Chart types: 51 = xlColumnClustered, 4 = xlLine, 5 = xlPie
            type_code = 51
            if chart_type.lower() == "line":
                type_code = 4
            elif chart_type.lower() == "pie":
                type_code = 5

            # Position chart to the right of data
            left_pos = source.Left + source.Width + 20
            top_pos = source.Top
            chart_obj = sheet.ChartObjects().Add(left_pos, top_pos, 440, 260)
            chart = chart_obj.Chart
            chart.SetSourceData(source)
            chart.ChartType = type_code
            chart.HasTitle = True
            chart.ChartTitle.Text = title

            return {"success": True, "chart_title": title, "type": chart_type}
        except Exception as e:
            return {"success": False, "error": str(e)}

excel_engine = ExcelEngine()
