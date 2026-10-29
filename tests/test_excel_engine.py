import pytest
from daemon.excel_engine import ExcelEngine

def test_excel_engine_simulated_read():
    engine = ExcelEngine()
    # In headless or test environment, read_sheet returns simulated data cleanly
    data = engine.read_sheet()
    assert "sheet_name" in data
    assert "headers" in data or "data" in data

def test_excel_engine_write_cell_range():
    engine = ExcelEngine()
    matrix = [
        ["Product", "Q1", "Q2"],
        ["SaaS Platform", 1000, 1500]
    ]
    res = engine.write_cell_range("A1", matrix, format_table=True)
    assert res["success"] is True

def test_excel_engine_financial_model():
    engine = ExcelEngine()
    res = engine.apply_financial_model(model_type="saas_revenue")
    assert res["success"] is True

def test_excel_engine_create_chart():
    engine = ExcelEngine()
    res = engine.create_chart(chart_type="column", data_range="A1:C2", title="Test Chart")
    assert res["success"] is True
