import pytest
from daemon.suite_orchestrator import SuiteOrchestrator

def test_suite_doc_to_slides():
    orchestrator = SuiteOrchestrator()
    res = orchestrator.doc_to_slides()
    assert res["success"] is True
    assert res["slides_created"] == 4

def test_suite_excel_to_report():
    orchestrator = SuiteOrchestrator()
    res = orchestrator.excel_to_report()
    assert res["success"] is True

def test_suite_email_summary():
    orchestrator = SuiteOrchestrator()
    res = orchestrator.email_summary(target_app="word")
    assert res["success"] is True
