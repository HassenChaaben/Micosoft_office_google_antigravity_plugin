import pytest
from daemon.outlook_engine import OutlookEngine

def test_outlook_engine_read_email():
    engine = OutlookEngine()
    data = engine.read_current_email()
    assert "subject" in data
    assert "body" in data

def test_outlook_engine_draft_reply():
    engine = OutlookEngine()
    res = engine.draft_reply("<p>Thank you for your email.</p>")
    assert res["success"] is True

def test_outlook_engine_create_email():
    engine = OutlookEngine()
    res = engine.create_email("test@example.com", "Test Subject", "<p>Hello</p>")
    assert res["success"] is True

def test_outlook_engine_schedule_meeting():
    engine = OutlookEngine()
    res = engine.schedule_meeting("Project Review", "2026-10-29 10:00", 45, "colleague@example.com")
    assert res["success"] is True
