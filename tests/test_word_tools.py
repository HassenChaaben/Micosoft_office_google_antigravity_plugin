import pytest
from daemon.word_tools import get_tool_definitions, get_tool, is_destructive, WORD_TOOLS


def test_tool_definitions_count_and_keys():
    tools = get_tool_definitions()
    assert len(tools) >= 15
    for t in tools:
        assert "name" in t
        assert "description" in t
        assert "inputSchema" in t
        assert "destructive" in t
        assert t["name"].startswith("word_")


def test_destructive_flags():
    assert is_destructive("word_replace_text") is True
    assert is_destructive("word_delete_paragraph") is True
    assert is_destructive("word_delete_table") is True
    assert is_destructive("word_clear_document") is True
    assert is_destructive("word_set_paragraph_text") is True
    assert is_destructive("word_read_table") is False
    assert is_destructive("word_insert_table") is False
    assert is_destructive("word_insert_paragraph") is False
    assert is_destructive("word_get_document_outline") is False
    assert is_destructive("word_batch_update_table_cells") is False
    assert is_destructive("word_fill_form_fields") is False
    assert is_destructive("word_insert_image") is False


def test_table_insertion_tool_schema():
    tool = get_tool("word_insert_table")
    assert tool is not None
    props = tool["inputSchema"]["properties"]
    assert "headers" in props
    assert "rows" in props
    assert "table_style" in props
    assert "headers" in tool["inputSchema"]["required"]
    assert "rows" in tool["inputSchema"]["required"]
