"""
Unit test per la gestione delle immagini allegate al prompt nel daemon Office.
Verifica il salvataggio su disco degli allegati base64 e l'iniezione nel prompt dell'agente.
"""

import os
import base64
import tempfile
import pytest
from daemon.main import _save_attached_images
from daemon.agent_session import AgentSession


def test_save_attached_images_decoding(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))

    # Pixel PNG 1x1 trasparente in base64
    sample_png_b64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
    raw_images = [
        {
            "name": "screenshot.png",
            "data": f"data:image/png;base64,{sample_png_b64}",
            "mime_type": "image/png"
        },
        {
            "name": "diagram.png",
            "data": sample_png_b64,
            "mime_type": "image/png"
        }
    ]

    saved = _save_attached_images(raw_images)
    assert len(saved) == 2
    for path in saved:
        assert os.path.isabs(path)
        assert os.path.exists(path)
        with open(path, "rb") as f:
            data = f.read()
            assert len(data) > 0


def test_save_attached_images_empty_and_invalid():
    assert _save_attached_images([]) == []
    assert _save_attached_images(None) == []
    assert _save_attached_images([{"data": ""}]) == []


def test_compose_message_with_image_paths():
    session = AgentSession()
    chat = {"id": "test-chat-1", "agy_conversation_id": None}
    fake_paths = [r"C:\Fake\Path\screenshot_1.png", r"C:\Fake\Path\chart_2.png"]

    msg_with_images, _, _ = session._compose_message(
        chat=chat,
        first_turn=False,
        prompt="Inserisci questa immagine nel documento",
        host="Word",
        doc_type=None,
        document_map="",
        doc_url=None,
        image_paths=fake_paths
    )

    assert "USER ATTACHED IMAGES / PICTURES:" in msg_with_images
    assert r"C:\Fake\Path\screenshot_1.png" in msg_with_images
    assert r"C:\Fake\Path\chart_2.png" in msg_with_images
    assert "word_insert_image" in msg_with_images

    msg_no_images, _, _ = session._compose_message(
        chat=chat,
        first_turn=False,
        prompt="Scrivi un paragrafo",
        host="Word",
        doc_type=None,
        document_map="",
        doc_url=None,
        image_paths=None
    )
    assert "USER ATTACHED IMAGES / PICTURES:" not in msg_no_images
