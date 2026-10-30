"""
Test unitari per il client Gemini Imagen 3 e per lo strumento word_generate_image.
"""

import os
import json
import io
import urllib.error
from unittest.mock import patch, MagicMock
import pytest

from daemon.office.imagen_client import generate_imagen3_image
from daemon.office.word_com_tools import word_generate_image, EXTRA_TOOLS


def test_generate_image_missing_key():
    """Verifica che senza chiave API la generazione fallisca con un messaggio chiaro."""
    res = generate_imagen3_image("un moderno data center", api_key=None)
    assert res["ok"] is False
    assert "chiave API" in res["error"] or "API key" in res["error"]


def test_generate_image_success(tmp_path):
    """Verifica che la risposta valida dell'API Imagen 3 crei il file immagine sul disco."""
    fake_png_base64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
    api_response = {
        "predictions": [
            {
                "bytesBase64Encoded": fake_png_base64,
                "mimeType": "image/png"
            }
        ]
    }

    mock_resp = MagicMock()
    mock_resp.read.return_value = json.dumps(api_response).encode("utf-8")
    mock_resp.__enter__.return_value = mock_resp

    with patch("urllib.request.urlopen", return_value=mock_resp), \
         patch("daemon.office.imagen_client.CACHE_DIR", str(tmp_path)):
        res = generate_imagen3_image("un moderno data center", aspect_ratio="16:9", api_key="AIzaSyDummyKey")
        assert res["ok"] is True
        assert os.path.exists(res["file_path"])
        assert res["file_path"].endswith(".png")
        assert res["mime_type"] == "image/png"


def test_generate_image_http_error():
    """Verifica la corretta gestione di errori HTTP da parte di Google API."""
    err_body = json.dumps({"error": {"message": "API key not valid"}}).encode("utf-8")
    mock_err = urllib.error.HTTPError(
        url="https://generativelanguage.googleapis.com/v1beta/models/imagen-3.0-generate-002:predict",
        code=400,
        msg="Bad Request",
        hdrs={},
        fp=io.BytesIO(err_body)
    )

    with patch("urllib.request.urlopen", side_effect=mock_err):
        res = generate_imagen3_image("test prompt", api_key="invalid_key")
        assert res["ok"] is False
        assert "API key not valid" in res["error"] or "400" in res["error"]


def test_word_generate_image_no_key():
    """Verifica che word_generate_image ritorni errore chiaro quando l'utente non ha impostato la chiave API."""
    mock_doc = MagicMock()
    with patch("daemon.auth.auth_bridge.get_api_key", return_value=None):
        res = word_generate_image(mock_doc, prompt="un diagramma ad albero")
        assert res["ok"] is False
        assert "chiave API" in res["error"] or "API key" in res["error"]


def test_word_generate_image_with_key_and_mock(tmp_path):
    """Verifica che con chiave API word_generate_image invochi Imagen 3 e inserisca l'immagine in Word."""
    mock_doc = MagicMock()
    fake_img = tmp_path / "test.png"
    fake_img.write_bytes(b"\x89PNG\r\n\x1a\n")

    with patch("daemon.auth.auth_bridge.get_api_key", return_value="AIzaSyValidKey"), \
         patch("daemon.office.imagen_client.generate_imagen3_image", return_value={"ok": True, "file_path": str(fake_img), "mime_type": "image/png"}), \
         patch("daemon.office.word_com_tools.word_insert_image", return_value={"ok": True, "anchor": "i0"}) as mock_insert:

        res = word_generate_image(mock_doc, prompt="un moderno data center", caption="Data Center Cloud4Job", location="end")
        assert res["ok"] is True
        mock_insert.assert_called_once()
        call_kwargs = mock_insert.call_args[1]
        assert call_kwargs["file_path"] == str(fake_img)
        assert call_kwargs["caption"] == "Data Center Cloud4Job"


def test_word_generate_image_registered_in_extra_tools():
    """Verifica che word_generate_image sia registrato nel catalogo strumenti EXTRA_TOOLS e in DAEMON_TOOLS."""
    assert "word_generate_image" in EXTRA_TOOLS
    tool_spec = EXTRA_TOOLS["word_generate_image"]
    assert tool_spec["destructive"] is False
    assert "prompt" in tool_spec["inputSchema"]["properties"]

    from daemon.mcp_server import DAEMON_TOOL_NAMES, COM_TOOL_NAMES
    assert "word_generate_image" in DAEMON_TOOL_NAMES
    assert "word_generate_image" in COM_TOOL_NAMES


def test_word_generate_image_width_px_compatibility(tmp_path):
    """Verifica che parametri come width_px e height_px siano accettati senza TypeError."""
    mock_doc = MagicMock()
    fake_img = tmp_path / "test.png"
    fake_img.write_bytes(b"\x89PNG\r\n\x1a\n")

    with patch("daemon.auth.auth_bridge.get_api_key", return_value="AIzaSyKey"), \
         patch("daemon.office.imagen_client.generate_imagen3_image", return_value={"ok": True, "file_path": str(fake_img), "mime_type": "image/png"}), \
         patch("daemon.office.word_com_tools.word_insert_image", return_value={"ok": True, "anchor": "i0"}) as mock_insert:

        res = word_generate_image(mock_doc, prompt="data center", width_px=350, height_px=200, unexpected_extra=123)
        assert res["ok"] is True
        mock_insert.assert_called_once()
        kwargs = mock_insert.call_args[1]
        assert kwargs["width_pt"] == 350
        assert kwargs["height_pt"] == 200


def test_generate_image_fallback_to_env_key(tmp_path, monkeypatch):
    """Verifica il fallback su os.environ se auth_bridge non ha una chiave memorizzata."""
    monkeypatch.setenv("GEMINI_API_KEY", "EnvGeminiKey123")
    fake_png_b64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
    api_response = {"predictions": [{"bytesBase64Encoded": fake_png_b64, "mimeType": "image/png"}]}

    mock_resp = MagicMock()
    mock_resp.read.return_value = json.dumps(api_response).encode("utf-8")
    mock_resp.__enter__.return_value = mock_resp

    with patch("daemon.auth.auth_bridge.get_api_key", return_value=None), \
         patch("urllib.request.urlopen", return_value=mock_resp) as mock_urlopen, \
         patch("daemon.office.imagen_client.CACHE_DIR", str(tmp_path)):

        res = generate_imagen3_image("fallback prompt")
        assert res["ok"] is True
        # Verifica che l'URL contenga la chiave dall'ambiente
        req_arg = mock_urlopen.call_args[0][0]
        assert "key=EnvGeminiKey123" in req_arg.full_url

