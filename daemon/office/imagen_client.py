"""
Client per la generazione di immagini AI tramite Google Gemini Imagen 3 (imagen-3.0-generate-002).
Utilizza l'endpoint REST standard di Google GenAI, attivo solo se l'utente fornisce una chiave API Gemini.
"""

import os
import json
import uuid
import base64
import urllib.request
import urllib.error
from typing import Dict, Any, Optional

CACHE_DIR = os.path.expandvars(r"%LOCALAPPDATA%\GeminiAntigravityOffice\cache")
IMAGEN_API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/imagen-3.0-generate-002:predict"


def generate_imagen3_image(
    prompt: str,
    aspect_ratio: str = "16:9",
    api_key: Optional[str] = None
) -> Dict[str, Any]:
    """
    Genera un'immagine AI fotorealistica o illustrativa usando Google Imagen 3.
    Restituisce il percorso del file PNG salvato nella cache locale.
    """
    clean_prompt = str(prompt or "").strip()
    if not clean_prompt:
        return {"ok": False, "error": "Il prompt per la generazione dell'immagine non può essere vuoto."}

    # Risoluzione della chiave API
    if not api_key:
        try:
            from daemon.auth import auth_bridge
            api_key = auth_bridge.get_api_key()
        except Exception:
            pass
        if not api_key:
            api_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")

    if not api_key or not str(api_key).strip():
        return {
            "ok": False,
            "error": "La generazione di immagini AI con Imagen 3 richiede una chiave API Gemini. "
                     "Aggiungi la tua chiave API nelle impostazioni del pannello laterale per attivare questa funzionalità."
        }

    valid_ratios = {"1:1", "16:9", "4:3", "3:4", "9:16"}
    ratio = aspect_ratio if aspect_ratio in valid_ratios else "16:9"

    url = f"{IMAGEN_API_ENDPOINT}?key={api_key.strip()}"
    req_payload = {
        "instances": [{"prompt": clean_prompt}],
        "parameters": {
            "sampleCount": 1,
            "aspectRatio": ratio,
            "outputMimeType": "image/png"
        }
    }

    try:
        data = json.dumps(req_payload).encode("utf-8")
        req = urllib.request.Request(
            url,
            data=data,
            headers={"Content-Type": "application/json"},
            method="POST"
        )

        with urllib.request.urlopen(req, timeout=45) as resp:
            resp_body = resp.read().decode("utf-8")
            res_json = json.loads(resp_body)

        predictions = res_json.get("predictions") or []
        if not predictions or not isinstance(predictions, list):
            return {"ok": False, "error": "Nessuna immagine restituita dall'API Imagen 3."}

        first_pred = predictions[0]
        b64_data = first_pred.get("bytesBase64Encoded")
        if not b64_data:
            return {"ok": False, "error": "Dati immagine non trovati nella risposta di Imagen 3."}

        mime_type = first_pred.get("mimeType", "image/png")
        ext = ".png" if "png" in mime_type else ".jpg"

        os.makedirs(CACHE_DIR, exist_ok=True)
        filename = f"imagen_{uuid.uuid4().hex[:12]}{ext}"
        out_path = os.path.join(CACHE_DIR, filename)

        img_bytes = base64.b64decode(b64_data)
        with open(out_path, "wb") as f:
            f.write(img_bytes)

        return {
            "ok": True,
            "file_path": os.path.abspath(out_path),
            "mime_type": mime_type,
            "prompt": clean_prompt,
            "aspect_ratio": ratio
        }

    except urllib.error.HTTPError as exc:
        err_msg = f"HTTP {exc.code} {exc.reason}"
        try:
            body = exc.read().decode("utf-8")
            err_json = json.loads(body)
            if "error" in err_json and "message" in err_json["error"]:
                err_msg = err_json["error"]["message"]
        except Exception:
            pass
        return {"ok": False, "error": f"Errore chiamata Imagen 3: {err_msg}"}

    except urllib.error.URLError as exc:
        return {"ok": False, "error": f"Errore di rete durante la connessione a Imagen 3: {exc.reason}"}

    except Exception as exc:
        return {"ok": False, "error": f"Eccezione imprevista durante la generazione dell'immagine: {exc}"}
