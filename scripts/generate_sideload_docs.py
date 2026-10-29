"""
Generate pre-configured Word, Excel, and PowerPoint launch documents with embedded WebExtension metadata.
This enables immediate loading of the Gemini Antigravity Add-in in Office.
"""

import os
import shutil
import zipfile

TEMPLATES_DIR = r"C:\Users\user\AppData\Local\npm-cache\_npx\f6a1d6caaa3cb423\node_modules\office-addin-dev-settings\templates"
WORD_TEMPLATE = os.path.join(TEMPLATES_DIR, "WordDocumentWithTaskPane.docx")
PPT_TEMPLATE = os.path.join(TEMPLATES_DIR, "PowerPointPresentationWithTaskPane.pptx")
EXCEL_TEMPLATE = os.path.join(TEMPLATES_DIR, "ExcelWorkbookWithTaskPane.xlsx")

PROJECT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
ASSETS_DIR = os.path.join(PROJECT_DIR, "assets")

TARGET_DIR = os.path.expandvars(r"%LOCALAPPDATA%\GeminiAntigravityOffice")
WORD_TARGET = os.path.join(TARGET_DIR, "Gemini_Antigravity_Word.docx")
PPT_TARGET = os.path.join(TARGET_DIR, "Gemini_Antigravity_PowerPoint.pptx")
EXCEL_TARGET = os.path.join(TARGET_DIR, "Gemini_Antigravity_Excel.xlsx")

ADDIN_ID = "7F89B2E3-8E5A-4B21-996A-12E4A9F200A1"
ADDIN_ID_BRACED = f"{{{ADDIN_ID}}}"

def patch_docx(source, destination):
    os.makedirs(os.path.dirname(destination), exist_ok=True)
    with zipfile.ZipFile(source, 'r') as zin, zipfile.ZipFile(destination, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            if item.filename == "word/webextensions/webextension.xml":
                text = content.decode('utf-8')
                text = text.replace("00000000-0000-0000-0000-000000000000", ADDIN_ID)
                content = text.encode('utf-8')
            zout.writestr(item, content)
    print(f"[OpenXML] Created {destination}")

def patch_pptx(source, destination):
    os.makedirs(os.path.dirname(destination), exist_ok=True)
    with zipfile.ZipFile(source, 'r') as zin, zipfile.ZipFile(destination, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            if "webextension" in item.filename:
                text = content.decode('utf-8')
                text = text.replace("00000000-0000-0000-0000-000000000000", ADDIN_ID)
                content = text.encode('utf-8')
            zout.writestr(item, content)
    print(f"[OpenXML] Created {destination}")

def patch_xlsx(source, destination):
    os.makedirs(os.path.dirname(destination), exist_ok=True)
    with zipfile.ZipFile(source, 'r') as zin, zipfile.ZipFile(destination, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            if "webextension" in item.filename:
                text = content.decode('utf-8')
                text = text.replace("00000000-0000-0000-0000-000000000000", ADDIN_ID)
                content = text.encode('utf-8')
            zout.writestr(item, content)
    print(f"[OpenXML] Created {destination}")

if __name__ == "__main__":
    if os.path.exists(WORD_TEMPLATE):
        patch_docx(WORD_TEMPLATE, WORD_TARGET)
        patch_docx(WORD_TEMPLATE, os.path.join(ASSETS_DIR, "Gemini_Antigravity_Word.docx"))
    if os.path.exists(PPT_TEMPLATE):
        patch_pptx(PPT_TEMPLATE, PPT_TARGET)
        patch_pptx(PPT_TEMPLATE, os.path.join(ASSETS_DIR, "Gemini_Antigravity_PowerPoint.pptx"))
    if os.path.exists(EXCEL_TEMPLATE):
        patch_xlsx(EXCEL_TEMPLATE, EXCEL_TARGET)
        patch_xlsx(EXCEL_TEMPLATE, os.path.join(ASSETS_DIR, "Gemini_Antigravity_Excel.xlsx"))
    print("[OpenXML] All sideload documents generated successfully.")
