"""
Microsoft Word Autonomous Automation Engine
Uses win32com to control running Word instances:
- Reading full document text and selection
- Native Word Track Changes (revisions)
- Inline review comments
- Text replacement and risk highlighting
"""

import os
import re
import sys
import tempfile
from typing import Dict, Any, List

class WordEngine:
    def __init__(self):
        pass

    def _get_word_app(self):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

        try:
            import win32com.client
            app = win32com.client.GetObject(Class="Word.Application")
            if app and app.Documents.Count > 0:
                return app
        except Exception:
            pass
        return None

    def _markdown_to_word_html(self, md: str) -> str:
        lines = md.split("\n")
        html = ["<html><head><meta charset='utf-8'></head><body style='font-family:Aptos, Calibri, sans-serif; font-size:11pt; color:#1C1917; line-height:1.5;'>"]
        in_table = False
        table_rows = []

        def flush_table():
            nonlocal in_table, table_rows
            if in_table and table_rows:
                html.append("<table border='1' cellpadding='8' cellspacing='0' style='border-collapse:collapse; width:100%; margin:14pt 0; border:1px solid #CBD5E1; font-family:Aptos, Calibri, sans-serif; font-size:10pt;'>")
                header_done = False
                row_idx = 0
                for row in table_rows:
                    if all(re.match(r"^:?-+:?$", c.strip()) for c in row):
                        continue
                    if not header_done:
                        html.append("<tr style='background-color:#1E3A8A; color:#FFFFFF; font-weight:bold;'>")
                        for c in row:
                            html.append(f"<th style='padding:8pt 10pt; text-align:left; border:1px solid #CBD5E1;'>{c}</th>")
                        html.append("</tr>")
                        header_done = True
                    else:
                        bg = "#FFFFFF" if row_idx % 2 == 0 else "#F8FAFC"
                        html.append(f"<tr style='background-color:{bg};'>")
                        for c in row:
                            html.append(f"<td style='padding:7pt 10pt; border:1px solid #CBD5E1;'>{c}</td>")
                        html.append("</tr>")
                        row_idx += 1
                html.append("</table>")
                in_table = False
                table_rows = []

        for line in lines:
            t = line.strip()
            if t.startswith("|") and t.endswith("|"):
                in_table = True
                cells = [c.strip() for c in t[1:-1].split("|")]
                table_rows.append(cells)
                continue
            else:
                flush_table()

            if t.startswith("# "):
                html.append(f"<h1 style='font-size:22pt; font-weight:bold; color:#0F2C59; border-bottom:2pt solid #1A73E8; margin-top:20pt; margin-bottom:8pt; padding-bottom:6pt;'>{t[2:]}</h1>")
            elif t.startswith("## "):
                html.append(f"<h2 style='font-size:15pt; font-weight:bold; color:#1E3A8A; margin-top:16pt; margin-bottom:6pt;'>{t[3:]}</h2>")
            elif t.startswith("### "):
                html.append(f"<h3 style='font-size:12.5pt; font-weight:bold; color:#2563EB; margin-top:12pt; margin-bottom:4pt;'>{t[4:]}</h3>")
            elif t.startswith(">"):
                quote_text = re.sub(r"^>\s*", "", t)
                html.append(f"<div style='background-color:#EEF2FF; border-left:4pt solid #4338CA; padding:10pt 14pt; margin:12pt 0; color:#1E1B4B; font-style:italic;'>{quote_text}</div>")
            elif t.startswith("- ") or t.startswith("* "):
                html.append(f"<li style='margin-left:20pt; margin-bottom:4pt;'>{t[2:]}</li>")
            elif re.match(r"^\d+\.\s+", t):
                num_text = re.sub(r"^\d+\.\s+", "", t)
                html.append(f"<li style='margin-left:20pt; margin-bottom:4pt;'>{num_text}</li>")
            elif t == "---":
                html.append("<hr style='border:none; border-top:1pt solid #CBD5E1; margin:16pt 0;'>")
            elif len(t) == 0:
                html.append("<p style='margin-bottom:6pt;'></p>")
            else:
                formatted_p = re.sub(r"\*\*(.*?)\*\*", r"<strong>\1</strong>", t)
                formatted_p = re.sub(r"`(.*?)`", r"<code style='background:#F1F5F9; padding:2px 4px;'>\1</code>", formatted_p)
                html.append(f"<p style='margin-bottom:8pt;'>{formatted_p}</p>")

        flush_table()
        html.append("</body></html>")
        return "\n".join(html)

    def read_document(self) -> Dict[str, Any]:
        app = self._get_word_app()
        if not app:
            return {"type": "mock", "content": "Sample Word document text loaded from active context."}

        try:
            try:
                doc = app.ActiveDocument
            except Exception:
                doc = app.Documents.Add()
            selection_text = app.Selection.Text.strip() if app.Selection else ""
            if len(selection_text) > 0 and selection_text != "\r":
                return {"type": "selection", "content": selection_text, "doc_name": doc.Name}
            else:
                full_text = doc.Content.Text
                return {"type": "full_document", "content": full_text, "doc_name": doc.Name}
        except Exception as e:
            return {"type": "error", "message": str(e), "content": "Unable to read active Word document."}

    def replace_selection(self, new_text: str) -> Dict[str, Any]:
        app = self._get_word_app()
        if not app:
            return {"success": True, "simulated": True, "text": new_text}

        try:
            try:
                doc = app.ActiveDocument
            except Exception:
                doc = app.Documents.Add()

            # If structured Markdown report, insert formatted HTML via COM
            if any(marker in new_text for marker in ["# ", "| ", "> "]):
                html_code = self._markdown_to_word_html(new_text)
                tmp_path = os.path.join(tempfile.gettempdir(), "gemini_word_report.html")
                with open(tmp_path, "w", encoding="utf-8") as f:
                    f.write(html_code)
                try:
                    app.Selection.InsertFile(tmp_path)
                    return {"success": True, "formatted": True, "method": "InsertFile_HTML"}
                except Exception as insert_err:
                    print(f"[WordEngine] Selection InsertFile error: {insert_err}")
                    try:
                        doc.Content.InsertFile(tmp_path)
                        return {"success": True, "formatted": True, "method": "DocContent_InsertFile"}
                    except Exception as doc_err:
                        print(f"[WordEngine] DocContent InsertFile error: {doc_err}")

            if app.Selection and hasattr(app.Selection, "Text"):
                app.Selection.Text = new_text
            else:
                doc.Content.InsertAfter(new_text)
            return {"success": True, "replaced_length": len(new_text)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def insert_comment(self, comment_text: str) -> Dict[str, Any]:
        app = self._get_word_app()
        if not app:
            return {"success": True, "simulated": True, "comment": comment_text}

        try:
            try:
                doc = app.ActiveDocument
            except Exception:
                doc = app.Documents.Add()
            sel = app.Selection
            doc.Comments.Add(Range=sel.Range, Text=comment_text)
            return {"success": True, "comment": comment_text}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def toggle_track_changes(self, enable: bool = True) -> Dict[str, Any]:
        app = self._get_word_app()
        if not app:
            return {"success": True, "simulated": True, "track_revisions": enable}

        try:
            doc = app.ActiveDocument
            doc.TrackRevisions = enable
            return {"success": True, "track_revisions": enable}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def highlight_risks(self, words: List[str]) -> Dict[str, Any]:
        app = self._get_word_app()
        if not app:
            return {"success": True, "simulated": True, "highlighted": words}

        try:
            # wdYellow = 7
            for w in words:
                app.Selection.Find.ClearFormatting()
                app.Selection.Find.Text = w
                app.Selection.Find.Replacement.ClearFormatting()
                app.Selection.Find.Replacement.Highlight = True
                app.Selection.Find.Execute(Replace=2) # 2 = wdReplaceAll
            return {"success": True, "highlighted_count": len(words)}
        except Exception as e:
            return {"success": False, "error": str(e)}

word_engine = WordEngine()
