"""
Microsoft Outlook Autonomous Automation Engine
Uses win32com to control running Outlook instances:
- Reading active email threads, subjects, senders, and message bodies
- Generating context-aware executive email replies
- Composing new rich HTML emails with embedded tables and summaries
- Scheduling calendar meetings and appointment invitations
"""

import re
import sys
from typing import Dict, Any, List, Optional, Union

class OutlookEngine:
    def __init__(self):
        pass

    def _get_outlook_app(self):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

        try:
            import win32com.client
            app = win32com.client.GetObject(Class="Outlook.Application")
            if app and (app.Explorers.Count > 0 or app.Inspectors.Count > 0):
                return app
        except Exception:
            pass
        return None

    def read_current_email(self) -> Dict[str, Any]:
        """
        Reads active email item from Inspector (open window) or Explorer (selected in list).
        """
        app = self._get_outlook_app()
        if not app:
            return {
                "type": "simulated",
                "subject": "Q4 Enterprise Strategy & Autonomous Office AI Deliverables",
                "sender_name": "Ergys Ristani",
                "sender_email": "ergys@cloud4job.com",
                "received_time": "2026-10-28 09:30 AM",
                "to": "Hassen Ben Chaaben <hassen.benchaaben01@gmail.com>",
                "body": (
                    "Hi Hassen,\n\n"
                    "Following up on our review of the Gemini Antigravity plugin for Word and PowerPoint. "
                    "Can we extend this to also control Excel financial models and Outlook correspondence? "
                    "We need to ensure it can auto-format tables and schedule calendar review syncs.\n\n"
                    "Best,\nErgys"
                )
            }

        try:
            item = None
            # 1. Try active open inspector window
            inspector = app.ActiveInspector()
            if inspector and inspector.CurrentItem:
                item = inspector.CurrentItem

            # 2. Try explorer selection list
            if not item:
                explorer = app.ActiveExplorer()
                if explorer and explorer.Selection and explorer.Selection.Count > 0:
                    item = explorer.Selection.Item(1)

            if not item:
                return {
                    "type": "empty",
                    "subject": "No email selected",
                    "body": "Select an email message in Outlook to analyze and draft replies."
                }

            return {
                "type": "live",
                "subject": getattr(item, "Subject", "Untitled Email"),
                "sender_name": getattr(item, "SenderName", ""),
                "sender_email": getattr(item, "SenderEmailAddress", ""),
                "to": getattr(item, "To", ""),
                "cc": getattr(item, "CC", ""),
                "received_time": str(getattr(item, "ReceivedTime", "")),
                "body": getattr(item, "Body", "")[:3000],
                "has_attachments": getattr(item, "Attachments", None) is not None and item.Attachments.Count > 0
            }
        except Exception as e:
            return {"type": "error", "message": str(e)}

    def draft_reply(self, reply_body_html: str, reply_all: bool = False) -> Dict[str, Any]:
        """
        Drafts a reply to the currently active email in Outlook with styled HTML.
        """
        app = self._get_outlook_app()
        if not app:
            return {
                "success": True,
                "simulated": True,
                "action": "reply_all" if reply_all else "reply",
                "reply_html": reply_body_html
            }

        try:
            item = None
            inspector = app.ActiveInspector()
            if inspector and inspector.CurrentItem:
                item = inspector.CurrentItem
            elif app.ActiveExplorer() and app.ActiveExplorer().Selection.Count > 0:
                item = app.ActiveExplorer().Selection.Item(1)

            if not item:
                # If no email selected, compose new mail draft
                return self.create_email(
                    to="team@cloud4job.com",
                    subject="Executive Reply Draft",
                    body_html=reply_body_html
                )

            reply_item = item.ReplyAll() if reply_all else item.Reply()
            
            # Format reply with clean typography
            formatted_reply = f"""
            <div style="font-family:Aptos, Calibri, Arial, sans-serif; font-size:11pt; color:#1C1917; line-height:1.5;">
              {reply_body_html}
            </div>
            <br>
            """ + getattr(reply_item, "HTMLBody", "")

            reply_item.HTMLBody = formatted_reply
            reply_item.Display() # Open draft window for review

            return {
                "success": True,
                "subject": reply_item.Subject,
                "recipients": reply_item.To
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def create_email(self, to: str, subject: str, body_html: str) -> Dict[str, Any]:
        """
        Composes a brand new email item in Outlook.
        """
        app = self._get_outlook_app()
        if not app:
            return {"success": True, "simulated": True, "to": to, "subject": subject}

        try:
            # 0 = olMailItem
            mail = app.CreateItem(0)
            mail.To = to
            mail.Subject = subject
            mail.HTMLBody = f"""
            <html>
              <body style="font-family:Aptos, Calibri, Arial, sans-serif; font-size:11pt; color:#1C1917; line-height:1.5;">
                {body_html}
              </body>
            </html>
            """
            mail.Display()
            return {"success": True, "subject": subject, "to": to}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def schedule_meeting(self, subject: str, start_time: str, duration_min: int = 30, attendees: Union[str, List[str]] = None) -> Dict[str, Any]:
        """
        Creates an Outlook Calendar Appointment / Meeting Invitation.
        """
        app = self._get_outlook_app()
        if not app:
            return {
                "success": True,
                "simulated": True,
                "meeting": {"subject": subject, "duration": duration_min}
            }

        try:
            # 1 = olAppointmentItem
            appt = app.CreateItem(1)
            appt.Subject = subject
            appt.Duration = duration_min
            appt.MeetingStatus = 1 # olMeeting (Sends invitations)
            
            if attendees:
                if isinstance(attendees, str):
                    attendees = [a.strip() for a in re.split(r"[,;]\s*", attendees) if a.strip()]
                for att in attendees:
                    appt.Recipients.Add(att)
                appt.Recipients.ResolveAll()

            appt.Body = f"Automated sync scheduled via Gemini Antigravity Suite.\nAgenda: {subject}"
            appt.Display()

            return {"success": True, "subject": subject, "duration_min": duration_min}
        except Exception as e:
            return {"success": False, "error": str(e)}

outlook_engine = OutlookEngine()
