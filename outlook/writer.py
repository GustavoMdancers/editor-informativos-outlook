import os
import tempfile
from typing import Any, Dict, Optional
import win32com.client

PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACH_CONTENT_LOCATION = "http://schemas.microsoft.com/mapi/proptag/0x3713001F"
PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"


def _copy_attachments_with_cid(src_item, dst_item) -> None:
    """
    Fallback attachment copier that preserves Content-ID (CID) and hidden inline flags.
    """
    try:
        att_count = src_item.Attachments.Count
        if att_count == 0:
            return
        tmp_dir = tempfile.mkdtemp(prefix="ebd_att_")
        try:
            for i in range(1, att_count + 1):
                try:
                    att = src_item.Attachments.Item(i)
                    fname = getattr(att, "FileName", f"att_{i}.dat")
                    tmp_path = os.path.join(tmp_dir, f"{i}_{fname}")
                    att.SaveAsFile(tmp_path)

                    cid = ""
                    try:
                        cid = att.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID)
                    except Exception:
                        cid = ""

                    new_att = dst_item.Attachments.Add(tmp_path, 1, 1, fname)
                    if cid:
                        try:
                            new_att.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, cid)
                            new_att.PropertyAccessor.SetProperty(PR_ATTACHMENT_HIDDEN, True)
                        except Exception:
                            pass
                    try:
                        os.remove(tmp_path)
                    except OSError:
                        pass
                except Exception:
                    continue
        finally:
            try:
                os.rmdir(tmp_dir)
            except OSError:
                pass
    except Exception:
        pass


def open_draft_in_outlook(
    entry_id: str,
    store_id: Optional[str],
    subject: str,
    html_body: str,
    to: str,
    cc: str = "",
    bcc: str = "",
    sender_mode: str = "shared",
    shared_mailbox: str = "sir.fibra",
) -> Dict[str, Any]:
    """
    Creates a new draft in Outlook Classic with the modified HTML body, subject,
    original recipients, inline images & attachments preserved, and calls .Display()
    (NEVER .Send()).
    """
    app = win32com.client.Dispatch("Outlook.Application")
    ns = app.GetNamespace("MAPI")

    src_item = None
    if entry_id:
        try:
            src_item = ns.GetItemFromID(entry_id, store_id) if store_id else ns.GetItemFromID(entry_id)
        except Exception:
            src_item = None

    draft = None
    # Strategy 1: Use Forward() from original item so Outlook natively clones all MAPI inline images (CID) & attachments
    if src_item is not None:
        try:
            draft = src_item.Forward()
        except Exception:
            draft = None

    # Strategy 2: Create a fresh MailItem and copy attachments with CID properties
    if draft is None:
        draft = app.CreateItem(0)  # olMailItem
        if src_item is not None:
            _copy_attachments_with_cid(src_item, draft)

    # Clear any pre-populated recipients from Forward() and set exact recipients
    try:
        while draft.Recipients.Count > 0:
            draft.Recipients.Remove(1)
    except Exception:
        pass

    draft.To = to or ""
    draft.CC = cc or ""
    draft.BCC = bcc or ""
    draft.Subject = subject or ""
    draft.HTMLBody = html_body or ""

    # Configure sender ("De") field
    if sender_mode == "shared" and shared_mailbox:
        # Check if there is a matching Account first, otherwise set SentOnBehalfOfName
        matched_account = None
        try:
            for i in range(1, ns.Accounts.Count + 1):
                acc = ns.Accounts.Item(i)
                smtp = (getattr(acc, "SmtpAddress", "") or "").lower()
                dname = (getattr(acc, "DisplayName", "") or "").lower()
                target = shared_mailbox.lower()
                if target == smtp or target == dname or smtp.startswith(target + "@"):
                    matched_account = acc
                    break
        except Exception:
            pass

        if matched_account is not None:
            try:
                draft.SendUsingAccount = matched_account
            except Exception:
                pass
        try:
            draft.SentOnBehalfOfName = shared_mailbox
        except Exception:
            pass
    elif sender_mode == "blank":
        try:
            draft.SentOnBehalfOfName = ""
        except Exception:
            pass

    # Display window for manual review and sending (NEVER call .Send())
    draft.Display(False)

    return {
        "ok": True,
        "subject": subject,
        "to": to,
    }
