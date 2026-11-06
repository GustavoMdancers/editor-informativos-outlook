import base64
import mimetypes
import os
import tempfile
from typing import Any, Dict, List, Optional
import win32com.client

from core.models import EmailItem

PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"


def _get_namespace():
    app = win32com.client.Dispatch("Outlook.Application")
    return app.GetNamespace("MAPI")


def _build_folder_path(folder) -> str:
    parts = []
    cur = folder
    while cur is not None:
        try:
            name = cur.Name
            parts.append(name)
            cur = cur.Parent
            # Stop when parent is the Namespace ("MAPI")
            if getattr(cur, "Name", None) == "MAPI" or not hasattr(cur, "Parent"):
                break
        except Exception:
            break
    return " > ".join(reversed(parts))


def _serialize_folder(folder, depth: int = 0, max_depth: int = 3) -> Dict[str, Any]:
    name = getattr(folder, "Name", "Pasta")
    entry_id = getattr(folder, "EntryID", "")
    store_id = getattr(folder, "StoreID", "")
    count = 0
    try:
        count = folder.Items.Count
    except Exception:
        pass

    children = []
    has_subfolders = False
    try:
        sub_count = folder.Folders.Count
        has_subfolders = sub_count > 0
        if depth < max_depth and has_subfolders:
            for i in range(1, sub_count + 1):
                try:
                    sf = folder.Folders.Item(i)
                    children.append(_serialize_folder(sf, depth + 1, max_depth))
                except Exception:
                    continue
    except Exception:
        pass

    return {
        "name": name,
        "entry_id": entry_id,
        "store_id": store_id,
        "path": _build_folder_path(folder),
        "count": count,
        "has_subfolders": has_subfolders,
        "children": children,
    }


def list_stores_and_folders(max_depth: int = 3) -> List[Dict[str, Any]]:
    """
    Lists all Outlook stores (mailboxes, shared mailboxes, PSTs) and their folder hierarchy.
    """
    ns = _get_namespace()
    stores_out: List[Dict[str, Any]] = []

    try:
        folders_count = ns.Folders.Count
        for i in range(1, folders_count + 1):
            try:
                root_folder = ns.Folders.Item(i)
                store_node = _serialize_folder(root_folder, depth=0, max_depth=max_depth)
                store_node["is_store"] = True
                stores_out.append(store_node)
            except Exception:
                continue
    except Exception as exc:
        raise RuntimeError(f"Não foi possível acessar as pastas do Outlook: {exc}")

    return stores_out


def list_subfolders(entry_id: str, store_id: Optional[str] = None) -> List[Dict[str, Any]]:
    ns = _get_namespace()
    folder = ns.GetFolderFromID(entry_id, store_id) if store_id else ns.GetFolderFromID(entry_id)
    out = []
    for i in range(1, folder.Folders.Count + 1):
        try:
            sf = folder.Folders.Item(i)
            out.append(_serialize_folder(sf, depth=0, max_depth=2))
        except Exception:
            continue
    return out


def pick_folder_native() -> Optional[Dict[str, Any]]:
    """
    Opens Outlook's native PickFolder() dialog and returns the selected folder info.
    """
    ns = _get_namespace()
    folder = ns.PickFolder()
    if folder is None:
        return None
    return {
        "name": folder.Name,
        "entry_id": folder.EntryID,
        "store_id": folder.StoreID,
        "path": _build_folder_path(folder),
        "count": getattr(folder.Items, "Count", 0),
    }


def list_accounts() -> List[Dict[str, str]]:
    """
    Returns available Outlook accounts and shared mailbox names found in the session.
    """
    ns = _get_namespace()
    accounts = []
    try:
        for i in range(1, ns.Accounts.Count + 1):
            acc = ns.Accounts.Item(i)
            smtp = getattr(acc, "SmtpAddress", "") or getattr(acc, "DisplayName", "")
            if smtp:
                accounts.append({"name": getattr(acc, "DisplayName", smtp), "smtp": smtp, "type": "account"})
    except Exception:
        pass

    # Also detect shared mailboxes mounted as top-level stores
    try:
        known_names = {a["name"].lower() for a in accounts} | {a["smtp"].lower() for a in accounts}
        for i in range(1, ns.Folders.Count + 1):
            root = ns.Folders.Item(i)
            rname = getattr(root, "Name", "")
            if rname and rname.lower() not in known_names:
                accounts.append({"name": rname, "smtp": rname, "type": "shared"})
    except Exception:
        pass

    return accounts


def _extract_recipients(item) -> tuple[str, str, str]:
    to_list, cc_list, bcc_list = [], [], []
    try:
        for i in range(1, item.Recipients.Count + 1):
            rec = item.Recipients.Item(i)
            addr = ""
            try:
                addr = rec.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
            except Exception:
                addr = getattr(rec, "Address", "") or getattr(rec, "Name", "")
            if not addr or "@" not in addr:
                addr = getattr(rec, "Name", "") or addr
            rtype = getattr(rec, "Type", 1)
            if rtype == 1:
                to_list.append(addr)
            elif rtype == 2:
                cc_list.append(addr)
            elif rtype == 3:
                bcc_list.append(addr)
    except Exception:
        pass

    to_str = "; ".join(to_list) if to_list else (getattr(item, "To", "") or "")
    cc_str = "; ".join(cc_list) if cc_list else (getattr(item, "CC", "") or "")
    bcc_str = "; ".join(bcc_list) if bcc_list else (getattr(item, "BCC", "") or "")
    return to_str, cc_str, bcc_str


def _extract_inline_images(item, html: str) -> tuple[Dict[str, str], int]:
    """
    Extracts inline CID attachments as base64 data URIs for browser preview,
    and returns (cid_map, non_inline_attachments_count).
    """
    cid_map: Dict[str, str] = {}
    regular_count = 0
    if not html:
        return cid_map, 0

    try:
        att_count = item.Attachments.Count
        if att_count == 0:
            return cid_map, 0

        tmp_dir = tempfile.mkdtemp(prefix="ebd_preview_")
        try:
            for i in range(1, att_count + 1):
                try:
                    att = item.Attachments.Item(i)
                    cid = ""
                    try:
                        cid = att.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID)
                    except Exception:
                        cid = ""
                    fname = getattr(att, "FileName", f"att_{i}.bin")
                    if cid and f"cid:{cid}" in html:
                        tmp_path = os.path.join(tmp_dir, f"{i}_{fname}")
                        att.SaveAsFile(tmp_path)
                        mime, _ = mimetypes.guess_type(fname)
                        if not mime:
                            mime = "image/png"
                        with open(tmp_path, "rb") as f:
                            b64 = base64.b64encode(f.read()).decode("ascii")
                        cid_map[cid] = f"data:{mime};base64,{b64}"
                        try:
                            os.remove(tmp_path)
                        except OSError:
                            pass
                    else:
                        regular_count += 1
                except Exception:
                    continue
        finally:
            try:
                os.rmdir(tmp_dir)
            except OSError:
                pass
    except Exception:
        pass

    return cid_map, regular_count


def load_emails_from_folder(
    entry_id: str,
    store_id: Optional[str] = None,
    max_items: int = 200,
    subject_filter: str = "",
) -> Dict[str, Any]:
    """
    Reads MailItems from the specified Outlook folder without modifying originals.
    """
    ns = _get_namespace()
    folder = ns.GetFolderFromID(entry_id, store_id) if store_id else ns.GetFolderFromID(entry_id)
    folder_path = _build_folder_path(folder)

    items = folder.Items
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass

    emails: List[EmailItem] = []
    subj_filter_low = (subject_filter or "").strip().lower()

    total_items = items.Count
    idx = 0
    for i in range(1, total_items + 1):
        if len(emails) >= max_items:
            break
        try:
            item = items.Item(i)
            # Class 43 == olMail
            if getattr(item, "Class", None) != 43:
                continue
            subj = getattr(item, "Subject", "") or ""
            if subj_filter_low and subj_filter_low not in subj.lower():
                continue

            item_entry_id = getattr(item, "EntryID", "")
            item_store_id = getattr(folder, "StoreID", "") or store_id or ""
            html_body = getattr(item, "HTMLBody", "") or ""
            text_body = getattr(item, "Body", "") or ""
            to_str, cc_str, bcc_str = _extract_recipients(item)

            sent_dt = getattr(item, "SentOn", None) or getattr(item, "ReceivedTime", None) or getattr(item, "CreationTime", None)
            sent_str = ""
            if sent_dt is not None:
                try:
                    sent_str = sent_dt.strftime("%d/%m/%Y %H:%M")
                except Exception:
                    sent_str = str(sent_dt)

            sender_name = getattr(item, "SentOnBehalfOfName", "") or getattr(item, "SenderName", "") or ""
            cid_map, reg_att_count = _extract_inline_images(item, html_body)

            emails.append(
                EmailItem(
                    id=f"e{idx}",
                    entry_id=item_entry_id,
                    store_id=item_store_id,
                    subject=subj,
                    to=to_str,
                    cc=cc_str,
                    bcc=bcc_str,
                    sent_on=sent_str,
                    sender_name=sender_name,
                    html_body=html_body,
                    text_body=text_body,
                    inline_images=cid_map,
                    attachments_count=reg_att_count,
                )
            )
            idx += 1
        except Exception:
            continue

    return {
        "folder_path": folder_path,
        "entry_id": entry_id,
        "store_id": getattr(folder, "StoreID", "") or store_id or "",
        "emails": emails,
    }
