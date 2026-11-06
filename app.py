import csv
import io
import json
import os
import socket
import sys
import threading
import time
import webbrowser
from datetime import datetime
from typing import Any, Dict, List

from flask import Flask, jsonify, request, Response, send_from_directory

from core.detector import analyze_batch
from core.html_editor import apply_changes_to_html
from core.models import EmailItem
from core.demo_data import get_demo_emails
from outlook.worker import com_worker
from outlook import reader as ol_reader
from outlook import writer as ol_writer


def get_base_dir() -> str:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def get_user_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = get_base_dir()
USER_DIR = get_user_dir()
STATIC_DIR = os.path.join(BASE_DIR, "static")
CONFIG_PATH = os.path.join(USER_DIR, "config.json")
BUNDLED_CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
LOG_DIR = os.path.join(USER_DIR, "logs")

app = Flask(__name__, static_folder=STATIC_DIR, static_url_path="/static")

# In-memory session state
SESSION_STATE: Dict[str, Any] = {
    "folder_path": "",
    "entry_id": "",
    "store_id": "",
    "is_demo": False,
    "emails": {},  # id -> EmailItem
    "logs": [],    # list of dict entries
    "last_heartbeat": time.time(),
}


def load_config() -> Dict[str, Any]:
    default_cfg = {
        "shared_mailbox": "sir.fibra",
        "last_folder_entry_id": "",
        "last_folder_store_id": "",
        "last_folder_path": "",
        "last_local_dir": "",
        "default_open_mode": "one",
        "default_sender_mode": "shared",
        "prefixes": ["REPROGRAMAÇÃO:", "CANCELAMENTO:"],
        "subject_bases": [
            "INFORMATIVO DE ATIVIDADE PROGRAMADA - AMERICAN TOWER",
            "MANUTENÇÃO PROGRAMADA",
        ],
        "scenario_notices": {
            "repro": "ATENÇÃO: A atividade informada abaixo foi REPROGRAMADA. Seguem as novas datas de execução.",
            "cancel": "ATENÇÃO: A atividade informada abaixo foi CANCELADA.",
        },
        "per_email_labels": [
            "Listagem dos equipamentos, clientes e/ou nodes afetados do cliente",
        ],
        "known_labels": [
            "Registro da atividade",
            "Localidade(s) de execução",
            "Data/Hora de início",
            "Data/Hora de término",
            "Descrição/Justificativa",
            "Listagem dos equipamentos, clientes e/ou nodes afetados do cliente",
        ],
    }
    path_to_read = CONFIG_PATH if os.path.exists(CONFIG_PATH) else BUNDLED_CONFIG_PATH
    if os.path.exists(path_to_read):
        try:
            with open(path_to_read, "r", encoding="utf-8") as f:
                data = json.load(f)
                default_cfg.update(data)
        except Exception:
            pass
    return default_cfg


def save_config(updates: Dict[str, Any]) -> Dict[str, Any]:
    cfg = load_config()
    cfg.update(updates)
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return cfg


def append_log_entry(entry: Dict[str, Any]) -> None:
    SESSION_STATE["logs"].append(entry)
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        day_str = datetime.now().strftime("%Y-%m-%d")
        csv_path = os.path.join(LOG_DIR, f"log_informativos_{day_str}.csv")
        file_exists = os.path.exists(csv_path)
        with open(csv_path, "a", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f, delimiter=";")
            if not file_exists:
                writer.writerow(["data_hora", "para", "assunto_original", "assunto_novo", "campos_alterados", "remetente", "status"])
            writer.writerow([
                entry.get("hora", ""),
                entry.get("para", ""),
                entry.get("orig", ""),
                entry.get("novo", ""),
                entry.get("campos", ""),
                entry.get("remetente", ""),
                entry.get("status", ""),
            ])
    except Exception:
        pass


@app.route("/")
def index():
    return send_from_directory(STATIC_DIR, "index.html")


@app.route("/api/config", methods=["GET", "POST"])
def api_config():
    if request.method == "POST":
        body = request.get_json(silent=True) or {}
        cfg = save_config(body)
        return jsonify({"ok": True, "config": cfg})

    cfg = load_config()
    accounts = []
    outlook_ok = True
    outlook_err = ""
    try:
        accounts = com_worker.run(ol_reader.list_accounts)
    except Exception as exc:
        outlook_ok = False
        outlook_err = str(exc)

    return jsonify({
        "ok": True,
        "config": cfg,
        "accounts": accounts,
        "outlook_ok": outlook_ok,
        "outlook_error": outlook_err,
    })


@app.route("/api/folders", methods=["GET"])
def api_folders():
    try:
        stores = com_worker.run(ol_reader.list_stores_and_folders, 3)
        return jsonify({"ok": True, "stores": stores})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/subfolders", methods=["POST"])
def api_subfolders():
    body = request.get_json(silent=True) or {}
    entry_id = body.get("entry_id", "")
    store_id = body.get("store_id", "")
    if not entry_id:
        return jsonify({"ok": False, "error": "entry_id obrigatório"}), 400
    try:
        children = com_worker.run(ol_reader.list_subfolders, entry_id, store_id)
        return jsonify({"ok": True, "children": children})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/pick-folder", methods=["POST"])
def api_pick_folder():
    try:
        picked = com_worker.run(ol_reader.pick_folder_native)
        if not picked:
            return jsonify({"ok": True, "cancelled": True})
        return jsonify({"ok": True, "cancelled": False, "folder": picked})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/pick-local-folder", methods=["POST"])
def api_pick_local_folder():
    """Opens a native Windows folder selection dialog on the local machine."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        cfg = load_config()
        initial_dir = cfg.get("last_local_dir") or os.path.expanduser("~")
        if not os.path.isdir(initial_dir):
            initial_dir = os.path.expanduser("~")

        root = tk.Tk()
        root.withdraw()
        try:
            root.attributes("-topmost", True)
        except Exception:
            pass
        selected = filedialog.askdirectory(
            parent=root,
            initialdir=initial_dir,
            title="Selecione a pasta com os arquivos .msg",
        )
        root.destroy()
        if not selected:
            return jsonify({"ok": True, "cancelled": True})
        norm_path = os.path.normpath(selected)
        return jsonify({"ok": True, "cancelled": False, "local_dir": norm_path})
    except Exception as exc:
        return jsonify({"ok": False, "error": f"Não foi possível abrir o seletor de pastas do Windows: {exc}"}), 500


@app.route("/api/load-folder", methods=["POST"])
def api_load_folder():
    body = request.get_json(silent=True) or {}
    is_demo = bool(body.get("demo", False))
    local_dir = (body.get("local_dir") or "").strip()
    entry_id = body.get("entry_id", "")
    store_id = body.get("store_id", "")
    max_items = int(body.get("max_items", 200))
    subject_filter = body.get("subject_filter", "")

    cfg = load_config()
    known_labels = cfg.get("known_labels", [])

    try:
        if is_demo:
            emails = get_demo_emails()
            folder_path = "Demonstração > sir.fibra > Informativos > Nov-2026 – ATV 655957"
            SESSION_STATE["is_demo"] = True
            SESSION_STATE["folder_path"] = folder_path
            SESSION_STATE["entry_id"] = "DEMO"
            SESSION_STATE["store_id"] = "DEMO"
        elif local_dir:
            res = com_worker.run(ol_reader.load_emails_from_local_folder, local_dir, max_items, subject_filter)
            emails: List[EmailItem] = res["emails"]
            folder_path = f"Disco Local: {res['folder_path']}"
            SESSION_STATE["is_demo"] = False
            SESSION_STATE["folder_path"] = folder_path
            SESSION_STATE["entry_id"] = ""
            SESSION_STATE["store_id"] = ""
            save_config({"last_local_dir": res["local_dir"]})
        else:
            if not entry_id:
                return jsonify({"ok": False, "error": "Selecione uma pasta do Outlook ou uma pasta local com arquivos .msg."}), 400
            res = com_worker.run(ol_reader.load_emails_from_folder, entry_id, store_id, max_items, subject_filter)
            emails: List[EmailItem] = res["emails"]
            folder_path = res["folder_path"]
            SESSION_STATE["is_demo"] = False
            SESSION_STATE["folder_path"] = folder_path
            SESSION_STATE["entry_id"] = res["entry_id"]
            SESSION_STATE["store_id"] = res["store_id"]
            save_config({
                "last_folder_entry_id": res["entry_id"],
                "last_folder_store_id": res["store_id"],
                "last_folder_path": folder_path,
            })

        analysis = analyze_batch(
            emails,
            known_labels=known_labels,
            per_email_labels=cfg.get("per_email_labels", []),
        )
        SESSION_STATE["emails"] = {e.id: e for e in emails}

        return jsonify({
            "ok": True,
            "is_demo": is_demo,
            "folder_path": folder_path,
            "emails": [e.to_dict() for e in emails],
            "fields": analysis["fields"],
            "subject": analysis["subject"],
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/preview", methods=["POST"])
def api_preview():
    body = request.get_json(silent=True) or {}
    email_id = body.get("email_id", "")
    field_updates = body.get("field_updates", {})
    replacements = body.get("replacements", [])
    notice = body.get("notice", "")
    highlight = bool(body.get("highlight", True))

    email_item: EmailItem = SESSION_STATE["emails"].get(email_id)
    if not email_item:
        return jsonify({"ok": False, "error": "E-mail não encontrado na sessão"}), 404

    cfg = load_config()
    base_html = email_item.get_preview_html()
    modified_html = apply_changes_to_html(
        base_html,
        field_updates=field_updates,
        replacements=replacements,
        known_labels=cfg.get("known_labels", []),
        highlight=highlight,
        notice=notice,
    )
    return jsonify({"ok": True, "html": modified_html})


@app.route("/api/open-draft", methods=["POST"])
def api_open_draft():
    """
    Opens a single modified email as an Outlook draft (.Display()) and logs the operation.
    """
    body = request.get_json(silent=True) or {}
    email_id = body.get("email_id", "")
    subject = body.get("subject", "")
    field_updates = body.get("field_updates", {})
    replacements = body.get("replacements", [])
    notice = body.get("notice", "")
    sender_mode = body.get("sender_mode", "shared")
    shared_mailbox = body.get("shared_mailbox", "sir.fibra")
    changes_list = body.get("changes", [])

    email_item: EmailItem = SESSION_STATE["emails"].get(email_id)
    if not email_item:
        return jsonify({"ok": False, "error": "E-mail não encontrado na sessão."}), 404

    cfg = load_config()
    if shared_mailbox and shared_mailbox != cfg.get("shared_mailbox"):
        save_config({"shared_mailbox": shared_mailbox, "default_sender_mode": sender_mode})

    # Apply edits to raw HTML (preserving original cid:... image references for Outlook)
    final_html_for_outlook = apply_changes_to_html(
        email_item.html_body,
        field_updates=field_updates,
        replacements=replacements,
        known_labels=cfg.get("known_labels", []),
        highlight=False,
        notice=notice,
    )
    # Also prepare preview HTML (with data URIs) in case we are in demo mode or fallback
    final_html_for_preview = apply_changes_to_html(
        email_item.get_preview_html(),
        field_updates=field_updates,
        replacements=replacements,
        known_labels=cfg.get("known_labels", []),
        highlight=False,
        notice=notice,
    )

    now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    remetente_str = shared_mailbox if sender_mode == "shared" else "(manual no Outlook)"

    try:
        com_worker.run(
            ol_writer.open_draft_in_outlook,
            email_item.entry_id,
            email_item.store_id,
            subject,
            final_html_for_outlook,
            email_item.to,
            email_item.cc,
            email_item.bcc,
            sender_mode,
            shared_mailbox,
            email_item.file_path,
        )
        log_entry = {
            "hora": now_str,
            "para": email_item.to,
            "orig": email_item.subject,
            "novo": subject,
            "campos": " | ".join(changes_list),
            "remetente": remetente_str,
            "status": "Rascunho aberto no Outlook",
        }
        append_log_entry(log_entry)
        return jsonify({
            "ok": True,
            "opened_in_outlook": True,
            "log_entry": log_entry,
            "logs_count": len(SESSION_STATE["logs"]),
        })
    except Exception as exc:
        # If in demo mode and Outlook isn't running/configured, let frontend show simulated window
        if SESSION_STATE.get("is_demo"):
            log_entry = {
                "hora": now_str,
                "para": email_item.to,
                "orig": email_item.subject,
                "novo": subject,
                "campos": " | ".join(changes_list),
                "remetente": remetente_str,
                "status": "Aberto em modo simulação (Outlook indisponível)",
            }
            append_log_entry(log_entry)
            return jsonify({
                "ok": True,
                "opened_in_outlook": False,
                "simulated_html": final_html_for_preview,
                "outlook_error": str(exc),
                "log_entry": log_entry,
                "logs_count": len(SESSION_STATE["logs"]),
            })

        log_entry = {
            "hora": now_str,
            "para": email_item.to,
            "orig": email_item.subject,
            "novo": subject,
            "campos": " | ".join(changes_list),
            "remetente": remetente_str,
            "status": f"Erro: {exc}",
        }
        append_log_entry(log_entry)
        return jsonify({"ok": False, "error": str(exc), "logs_count": len(SESSION_STATE["logs"])}), 500


@app.route("/api/logs/csv", methods=["GET"])
def api_logs_csv():
    output = io.StringIO()
    output.write("\ufeff")
    writer = csv.writer(output, delimiter=";")
    writer.writerow(["data_hora", "para", "assunto_original", "assunto_novo", "campos_alterados", "remetente", "status"])
    for l in SESSION_STATE["logs"]:
        writer.writerow([
            l.get("hora", ""),
            l.get("para", ""),
            l.get("orig", ""),
            l.get("novo", ""),
            l.get("campos", ""),
            l.get("remetente", ""),
            l.get("status", ""),
        ])
    csv_bytes = output.getvalue().encode("utf-8")
    return Response(
        csv_bytes,
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": "attachment; filename=log_informativos.csv"},
    )


@app.route("/api/heartbeat", methods=["POST"])
def api_heartbeat():
    SESSION_STATE["last_heartbeat"] = time.time()
    return jsonify({"ok": True})


@app.route("/api/shutdown", methods=["POST"])
def api_shutdown():
    def _exit():
        time.sleep(0.4)
        os._exit(0)

    threading.Thread(target=_exit, daemon=True).start()
    return jsonify({"ok": True})


def find_free_port(preferred: int = 54123) -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        if s.connect_ex(("127.0.0.1", preferred)) != 0:
            return preferred
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def main():
    no_browser = "--no-browser" in sys.argv
    port = find_free_port()
    url = f"http://127.0.0.1:{port}"
    print(f"Iniciando Editor de Informativos em {url} ...")

    if not no_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()

    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False, threaded=True)


if __name__ == "__main__":
    main()
