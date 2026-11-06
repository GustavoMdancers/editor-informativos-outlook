from dataclasses import dataclass, field
from typing import Dict, List, Optional, Any


@dataclass
class ExtractedField:
    label: str
    value: str
    mode: str = "inline"  # "inline" | "next" | "table" | "text"


@dataclass
class EmailItem:
    id: str
    entry_id: str
    store_id: str
    subject: str
    to: str
    cc: str
    bcc: str
    sent_on: str
    sender_name: str
    html_body: str
    text_body: str
    fields: List[ExtractedField] = field(default_factory=list)
    inline_images: Dict[str, str] = field(default_factory=dict)  # cid -> data URI for browser preview
    attachments_count: int = 0

    def to_dict(self) -> Dict[str, Any]:
        return {
            "id": self.id,
            "entry_id": self.entry_id,
            "store_id": self.store_id,
            "subject": self.subject,
            "to": self.to,
            "cc": self.cc,
            "bcc": self.bcc,
            "sent_on": self.sent_on,
            "sender_name": self.sender_name,
            "html": self.html_body,
            "preview_html": self.get_preview_html(),
            "text": self.text_body,
            "fields": [{"label": f.label, "value": f.value, "mode": f.mode} for f in self.fields],
            "attachments_count": self.attachments_count,
        }

    def get_preview_html(self) -> str:
        """Replace cid:... references with base64 data URIs so the browser preview renders inline images."""
        html = self.html_body or ""
        for cid, data_uri in self.inline_images.items():
            html = html.replace(f"cid:{cid}", data_uri)
        return html


@dataclass
class CommonFieldInfo:
    label: str
    unique_values: List[str]
    varied: bool
    current_value: Optional[str]
    field_type: str  # "date" | "long" | "text"
    present_count: int
    total_count: int

    def to_dict(self) -> Dict[str, Any]:
        return {
            "label": self.label,
            "uniq": self.unique_values,
            "varied": self.varied,
            "cur": self.current_value,
            "nv": "" if self.varied else (self.current_value or ""),
            "apply": not self.varied,
            "type": "long" if self.varied else self.field_type,
            "present_count": self.present_count,
            "total_count": self.total_count,
        }
