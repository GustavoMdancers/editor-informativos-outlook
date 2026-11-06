import re
from typing import Dict, List, Optional, Tuple, Any
from bs4 import BeautifulSoup, NavigableString, Tag

from core.models import EmailItem, ExtractedField, CommonFieldInfo

PFX_RE = re.compile(r"^(REPROGRAMAÇÃO|REPROGRAMACAO|CANCELAMENTO|ATUALIZAÇÃO|ATUALIZACAO|ALTERAÇÃO|ALTERACAO):\s*", re.IGNORECASE)
DATE_RE = re.compile(r"^(\d{2})/(\d{2})/(\d{4})\s*[–\-]\s*(\d{2})h(\d{2})$")
LABEL_LINE_RE = re.compile(r"^([A-Za-zÀ-ÿ0-9\(\)\/\s\-\.,]{2,90}?):\s*(.*)$")

STOP_PHRASES = [
    "aproveitamos a oportunidade",
    "atenciosamente",
    "seguem abaixo dados",
    "prezado cliente",
    "a american tower",
    "centro de operações",
    "centro de operacoes",
]


def norm_space(s: str) -> str:
    if not s:
        return ""
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()


def norm_label_key(s: str) -> str:
    return norm_space(s).rstrip(":").strip().lower()


def is_stop_paragraph(text: str) -> bool:
    low = norm_space(text).lower()
    if not low:
        return True
    for phrase in STOP_PHRASES:
        if low.startswith(phrase):
            return True
    return False


def text_with_breaks(el: Tag) -> str:
    clone = BeautifulSoup(str(el), "lxml")
    for br in clone.find_all("br"):
        br.replace_with("\n")
    lines = [norm_space(line) for line in clone.get_text().split("\n")]
    return "\n".join([l for l in lines if l])


def find_bold_label_in_p(p: Tag, known_labels_map: Dict[str, str]) -> Optional[Tuple[Optional[Tag], str, str]]:
    """
    Returns (label_tag_or_None, canonical_label, inline_value_after_label) if paragraph starts with a label.
    Checks:
    1. Leading <b> or <strong> tag ending with ':' (or with ':' right after it)
    2. Known labels or general 'Label: value' pattern at start of paragraph text
    """
    full_text = text_with_breaks(p)
    if not full_text or is_stop_paragraph(full_text):
        return None

    # Skip leading whitespace / empty spans
    first_el = None
    for child in p.children:
        if isinstance(child, NavigableString):
            if not norm_space(str(child)):
                continue
            break
        elif isinstance(child, Tag):
            if child.name in ("b", "strong") or (child.find(["b", "strong"]) and norm_space(child.get_text()) == norm_space(child.find(["b", "strong"]).get_text())):
                first_el = child
                break
            if not norm_space(child.get_text()):
                continue
            break

    if first_el is not None:
        b_text = norm_space(first_el.get_text())
        # Case A: <b>Label:</b> value
        if b_text.endswith(":") and len(b_text) < 120:
            lbl = b_text[:-1].strip()
            if not is_stop_paragraph(lbl):
                after_parts = []
                for sib in first_el.next_siblings:
                    if isinstance(sib, NavigableString):
                        after_parts.append(str(sib))
                    elif isinstance(sib, Tag):
                        after_parts.append(text_with_breaks(sib))
                inline_val = norm_space(" ".join(after_parts))
                return (first_el, known_labels_map.get(lbl.lower(), lbl), inline_val)
        # Case B: <b>Label</b>: value
        elif len(b_text) < 120:
            rest_parts = []
            for sib in first_el.next_siblings:
                if isinstance(sib, NavigableString):
                    rest_parts.append(str(sib))
                elif isinstance(sib, Tag):
                    rest_parts.append(text_with_breaks(sib))
            rest_str = norm_space(" ".join(rest_parts))
            if rest_str.startswith(":"):
                lbl = b_text.strip()
                inline_val = norm_space(rest_str[1:])
                return (first_el, known_labels_map.get(lbl.lower(), lbl), inline_val)

    # Fallback: check if paragraph starts with a known label or regex pattern 'Label: value'
    first_line = full_text.split("\n")[0]
    for kl_low, kl_canon in known_labels_map.items():
        if first_line.lower().startswith(kl_low + ":"):
            after = norm_space(first_line[len(kl_low) + 1 :])
            extra_lines = full_text.split("\n")[1:]
            if extra_lines:
                after = "\n".join([after] + extra_lines if after else extra_lines)
            return (None, kl_canon, after)

    m = LABEL_LINE_RE.match(first_line)
    if m:
        candidate_lbl = norm_space(m.group(1))
        # Ensure it's not a regular sentence or URL
        if len(candidate_lbl) <= 85 and not candidate_lbl.lower().startswith(("http", "mailto", "prezado", "atenciosamente")):
            # If paragraph has bold or matches typical field structure
            if p.find(["b", "strong"]) or candidate_lbl.lower() in known_labels_map:
                after = norm_space(m.group(2))
                extra_lines = full_text.split("\n")[1:]
                if extra_lines:
                    after = "\n".join([after] + extra_lines if after else extra_lines)
                return (None, known_labels_map.get(candidate_lbl.lower(), candidate_lbl), after)

    return None


def extract_fields_from_html(html: str, known_labels: Optional[List[str]] = None) -> List[ExtractedField]:
    if not html:
        return []

    known_map: Dict[str, str] = {}
    if known_labels:
        for kl in known_labels:
            known_map[norm_label_key(kl)] = kl.strip().rstrip(":")

    soup = BeautifulSoup(html, "lxml")
    body = soup.body or soup

    fields: List[ExtractedField] = []
    seen_labels = set()

    # 1. Check paragraphs (<p>, <div>)
    paragraphs = body.find_all(["p"])
    if not paragraphs:
        paragraphs = body.find_all(["div", "li"])

    i = 0
    while i < len(paragraphs):
        p = paragraphs[i]
        res = find_bold_label_in_p(p, known_map)
        if res is not None:
            el, label, inline_val = res
            if label not in seen_labels:
                if inline_val:
                    fields.append(ExtractedField(label=label, value=inline_val, mode="inline"))
                    seen_labels.add(label)
                else:
                    # Value is in subsequent paragraph(s) until next label or empty/stop paragraph
                    val_lines = []
                    j = i + 1
                    while j < len(paragraphs):
                        next_p = paragraphs[j]
                        if find_bold_label_in_p(next_p, known_map) is not None:
                            break
                        next_txt = text_with_breaks(next_p)
                        if not next_txt or is_stop_paragraph(next_txt):
                            break
                        val_lines.append(next_txt)
                        j += 1
                    fields.append(ExtractedField(label=label, value="\n".join(val_lines), mode="next"))
                    seen_labels.add(label)
                    i = j - 1
        i += 1

    # 2. Check 2-column tables if any fields are inside tables
    for tr in body.find_all("tr"):
        tds = tr.find_all(["td", "th"], recursive=False)
        if len(tds) >= 2:
            raw_lbl = norm_space(tds[0].get_text())
            if raw_lbl.endswith(":") or norm_label_key(raw_lbl) in known_map:
                lbl = raw_lbl.rstrip(":").strip()
                lbl = known_map.get(lbl.lower(), lbl)
                if lbl and lbl not in seen_labels and len(lbl) < 100 and not is_stop_paragraph(lbl):
                    val = text_with_breaks(tds[1])
                    fields.append(ExtractedField(label=lbl, value=val, mode="table"))
                    seen_labels.add(lbl)

    return fields



def strip_subject_prefix(subject: str) -> str:
    """Removes a leading scenario prefix (REPROGRAMAÇÃO:, CANCELAMENTO: ...) from a subject."""
    return PFX_RE.sub("", subject or "").strip()


def analyze_batch(
    emails: List[EmailItem],
    known_labels: Optional[List[str]] = None,
    per_email_labels: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    Extracts fields for each email and identifies common fields across the batch.

    Subject model: final subject = [scenario prefix] + base subject.
    The base subject is the same for the whole batch (e.g. "INFORMATIVO DE ATIVIDADE
    PROGRAMADA - AMERICAN TOWER" or "MANUTENÇÃO PROGRAMADA"). If the batch mixes
    different bases, mode "keep" preserves each email's own base and only the prefix changes.

    Labels in `per_email_labels` (e.g. the equipment/client listing) are always treated
    as per-recipient values, locked from batch editing even if they happen to be equal.
    """
    if not emails:
        return {
            "fields": [],
            "subject": {"prefix": "", "fixed": "", "mode": "set", "bases": []},
        }

    for e in emails:
        e.fields = extract_fields_from_html(e.html_body, known_labels)

    forced_keys = {norm_label_key(l) for l in (per_email_labels or [])}

    total = len(emails)
    ordered_labels: List[str] = []
    for e in emails:
        for f in e.fields:
            if f.label not in ordered_labels:
                ordered_labels.append(f.label)

    common_fields: List[CommonFieldInfo] = []
    for lbl in ordered_labels:
        vals = []
        for e in emails:
            found = next((f.value for f in e.fields if f.label == lbl), None)
            if found is not None:
                vals.append(found)
        if not vals:
            continue
        uniq = list(dict.fromkeys(vals))
        varied = len(uniq) > 1 or norm_label_key(lbl) in forced_keys
        cur = None if varied else uniq[0]
        sample = uniq[0] if uniq else ""
        if DATE_RE.match(sample):
            ftype = "date"
        elif len(sample) > 60 or "\n" in sample:
            ftype = "long"
        else:
            ftype = "text"

        common_fields.append(
            CommonFieldInfo(
                label=lbl,
                unique_values=uniq,
                varied=varied,
                current_value=cur,
                field_type=ftype,
                present_count=len(vals),
                total_count=total,
            )
        )

    # Subject: existing prefix (if shared by all) + base subject
    subjects = [e.subject or "" for e in emails]
    found_prefixes = []
    for s in subjects:
        m = PFX_RE.match(s)
        found_prefixes.append((m.group(1).upper() + ":") if m else "")
    prefix = found_prefixes[0] if len(set(found_prefixes)) == 1 else ""

    bases = list(dict.fromkeys(strip_subject_prefix(s) for s in subjects))
    same_base = len(bases) == 1

    return {
        "fields": [cf.to_dict() for cf in common_fields],
        "subject": {
            "prefix": prefix,
            "fixed": bases[0] if same_base else "",
            "mode": "set" if same_base else "keep",
            "bases": bases,
        },
    }
