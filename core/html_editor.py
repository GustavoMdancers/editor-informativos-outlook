import re
from typing import Dict, List, Optional
from bs4 import BeautifulSoup, NavigableString, Tag

from core.detector import find_bold_label_in_p, norm_space, norm_label_key, text_with_breaks, is_stop_paragraph


def _build_value_node(soup: BeautifulSoup, value: str, highlight: bool) -> Tag:
    container = soup.new_tag("mark" if highlight else "span")
    if highlight:
        container["class"] = "chg"
    lines = value.split("\n")
    for idx, line in enumerate(lines):
        if idx > 0:
            container.append(soup.new_tag("br"))
        container.append(NavigableString(line))
    return container


def apply_changes_to_html(
    html: str,
    field_updates: Dict[str, str],
    replacements: List[Dict[str, str]],
    known_labels: Optional[List[str]] = None,
    highlight: bool = False,
    notice: str = "",
) -> str:
    """
    Applies field value updates and free-text replacements directly to the email HTML,
    preserving original styles, inline images, and formatting.
    """
    if not html:
        return ""

    known_map: Dict[str, str] = {}
    if known_labels:
        for kl in known_labels:
            known_map[norm_label_key(kl)] = kl.strip().rstrip(":")
    for k in field_updates.keys():
        known_map[norm_label_key(k)] = k

    soup = BeautifulSoup(html, "lxml")
    body = soup.body or soup

    paragraphs = body.find_all(["p"])
    if not paragraphs:
        paragraphs = body.find_all(["div", "li"])

    i = 0
    while i < len(paragraphs):
        p = paragraphs[i]
        res = find_bold_label_in_p(p, known_map)
        if res is not None:
            el, label, inline_val = res
            if label in field_updates:
                new_val = field_updates[label]
                if inline_val:
                    # Inline value in the same paragraph
                    if el is not None:
                        # Preserve leading whitespace/nbsp after the bold label
                        lead = " "
                        nxt = el.next_sibling
                        if isinstance(nxt, NavigableString):
                            m = re.match(r"^([\s\u00a0:]*)", str(nxt))
                            if m and m.group(1):
                                lead = m.group(1)
                        # Remove everything after `el`
                        for sib in list(el.next_siblings):
                            sib.extract()
                        p.append(NavigableString(lead))
                        p.append(_build_value_node(soup, new_val, highlight))
                    else:
                        # Plain paragraph starting with 'Label:'
                        raw_txt = p.get_text()
                        colon_idx = raw_txt.find(":")
                        prefix_txt = raw_txt[: colon_idx + 1] if colon_idx != -1 else (label + ":")
                        p.clear()
                        b_tag = soup.new_tag("b")
                        b_tag.string = prefix_txt
                        p.append(b_tag)
                        p.append(NavigableString(" "))
                        p.append(_build_value_node(soup, new_val, highlight))
                else:
                    # Value in subsequent paragraph(s)
                    val_ps = []
                    j = i + 1
                    while j < len(paragraphs):
                        next_p = paragraphs[j]
                        if find_bold_label_in_p(next_p, known_map) is not None:
                            break
                        next_txt = text_with_breaks(next_p)
                        if not next_txt or is_stop_paragraph(next_txt):
                            break
                        val_ps.append(next_p)
                        j += 1

                    if val_ps:
                        first_vp = val_ps[0]
                        for extra_vp in val_ps[1:]:
                            extra_vp.decompose()
                        # Preserve inner font/span style if present, or replace content of first_vp
                        first_vp.clear()
                        first_vp.append(_build_value_node(soup, new_val, highlight))
                    else:
                        new_p = soup.new_tag("p")
                        new_p.append(_build_value_node(soup, new_val, highlight))
                        p.insert_after(new_p)
                    i = j - 1
        i += 1

    # Also check 2-column tables
    for tr in body.find_all("tr"):
        tds = tr.find_all(["td", "th"], recursive=False)
        if len(tds) >= 2:
            raw_lbl = norm_space(tds[0].get_text()).rstrip(":").strip()
            canon_lbl = known_map.get(raw_lbl.lower(), raw_lbl)
            if canon_lbl in field_updates:
                tds[1].clear()
                tds[1].append(_build_value_node(soup, field_updates[canon_lbl], highlight))

    # Apply free-text find & replace across text nodes
    active_reps = [r for r in (replacements or []) if r.get("find")]
    if active_reps:
        for r in active_reps:
            find_str = r["find"]
            rep_str = r.get("rep", "")
            for txt_node in list(body.find_all(string=True)):
                if isinstance(txt_node, NavigableString) and find_str in str(txt_node):
                    parent = txt_node.parent
                    if parent and parent.name in ("script", "style"):
                        continue
                    parts = str(txt_node).split(find_str)
                    new_nodes = []
                    for idx, pt in enumerate(parts):
                        if idx > 0:
                            if highlight:
                                m_tag = soup.new_tag("mark")
                                m_tag["class"] = "chg"
                                m_tag.string = rep_str
                                new_nodes.append(m_tag)
                            else:
                                new_nodes.append(NavigableString(rep_str))
                        if pt:
                            new_nodes.append(NavigableString(pt))
                    for nn in new_nodes:
                        txt_node.insert_before(nn)
                    txt_node.extract()

    # Optional scenario notice (e.g. "A atividade abaixo foi REPROGRAMADA"),
    # inserted right after the greeting paragraph ("Prezado cliente,").
    notice = (notice or "").strip()
    if notice:
        greeting = None
        for p in body.find_all("p"):
            if norm_space(p.get_text()).lower().startswith("prezad"):
                greeting = p
                break
        notice_p = soup.new_tag("p")
        b_tag = soup.new_tag("b")
        b_tag.append(_build_value_node(soup, notice, highlight))
        notice_p.append(b_tag)
        if greeting is not None:
            spacer = soup.new_tag("p")
            spacer.append(NavigableString("\u00a0"))
            greeting.insert_after(spacer)
            spacer.insert_after(notice_p)
        else:
            body.insert(0, notice_p)

    if highlight:
        head = soup.head
        if not head:
            head = soup.new_tag("head")
            if soup.html:
                soup.html.insert(0, head)
        style_tag = soup.new_tag("style")
        style_tag.string = "mark.chg{background:#fff1a8;outline:1px dashed #d9a400;padding:0 2px;color:#000}"
        head.append(style_tag)

    return str(soup)
