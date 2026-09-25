"""
devis_engine.py
================
Kernlogica om tekst uit een ongeformatteerd Word-document te gieten in de
opmaakstructuur van een referentie-sjabloon (devis-documenten met titels,
subtitels A/B/C, verdiepingsniveaus, ruimtenamen, Contenant/Contenu-bedragen
en bullets). Wordt gebruikt door app.py (Streamlit-interface).

Behoudt de lege spacer-alinea's (witruimte) die in het sjabloon tussen
structuurelementen staan (bv. na een verdiepingsniveau, na een ruimtenaam,
na het laatste bullet-item van een Contenant/Contenu-blok).
"""

import copy
import re
import zipfile
from io import BytesIO
from lxml import etree
from docx import Document

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

FLOOR_PATTERNS = re.compile(
    r'^(REZ-DE-CHAUSSEE|PREMIER ETAGE|DEUXIEME ETAGE|TROISIEME ETAGE|EXTERIEURS|SOUS-SOL)$',
    re.IGNORECASE
)
MAIN_TITLE_PATTERNS = [
    r'^LOCAUX CONCERNES PAR NOTRE INTERVENTION$',
    r'^DESCRIPTION DE NOTRE INTERVENTION$',
]


def strip_underline_unicode(s: str) -> str:
    return s.replace('\u0332', '')


def load_paragraphs_from_bytes(docx_bytes):
    with zipfile.ZipFile(BytesIO(docx_bytes)) as z:
        xml_bytes = z.read('word/document.xml')
    tree = etree.fromstring(xml_bytes)
    body = tree.find('w:body', NS)
    paras = body.findall('w:p', NS)
    return tree, body, paras


def para_text(p):
    return ''.join(t.text or '' for t in p.findall('.//w:t', NS))


def find_role_source_indices(paras):
    """
    Vindt voor elke rol de paragraaf-index EN het aantal lege spacer-
    paragrafen die in het sjabloon onmiddellijk NA dat element volgen.
    Retourneert {rol: (paragraaf_index, aantal_spacers_erna)}.
    """
    all_texts = [para_text(p).strip() for p in paras]
    non_empty_positions = [i for i, t in enumerate(all_texts) if t]

    def count_trailing_spacers(idx):
        count = 0
        j = idx + 1
        while j < len(all_texts) and all_texts[j] == '':
            count += 1
            j += 1
        return count

    roles = {}
    for pos_idx, orig_i in enumerate(non_empty_positions):
        text = all_texts[orig_i]
        clean = strip_underline_unicode(text)
        next_orig_i = non_empty_positions[pos_idx + 1] if pos_idx + 1 < len(non_empty_positions) else None
        next_clean = strip_underline_unicode(all_texts[next_orig_i]).strip() if next_orig_i is not None else ''

        if 'titre_niveau1' not in roles and any(re.match(pat, clean) for pat in MAIN_TITLE_PATTERNS):
            roles['titre_niveau1'] = (orig_i, count_trailing_spacers(orig_i))
            continue
        if 'titre_souligne' not in roles and 'titre_niveau1' in roles and orig_i != roles['titre_niveau1'][0]:
            if any(re.match(pat, clean) for pat in MAIN_TITLE_PATTERNS):
                roles['titre_souligne'] = (orig_i, count_trailing_spacers(orig_i))
                continue
        if 'sous_titre_ABC' not in roles and re.match(r'^[A-E]\.\s', clean):
            roles['sous_titre_ABC'] = (orig_i, count_trailing_spacers(orig_i))
            continue
        if 'sous_niveau_etage' not in roles and FLOOR_PATTERNS.match(clean):
            roles['sous_niveau_etage'] = (orig_i, count_trailing_spacers(orig_i))
            continue
        if 'label_contenant' not in roles and re.match(r'^(Contenant|Contenu)\b', clean):
            roles['label_contenant'] = (orig_i, count_trailing_spacers(orig_i))
            continue
        if 'montant' not in roles and re.match(r'^[\d\.,]+\s*€$', clean):
            roles['montant'] = (orig_i, count_trailing_spacers(orig_i))
            continue
        if 'nom_piece' not in roles:
            if re.match(r'^(Contenant|Contenu)\b', next_clean) and ':' not in clean:
                roles['nom_piece'] = (orig_i, count_trailing_spacers(orig_i))
                continue
        if 'liste_intro' not in roles and ':' in clean and 'sous_titre_ABC' not in roles and 'titre_souligne' not in roles:
            roles['liste_intro'] = (orig_i, count_trailing_spacers(orig_i))
            continue
        if 'bullet_item' not in roles and ':' in clean:
            roles['bullet_item'] = (orig_i, count_trailing_spacers(orig_i))
            continue

    required = ['titre_niveau1', 'titre_souligne', 'sous_titre_ABC', 'sous_niveau_etage',
                'nom_piece', 'label_contenant', 'montant', 'bullet_item', 'liste_intro']
    missing = [r for r in required if r not in roles]
    if missing:
        raise ValueError(
            f"Kon deze structuur-onderdelen niet herkennen in het sjabloon: {missing}. "
            f"Zorg dat het sjabloon minstens een voorbeeld van elk onderdeel bevat "
            f"(titel, onderlijnde titel, subtitel A./B./..., verdiepingsniveau, "
            f"ruimtenaam gevolgd door Contenant/Contenu, en een bullet met ':')."
        )
    return roles


def find_first_empty_paragraph(paras):
    """Vindt de eerste lege paragraaf in het sjabloon, te gebruiken als
    universeel formaat voor spacer-regels (witruimte tussen elementen)."""
    for p in paras:
        if not para_text(p).strip():
            return p
    return None


def classify_line(text, idx, all_lines):
    clean = strip_underline_unicode(text).strip()
    if any(re.match(pat, clean) for pat in MAIN_TITLE_PATTERNS):
        return 'titre_souligne' if idx > 0 else 'titre_niveau1'
    if FLOOR_PATTERNS.match(clean):
        return 'sous_niveau_etage'
    if re.match(r'^[A-E]\.\s', clean):
        return 'sous_titre_ABC'
    if re.match(r'^(Contenant|Contenu)\b', clean):
        return 'label_contenant'
    if idx <= 6 and ':' in clean and not re.match(r'^[A-E]\.\s', clean):
        return 'liste_intro'
    nxt = all_lines[idx + 1] if idx + 1 < len(all_lines) else ''
    nxt_clean = strip_underline_unicode(nxt).strip()
    if re.match(r'^(Contenant|Contenu)\b', nxt_clean) and ':' not in clean:
        return 'nom_piece'
    return 'bullet_item'


def split_title_amount(text):
    m = re.search(r'([\d\.,]+\s*€)\s*$', text)
    if m and len(text) - m.end() <= 2:
        title_part = text[:m.start()].strip(' .…\t')
        return title_part, m.group(1).strip()
    return text, None


def split_label_amount(text):
    m = re.match(r'^(Contenant|Contenu)\b.*?([\d\.,]+\s*€)?\s*$', text)
    label = m.group(1)
    amount = m.group(2)
    return label, amount


def build_final_items(lines):
    texts = [strip_underline_unicode(l).strip() for l in lines]
    roled = [(t, classify_line(t, i, texts)) for i, t in enumerate(texts)]
    final_items = []
    for text, role in roled:
        if role == 'sous_titre_ABC' and '€' in text:
            title_part, amount_part = split_title_amount(text)
            final_items.append((title_part, 'sous_titre_ABC'))
            if amount_part:
                final_items.append((amount_part, 'montant'))
        elif role == 'label_contenant':
            label, amount = split_label_amount(text)
            final_items.append((label, 'label_contenant'))
            if amount:
                final_items.append((amount, 'montant'))
        else:
            final_items.append((text, role))
    return final_items


def clone_paragraph_with_text(source_para, new_text):
    p_new = copy.deepcopy(source_para)
    runs = p_new.findall('w:r', NS)
    text_runs = [r for r in runs if r.find('w:t', NS) is not None]
    if not text_runs:
        return p_new
    keep = text_runs[0]
    for r in runs:
        if r is not keep:
            p_new.remove(r)
    t_el = keep.find('w:t', NS)
    t_el.text = new_text
    t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return p_new


def read_input_lines_from_bytes(docx_bytes):
    doc = Document(BytesIO(docx_bytes))
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]


def fill_template_from_bytes(template_bytes, input_bytes):
    """
    Hoofdfunctie voor de app: neemt de bytes van het sjabloonbestand en het
    ongeformatteerde inputbestand, en retourneert de bytes van het resulterende
    .docx-bestand plus het aantal verwerkte alinea's (inclusief spacers).

    Voegt na het LAATSTE item van elke opeenvolgende reeks van dezelfde rol
    de lege spacer-alinea's toe die in het sjabloon op die plek stonden, zodat
    de witruimte tussen structuurelementen (bv. na een ruimtenaam, na het
    laatste bullet-item vóór Contenu) behouden blijft.
    """
    tree, body, paras = load_paragraphs_from_bytes(template_bytes)
    role_indices = find_role_source_indices(paras)
    spacer_template = find_first_empty_paragraph(paras)

    lines = read_input_lines_from_bytes(input_bytes)
    final_items = build_final_items(lines)

    new_body = etree.Element(f'{W}body', nsmap=NS)
    n = len(final_items)
    for i, (text, role) in enumerate(final_items):
        role_use = role if role in role_indices else 'bullet_item'
        source_idx, n_spacers = role_indices[role_use]
        new_body.append(clone_paragraph_with_text(paras[source_idx], text))

        next_role = final_items[i + 1][1] if i + 1 < n else None
        is_last_of_group = (next_role != role)
        if is_last_of_group and n_spacers > 0 and spacer_template is not None:
            for _ in range(n_spacers):
                new_body.append(copy.deepcopy(spacer_template))

    sect_pr = body.find('w:sectPr', NS)
    if sect_pr is not None:
        new_body.append(copy.deepcopy(sect_pr))

    new_tree = copy.deepcopy(tree)
    old_body = new_tree.find('w:body', NS)
    new_tree.remove(old_body)
    new_tree.append(new_body)

    new_doc_xml = etree.tostring(new_tree, xml_declaration=True, encoding='UTF-8', standalone=True)

    out_buffer = BytesIO()
    with zipfile.ZipFile(BytesIO(template_bytes), 'r') as zin:
        with zipfile.ZipFile(out_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'word/document.xml':
                    zout.writestr(item, new_doc_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))

    out_buffer.seek(0)
    return out_buffer.read(), len(final_items)
