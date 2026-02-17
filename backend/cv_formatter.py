"""
cv_formatter.py — CV Maltem Africa — 100% fidèle référence ZAID
"""
import os, io, re
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

RED   = RGBColor(0xC0, 0x00, 0x00)
TITLE = RGBColor(0xBE, 0x3B, 0x4E)
BLACK = RGBColor(0x11, 0x11, 0x11)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREY  = RGBColor(0x60, 0x60, 0x60)

ASSETS  = os.path.join(os.path.dirname(__file__), "assets")
LOGO    = os.path.join(ASSETS, "logo_maltem.png")
DECO_TL = os.path.join(ASSETS, "deco_top_left_white.png")
DECO_BR = os.path.join(ASSETS, "deco_bottom_right_white.png")


def sf(run, size_pt=10, bold=False, italic=False, underline=False, color=None, font="Century Gothic"):
    run.font.name=font; run.font.size=Pt(size_pt); run.font.bold=bold
    run.font.italic=italic; run.font.underline=underline
    if color: run.font.color.rgb=color
    rPr=run._r.get_or_add_rPr()
    rF=OxmlElement('w:rFonts')
    for a in ('w:ascii','w:hAnsi','w:eastAsia','w:cs'): rF.set(qn(a),font)
    ex=rPr.find(qn('w:rFonts'))
    if ex is not None: rPr.remove(ex)
    rPr.insert(0,rF)

def sp(para, before=0, after=0, line=240):
    pPr=para._p.get_or_add_pPr()
    s=pPr.find(qn('w:spacing'))
    if s is None: s=OxmlElement('w:spacing'); pPr.append(s)
    s.set(qn('w:before'),str(int(before*20))); s.set(qn('w:after'),str(int(after*20)))
    s.set(qn('w:line'),str(line)); s.set(qn('w:lineRule'),'auto')

def ind(para, left=0):
    pPr=para._p.get_or_add_pPr()
    i=pPr.find(qn('w:ind'))
    if i is None: i=OxmlElement('w:ind'); pPr.append(i)
    if left: i.set(qn('w:left'),str(left))
    elif qn('w:left') in i.attrib: del i.attrib[qn('w:left')]

def bdr(para, top=False, bottom=False, color="auto", sz="12"):
    pPr=para._p.get_or_add_pPr()
    ex=pPr.find(qn('w:pBdr'))
    if ex is not None: pPr.remove(ex)
    pBdr=OxmlElement('w:pBdr')
    for side,do in [('top',top),('bottom',bottom)]:
        if do:
            e=OxmlElement(f'w:{side}')
            e.set(qn('w:val'),'single'); e.set(qn('w:sz'),sz)
            e.set(qn('w:space'),'1'); e.set(qn('w:color'),color)
            pBdr.append(e)
    pPr.append(pBdr)

def get_rId_from_run(run):
    """Extraire le rId d'une image ajoutée via add_picture."""
    m = re.search(r'r:embed="(rId\d+)"', run._r.xml)
    return m.group(1) if m else None


def make_anchor_from_inline(run, cx, cy, posH_offset, posV_offset, doc_id,
                              relH="column", relV="paragraph", behindDoc="0"):
    """Convertit un inline drawing en anchor drawing."""
    r_xml = run._r.xml
    m = re.search(r'r:embed="(rId\d+)"', r_xml)
    if not m: return None
    rId = m.group(1)

    anchor_xml = f'''<w:drawing
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"
    relativeHeight="251658240" behindDoc="{behindDoc}" locked="0" layoutInCell="1" allowOverlap="1">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="{relH}"><wp:posOffset>{posH_offset}</wp:posOffset></wp:positionH>
    <wp:positionV relativeFrom="{relV}"><wp:posOffset>{posV_offset}</wp:posOffset></wp:positionV>
    <wp:extent cx="{cx}" cy="{cy}"/>
    <wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:wrapNone/>
    <wp:docPr id="{doc_id}" name="img{doc_id}"/>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:nvPicPr>
            <pic:cNvPr id="0" name="img{doc_id}"/>
            <pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr>
          </pic:nvPicPr>
          <pic:blipFill>
            <a:blip r:embed="{rId}"/>
            <a:stretch><a:fillRect/></a:stretch>
          </pic:blipFill>
          <pic:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:anchor>
</w:drawing>'''
    return etree.fromstring(anchor_xml)


def make_red_bar_xml():
    return etree.fromstring('''<w:drawing
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"
    relativeHeight="251659264" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="page"><wp:align>left</wp:align></wp:positionH>
    <wp:positionV relativeFrom="paragraph"><wp:posOffset>-647700</wp:posOffset></wp:positionV>
    <wp:extent cx="7560310" cy="72390"/>
    <wp:effectExtent l="0" t="0" r="2540" b="3810"/>
    <wp:wrapNone/>
    <wp:docPr id="1" name="RedBar"/>
    <a:graphic>
      <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <wps:wsp>
          <wps:cNvSpPr><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>
          <wps:spPr bwMode="auto">
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="7560310" cy="72390"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="E9272D"/></a:solidFill>
            <a:ln><a:noFill/></a:ln>
          </wps:spPr>
          <wps:bodyPr rot="0" wrap="square" anchor="t"><a:noAutofit/></wps:bodyPr>
        </wps:wsp>
      </a:graphicData>
    </a:graphic>
  </wp:anchor>
</w:drawing>''')


def make_badge_xml(periode_text, doc_id=100):
    """Badge rouge arrondi avec texte blanc - construction correcte via etree."""
    cx = 1097280
    cy = 246888

    base_xml = f'''<w:drawing
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"
    relativeHeight="251664384" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="column"><wp:posOffset>-50800</wp:posOffset></wp:positionH>
    <wp:positionV relativeFrom="paragraph"><wp:posOffset>100000</wp:posOffset></wp:positionV>
    <wp:extent cx="{cx}" cy="{cy}"/>
    <wp:effectExtent l="0" t="0" r="1270" b="5715"/>
    <wp:wrapNone/>
    <wp:docPr id="{doc_id}" name="badge{doc_id}"/>
    <a:graphic>
      <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <wps:wsp>
          <wps:cNvSpPr><a:spLocks/></wps:cNvSpPr>
          <wps:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
            <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>
            <a:solidFill><a:srgbClr val="C00000"/></a:solidFill>
            <a:ln><a:noFill/></a:ln>
          </wps:spPr>
          <wps:txbx>
            <w:txbxContent>
              <w:p>
                <w:pPr><w:jc w:val="center"/>
                  <w:rPr><w:rFonts w:ascii="Century Gothic" w:hAnsi="Century Gothic" w:cs="Century Gothic"/>
                    <w:color w:val="FFFFFF"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>
                </w:pPr>
                <w:r><w:rPr>
                    <w:rFonts w:ascii="Century Gothic" w:hAnsi="Century Gothic" w:cs="Century Gothic"/>
                    <w:color w:val="FFFFFF"/><w:sz w:val="16"/><w:szCs w:val="16"/>
                  </w:rPr><w:t xml:space="preserve">BADGE_PLACEHOLDER</w:t></w:r>
              </w:p>
            </w:txbxContent>
          </wps:txbx>
          <wps:bodyPr rot="0" anchor="ctr"><a:noAutofit/></wps:bodyPr>
        </wps:wsp>
      </a:graphicData>
    </a:graphic>
  </wp:anchor>
</w:drawing>'''
    root = etree.fromstring(base_xml)
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    for wt in root.iter(f"{{{W}}}t"):
        if wt.text == "BADGE_PLACEHOLDER":
            wt.text = periode_text
            break
    return root


def build_header(doc, cv_data):
    p_main = doc.add_paragraph()
    p_main.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sp(p_main, before=0, after=0)

    r_empty = p_main.add_run(" ")
    sf(r_empty, size_pt=10)
    p_main._p.insert(1, make_red_bar_xml())

    r_logo = p_main.add_run()
    r_logo.add_picture(LOGO, width=Cm(6.76))
    logo_anchor = make_anchor_from_inline(r_logo, 2433960, 635257, 4852434, -482157, 44,
                                          relH="page", relV="paragraph", behindDoc="0")
    if logo_anchor is not None:
        r_logo._r.getparent().remove(r_logo._r)
        p_main._p.insert(2, logo_anchor)

    r_tl = p_main.add_run()
    r_tl.add_picture(DECO_TL, width=Cm(1.9))
    tl_anchor = make_anchor_from_inline(r_tl, 682625, 1567180, -106680, -106680, 50,
                                         relH="page", relV="paragraph", behindDoc="1")
    if tl_anchor is not None:
        r_tl._r.getparent().remove(r_tl._r)
        p_main._p.insert(3, tl_anchor)

    # Ligne noire au-dessus du nom
    p_l1 = doc.add_paragraph(); sp(p_l1, before=6, after=4)
    bdr(p_l1, bottom=True, color="231F20", sz="6")

    # Nom + expérience
    p_nom = doc.add_paragraph(); p_nom.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sp(p_nom, before=4, after=2)
    nom = cv_data.get("nom_prenom",""); exp = cv_data.get("annees_experience","")
    r_n = p_nom.add_run(nom); sf(r_n, size_pt=14)
    if exp: r_e = p_nom.add_run(f" – {exp}"); sf(r_e, size_pt=14)

    # Titre poste — gras, centré
    p_t = doc.add_paragraph(); p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sp(p_t, before=0, after=4)
    sf(p_t.add_run(cv_data.get("titre_poste","")), size_pt=14, bold=True)

    # Ligne noire en dessous
    p_l2 = doc.add_paragraph(); sp(p_l2, before=4, after=8)
    bdr(p_l2, bottom=True, color="231F20", sz="6")


def section_title(doc, title):
    """Titre de section : texte gras souligné + ligne horizontale en dessous."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    sp(p, before=10, after=2)
    r = p.add_run(title.upper())
    sf(r, size_pt=10, bold=True, color=BLACK)
    bdr(p, bottom=True, color="231F20", sz="6")
    return p


def subsection_title(doc, title):
    """Sous-titre de catégorie compétences : gras, sans ligne de séparation."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    sp(p, before=6, after=0)
    sf(p.add_run(title), size_pt=10, bold=True, color=BLACK)
    return p


def exp_title(doc):
    """Titre EXPÉRIENCES PROFESSIONNELLES : centré rouge entre deux lignes."""
    p_top = doc.add_paragraph(); sp(p_top, before=10, after=0)
    bdr(p_top, bottom=True, color="231F20", sz="6")
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sp(p, before=4, after=4)
    sf(p.add_run("EXPÉRIENCES PROFESSIONNELLES"), size_pt=12, bold=True, color=TITLE)
    p_bot = doc.add_paragraph(); sp(p_bot, before=0, after=8)
    bdr(p_bot, bottom=True, color="231F20", sz="6")


def exp_badge_line(doc, periode, entreprise, badge_id=100):
    """
    Ligne badge : rectangle rouge arrondi à gauche + entreprise rouge gras à droite.
    """
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    sp(p, before=10, after=2)
    ind(p, left=1800)
    if entreprise:
        # FIX: entreprise en gras + majuscules comme dans la référence
        sf(p.add_run(entreprise.upper()), size_pt=10, bold=True, color=RED)
    if periode:
        try:
            badge = make_badge_xml(periode, doc_id=badge_id)
            p._p.insert(1, badge)
        except Exception:
            r_per = p.add_run(f"[{periode}]  ")
            sf(r_per, size_pt=9, bold=True, color=RED)
    return p


def exp_poste_line(doc, poste):
    """Poste centré et gras sous le badge — aligné au centre comme dans la référence."""
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sp(p, before=2, after=4)
    sf(p.add_run(poste), size_pt=10, bold=True, color=BLACK)
    return p


def exp_label(doc, label):
    """Label de sous-section expérience : gras, sans ligne de séparation sous lui."""
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    sp(p, before=6, after=2)
    sf(p.add_run(label), size_pt=10, bold=True, color=BLACK)
    return p


def exp_body(doc, text, indent_twips=426):
    """Corps de texte justifié avec indentation légère."""
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    sp(p, before=0, after=2); ind(p, left=indent_twips)
    sf(p.add_run(text), size_pt=10, color=BLACK)
    return p


def exp_item_with_bold_prefix(doc, text, indent_twips=426):
    """
    Item de liste avec le début en gras jusqu'au premier ':' ou '–'.
    Ex: "Optimisation & Diagnostic JVM : Analyse..." → début en gras, suite normale.
    Correspond exactement au style du CV référence pour les réalisations.
    """
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    sp(p, before=1, after=1); ind(p, left=indent_twips)

    # Tiret Maltem style (espace + tiret)
    r_dash = p.add_run("- ")
    sf(r_dash, size_pt=10, color=BLACK)

    # Chercher le premier ':' ou '–' pour séparer partie gras / partie normale
    split_idx = -1
    for sep in [':', ' –', ' - ']:
        idx = text.find(sep)
        if idx != -1:
            split_idx = idx + len(sep)
            break

    if split_idx > 0:
        bold_part = text[:split_idx]
        normal_part = text[split_idx:]
        r_bold = p.add_run(bold_part)
        sf(r_bold, size_pt=10, bold=True, color=BLACK)
        if normal_part.strip():
            r_normal = p.add_run(normal_part)
            sf(r_normal, size_pt=10, color=BLACK)
    else:
        r_full = p.add_run(text)
        sf(r_full, size_pt=10, color=BLACK)

    return p


def exp_item(doc, text, indent_twips=426):
    """Item de liste simple avec tiret Maltem style."""
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    sp(p, before=1, after=1); ind(p, left=indent_twips)
    r_dash = p.add_run("- ")
    sf(r_dash, size_pt=10, color=BLACK)
    r_text = p.add_run(text)
    sf(r_text, size_pt=10, color=BLACK)
    return p


# ─────────────────────────────────────────────────────────────────────────────
# COMPÉTENCES EN 2 COLONNES (tableau) — comme dans le CV référence
# ─────────────────────────────────────────────────────────────────────────────

def build_competences_table(doc, competences):
    """
    Affiche les compétences en tableau 2 colonnes, sans bordures visibles,
    exactement comme dans le CV référence Maltem.
    Chaque colonne contient des sous-titres gras + items avec tirets.
    """
    if not competences:
        return

    section_title(doc, "COMPÉTENCES")

    # Diviser les catégories en 2 colonnes équilibrées
    mid = (len(competences) + 1) // 2
    col_left  = competences[:mid]
    col_right = competences[mid:]

    # Largeur des colonnes (page 21cm - 3.6cm marges = 17.4cm / 2 = 8.7cm)
    col_width = Cm(8.5)

    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = 'Table Grid'

    # Supprimer toutes les bordures du tableau
    tbl_el = tbl._tbl
    tblPr = tbl_el.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl_el.insert(0, tblPr)
    tblBorders = OxmlElement('w:tblBorders')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'none')
        b.set(qn('w:sz'), '0')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), 'auto')
        tblBorders.append(b)
    tblPr.append(tblBorders)

    # Définir la largeur des colonnes
    for i, cell in enumerate(tbl.rows[0].cells):
        tc = cell._tc
        tcPr = tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = OxmlElement('w:tcPr')
            tc.insert(0, tcPr)
        tcW = OxmlElement('w:tcW')
        tcW.set(qn('w:w'), str(int(col_width.emu / 914.4)))  # EMU → twips
        tcW.set(qn('w:type'), 'dxa')
        tcPr.append(tcW)
        # Supprimer les bordures de cellule
        tcBorders = OxmlElement('w:tcBorders')
        for side in ['top', 'left', 'bottom', 'right']:
            b = OxmlElement(f'w:{side}')
            b.set(qn('w:val'), 'none')
            tcBorders.append(b)
        tcPr.append(tcBorders)

    def fill_column(cell, categories):
        """Remplir une cellule du tableau avec les catégories de compétences."""
        # Vider le paragraphe vide par défaut
        for para in cell.paragraphs:
            para.clear()

        first = True
        for cat in categories:
            # Sous-titre catégorie en gras
            if cat.get("categorie"):
                if first:
                    p = cell.paragraphs[0]
                    first = False
                else:
                    p = cell.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                sp(p, before=6, after=0)
                sf(p.add_run(cat["categorie"]), size_pt=10, bold=True, color=BLACK)

            # Items avec tiret
            for item in cat.get("items", []):
                if first:
                    p = cell.paragraphs[0]
                    first = False
                else:
                    p = cell.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                sp(p, before=0, after=0)
                ind(p, left=240)
                r_dash = p.add_run("- ")
                sf(r_dash, size_pt=10, color=BLACK)
                r_item = p.add_run(item)
                sf(r_item, size_pt=10, color=BLACK)

    fill_column(tbl.rows[0].cells[0], col_left)
    fill_column(tbl.rows[0].cells[1], col_right)

    # Espace après le tableau
    p_after = doc.add_paragraph()
    sp(p_after, before=4, after=4)


def build_footer(doc):
    section = doc.sections[0]; footer = section.footer
    for para in footer.paragraphs: para.clear()
    fp = footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    sp(fp, before=0, after=0)

    pPr = fp._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr'); top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single'); top.set(qn('w:sz'), '6')
    top.set(qn('w:space'), '1'); top.set(qn('w:color'), 'auto')
    pBdr.append(top); pPr.append(pBdr)

    tabs_el = OxmlElement('w:tabs')
    for val, pos in [('center', '4536'), ('right', '9026')]:
        t = OxmlElement('w:tab'); t.set(qn('w:val'), val); t.set(qn('w:pos'), pos)
        tabs_el.append(t)
    pPr.append(tabs_el)

    sf(fp.add_run("MALTEM AFRICA"), size_pt=8, color=GREY)
    fp.add_run("\t").font.size = Pt(8)

    def field(para, instr):
        r = para.add_run(); r.font.size = Pt(8); r.font.color.rgb = GREY
        f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
        it = OxmlElement('w:instrText'); it.text = instr
        f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'end')
        r._r.extend([f1, it, f2])

    sf(fp.add_run("\tPage "), size_pt=8, color=GREY)
    field(fp, ' PAGE ')
    sf(fp.add_run(" sur "), size_pt=8, color=GREY)
    field(fp, ' NUMPAGES ')

    # Deco bas-droite
    r_br = fp.add_run()
    r_br.add_picture(DECO_BR, width=Cm(1.87))
    br_anchor = make_anchor_from_inline(r_br, 673100, 1567180, 6886900, -1200000, 51,
                                        relH="page", relV="paragraph", behindDoc="1")
    if br_anchor is not None:
        r_br._r.getparent().remove(r_br._r)
        fp._p.insert(1, br_anchor)

    # Deco haut-gauche (depuis le footer avec position négative)
    r_tl = fp.add_run()
    r_tl.add_picture(DECO_TL, width=Cm(1.9))
    tl_anchor = make_anchor_from_inline(r_tl, 682625, 1567180, -106680, -9500000, 52,
                                         relH="page", relV="paragraph", behindDoc="1")
    if tl_anchor is not None:
        r_tl._r.getparent().remove(r_tl._r)
        fp._p.insert(1, tl_anchor)


def generate_maltem_cv(cv_data: dict, output_path: str) -> str:
    doc = Document()
    s = doc.sections[0]
    s.page_height = Cm(29.7); s.page_width = Cm(21.0)
    s.left_margin = Cm(1.8);  s.right_margin = Cm(1.8)
    s.top_margin  = Cm(1.8);  s.bottom_margin = Cm(1.8)
    s.footer_distance = Cm(0.8)
    doc.styles['Normal'].font.name = 'Century Gothic'
    doc.styles['Normal'].font.size = Pt(10)

    # ── EN-TÊTE ──────────────────────────────────────────────────────────────
    build_header(doc, cv_data)

    # ── À PROPOS ─────────────────────────────────────────────────────────────
    a_propos = cv_data.get("a_propos", "")
    if a_propos:
        section_title(doc, "À PROPOS")
        # Traiter les lignes séparées par '\n' ou liste
        lignes = a_propos if isinstance(a_propos, list) else [a_propos]
        for ligne in lignes:
            if ligne.strip():
                p_ap = doc.add_paragraph(); p_ap.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                sp(p_ap, before=4, after=2); ind(p_ap, left=426)
                r_dash = p_ap.add_run("- ")
                sf(r_dash, size_pt=10, color=BLACK)
                sf(p_ap.add_run(ligne.strip()), size_pt=10, color=BLACK)
        p_space = doc.add_paragraph(); sp(p_space, before=0, after=6)

    # ── COMPÉTENCES (2 colonnes) ──────────────────────────────────────────────
    competences = cv_data.get("competences", [])
    if competences:
        build_competences_table(doc, competences)

    # ── FORMATION ────────────────────────────────────────────────────────────
    formations = cv_data.get("formations", [])
    if formations:
        section_title(doc, "FORMATION")
        for form in formations:
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            sp(p, before=4, after=2)
            annee = form.get("annee", "")
            diplome = form.get("diplome", "")
            etablissement = form.get("etablissement", "")
            if annee:
                r_yr = p.add_run(f"{annee}  ")
                sf(r_yr, size_pt=10, bold=True, color=BLACK)
            txt = diplome
            if etablissement: txt += f" – {etablissement}"
            sf(p.add_run(txt), size_pt=10, color=BLACK)

    # ── CERTIFICATIONS ───────────────────────────────────────────────────────
    certs = cv_data.get("certifications", [])
    if certs:
        section_title(doc, "CERTIFICATIONS")
        for cert in certs:
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            sp(p, before=0, after=0); ind(p, left=426)
            r_dash = p.add_run("- ")
            sf(r_dash, size_pt=10, color=BLACK)
            sf(p.add_run(cert), size_pt=10, color=BLACK)

    # ── LANGUES ──────────────────────────────────────────────────────────────
    langues = cv_data.get("langues", [])
    if langues:
        section_title(doc, "LANGUES")
        for langue in langues:
            p = doc.add_paragraph(); sp(p, before=0, after=0); ind(p, left=426)
            r_dash = p.add_run("- ")
            sf(r_dash, size_pt=10, color=BLACK)
            sf(p.add_run(langue), size_pt=10, color=BLACK)

    # ── AUTRES RÉFÉRENCES ────────────────────────────────────────────────────
    refs = cv_data.get("autres_references", [])
    if refs:
        section_title(doc, "AUTRES RÉFÉRENCES")
        for ref in refs:
            p = doc.add_paragraph(); sp(p, before=0, after=0)
            sf(p.add_run(f"{ref.get('entreprise', '')} : "), size_pt=10, bold=True, color=RED)
            sf(p.add_run(ref.get("poste", "")), size_pt=10, color=BLACK)

    # ── PROJETS MARQUANTS ────────────────────────────────────────────────────
    projets = cv_data.get("projets_marquants", [])
    if projets:
        section_title(doc, "PROJETS MARQUANTS")
        for projet in projets:
            p = doc.add_paragraph(); sp(p, before=0, after=0); ind(p, left=426)
            r_dash = p.add_run("- ")
            sf(r_dash, size_pt=10, color=BLACK)
            sf(p.add_run(projet), size_pt=10, color=BLACK)

    # ── EXPÉRIENCES PROFESSIONNELLES ─────────────────────────────────────────
    experiences = cv_data.get("experiences", [])
    if experiences:
        exp_title(doc)
        badge_id = 100

        for exp in experiences:
            periode      = exp.get("periode", "")
            entreprise   = exp.get("entreprise", "")
            poste        = exp.get("poste", "")
            direction    = exp.get("direction", "")
            contexte     = exp.get("contexte", "")
            objectifs    = exp.get("objectifs", [])
            missions     = exp.get("missions", [])
            realisations = exp.get("realisations", [])
            resultats    = exp.get("resultats", [])
            environnement = exp.get("environnement", "")

            # Badge période + entreprise gras rouge en majuscules
            exp_badge_line(doc, periode, entreprise, badge_id=badge_id)
            badge_id += 1

            # Poste centré gras
            if poste:
                exp_poste_line(doc, poste)

            # Direction (si présente) centré gras
            if direction:
                p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                sp(p, before=0, after=2)
                sf(p.add_run(direction), size_pt=10, bold=True, color=BLACK)

            # Contexte & enjeux — label gras + body justifié
            if contexte:
                exp_label(doc, "Contexte & enjeux :")
                if isinstance(contexte, list):
                    for c in contexte:
                        exp_item(doc, c)
                else:
                    exp_body(doc, contexte)

            # Objectifs
            if objectifs:
                exp_label(doc, "Objectifs :")
                for o in objectifs:
                    exp_item(doc, o)

            # Réalisations — avec premier terme en gras avant ':'
            reals = realisations if realisations else missions
            if reals:
                exp_label(doc, "Réalisations :")
                for r in reals:
                    exp_item_with_bold_prefix(doc, r)

            # Résultats / impacts
            if resultats:
                exp_label(doc, "Résultats / impacts :")
                for r in resultats:
                    exp_item(doc, r)

            # Environnement
            if environnement:
                exp_label(doc, "Environnement :")
                exp_body(doc, environnement)

            # Espace entre expériences
            p_sep = doc.add_paragraph(); sp(p_sep, before=4, after=0)

    # ── PIED DE PAGE ─────────────────────────────────────────────────────────
    build_footer(doc)

    doc.save(output_path)
    return output_path
