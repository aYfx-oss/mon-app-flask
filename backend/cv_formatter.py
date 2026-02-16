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
    # Trouver l'embed rId
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
    # Paragraphe avec barre rouge + logo + deco TL
    p_main = doc.add_paragraph()
    p_main.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sp(p_main, before=0, after=0)

    # Barre rouge
    r_empty = p_main.add_run(" ")
    sf(r_empty, size_pt=10)
    p_main._p.insert(1, make_red_bar_xml())

    # Logo via add_picture inline puis convertir en anchor
    r_logo = p_main.add_run()
    r_logo.add_picture(LOGO, width=Cm(6.76))
    logo_anchor = make_anchor_from_inline(r_logo, 2433960, 635257, 4852434, -482157, 44,
                                          relH="page", relV="paragraph", behindDoc="0")
    if logo_anchor is not None:
        r_logo._r.getparent().remove(r_logo._r)
        p_main._p.insert(2, logo_anchor)
    
    # Deco haut-gauche via add_picture inline puis anchor
    r_tl = p_main.add_run()
    r_tl.add_picture(DECO_TL, width=Cm(1.9))
    tl_anchor = make_anchor_from_inline(r_tl, 682625, 1567180, -106680, -106680, 50,
                                         relH="page", relV="paragraph", behindDoc="1")
    if tl_anchor is not None:
        r_tl._r.getparent().remove(r_tl._r)
        p_main._p.insert(3, tl_anchor)

    # Ligne noire au-dessus
    p_l1 = doc.add_paragraph(); sp(p_l1, before=6, after=4)
    bdr(p_l1, bottom=True, color="231F20", sz="6")

    # Nom + exp
    p_nom = doc.add_paragraph(); p_nom.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sp(p_nom, before=4, after=2)
    nom = cv_data.get("nom_prenom",""); exp = cv_data.get("annees_experience","")
    r_n = p_nom.add_run(nom); sf(r_n, size_pt=14)
    if exp: r_e = p_nom.add_run(f"– {exp}"); sf(r_e, size_pt=14)

    # Titre
    p_t = doc.add_paragraph(); p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sp(p_t, before=0, after=4)
    sf(p_t.add_run(cv_data.get("titre_poste","")), size_pt=14, bold=True)

    # Ligne noire en dessous
    p_l2 = doc.add_paragraph(); sp(p_l2, before=4, after=8)
    bdr(p_l2, bottom=True, color="231F20", sz="6")


def section_title(doc, title):
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    sp(p, before=8, after=0); ind(p, left=426)
    sf(p.add_run(title.upper()), size_pt=10, bold=True, color=BLACK)
    bdr(p, bottom=True, color="auto", sz="12")
    return p


def subsection_title(doc, title):
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    sp(p, before=8, after=0); ind(p, left=340)
    sf(p.add_run(title), size_pt=10, bold=True)
    bdr(p, bottom=True, color="auto", sz="12")
    return p


def exp_title(doc):
    p_top=doc.add_paragraph(); sp(p_top, before=10, after=0)
    bdr(p_top, bottom=True, color="231F20", sz="6")
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    sp(p, before=4, after=4)
    sf(p.add_run("EXPÉRIENCES PROFESSIONNELLES"), size_pt=12, color=TITLE)
    p_bot=doc.add_paragraph(); sp(p_bot, before=0, after=8)
    bdr(p_bot, bottom=True, color="231F20", sz="6")


def exp_badge_line(doc, periode, entreprise, badge_id=100):
    """
    Ligne badge : rectangle rouge arrondi à gauche + entreprise rouge à droite.
    Le badge est ancré flottant, le paragraphe a juste l'indentation pour laisser la place.
    """
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    sp(p, before=10, after=2)
    # Indentation gauche pour laisser place au badge (3.12cm ≈ 1769 twips)
    ind(p, left=1800)
    # Entreprise en rouge après le badge
    if entreprise:
        sf(p.add_run(entreprise), size_pt=10, color=RED)
    # Insérer le badge ancré AVANT le run de l'entreprise
    if periode:
        try:
            badge = make_badge_xml(periode, doc_id=badge_id)
            p._p.insert(1, badge)
        except Exception as e:
            # Fallback texte
            r_per = p.add_run(f"[{periode}]  ")
            sf(r_per, size_pt=9, bold=True, color=RED)
    return p


def exp_label(doc, label, indent_twips=2124):
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    sp(p, before=6, after=0); ind(p, left=indent_twips)
    sf(p.add_run(label), size_pt=10, bold=True, color=BLACK)
    return p


def exp_body(doc, text, indent_twips=2124, bold=False):
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    sp(p, before=0, after=0); ind(p, left=indent_twips)
    sf(p.add_run(text), size_pt=10, bold=bold, color=BLACK)
    return p


def exp_item(doc, text, indent_twips=2500):
    p=doc.add_paragraph(style='List Bullet'); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    sp(p, before=0, after=0)
    sf(p.add_run(text), size_pt=10, color=BLACK)
    pPr=p._p.get_or_add_pPr(); i=pPr.find(qn('w:ind'))
    if i is None: i=OxmlElement('w:ind'); pPr.append(i)
    i.set(qn('w:left'), str(indent_twips))
    return p


def build_footer(doc):
    section=doc.sections[0]; footer=section.footer
    for para in footer.paragraphs: para.clear()
    fp=footer.paragraphs[0]; fp.alignment=WD_ALIGN_PARAGRAPH.LEFT
    sp(fp, before=0, after=0)

    pPr=fp._p.get_or_add_pPr()
    pBdr=OxmlElement('w:pBdr'); top=OxmlElement('w:top')
    top.set(qn('w:val'),'single'); top.set(qn('w:sz'),'6')
    top.set(qn('w:space'),'1'); top.set(qn('w:color'),'auto')
    pBdr.append(top); pPr.append(pBdr)

    tabs_el=OxmlElement('w:tabs')
    for val,pos in [('center','4536'),('right','9026')]:
        t=OxmlElement('w:tab'); t.set(qn('w:val'),val); t.set(qn('w:pos'),pos); tabs_el.append(t)
    pPr.append(tabs_el)

    sf(fp.add_run("MALTEM AFRICA"), size_pt=8, color=GREY)
    fp.add_run("\t").font.size=Pt(8)

    def field(para, instr):
        r=para.add_run(); r.font.size=Pt(8); r.font.color.rgb=GREY
        f1=OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'),'begin')
        it=OxmlElement('w:instrText'); it.text=instr
        f2=OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'),'end')
        r._r.extend([f1,it,f2])

    sf(fp.add_run("\tPage "), size_pt=8, color=GREY)
    field(fp, ' PAGE '); sf(fp.add_run(" sur "), size_pt=8, color=GREY); field(fp, ' NUMPAGES ')

    # Deco bas-droite dans footer (bas droite de chaque page)
    r_br=fp.add_run()
    r_br.add_picture(DECO_BR, width=Cm(1.87))
    br_anchor=make_anchor_from_inline(r_br, 673100, 1567180, 6886900, -1200000, 51,
                                      relH="page", relV="paragraph", behindDoc="1")
    if br_anchor is not None:
        r_br._r.getparent().remove(r_br._r)
        fp._p.insert(1, br_anchor)

    # Deco haut-gauche dans footer (haut de chaque page via position négative)
    r_tl=fp.add_run()
    r_tl.add_picture(DECO_TL, width=Cm(1.9))
    # posV très négatif pour remonter en haut de page depuis le footer
    # footer est à ~27cm du haut, deco doit être à ~0cm => offset = -27cm = -9720000 EMU
    tl_anchor=make_anchor_from_inline(r_tl, 682625, 1567180, -106680, -9500000, 52,
                                       relH="page", relV="paragraph", behindDoc="1")
    if tl_anchor is not None:
        r_tl._r.getparent().remove(r_tl._r)
        fp._p.insert(1, tl_anchor)


def generate_maltem_cv(cv_data: dict, output_path: str) -> str:
    doc=Document()
    s=doc.sections[0]
    s.page_height=Cm(29.7); s.page_width=Cm(21.0)
    s.left_margin=Cm(1.8);  s.right_margin=Cm(1.8)
    s.top_margin=Cm(1.8);   s.bottom_margin=Cm(1.8)
    s.footer_distance=Cm(0.8)
    doc.styles['Normal'].font.name='Century Gothic'
    doc.styles['Normal'].font.size=Pt(10)

    build_header(doc, cv_data)

    a_propos=cv_data.get("a_propos","")
    section_title(doc, "À PROPOS")
    if a_propos:
        p_ap=doc.add_paragraph(); p_ap.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        sp(p_ap, before=4, after=8)
        ind(p_ap, left=426)
        sf(p_ap.add_run(a_propos), size_pt=10)
    else:
        p_ap=doc.add_paragraph(); sp(p_ap, before=4, after=8)

    competences=cv_data.get("competences",[])
    if competences:
        section_title(doc, "COMPÉTENCES")
        for cat in competences:
            if cat.get("categorie"): subsection_title(doc, cat["categorie"])
            for item in cat.get("items",[]):
                p=doc.add_paragraph(style='List Bullet')
                p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; sp(p, before=8, after=0)
                sf(p.add_run(item), size_pt=10)

    formations=cv_data.get("formations",[])
    if formations:
        section_title(doc, "FORMATION")
        for form in formations:
            p=doc.add_paragraph(style='List Bullet')
            p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; sp(p, before=8, after=0)
            txt=""
            if form.get("diplome"): txt+=form["diplome"]
            if form.get("etablissement"): txt+=f" – {form['etablissement']}"
            if form.get("annee"): txt+=f" ({form['annee']})"
            sf(p.add_run(txt), size_pt=10, bold=True)

    certs=cv_data.get("certifications",[])
    if certs:
        section_title(doc, "CERTIFICATIONS")
        for cert in certs:
            p=doc.add_paragraph(style='List Bullet')
            p.alignment=WD_ALIGN_PARAGRAPH.LEFT; sp(p, before=0, after=0)
            sf(p.add_run(cert), size_pt=10, color=BLACK)

    langues=cv_data.get("langues",[])
    if langues:
        section_title(doc, "LANGUES")
        for langue in langues:
            p=doc.add_paragraph(); sp(p, before=0, after=0)
            sf(p.add_run(langue), size_pt=10, color=BLACK)

    experiences=cv_data.get("experiences",[])
    if experiences:
        exp_title(doc)
        badge_id=100
        for exp in experiences:
            periode=exp.get("periode",""); entreprise=exp.get("entreprise","")
            poste=exp.get("poste",""); direction=exp.get("direction","")
            contexte=exp.get("contexte",""); objectifs=exp.get("objectifs",[])
            missions=exp.get("missions",[]); realisations=exp.get("realisations",[])
            resultats=exp.get("resultats",[]); environnement=exp.get("environnement","")

            exp_badge_line(doc, periode, entreprise, badge_id=badge_id); badge_id+=1

            if poste:
                p=doc.add_paragraph(); sp(p, before=0, after=0); ind(p, left=2124)
                sf(p.add_run(poste), size_pt=10, bold=True, color=BLACK)
            if direction:
                p=doc.add_paragraph(); sp(p, before=0, after=0); ind(p, left=2124)
                sf(p.add_run(direction), size_pt=10, bold=True, color=BLACK)
            if contexte: exp_label(doc, "Contexte & enjeux"); exp_body(doc, contexte)
            if objectifs:
                exp_label(doc, "Objectifs")
                for o in objectifs: exp_item(doc, o)
            reals=realisations if realisations else missions
            if reals:
                exp_label(doc, "Réalisations")
                for r in reals: exp_item(doc, r)
            if resultats:
                exp_label(doc, "Résultats / impacts")
                for r in resultats: exp_item(doc, r)
            if environnement: exp_label(doc, "Environnement"); exp_body(doc, environnement)

    refs=cv_data.get("autres_references",[])
    if refs:
        section_title(doc, "AUTRES RÉFÉRENCES")
        for ref in refs:
            p=doc.add_paragraph(); sp(p, before=0, after=0)
            sf(p.add_run(f"{ref.get('entreprise','')} : "), size_pt=10, bold=True, color=RED)
            sf(p.add_run(ref.get("poste","")), size_pt=10, color=BLACK)

    projets=cv_data.get("projets_marquants",[])
    if projets:
        section_title(doc, "PROJETS MARQUANTS")
        for projet in projets:
            p=doc.add_paragraph(style='List Bullet'); sp(p, before=0, after=0)
            sf(p.add_run(projet), size_pt=10, color=BLACK)

    build_footer(doc)
    doc.save(output_path)
    return output_path
