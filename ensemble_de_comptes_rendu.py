import os
import re
from datetime import datetime
from reportlab.platypus import BaseDocTemplate, Frame, Paragraph, Spacer, PageBreak, PageTemplate
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm

# 1) Configuration
dossier_txt = r"C:/Users/hmagallanesg/Documents/Ctes Rendu"
output_pdf  = r"C:/Users/hmagallanesg/Documents/ensemble_comptes_rendus.pdf"

# 2) Récupération et tri des .txt
def extraire_date(nom):
    m = re.search(r"(\d{4}-\d{2}-\d{2})", nom)
    return datetime.strptime(m.group(1), "%Y-%m-%d").date() if m else None

fichiers = []
for f in os.listdir(dossier_txt):
    if f.lower().endswith(".txt"):
        d = extraire_date(f)
        if d:
            fichiers.append((f, d))
if not fichiers:
    raise RuntimeError("Aucun fichier .txt daté trouvé dans le dossier spécifié.")
fichiers.sort(key=lambda x: x[1])

date_min   = fichiers[0][1].isoformat()
date_max   = fichiers[-1][1].isoformat()
aujourdhui = datetime.today().date().isoformat()

# 3) Styles
doc_styles   = getSampleStyleSheet()
style_title  = doc_styles['Title']
style_normal = doc_styles['Normal']
style_heading= ParagraphStyle(
    name='HeadingReport',
    parent=doc_styles['Heading1'],
    fontSize=14,
    leading=16,
    spaceAfter=12,
    outlineLevel=0
)

# 4) Table des matières
toc = TableOfContents()
toc.levelStyles = [
    ParagraphStyle(
        name='TOCEntry',
        fontName='Times-Roman',
        fontSize=12,
        leftIndent=20,
        firstLineIndent=-20,
        spaceBefore=5,
        leading=12
    )
]

# 5) Template avec marges et notification TOC
class MyDocTemplate(BaseDocTemplate):
    def __init__(self, filename, **kwargs):
        super().__init__(filename, **kwargs)
        frame = Frame(
            self.leftMargin,
            self.bottomMargin,
            self.width,
            self.height,
            id='normal'
        )
        template = PageTemplate(id='all', frames=[frame])
        self.addPageTemplates([template])

    def afterFlowable(self, flowable):
        if isinstance(flowable, Paragraph) and getattr(flowable.style, 'outlineLevel', None) == 0:
            text = flowable.getPlainText()
            self.notify('TOCEntry', (0, text, self.page))

# 6) Construction du document
doc = MyDocTemplate(
    output_pdf,
    pagesize=A4,
    leftMargin=2*cm,
    rightMargin=2*cm,
    topMargin=2*cm,
    bottomMargin=2*cm
)

story = []

# 6.1) Première page
titre = f"Ensemble des comptes rendus du {date_min} au {date_max}, réalisé le {aujourdhui}."
story.append(Paragraph(titre, style_title))
story.append(PageBreak())

# 6.2) Index
story.append(Paragraph("Index des comptes rendus", style_title))
story.append(Spacer(1, 0.5 * cm))
story.append(toc)
story.append(PageBreak())

# 6.3) Comptes rendus détaillés
for nom_f, _ in fichiers:
    chemin = os.path.join(dossier_txt, nom_f)
    # Titre de section pour le TOC
    story.append(Paragraph(nom_f, style_heading))
    # Lecture du texte brut
    with open(chemin, 'r', encoding='utf-8') as fd:
        raw = fd.read()
    # Supprimer balises HTML (<br/>, etc.)
    cleaned = re.sub(r"<[^>]+>", "", raw)
    # Fractionnement en lignes pour wrapping automatique
    for ligne in cleaned.splitlines():
        if ligne.strip():
            story.append(Paragraph(ligne.strip(), style_normal))
    story.append(PageBreak())

# 7) Génération en double passe pour la TOC
doc.multiBuild(story)
print(f"PDF généré : {output_pdf}")
