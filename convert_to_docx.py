from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re

doc = Document()

# ── Page margins
for section in doc.sections:
    section.top_margin    = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin   = Cm(3)
    section.right_margin  = Cm(2.5)

# ── Styles helpers
def set_font(run, name="Times New Roman", size=12, bold=False, italic=False, color=None):
    run.font.name  = name
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)

def add_heading(text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14 if level == 1 else 10)
    p.paragraph_format.space_after  = Pt(6)
    run = p.add_run(text)
    colors = {
        1: (0x1F, 0x49, 0x7D),   # dark blue
        2: (0x2E, 0x75, 0xB6),   # medium blue
        3: (0x70, 0xAD, 0x47),   # green
        4: (0xC0, 0x50, 0x20),   # dark orange
        5: (0x7B, 0x2C, 0x2C),   # dark red
    }
    sizes = {1: 16, 2: 14, 3: 13, 4: 12, 5: 11}
    set_font(run, size=sizes.get(level, 12), bold=True, color=colors.get(level, (0,0,0)))
    return p

def add_body(text, italic=False, bold=False, before=0, after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    p.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(text)
    set_font(run, italic=italic, bold=bold)
    return p

def add_ref(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(3)
    p.paragraph_format.left_indent  = Cm(0.5)
    p.paragraph_format.first_line_indent = Cm(-0.5)
    run = p.add_run(text)
    set_font(run, size=10, italic=True)
    return p

def add_table_method(rows_data):
    table = doc.add_table(rows=len(rows_data)+1, cols=2)
    table.style = 'Table Grid'
    # header row
    hdr = table.rows[0].cells
    for i, txt in enumerate(["Élément", "Détail"]):
        hdr[i].text = txt
        for para in hdr[i].paragraphs:
            for run in para.runs:
                set_font(run, bold=True, size=10, color=(0xFF, 0xFF, 0xFF))
        tc = hdr[i]._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '2E75B6')
        tcPr.append(shd)
    # data rows
    for ri, (label, value) in enumerate(rows_data):
        row = table.rows[ri+1].cells
        row[0].text = label
        row[1].text = value
        for para in row[0].paragraphs:
            for run in para.runs:
                set_font(run, bold=True, size=10)
        for para in row[1].paragraphs:
            for run in para.runs:
                set_font(run, size=10)
        if ri % 2 == 0:
            for cell in row:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'), 'DEEAF1')
                tcPr.append(shd)
    doc.add_paragraph()

# ═══════════════════════════════════════════════════════════
# PAGE DE TITRE
# ═══════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(40)
run = p.add_run("Rapport de Stage")
set_font(run, size=22, bold=True, color=(0x1F, 0x49, 0x7D))

p = doc.add_paragraph()
p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Master en Management de la Qualité et de la Sécurité des Soins")
set_font(run, size=14, bold=True, color=(0x2E, 0x75, 0xB6))

doc.add_paragraph()

p = doc.add_paragraph()
p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Institut Supérieur des Professions Infirmières et Techniques de Santé d'Oujda")
set_font(run, size=12, bold=True)

p = doc.add_paragraph()
p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Royaume du Maroc — Ministère de la Santé et de la Protection Sociale")
set_font(run, size=11)

doc.add_paragraph()

p = doc.add_paragraph()
p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Site de stage : CHU Mohammed VI d'Oujda — Bloc Opératoire Central (BOC)")
set_font(run, size=11, italic=True)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════
# Contenu
# ═══════════════════════════════════════════════════════════

# ────────────────────────────────────────────────────────────
# OBJECTIF 1
# ────────────────────────────────────────────────────────────
add_heading("OBJECTIF 1 : Gérer les activités de soins dans le respect des règles éthiques et humanistes", 1)

# ── Act 1
add_heading("Activité 1 — Identification de la mission et des objectifs d'une structure/unité de soins et la situer par rapport à l'établissement dont elle fait partie", 2)
add_heading("1.1 Introduction", 3)
add_body(
    "Toute organisation sanitaire s'inscrit dans une architecture institutionnelle hiérarchisée où chaque unité de soins "
    "remplit une mission spécifique en cohérence avec les orientations stratégiques de l'établissement. La compréhension "
    "de la mission et des objectifs d'une structure de soins constitue le point de départ indispensable de toute démarche "
    "de management de la qualité et de sécurité des soins. La mission d'une unité de soins se définit comme sa raison "
    "d'être au sein de l'établissement, exprimée en termes de services rendus à la population et d'objectifs organisationnels."
)
add_body(
    "Selon l'article 2 du Décret n° 2-06-656 du 13 avril 2007 relatif à l'organisation hospitalière au Maroc, les hôpitaux "
    "sont des établissements de santé ayant pour mission de dispenser, avec ou sans hébergement, des prestations de diagnostic, "
    "de soins et de services aux malades, blessés et parturientes, et d'assurer la permanence des soins ainsi que l'aide "
    "médicale urgente [1]. Le CHU Mohammed VI d'Oujda, créé par le décret n° 2-13-407 du 10 juillet 2013 et inauguré "
    "le 23 juillet 2014, assure quatre missions fondamentales : la prestation de soins spécialisés, la formation des "
    "professionnels de santé, la recherche et l'expertise, et le soutien aux politiques publiques de santé [2]."
)
add_body(
    "Selon Mintzberg (1982), la structure d'une organisation est « la somme totale des moyens employés pour diviser le travail "
    "en tâches distinctes et pour ensuite assurer la coordination nécessaire entre ces tâches » [3]. Dans ce cadre, le bloc "
    "opératoire central (BOC) occupe une position stratégique au sein du dispositif hospitalier du CHU. La loi n° 70-13 "
    "relative aux établissements et entreprises publics renforce cette logique en imposant une cohérence entre les missions "
    "de chaque unité et le Projet d'Établissement Hospitalier (PEH) [4]."
)
add_body("Références :", bold=True, after=2)
for ref in [
    "[1] Décret n° 2-06-656 du 13 avril 2007 relatif à l'organisation hospitalière. Bulletin Officiel du Royaume du Maroc, n° 5514.",
    "[2] Décret n° 2-13-407 du 10 juillet 2013 portant création du CHU Mohammed VI d'Oujda. Bulletin Officiel du Royaume du Maroc.",
    "[3] Mintzberg, H. (1982). Structure et dynamique des organisations. Éditions d'Organisation.",
    "[4] Loi n° 70-13 relative aux établissements et entreprises publics. (2015). Bulletin Officiel du Royaume du Maroc, n° 6340.",
]:
    add_ref(ref)

add_heading("1.2 Méthodologie", 3)
add_table_method([
    ("Lieu de l'étude",         "Bloc Opératoire Central (BOC) — CHU Mohammed VI d'Oujda"),
    ("Méthode de collecte",     "Investigation documentaire + Entretien semi-directif"),
    ("Outil de collecte",       "Guide d'entretien + Grille d'exploitation documentaire (PEH 2023–2025, Loi 70-13, organigramme)"),
    ("Population cible",        "Chef de service du BOC"),
])
add_body(
    "L'activité a été réalisée en mobilisant deux outils complémentaires. D'une part, une investigation documentaire a permis "
    "d'examiner les référentiels internes et les documents réglementaires liés à l'organisation du service, notamment le Projet "
    "d'Établissement Hospitalier (PEH 2023–2025) et la loi n° 70-13. D'autre part, un entretien semi-directif mené avec le chef "
    "du service du BOC a offert un éclairage direct sur le fonctionnement réel de l'unité, ses objectifs à court, moyen et long "
    "terme, ainsi que sa position au sein de l'organigramme du CHU. La combinaison des deux méthodes garantit une triangulation "
    "des données et une meilleure validité des conclusions."
)

# ── Act 2
add_heading("Activité 2 — Analyse des déterminants socioéconomiques et culturels des usagers de la structure/unité de soins", 2)
add_heading("2.1 Introduction", 3)
add_body(
    "La compréhension des déterminants socioéconomiques et culturels des usagers constitue un fondement éthique et opérationnel "
    "de la gestion des soins. Elle permet d'adapter la prise en charge aux besoins réels de la population, de réduire les "
    "inégalités de santé et d'améliorer l'expérience patient."
)
add_body(
    "Selon la définition de l'OMS, les déterminants de la santé sont les « facteurs personnels, sociaux, économiques et "
    "environnementaux qui déterminent l'état de santé des individus ou des populations » [1]. Le modèle de Dahlgren et "
    "Whitehead (1991) met en évidence quatre niveaux de déterminants : les caractéristiques individuelles (biologiques, "
    "comportementales, socioéconomiques), les milieux de vie (familial, scolaire, professionnel), les systèmes institutionnels "
    "(éducation, santé) et le contexte global (politique, économique, culturel) [2]."
)
add_body(
    "Dans le contexte marocain, l'Enquête Nationale sur la Population et la Santé Familiale (ENPSF 2018) révèle des disparités "
    "importantes d'accès aux soins selon le niveau de revenu, le lieu de résidence et le niveau d'instruction [3]. Au niveau "
    "de la région de l'Oriental, le profil des usagers du CHU Mohammed VI d'Oujda est marqué par une proportion significative "
    "de patients relevant du RAMED, un niveau d'instruction variable et une diversité culturelle influençant les comportements "
    "de recours aux soins, la communication thérapeutique et l'observance post-opératoire [4]."
)
add_body("Références :", bold=True, after=2)
for ref in [
    "[1] Organisation mondiale de la Santé. (2008). Combler le fossé en une génération. Commission sur les déterminants sociaux de la santé.",
    "[2] Dahlgren, G., & Whitehead, M. (1991). Policies and Strategies to Promote Social Equity in Health. Institute for Future Studies.",
    "[3] Ministère de la Santé du Maroc. (2018). Enquête Nationale sur la Population et la Santé Familiale (ENPSF 2018). DPRF.",
    "[4] Haut-Commissariat au Plan. (2022). Monographie régionale : Région de l'Oriental. HCP Maroc.",
]:
    add_ref(ref)

add_heading("2.2 Méthodologie", 3)
add_table_method([
    ("Lieu de l'étude",     "Bloc Opératoire Central (BOC) — CHU Mohammed VI d'Oujda"),
    ("Méthode de collecte", "Enquête par questionnaire"),
    ("Outil de collecte",   "Questionnaire auto-administré (données socioéconomiques et culturelles)"),
    ("Population cible",    "Patients programmés au BOC — échantillonnage par convenance"),
])
add_body(
    "L'analyse des déterminants socioéconomiques et culturels des usagers a été menée à l'aide d'un questionnaire distribué "
    "aux patients du BOC, sélectionnés selon un échantillonnage par convenance. Cet outil a permis de recueillir des données "
    "sur le profil social, économique et culturel des usagers (niveau d'instruction, couverture médicale, situation "
    "professionnelle, langue parlée, lieu de résidence) afin d'identifier les facteurs influençant leur accès aux soins "
    "et leur perception des services. Les données ont été traitées par statistiques descriptives et présentées sous forme "
    "de tableaux et graphiques."
)

# ── Act 3
add_heading("Activité 3 — Étude du circuit du patient/usager", 2)
add_heading("3.1 Introduction", 3)
add_body(
    "Le parcours de soins se décrit comme une série d'étapes que le patient traverse en fonction de l'évolution de sa "
    "maladie et de sa situation particulière, depuis son admission à l'hôpital jusqu'à sa sortie, en intégrant les "
    "différentes structures et professionnels impliqués dans sa prise en charge (Claverane & Pascal, 2004) [1]. Dans "
    "le contexte du bloc opératoire, ce parcours prend le nom de parcours périopératoire et revêt une complexité "
    "particulière en raison de l'implication de multiples acteurs."
)
add_body(
    "Le parcours périopératoire se décompose en trois phases interdépendantes : la phase préopératoire (de la consultation "
    "chirurgicale jusqu'à l'entrée dans le sas du bloc), la phase peropératoire (de l'induction anesthésique jusqu'à la "
    "fermeture du site opératoire, incluant la check-list sécurité OMS/HAS), et la phase postopératoire immédiate "
    "(de la sortie de salle jusqu'au transfert de la SSPI vers l'unité de soins) [2]."
)
add_body(
    "La méthode du patient-traceur, développée en France dans le cadre de la certification HAS et s'inspirant de la "
    "Joint Commission américaine, permet d'analyser de manière rétrospective le parcours réel d'un patient, d'évaluer "
    "les processus de soins et d'identifier les points d'amélioration. Elle constitue également une méthode reconnue "
    "de développement professionnel continu (DPC) (HAS, 2012) [3][4]."
)
add_body("Références :", bold=True, after=2)
for ref in [
    "[1] Claverane, R., & Pascal, J. (2004). Le parcours de soins du patient. Revue Hospitalière de France, 498, 12–18.",
    "[2] Société Française d'Anesthésie et de Réanimation. (2018). Recommandations sur le parcours périopératoire. SFAR.",
    "[3] Joint Commission International. (2011). Tracer Methodology: Tips and Strategies for Continuous Systems Improvement. JCI.",
    "[4] Haute Autorité de Santé. (2012). Le patient-traceur en établissement de santé : guide méthodologique. HAS.",
]:
    add_ref(ref)

add_heading("3.2 Méthodologie", 3)
add_table_method([
    ("Lieu de l'étude",     "Bloc Opératoire Central (BOC) — CHU Mohammed VI d'Oujda"),
    ("Méthode de collecte", "Observation directe (méthode du patient-traceur) + Entretien semi-directif"),
    ("Outil de collecte",   "Grille d'observation du parcours périopératoire + Guide d'entretien"),
    ("Population cible",    "Major du service BOC + patients programmés (avec consentement)"),
])
add_body(
    "L'étude du circuit patient a été réalisée à partir d'une observation directe suivant la méthode du patient-traceur, "
    "permettant de suivre le parcours réel d'un usager au sein de l'unité, de l'admission préopératoire jusqu'au transfert "
    "en SSPI. Un chronométrage des différentes étapes a permis d'identifier les temps d'attente et les éventuels goulets "
    "d'étranglement. Cette démarche a été complétée par un entretien semi-directif avec le major du service. Les données "
    "ont été synthétisées sous forme d'un diagramme de flux matérialisant les étapes successives, les acteurs impliqués "
    "et les interfaces critiques du parcours."
)

# ── Act 4
add_heading("Activité 4 — Analyse du modèle organisationnel adopté au sein d'une structure/unité de soins", 2)
add_heading("4.1 Introduction", 3)
add_body(
    "L'organisation peut être définie comme « une collectivité axée sur la poursuite de buts relativement spécifiques "
    "et manifestant une structure sociale hautement formalisée » (Scott, 1987) [1]. Dans le champ hospitalier, l'analyse "
    "du modèle organisationnel vise à comprendre comment le travail est divisé, coordonné et contrôlé, et comment cette "
    "organisation contribue à la qualité et à la sécurité des soins."
)
add_body(
    "Henry Mintzberg (1982) définit la structure comme « la somme totale des moyens employés pour diviser le travail en "
    "tâches distinctes et pour ensuite assurer la coordination nécessaire entre ces tâches » [2]. Il identifie six "
    "composantes fondamentales (sommet stratégique, centre opérationnel, ligne hiérarchique, technostructure, support "
    "logistique, idéologie) et six mécanismes de coordination (ajustement mutuel, supervision directe, standardisation "
    "des procédés, des résultats, des compétences et des normes)."
)
add_body(
    "Le bloc opératoire central s'apparente à une organisation professionnelle selon cette typologie : il repose sur "
    "l'expertise pointue de ses acteurs (chirurgiens, anesthésistes, IBODE, IADE), dont l'autonomie professionnelle "
    "leur permet de faire face à des situations complexes. Le mécanisme de coordination dominant est la standardisation "
    "des compétences, complétée par une supervision directe lors des procédures à haut risque [3][4]."
)
add_body("Références :", bold=True, after=2)
for ref in [
    "[1] Scott, W. R. (1987). Organizations: Rational, Natural, and Open Systems. Prentice-Hall.",
    "[2] Mintzberg, H. (1982). Structure et dynamique des organisations. Éditions d'Organisation.",
    "[3] Gittell, J. H. (2009). High Performance Healthcare. McGraw-Hill.",
    "[4] Loi n° 70-13 relative aux établissements et entreprises publics. (2015). Bulletin Officiel du Royaume du Maroc.",
]:
    add_ref(ref)

add_heading("4.2 Méthodologie", 3)
add_table_method([
    ("Lieu de l'étude",     "Bloc Opératoire Central (BOC) — CHU Mohammed VI d'Oujda"),
    ("Méthode de collecte", "Observation directe + Entretien semi-directif"),
    ("Outil de collecte",   "Grille d'observation organisationnelle + Guide d'entretien + Organigramme du service"),
    ("Population cible",    "Infirmier chef du BOC"),
])
add_body(
    "L'analyse du modèle organisationnel du BOC a été conduite par une observation directe des pratiques internes "
    "et un entretien semi-directif avec l'infirmier chef. Cette double approche a permis de comprendre la structuration "
    "des tâches, la répartition des responsabilités et les modes de coordination entre acteurs. Les données ont été "
    "analysées à la lumière du cadre de Mintzberg, permettant d'identifier la configuration dominante, les mécanismes "
    "de coordination en vigueur et les éventuels dysfonctionnements organisationnels."
)

# ── Act 5
add_heading("Activité 5 — Identification du processus, outils et supports de planification, mise en œuvre et d'évaluation des activités de soins", 2)
add_heading("5.1 Introduction", 3)
add_body(
    "La planification des activités de soins constitue une fonction managériale fondamentale visant à anticiper les besoins, "
    "à coordonner les ressources et à garantir la continuité et la qualité de la prise en charge. Elle s'inscrit dans le cycle "
    "PDCA (Plan-Do-Check-Act) développé par Deming, adopté comme référence dans les normes ISO 9001:2015 et largement "
    "utilisé dans les établissements de santé pour l'amélioration continue des processus [1]."
)
add_body(
    "Au bloc opératoire, la planification revêt une dimension opérationnelle critique : la programmation des interventions "
    "chirurgicales, la gestion des ressources humaines par poste, la disponibilité des dispositifs médicaux et la coordination "
    "avec les services amont et aval conditionnent directement la performance de l'unité. Les outils comprennent notamment "
    "le programme opératoire hebdomadaire, les fiches de poste, les protocoles de prise en charge et les outils de "
    "traçabilité per-opératoire (check-list OMS, fiches d'anesthésie, comptes rendus opératoires) [2][3]."
)
add_body(
    "La théorie de la fixation d'objectifs de Locke et Latham (1990) souligne que des objectifs spécifiques, mesurables "
    "et temporellement définis sont les principaux déterminants de la performance organisationnelle [4]."
)
add_body("Références :", bold=True, after=2)
for ref in [
    "[1] Organisation internationale de normalisation. (2015). ISO 9001:2015 — Systèmes de management de la qualité. ISO.",
    "[2] Société Française d'Anesthésie et de Réanimation. (2018). Recommandations sur l'organisation du bloc opératoire. SFAR.",
    "[3] Ministère de la Santé du Maroc. (2023). Projet d'Établissement Hospitalier 2023–2025 — CHU Mohammed VI d'Oujda. DHSA.",
    "[4] Locke, E. A., & Latham, G. P. (1990). A Theory of Goal Setting and Task Performance. Prentice Hall.",
]:
    add_ref(ref)

add_heading("5.2 Méthodologie", 3)
add_table_method([
    ("Lieu de l'étude",     "Bloc Opératoire Central (BOC) — CHU Mohammed VI d'Oujda"),
    ("Méthode de collecte", "Entretien semi-directif + Investigation documentaire"),
    ("Outil de collecte",   "Guide d'entretien + Grille d'exploitation documentaire (planning opératoire, PEH, protocoles)"),
    ("Population cible",    "Infirmier chef du BOC"),
])
add_body(
    "L'identification des processus, outils et supports a été réalisée par des entretiens semi-directifs avec l'infirmier "
    "chef du BOC, permettant de recueillir des données qualitatives sur les pratiques de planification, les outils utilisés "
    "et les mécanismes de suivi. Une investigation documentaire complémentaire a permis d'examiner les supports formels "
    "disponibles (programmes opératoires, protocoles, PEH) et de les comparer aux référentiels en vigueur."
)

# ── Act 6
add_heading("Activité 6 — Analyse de l'évolution des indicateurs", 2)
add_heading("6.1 Introduction", 3)
add_body(
    "Les indicateurs de performance permettent de quantifier l'atteinte des objectifs et de suivre l'évolution des "
    "activités dans le temps. Avedis Donabedian (1980) propose un cadre tripartite distinguant les indicateurs de "
    "structure (ressources disponibles), de processus (actions mises en œuvre) et de résultats (effets sur la santé "
    "des patients) [1]. Ce modèle demeure la référence conceptuelle dominante en évaluation sanitaire."
)
add_body(
    "Au niveau du bloc opératoire, les indicateurs clés incluent : le nombre total d'interventions par période, "
    "le taux d'occupation des salles (TOS), le taux de déprogrammation, la répartition des actes par spécialité, "
    "ainsi que les indicateurs de sécurité (taux d'infections du site opératoire, taux de réintervention). "
    "L'analyse longitudinale de ces indicateurs permet de détecter les tendances et d'orienter les décisions "
    "managériales vers les actions correctives appropriées [2][3]."
)
add_body("Références :", bold=True, after=2)
for ref in [
    "[1] Donabedian, A. (1980). Explorations in Quality Assessment and Monitoring: Vol. 1. Health Administration Press.",
    "[2] Ministère de la Santé du Maroc. (2023). Projet d'Établissement Hospitalier 2023–2025 — CHU Mohammed VI d'Oujda.",
    "[3] Organisation mondiale de la Santé. (2018). Handbook on health inequality monitoring. WHO.",
    "[4] Hurst, J., & Jee-Hughes, M. (2001). Performance Measurement in OECD Health Systems. OECD.",
]:
    add_ref(ref)

add_heading("6.2 Méthodologie", 3)
add_table_method([
    ("Lieu de l'étude",     "Hôpital des Spécialités — CHU Mohammed VI d'Oujda"),
    ("Méthode de collecte", "Analyse statistique des données d'activité"),
    ("Outil de collecte",   "Registres d'activité, tableaux de bord, rapports statistiques annuels"),
    ("Population cible",    "Technicien en statistiques sanitaires du CHU"),
])
add_body(
    "L'analyse de l'évolution des indicateurs a été réalisée à partir des données statistiques fournies par le "
    "technicien en statistiques sanitaires du CHU Mohammed VI d'Oujda. Ces données, issues de l'activité du bloc "
    "opératoire sur trois années consécutives, ont permis d'analyser les tendances quantitatives, d'identifier les "
    "variations significatives et les anomalies. Les données ont été traitées via Excel et représentées graphiquement "
    "(courbes d'évolution, histogrammes) pour faciliter la lecture et l'interprétation."
)

doc.add_page_break()

# ────────────────────────────────────────────────────────────
# OBJECTIF 2
# ────────────────────────────────────────────────────────────
add_heading("OBJECTIF 2 : Promouvoir un environnement de soin sécuritaire, confortable et humaniste", 1)

activities_obj2 = [
    {
        "title": "Activité 1 — Identification et analyse des procédures, méthodes et des outils d'évaluation de la qualité et sécurité des soins",
        "intro_paras": [
            "La qualité des soins est définie par l'OMS comme « le degré auquel les services de santé destinés aux individus et aux populations augmentent la probabilité d'atteindre les résultats de santé souhaités et sont conformes aux connaissances professionnelles actuelles » [1]. Dans le contexte du bloc opératoire, la qualité et la sécurité des soins revêtent une dimension critique en raison de la nature invasive des actes pratiqués.",
            "Les démarches d'évaluation reposent sur une pluralité d'outils. La check-list « Sécurité du patient au bloc opératoire », développée par l'OMS (2009) et adaptée par la HAS, constitue l'outil de référence international pour la prévention des événements indésirables per-opératoires [2]. Au Maroc, le Ministère de la Santé a engagé depuis 2012 une démarche nationale d'amélioration de la qualité hospitalière à travers le référentiel d'accréditation hospitalière [3]. Le CHU Mohammed VI d'Oujda dispose d'une cellule qualité institutionnelle assurant la coordination des démarches qualité, la gestion des événements indésirables et le suivi des indicateurs qualité et sécurité des soins [4].",
        ],
        "refs": [
            "[1] Organisation mondiale de la Santé. (2006). Quality of Care: A Process for Making Strategic Choices in Health Systems. WHO.",
            "[2] Organisation mondiale de la Santé. (2009). Liste de vérification de la sécurité chirurgicale : Manuel d'application. WHO.",
            "[3] Ministère de la Santé du Maroc. (2012). Référentiel national d'accréditation des hôpitaux. DHSA.",
            "[4] Institute of Medicine. (2001). Crossing the Quality Chasm. National Academy Press.",
        ],
        "method": [
            ("Lieu de l'étude",     "CHU Mohammed VI d'Oujda — Cellule Qualité + BOC"),
            ("Méthode de collecte", "Entretien semi-directif + Investigation documentaire"),
            ("Outil de collecte",   "Guide d'entretien + Grille d'analyse documentaire (procédures qualité, rapports d'audit)"),
            ("Population cible",    "Responsable qualité du CHU Mohammed VI d'Oujda"),
        ],
        "method_text": "La méthodologie adoptée s'est appuyée sur la réalisation d'un entretien semi-directif avec le responsable qualité du CHU Mohammed VI d'Oujda. Ces échanges ont permis de collecter des données précises sur les procédures, méthodes et outils utilisés pour l'évaluation de la qualité et de la sécurité des soins, tant au niveau institutionnel qu'à l'échelle du bloc opératoire. Une investigation documentaire complémentaire a permis d'examiner les procédures formalisées (check-listes, protocoles, référentiel d'accréditation, rapports d'audit interne). Cette double approche a permis d'obtenir une vision globale et contextualisée, identifiant les référentiels en vigueur, les pratiques concrètes et les défis de leur mise en œuvre.",
    },
    {
        "title": "Activité 2 — Analyse des systèmes de tri, de collecte, de traitement et d'élimination des déchets médicaux conformes aux normes réglementaires",
        "intro_paras": [
            "La gestion des déchets d'activités de soins à risques infectieux (DASRI) constitue un enjeu majeur de santé publique et d'environnement hospitalier. Le bloc opératoire est l'une des unités les plus productrices de déchets à risque en milieu hospitalier, du fait de la nature des actes pratiqués et de la quantité de dispositifs médicaux à usage unique utilisés [1].",
            "Le cadre réglementaire marocain est défini par la loi n° 28-00 relative à la gestion des déchets et à leur élimination (2006), qui impose des obligations strictes en matière de tri à la source, de conditionnement, de stockage, de transport et de traitement des DASRI [2]. Le Guide national de gestion des déchets des activités de soins du Ministère de la Santé précise les bonnes pratiques à chaque étape de la filière [3]. Au CHU Mohammed VI d'Oujda, la gestion des déchets est assurée en partie par une société de sous-traitance spécialisée, sous la supervision du service d'hygiène hospitalière [4].",
        ],
        "refs": [
            "[1] Organisation mondiale de la Santé. (2014). Safe management of wastes from health-care activities (2e éd.). WHO Press.",
            "[2] Loi n° 28-00 relative à la gestion des déchets et à leur élimination. (2006). Bulletin Officiel du Royaume du Maroc, n° 5480.",
            "[3] Ministère de la Santé du Maroc. (2020). Guide national de gestion des déchets des activités de soins. DELM.",
            "[4] Organisation mondiale de la Santé. (2022). Global report on infection prevention and control. WHO.",
        ],
        "method": [
            ("Lieu de l'étude",     "BOC + Service d'hygiène — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Entretiens semi-directifs + Observation directe"),
            ("Outil de collecte",   "Guides d'entretien + Grille d'observation (tri, stockage, élimination)"),
            ("Population cible",    "Technicien d'hygiène de santé environnement + Responsable société Alliance"),
        ],
        "method_text": "Pour analyser la gestion des déchets médicaux au sein de la structure hospitalière, une approche qualitative combinant deux outils de collecte a été adoptée. Des entretiens semi-directifs ont été menés avec le technicien d'hygiène de santé environnement et le responsable de la société de sous-traitance Alliance, afin de comprendre les pratiques en place et les responsabilités de chaque acteur. Parallèlement, une grille d'observation au niveau du bloc opératoire central a permis d'évaluer concrètement les modalités de tri, de stockage et d'élimination des déchets, en identifiant les éventuelles non-conformités aux normes réglementaires. Cette double approche a permis de croiser les données et d'établir un diagnostic synthétique et fiable.",
    },
    {
        "title": "Activité 3 — Analyse des mécanismes d'identification et de contrôle des risques liés aux activités de soins",
        "intro_paras": [
            "La gestion des risques en milieu de soins vise à identifier, évaluer, hiérarchiser et traiter les risques susceptibles d'affecter la sécurité des patients et du personnel. Le bloc opératoire est reconnu comme l'une des unités à plus haut risque au sein de l'hôpital : les risques y sont multiples (anesthésiques, chirurgicaux, infectieux, transfusionnels, confusion d'identité) [1].",
            "La méthode AMDEC (Healthcare Failure Mode and Effect Analysis — HFMEA) est l'outil de référence pour l'analyse prospective des risques liés aux processus chirurgicaux [2]. Elle permet de calculer un Indice de Priorité de Risque (IPR = Gravité × Occurrence × Détectabilité). Parallèlement, l'analyse rétrospective des événements indésirables par la Root Cause Analysis (RCA) constitue le mécanisme institutionnel de retour d'expérience [3]. La culture de sécurité, définie par Reason (2000) comme les valeurs et comportements collectifs engagés envers la sécurité, est le principal levier d'amélioration durable [4].",
        ],
        "refs": [
            "[1] Vincent, C. (2010). Patient Safety (2e éd.). Wiley-Blackwell.",
            "[2] DeRosier, J., et al. (2002). Using health care failure mode and effect analysis. Joint Commission Journal on Quality Improvement, 28(5), 248–267.",
            "[3] Organisation internationale de normalisation. (2018). ISO 31000:2018 — Management du risque. ISO.",
            "[4] Reason, J. (2000). Human error: Models and management. British Medical Journal, 320(7237), 768–770.",
        ],
        "method": [
            ("Lieu de l'étude",     "Bloc Opératoire Central (BOC) — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Entretien semi-directif + Analyse documentaire des EIAS déclarés"),
            ("Outil de collecte",   "Guide d'entretien + Registre des événements indésirables + Questionnaire HSOPSC"),
            ("Population cible",    "Responsable qualité/gestion des risques du CHU + Infirmier chef du BOC"),
        ],
        "method_text": "L'analyse des mécanismes de gestion des risques au BOC a été conduite en deux temps. Un entretien semi-directif avec le responsable qualité a permis d'identifier les procédures formalisées de déclaration et de traitement des événements indésirables, ainsi que les outils de cartographie des risques. La revue des fiches de déclaration sur 12 mois a permis d'analyser la nature, la fréquence et la gravité des incidents survenus. Les données ont été traitées par construction d'une matrice de criticité (probabilité × gravité) et d'un plan de maîtrise des risques avec actions priorisées.",
    },
    {
        "title": "Activité 4 — Analyse des protocoles et des mesures d'hygiène pour prévenir les infections nosocomiales",
        "intro_paras": [
            "Les infections associées aux soins (IAS) constituent l'une des complications les plus fréquentes et les plus évitables en milieu hospitalier. Les infections du site opératoire (ISO) représentent l'une des formes les plus graves d'IAS, entraînant une prolongation du séjour, une augmentation de la mortalité et des surcoûts considérables [1]. L'OMS (2022) estime que dans les pays à revenu intermédiaire, la prévalence des IAS peut atteindre 15,5 % des patients hospitalisés [2].",
            "La prévention des ISO repose sur un ensemble de précautions spécifiques : préparation préopératoire du patient (antisepsie cutanée, antibioprophylaxie), règles d'hygiène architecturale, tenue vestimentaire adaptée et respect strict de l'hygiène des mains selon les 5 moments de l'OMS [3]. L'hygiène des mains demeure l'intervention la plus efficace pour réduire la transmission croisée (Pittet et al., 2000) [4]. Au CHU Mohammed VI d'Oujda, le CLIN coordonne la politique de prévention et surveille les taux d'IAS.",
        ],
        "refs": [
            "[1] Société Française d'Hygiène Hospitalière. (2017). Prévention des infections du site opératoire. SF2H.",
            "[2] Organisation mondiale de la Santé. (2022). Global report on infection prevention and control. WHO.",
            "[3] Organisation mondiale de la Santé. (2016). Lignes directrices pour la prévention des ISO. WHO.",
            "[4] Pittet, D., et al. (2000). Effectiveness of a hospital-wide programme to improve hand hygiene. The Lancet, 356(9238), 1307–1312.",
        ],
        "method": [
            ("Lieu de l'étude",     "Bloc Opératoire Central (BOC) — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Observation directe + Entretien semi-directif + Analyse documentaire"),
            ("Outil de collecte",   "Grille d'audit hygiène (5 moments OMS) + Guide d'entretien + Protocoles d'hygiène"),
            ("Population cible",    "Infirmier hygiéniste / Référent CLIN du BOC + Personnel soignant"),
        ],
        "method_text": "L'analyse des pratiques d'hygiène au BOC a été menée par une combinaison d'observation directe et d'investigation documentaire. Un audit de l'hygiène des mains a été réalisé par observation des soignants lors des soins, en utilisant la grille standardisée des 5 moments de l'OMS, sur un minimum de 100 opportunités. Les protocoles d'hygiène disponibles (lavage chirurgical, habillage stérile, désinfection des salles) ont été analysés selon leur conformité aux recommandations nationales et internationales. Un entretien avec le référent hygiène a complété la collecte en fournissant les données de surveillance des taux d'ISO.",
    },
    {
        "title": "Activité 5 — Analyse des procédures d'accueil, d'orientation et de gestion des patients et leur famille",
        "intro_paras": [
            "L'accueil du patient constitue le premier contact entre l'usager et l'institution hospitalière et conditionne la qualité de la relation thérapeutique. Dans le contexte du bloc opératoire, l'accueil préopératoire — information sur l'intervention, vérification du consentement éclairé, communication empathique — a un impact direct sur le niveau d'anxiété du patient et ses résultats postopératoires [1].",
            "La loi n° 131-13 relative à l'exercice de la médecine au Maroc consacre le droit du patient à l'information et au consentement éclairé avant tout acte médical ou chirurgical [2]. La théorie des soins centrés sur la personne (Patient-Centered Care) de l'Institute of Medicine place les préférences et les besoins du patient au cœur de la démarche de soins [3]. Dans le contexte culturel marocain, la famille joue un rôle central de soutien, nécessitant des procédures spécifiques d'information et d'orientation en raison de l'accès restreint au bloc pour raisons d'hygiène [4].",
        ],
        "refs": [
            "[1] Kindler, C. H., et al. (2000). The visual analog scale allows effective measurement of preoperative anxiety. Anesthesia & Analgesia, 90(3), 706–712.",
            "[2] Loi n° 131-13 relative à l'exercice de la médecine. (2015). Bulletin Officiel du Royaume du Maroc, n° 6328.",
            "[3] Institute of Medicine. (2001). Crossing the Quality Chasm. National Academy Press.",
            "[4] Jenkinson, C., et al. (2002). Patients' experiences and satisfaction with health care. Quality and Safety in Health Care, 11(4), 335–339.",
        ],
        "method": [
            ("Lieu de l'étude",     "Bloc Opératoire Central (BOC) — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Observation directe + Questionnaire de satisfaction + Analyse documentaire"),
            ("Outil de collecte",   "Grille d'observation de l'accueil + Questionnaire patients + Procédures d'accueil"),
            ("Population cible",    "Patients admis au BOC + Personnel d'accueil (infirmiers de pré-anesthésie)"),
        ],
        "method_text": "L'analyse des procédures d'accueil au BOC a été conduite par une triple approche. Une observation directe de la prise en charge à l'arrivée au bloc (sas d'accueil préopératoire) a permis d'évaluer la qualité de l'accueil, la vérification d'identité, la communication préopératoire et la gestion des familles. Un questionnaire de satisfaction a été administré aux patients à leur retour en unité. Une analyse documentaire des procédures formalisées (loi 131-13, protocoles internes) a complété la démarche pour évaluer la conformité des pratiques aux exigences légales.",
    },
    {
        "title": "Activité 6 — Identification des insuffisances et de leurs mesures correctives au niveau de la gestion des patients/usagers de l'unité",
        "intro_paras": [
            "L'identification des insuffisances dans la gestion des patients est une démarche d'amélioration continue inscrite dans le cycle PDCA et la logique de l'apprentissage organisationnel. Selon Argyris et Schön (1978), une organisation apprenante est capable de détecter et de corriger ses erreurs en remettant en question ses routines de fonctionnement [1]. Au BOC, les insuffisances peuvent se manifester à différentes étapes : retards dans la préparation préopératoire, non-conformités à la check-list, défauts de traçabilité ou insuffisances dans la transmission inter-équipes [2].",
            "La méthode des 5 Pourquoi et le diagramme d'Ishikawa constituent les outils pratiques de référence pour l'analyse des causes profondes des dysfonctionnements [3]. La priorisation des actions correctives s'appuie sur une matrice impact/effort. La HAS (2021) recommande l'intégration systématique de ces analyses dans le fonctionnement des revues de morbi-mortalité (RMM) des blocs opératoires [4].",
        ],
        "refs": [
            "[1] Argyris, C., & Schön, D. A. (1978). Organizational Learning: A Theory of Action Perspective. Addison-Wesley.",
            "[2] Haute Autorité de Santé. (2019). Manuel de certification des établissements de santé V2020. HAS.",
            "[3] Ishikawa, K. (1986). Guide to Quality Control (2e éd.). Asian Productivity Organization.",
            "[4] Haute Autorité de Santé. (2021). Revue de mortalité et de morbidité (RMM) : guide méthodologique. HAS.",
        ],
        "method": [
            ("Lieu de l'étude",     "Bloc Opératoire Central (BOC) — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Analyse documentaire des EIAS + Focus group + Entretiens"),
            ("Outil de collecte",   "Registre des incidents + Guide de focus group + Diagramme d'Ishikawa"),
            ("Population cible",    "Équipe soignante du BOC (IBODE, IADE, infirmier chef)"),
        ],
        "method_text": "L'identification des insuffisances dans la gestion des patients au BOC a été conduite par une analyse systématique des fiches d'événements indésirables et des réclamations sur 12 mois. Un focus group réunissant des infirmiers de bloc, des IADE et l'infirmier chef a permis de recueillir les perceptions de l'équipe sur les dysfonctionnements récurrents. Pour chaque insuffisance majeure identifiée, un diagramme d'Ishikawa a été construit selon les 5M. Les insuffisances ont été hiérarchisées par matrice impact/effort, et un plan d'action SMART a été élaboré avec des responsables désignés et des indicateurs de suivi.",
    },
]

for idx, act in enumerate(activities_obj2, 1):
    add_heading(act["title"], 2)
    add_heading(f"{idx}.1 Introduction", 3)
    for para in act["intro_paras"]:
        add_body(para)
    add_body("Références :", bold=True, after=2)
    for ref in act["refs"]:
        add_ref(ref)
    add_heading(f"{idx}.2 Méthodologie", 3)
    add_table_method(act["method"])
    add_body(act["method_text"])

doc.add_page_break()

# ────────────────────────────────────────────────────────────
# OBJECTIF 3
# ────────────────────────────────────────────────────────────
add_heading("OBJECTIF 3 : Gérer les ressources matérielles, financières et médicaments au sein de la structure/unité de soins", 1)

activities_obj3 = [
    {
        "title": "Activité 1 — Étude des procédures d'identification des besoins de l'unité en ressources matérielles et financières",
        "intro_paras": [
            "La gestion des ressources matérielles et financières vise à optimiser l'utilisation des ressources disponibles pour répondre aux besoins de santé dans un contexte de contrainte budgétaire croissante. Le bloc opératoire est l'une des unités les plus consommatrices de ressources au sein du CHU : dispositifs médicaux à usage unique, implants chirurgicaux, produits anesthésiques et équipements biomédicaux représentent une part substantielle du budget [1].",
            "L'identification des besoins constitue la première étape du processus budgétaire. Elle repose sur l'analyse des consommations passées et la projection des activités futures. Dans les établissements publics marocains, ce processus est encadré par la loi organique des finances n° 130-13 (2015) et les circulaires de programmation budgétaire [2]. Au CHU Mohammed VI d'Oujda, la Direction des Affaires Administratives et Économiques (DAAE) joue un rôle central dans la consolidation et l'arbitrage des besoins exprimés par les unités [3]. La classification ABC des ressources selon leur valeur de consommation est un outil de gestion des stocks adapté au contexte hospitalier [4].",
        ],
        "refs": [
            "[1] Association Française des Gestionnaires de Blocs Opératoires. (2019). Guide de gestion des ressources au bloc opératoire. AFGBO.",
            "[2] Loi organique n° 130-13 relative à la loi de finances. (2015). Bulletin Officiel du Royaume du Maroc, n° 6340.",
            "[3] Ministère de la Santé du Maroc. (2019). Manuel de gestion des ressources matérielles dans les établissements de santé publics. DRH.",
            "[4] Hicks, D. A. (2007). Hospital Financial Management. Jossey-Bass.",
        ],
        "method": [
            ("Lieu de l'étude",     "BOC + DAAE — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Investigation documentaire + Entretien semi-directif"),
            ("Outil de collecte",   "Grille d'exploitation documentaire (bons de commande, états de consommation) + Guide d'entretien"),
            ("Population cible",    "Infirmier chef du BOC + Responsable des services économiques (DAAE)"),
        ],
        "method_text": "L'étude des procédures d'identification des besoins a été conduite par une double approche documentaire et terrain. L'investigation documentaire a porté sur les bons de commande, états de consommation et inventaires des deux dernières années afin d'établir le profil de consommation du BOC. Des entretiens ont été menés avec l'infirmier chef et le responsable de la DAAE pour comprendre le processus formel d'expression, de consolidation et d'arbitrage des besoins. Les données ont été traitées par une analyse ABC hiérarchisant les ressources selon leur valeur de consommation.",
    },
    {
        "title": "Activité 2 — Analyse des procédures de prévision, de commande, de réception, d'utilisation et de réforme des équipements et du matériel",
        "intro_paras": [
            "Le cycle de vie d'un équipement médical au BOC comprend plusieurs phases interdépendantes : programmation, acquisition, réception, mise en service, utilisation, maintenance et réforme. L'OMS (2011) estime que dans les pays à revenu intermédiaire, entre 40 et 70 % des équipements médicaux sont hors service ou sous-utilisés en raison d'une gestion déficiente du cycle de vie [1].",
            "Les procédures d'acquisition sont régies par le décret n° 2-12-349 portant réglementation des marchés publics (2013), qui impose des procédures d'appel d'offres pour les achats dépassant les seuils définis [2]. La maintenance préventive et corrective, encadrée par des contrats de maintenance biomédicale, est un facteur clé de la longévité des équipements et de la réduction des pannes imprévues au BOC [3]. La réforme des équipements obsolètes obéit à des procédures réglementaires spécifiques du Ministère de l'Économie et des Finances [4].",
        ],
        "refs": [
            "[1] Organisation mondiale de la Santé. (2011). Introduction to Medical Equipment Inventory Management. WHO.",
            "[2] Décret n° 2-12-349 portant réglementation des marchés publics. (2013). Bulletin Officiel du Royaume du Maroc, n° 6140.",
            "[3] Wang, B., et al. (2013). An estimate of patient safety risk reduction with improved medical equipment management. Biomedical Instrumentation & Technology, 47(2), 145–155.",
            "[4] Cohen, T. (2013). Introduction to Biomedical Equipment Technology (5e éd.). Pearson.",
        ],
        "method": [
            ("Lieu de l'étude",     "BOC + Service biomédical — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Investigation documentaire + Entretien semi-directif + Inventaire physique"),
            ("Outil de collecte",   "Registre des équipements + Contrats de maintenance + Guide d'entretien"),
            ("Population cible",    "Infirmier chef du BOC + Technicien biomédical + Responsable DAAE"),
        ],
        "method_text": "L'analyse du cycle de gestion des équipements a été conduite par une triple approche. Un inventaire physique des équipements du BOC a permis d'évaluer leur état, leur âge et leur conformité aux normes de sécurité. Une investigation documentaire a porté sur les procédures de commande, de réception, de maintenance et de réforme. Des entretiens avec l'infirmier chef, le technicien biomédical et le responsable DAAE ont permis de comprendre les pratiques réelles à chaque étape du cycle. Les résultats ont été synthétisés sous forme d'une matrice de cycle de vie par catégorie d'équipement, avec calcul du taux de disponibilité.",
    },
    {
        "title": "Activité 3 — Étude du cycle de gestion des médicaments au sein d'une structure/unité de soins",
        "intro_paras": [
            "La gestion du médicament à l'hôpital est un processus complexe et à haut risque. Au bloc opératoire, la gestion des produits anesthésiques, des antibiotiques de prophylaxie et des médicaments d'urgence est soumise à des règles strictes en raison de leur potentiel létal en cas d'erreur [1]. Le cycle de gestion comprend quatre étapes : sélection, approvisionnement, distribution et utilisation (MSH, 2012) [2].",
            "Au Maroc, la pharmacie hospitalière est régie par la loi n° 17-04 portant code du médicament et de la pharmacie, qui impose des règles strictes de dispensation, de traçabilité et de conservation des médicaments, notamment pour les stupéfiants et psychotropes [3]. L'OMS (2019) estime que les erreurs médicamenteuses coûtent au système de santé mondial environ 42 milliards USD par an [4]. La règle des 5B constitue la principale barrière de sécurité dans les unités non entièrement informatisées.",
        ],
        "refs": [
            "[1] Cohen, M. R. (Ed.). (2011). Medication Errors (2e éd.). American Pharmacists Association.",
            "[2] Management Sciences for Health. (2012). MDS-3: Managing Access to Medicines and Health Technologies. Kumarian Press.",
            "[3] Loi n° 17-04 portant code du médicament et de la pharmacie. (2006). Bulletin Officiel du Royaume du Maroc.",
            "[4] Organisation mondiale de la Santé. (2019). Medication without harm: WHO global patient safety challenge. WHO.",
        ],
        "method": [
            ("Lieu de l'étude",     "BOC + Pharmacie hospitalière — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Observation directe + Entretien semi-directif + Analyse documentaire"),
            ("Outil de collecte",   "Grille d'audit circuit médicament (règle des 5B) + Guide d'entretien + Registres de pharmacie"),
            ("Population cible",    "Pharmacien responsable + Infirmier chef du BOC + IADE"),
        ],
        "method_text": "L'étude du cycle de gestion des médicaments au BOC a été réalisée par une cartographie complète du circuit médicamenteux, de la commande jusqu'à l'administration en salle. Une observation directe des pratiques de préparation des drogues anesthésiques a été conduite avec une grille d'audit des 5B. Des entretiens avec le pharmacien et les IADE ont permis d'identifier les points de risque, notamment la gestion des stupéfiants et des médicaments à haut risque. Une analyse des fiches de déclaration d'erreurs médicamenteuses sur 12 mois a complété la démarche.",
    },
    {
        "title": "Activité 4 — Évaluation des indicateurs de suivi et d'évaluation de la gestion du médicament d'une unité de soin",
        "intro_paras": [
            "Le suivi et l'évaluation de la gestion du médicament reposent sur des indicateurs couvrant : la disponibilité (taux de rupture en médicaments essentiels), la sécurité (taux d'erreurs médicamenteuses), l'utilisation rationnelle (conformité de l'antibioprophylaxie) et l'économie (coût médicamenteux par acte chirurgical) [1].",
            "L'OMS (1999) a développé des indicateurs standardisés pour l'évaluation de la gestion des médicaments dans les structures de santé des pays à revenu intermédiaire [2]. L'ISMP (2018) a établi une liste de médicaments à haut risque nécessitant une surveillance renforcée, particulièrement pertinente au BOC [3]. L'analyse ABC/VEN (Vital, Essential, Non-essential) du portefeuille médicamenteux permet d'identifier les médicaments prioritaires en termes de gestion des approvisionnements [4].",
        ],
        "refs": [
            "[1] Organisation mondiale de la Santé. (1999). Indicators for monitoring national drug policies (2e éd.). WHO/DAP.",
            "[2] Malone, D. C., et al. (2004). Identification of serious drug-drug interactions. Journal of the American Pharmacists Association, 44(2), 142–151.",
            "[3] Institute for Safe Medication Practices. (2018). ISMP List of High-Alert Medications in Acute Care Settings. ISMP.",
            "[4] Centre Anti-Poison et de Pharmacovigilance du Maroc. (2022). Rapport annuel de pharmacovigilance. CAPM.",
        ],
        "method": [
            ("Lieu de l'étude",     "BOC + Pharmacie hospitalière — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Analyse des données de consommation + Entretien semi-directif"),
            ("Outil de collecte",   "Registres de consommation pharmaceutique + Tableau de bord pharmacie + Guide d'entretien"),
            ("Population cible",    "Pharmacien responsable du CHU"),
        ],
        "method_text": "L'évaluation des indicateurs de gestion du médicament au BOC a été réalisée à partir des données de consommation pharmaceutique des 12 derniers mois. Une analyse ABC/VEN du portefeuille médicamenteux du BOC a permis de classer les médicaments selon leur criticité. Les indicateurs calculés incluent : taux de disponibilité des médicaments essentiels, taux de rupture, coût moyen par acte chirurgical et conformité de l'antibioprophylaxie aux recommandations SFAR. Un entretien avec le pharmacien responsable a validé les données et permis d'identifier les marges d'amélioration.",
    },
]

for idx, act in enumerate(activities_obj3, 1):
    add_heading(act["title"], 2)
    add_heading(f"{idx}.1 Introduction", 3)
    for para in act["intro_paras"]:
        add_body(para)
    add_body("Références :", bold=True, after=2)
    for ref in act["refs"]:
        add_ref(ref)
    add_heading(f"{idx}.2 Méthodologie", 3)
    add_table_method(act["method"])
    add_body(act["method_text"])

doc.add_page_break()

# ────────────────────────────────────────────────────────────
# OBJECTIF 4
# ────────────────────────────────────────────────────────────
add_heading("OBJECTIF 4 : Gérer les ressources humaines au sein d'une structure/unité de soins", 1)

activities_obj4 = [
    {
        "title": "Activité 1 — Analyse de la démarche et des outils requis pour estimer les besoins en ressources humaines",
        "intro_paras": [
            "La planification des ressources humaines en santé (PRHS) est définie par l'OMS comme le processus permettant de déterminer les besoins futurs en personnels de santé, en quantité, qualité et compétences, pour répondre aux besoins de soins de la population [1]. Au bloc opératoire, l'estimation des besoins est complexe du fait de la diversité des profils requis (IBODE, IADE, aides-soignants, brancardiers), chacun avec des compétences spécifiques et non substituables.",
            "La méthode WISN (Workload Indicators of Staffing Need), développée par l'OMS (2010), est l'outil de référence pour estimer les besoins en effectifs à partir de l'analyse réelle des activités et des charges de travail [2]. Elle permet de calculer un ratio besoin/disponibilité par catégorie professionnelle. Au Maroc, le Ministère de la Santé a adopté la méthode WISN dans les établissements hospitaliers dans le cadre de la réforme des ressources humaines en santé [3][4].",
        ],
        "refs": [
            "[1] Organisation mondiale de la Santé. (2016). Health Workforce 2030: Towards a Global Strategy on Human Resources for Health. WHO.",
            "[2] Organisation mondiale de la Santé. (2010). Workload Indicators of Staffing Need (WISN): User's manual. WHO.",
            "[3] Ministère de la Santé du Maroc. (2021). Cadre réglementaire des ressources humaines de la santé. Direction des Ressources Humaines.",
            "[4] Buchan, J., et al. (2022). Sustain and Retain in 2022 and Beyond. International Centre on Nurse Migration.",
        ],
        "method": [
            ("Lieu de l'étude",     "Bloc Opératoire Central (BOC) — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Investigation documentaire + Entretien semi-directif"),
            ("Outil de collecte",   "Grille WISN + Données d'activité + Fiches de poste + Guide d'entretien"),
            ("Population cible",    "Infirmier chef du BOC + Direction des Ressources Humaines du CHU"),
        ],
        "method_text": "L'analyse a été conduite en plusieurs étapes. Une investigation documentaire a permis de recueillir les données d'activité du BOC (nombre et durée des interventions par spécialité sur 12 mois), les fiches de poste et les effectifs en poste. L'application de la méthode WISN a permis de calculer le temps de travail disponible par catégorie et de le comparer aux besoins théoriques. Des entretiens avec l'infirmier chef et la DRH ont fourni des données sur les contraintes réglementaires (statuts particuliers) et les tensions récurrentes en matière d'effectifs.",
    },
    {
        "title": "Activité 2 — Analyse de la charge de travail pour chacun des emplois définis pour une structure/unité de soins",
        "intro_paras": [
            "La charge de travail est un concept multidimensionnel englobant la charge quantitative, qualitative et émotionnelle. Au bloc opératoire, elle est directement liée au programme opératoire, à l'imprévisibilité des urgences et à la multiplicité des rôles exercés simultanément [1]. Une charge excessive est associée à une augmentation des erreurs, des incidents peropératoires et du burnout (Aiken et al., 2014) [2].",
            "Le Nursing Activities Score (NAS) et le Therapeutic Intervention Scoring System (TISS-28) constituent des références pour la mesure objective de la charge de travail par acte en unités techniques [3]. La théorie de la demande-contrôle-soutien de Karasek et Theorell (1990) apporte un éclairage psychosocial sur les déterminants de la charge perçue, soulignant l'importance du soutien hiérarchique et de l'autonomie [4].",
        ],
        "refs": [
            "[1] Association Française des Infirmiers de Bloc Opératoire. (2020). Référentiel de compétences IBODE. AFIBO.",
            "[2] Aiken, L. H., et al. (2014). Nurse staffing and education and hospital mortality in nine European countries. The Lancet, 383(9931), 1824–1830.",
            "[3] Miranda, D. R., et al. (1996). Simplified Therapeutic Intervention Scoring System: The TISS-28 items. Critical Care Medicine, 24(1), 64–73.",
            "[4] Karasek, R., & Theorell, T. (1990). Healthy Work: Stress, Productivity, and the Reconstruction of Working Life. Basic Books.",
        ],
        "method": [
            ("Lieu de l'étude",     "Bloc Opératoire Central (BOC) — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Auto-enregistrement des activités + Observation directe + Questionnaire"),
            ("Outil de collecte",   "Grille de relevé d'activités + Questionnaire de charge de travail perçue"),
            ("Population cible",    "Infirmiers de bloc (IBODE) + IADE + Aides-soignants de bloc"),
        ],
        "method_text": "L'analyse a été réalisée selon une double approche objective et subjective. Un relevé d'activités auto-administré a été conduit sur deux semaines, demandant à chaque professionnel de consigner le temps consacré à chaque type de tâche par poste. Une observation directe chronométrée a complété ce relevé sur un échantillon de journées opératoires. Un questionnaire évaluant la charge de travail perçue a été administré à l'ensemble du personnel. Les résultats ont été analysés par comparaison aux normes de référence et identification des pics de charge.",
    },
    {
        "title": "Activité 3 — Analyse de la démarche de répartition et d'affectation des ressources humaines",
        "intro_paras": [
            "La répartition et l'affectation des ressources humaines visent à aligner les compétences disponibles avec les besoins opérationnels. Au BOC, la constitution des équipes par salle doit tenir compte des spécialités chirurgicales, des niveaux de compétence requis et de la disponibilité en cas d'urgence [1].",
            "La théorie de la contingence stipule qu'il n'existe pas de mode d'organisation universel : la structure optimale s'adapte au contexte spécifique de l'unité [2]. Les plannings du BOC sont élaborés en tenant compte des contraintes réglementaires du statut général de la fonction publique marocaine (repos compensateurs, durée des gardes, congés) [3]. L'implication des soignants dans l'élaboration des plannings améliore la satisfaction au travail et réduit l'absentéisme (Bailyn et al., 2007) [4].",
        ],
        "refs": [
            "[1] Société Française d'Anesthésie et de Réanimation. (2018). Recommandations sur l'organisation du bloc opératoire. SFAR.",
            "[2] Lawrence, P. R., & Lorsch, J. W. (1967). Organization and Environment. Harvard Business School Press.",
            "[3] Décret n° 2-14-513 fixant le statut particulier du corps des infirmiers et techniciens de santé. (2015). Bulletin Officiel du Royaume du Maroc.",
            "[4] Bailyn, L., et al. (2007). Self-scheduling for hospital nurses. Journal of Nursing Management, 15(1), 72–77.",
        ],
        "method": [
            ("Lieu de l'étude",     "Bloc Opératoire Central (BOC) — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Analyse documentaire + Entretien semi-directif + Questionnaire"),
            ("Outil de collecte",   "Tableaux de service des 3 derniers mois + Guide d'entretien + Questionnaire de satisfaction"),
            ("Population cible",    "Infirmier chef du BOC + Personnel soignant du BOC"),
        ],
        "method_text": "L'analyse a été conduite par une revue des tableaux de service des trois derniers mois, évaluant la régularité de la répartition des postes par profil de compétences. Des entretiens avec l'infirmier chef ont permis de comprendre les critères formels et informels gouvernant les affectations. Un questionnaire de satisfaction vis-à-vis des plannings a été administré aux soignants. L'analyse des données d'absentéisme et d'heures supplémentaires a identifié les périodes de tension.",
    },
    {
        "title": "Activité 4 — Identification des mécanismes de collaboration et de leadership au sein de l'équipe de travail",
        "intro_paras": [
            "Le bloc opératoire est un environnement multidisciplinaire réunissant chirurgiens, anesthésistes, IBODE et IADE dans un espace confiné, sous pression temporelle et avec un enjeu vital pour le patient. La qualité de la collaboration interprofessionnelle est un déterminant majeur de la sécurité chirurgicale [1].",
            "Le leadership transformationnel (Bass & Riggio, 2006), caractérisé par l'inspiration et le développement des collaborateurs, est associé à de meilleurs résultats pour les patients [2]. La théorie de la coordination relationnelle de Gittell (2009), mesurant la fréquence, la ponctualité et la précision de la communication, est pertinente pour l'analyse des équipes de bloc [3]. L'OMS (2010) insiste sur la pratique collaborative interprofessionnelle (IPCP) comme levier de réduction des complications [4].",
        ],
        "refs": [
            "[1] Makary, M. A., et al. (2006). Operating room teamwork among physicians and nurses. Journal of the American College of Surgeons, 202(5), 746–752.",
            "[2] Bass, B. M., & Riggio, R. E. (2006). Transformational Leadership (2e éd.). Lawrence Erlbaum Associates.",
            "[3] Gittell, J. H. (2009). High Performance Healthcare. McGraw-Hill.",
            "[4] Organisation mondiale de la Santé. (2010). Framework for Action on Interprofessional Education and Collaborative Practice. WHO.",
        ],
        "method": [
            ("Lieu de l'étude",     "Bloc Opératoire Central (BOC) — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Observation directe + Questionnaire + Entretien semi-directif"),
            ("Outil de collecte",   "Grille d'observation + Questionnaire MLQ (Bass & Avolio) + Guide d'entretien"),
            ("Population cible",    "Infirmier chef + Équipe soignante du BOC (IBODE, IADE)"),
        ],
        "method_text": "L'identification des mécanismes de collaboration et de leadership a été réalisée par une combinaison d'outils. Une observation directe lors des interventions chirurgicales et des briefings pré-opératoires a permis d'évaluer la coordination interprofessionnelle. Le questionnaire MLQ de Bass & Avolio a été administré pour évaluer le profil de leadership de l'infirmier chef. Des entretiens semi-directifs avec des membres de l'équipe ont complété la démarche.",
    },
    {
        "title": "Activité 5 — Identification des approches utilisées dans la gestion des conflits interpersonnels",
        "intro_paras": [
            "Les conflits interpersonnels sont inhérents aux organisations de santé, notamment au BOC où les tensions entre chirurgiens et anesthésistes, ou au sein de l'équipe infirmière, constituent des sources fréquentes de conflits [1]. Thomas (1976) distingue cinq styles : évitement, compétition, accommodation, compromis et collaboration [2].",
            "Une gestion inefficace des conflits est associée à un turnover élevé, un absentéisme accru et une dégradation de la qualité des soins (Johansen, 2012) [3]. La communication non violente (CNV) de Rosenberg (2003), fondée sur l'observation factuelle et l'expression des besoins, constitue une approche adaptée au management des conflits en équipe soignante [4].",
        ],
        "refs": [
            "[1] Sexton, J. B., et al. (2000). Error, stress, and teamwork in medicine and aviation. British Medical Journal, 320(7237), 745–749.",
            "[2] Thomas, K. W. (1976). Conflict and conflict management. In Handbook of Industrial and Organizational Psychology. Rand McNally.",
            "[3] Johansen, M. L. (2012). Keeping the peace: Conflict management strategies for nurse managers. Nursing Management, 43(2), 50–54.",
            "[4] Rosenberg, M. B. (2003). Nonviolent Communication: A Language of Life. PuddleDancer Press.",
        ],
        "method": [
            ("Lieu de l'étude",     "Bloc Opératoire Central (BOC) — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Questionnaire + Entretiens confidentiels + Analyse documentaire"),
            ("Outil de collecte",   "Thomas-Kilmann Conflict Mode Instrument (TKI) + Guide d'entretien + Registre des plaintes internes"),
            ("Population cible",    "Personnel soignant du BOC (IBODE, IADE, infirmier chef)"),
        ],
        "method_text": "L'identification des approches de gestion des conflits a été conduite par le questionnaire TKI administré à l'ensemble du personnel, établissant la distribution des styles de gestion des conflits. Des entretiens confidentiels ont été menés avec des soignants ayant vécu des situations conflictuelles récentes. Une analyse des registres de plaintes internes sur deux ans a complété la démarche. Les résultats ont servi à cartographier les zones à risque et à proposer des mécanismes de médiation adaptés.",
    },
    {
        "title": "Activité 6 — Étude du système d'évaluation du rendement du personnel adopté au sein de la structure/unité de soins",
        "intro_paras": [
            "L'évaluation du rendement vise à mesurer la contribution individuelle aux objectifs de l'organisation, identifier les besoins de développement et alimenter les décisions de promotion et de formation. Dans la fonction publique marocaine, elle est régie par le statut général et les statuts particuliers de chaque corps de santé [1].",
            "La théorie de la fixation d'objectifs de Locke et Latham (1990) établit que des objectifs spécifiques et atteignables sont les déterminants majeurs de la performance individuelle [2]. L'approche à 360 degrés intégrant les évaluations du supérieur, des pairs et des patients est reconnue pour son équité (Bracken et al., 2001) [3]. La théorie de l'autodétermination de Deci et Ryan (1985) souligne que la motivation intrinsèque est plus durable que la motivation extrinsèque pour les professionnels de santé [4].",
        ],
        "refs": [
            "[1] Décret n° 2-11-681 portant statut général de la fonction publique. (2011). Bulletin Officiel du Royaume du Maroc.",
            "[2] Locke, E. A., & Latham, G. P. (1990). A Theory of Goal Setting and Task Performance. Prentice Hall.",
            "[3] Bracken, D. W., et al. (Eds.). (2001). The Handbook of Multisource Feedback. Jossey-Bass.",
            "[4] Deci, E. L., & Ryan, R. M. (1985). Intrinsic Motivation and Self-Determination in Human Behavior. Plenum Press.",
        ],
        "method": [
            ("Lieu de l'étude",     "BOC + Direction des Ressources Humaines — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Analyse documentaire + Entretien semi-directif + Questionnaire"),
            ("Outil de collecte",   "Formulaires d'évaluation annuelle + Guide d'entretien + Questionnaire de satisfaction"),
            ("Population cible",    "Infirmier chef du BOC + Direction des Ressources Humaines du CHU"),
        ],
        "method_text": "L'étude du système d'évaluation a été réalisée par une analyse documentaire des formulaires d'évaluation en vigueur et des fiches d'évaluation anonymisées des deux dernières années. Des entretiens ont été conduits avec la DRH et l'infirmier chef. Un questionnaire a évalué la satisfaction des soignants vis-à-vis du système et leur perception de son équité. Les résultats ont permis d'analyser la validité des critères d'évaluation et d'identifier des pistes d'amélioration.",
    },
    {
        "title": "Activité 7 — Élaboration d'un plan de formation continue en fonction des besoins pré-identifiés au niveau d'une unité/structure",
        "intro_paras": [
            "La formation continue au BOC est un impératif rendu d'autant plus pressant par l'évolution rapide des techniques chirurgicales et anesthésiques. L'approche par compétences (CBE), recommandée par l'OMS (2013), place le développement des compétences au cœur des dispositifs de formation continue [1].",
            "L'ingénierie de formation repose sur l'analyse des besoins (TNA), la définition d'objectifs SMART, le choix des modalités et l'évaluation selon les quatre niveaux de Kirkpatrick [2]. La simulation en santé est reconnue comme méthode de référence pour la formation des équipes de bloc opératoire (HAS, 2012) [3]. Au Maroc, la circulaire ministérielle impose l'élaboration annuelle d'un plan de formation basé sur les besoins identifiés [4].",
        ],
        "refs": [
            "[1] Organisation mondiale de la Santé. (2013). Transforming and Scaling Up Health Professionals' Education and Training. WHO.",
            "[2] Kirkpatrick, D. L., & Kirkpatrick, J. D. (2006). Evaluating Training Programs: The Four Levels (3e éd.). Berrett-Koehler.",
            "[3] Haute Autorité de Santé. (2012). Guide de bonnes pratiques en matière de simulation en santé. HAS.",
            "[4] Ministère de la Santé du Maroc. (2022). Circulaire relative à l'élaboration des plans de formation continue. DRH.",
        ],
        "method": [
            ("Lieu de l'étude",     "Bloc Opératoire Central (BOC) — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Questionnaire d'auto-évaluation + Analyse des évaluations de rendement + Entretiens"),
            ("Outil de collecte",   "Grille d'auto-évaluation des compétences (référentiel IBODE/IADE) + Fiches d'évaluation annuelle"),
            ("Population cible",    "Personnel soignant du BOC + Infirmier chef"),
        ],
        "method_text": "L'élaboration du plan de formation a suivi trois étapes. L'analyse des besoins a combiné un questionnaire d'auto-évaluation, l'analyse des évaluations annuelles et la revue des incidents liés à des déficits de compétences sur 12 mois. La priorisation a croisé les besoins avec les priorités stratégiques du CHU (PEH 2023–2025) et les contraintes budgétaires. La planification a abouti à un plan structuré comprenant : intitulé, objectifs pédagogiques, population cible, modalité, calendrier, budget estimé et indicateurs d'évaluation.",
    },
    {
        "title": "Activité 8 — Préparation et animation d'une activité de formation continue au sein d'une structure de soins",
        "intro_paras": [
            "L'animation d'une formation continue requiert la maîtrise des principes de l'andragogie. Malcolm Knowles (1980) identifie six postulats de l'apprenant adulte : besoin de comprendre le pourquoi, autonomie, valorisation de l'expérience, orientation vers des problèmes concrets et motivation intrinsèque [1]. Ces principes sont pertinents pour la formation des professionnels du BOC, experts expérimentés dont l'engagement est conditionné par l'ancrage clinique de la formation.",
            "Le modèle ADDIE constitue le cadre de référence pour la conception et l'évaluation des actions de formation [2]. Les méthodes actives (études de cas cliniques, simulation haute-fidélité) favorisent le transfert des apprentissages (Kolb, 1984) [3]. La taxonomie révisée de Bloom permet de définir des objectifs couvrant les dimensions cognitives, psychomotrices et affectives [4].",
        ],
        "refs": [
            "[1] Knowles, M. S. (1980). The Modern Practice of Adult Education (2e éd.). Cambridge Books.",
            "[2] Issenberg, S. B., et al. (2005). Features and uses of high-fidelity medical simulations that lead to effective learning. Medical Education, 39(1), 10–20.",
            "[3] Kolb, D. A. (1984). Experiential Learning: Experience as the Source of Learning and Development. Prentice Hall.",
            "[4] Anderson, L. W., & Krathwohl, D. R. (Eds.). (2001). A Taxonomy for Learning, Teaching, and Assessing. Longman.",
        ],
        "method": [
            ("Lieu de l'étude",     "Bloc Opératoire Central (BOC) — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Évaluation pré/post-formation + Observation + Questionnaire de satisfaction"),
            ("Outil de collecte",   "QCM pré/post-test + Grille d'évaluation des compétences + Questionnaire Kirkpatrick N1"),
            ("Population cible",    "Personnel soignant du BOC ciblé par l'action de formation"),
        ],
        "method_text": "La préparation a suivi le modèle ADDIE. La conception a inclus la définition des objectifs selon Bloom, la sélection du contenu fondé sur les données probantes, le choix des méthodes actives (exposé interactif, étude de cas clinique, démonstration gestuelle) et la conception des supports. La mise en œuvre a été conduite en salle de réunion du BOC. L'évaluation des acquis a été réalisée par QCM pré/post-formation (Kirkpatrick N2) et questionnaire de satisfaction (N1). Un suivi à J+30 par observation des pratiques a évalué le transfert des apprentissages (N3).",
    },
]

for idx, act in enumerate(activities_obj4, 1):
    add_heading(act["title"], 2)
    add_heading(f"{idx}.1 Introduction", 3)
    for para in act["intro_paras"]:
        add_body(para)
    add_body("Références :", bold=True, after=2)
    for ref in act["refs"]:
        add_ref(ref)
    add_heading(f"{idx}.2 Méthodologie", 3)
    add_table_method(act["method"])
    add_body(act["method_text"])

doc.add_page_break()

# ────────────────────────────────────────────────────────────
# OBJECTIF 5
# ────────────────────────────────────────────────────────────
add_heading("OBJECTIF 5 : Gérer le système d'information d'une structure/unité de soins", 1)

activities_obj5 = [
    {
        "title": "Activité 1 — Identification et analyse du système d'information déployé au niveau de la structure/unité de soins",
        "intro_paras": [
            "Le système d'information hospitalier (SIH) est l'ensemble des ressources permettant de recueillir, stocker, traiter et utiliser les informations nécessaires à la gestion d'un établissement de santé. Au BOC, il couvre la gestion du programme opératoire, la traçabilité des DMI, le dossier anesthésique et le suivi des indicateurs de performance [1].",
            "La stratégie nationale de e-santé du Maroc (2020–2025) prévoit le déploiement progressif d'un SIH intégré dans les CHU [2]. La théorie de l'acceptation des technologies (TAM) de Davis (1989) identifie l'utilité perçue et la facilité d'utilisation comme déterminants de l'adoption du SIH par les professionnels [3]. La loi n° 09-08 encadre la collecte et le traitement des données de santé numériques [4].",
        ],
        "refs": [
            "[1] Shortliffe, E. H., & Cimino, J. J. (Eds.). (2014). Biomedical Informatics (4e éd.). Springer.",
            "[2] Ministère de la Santé du Maroc. (2020). Stratégie nationale de e-santé 2020–2025. DPRF.",
            "[3] Davis, F. D. (1989). Perceived usefulness, perceived ease of use, and user acceptance of information technology. MIS Quarterly, 13(3), 319–340.",
            "[4] Loi n° 09-08 relative à la protection des personnes physiques à l'égard des traitements des données à caractère personnel. (2009). Bulletin Officiel du Royaume du Maroc.",
        ],
        "method": [
            ("Lieu de l'étude",     "BOC + Direction informatique — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Inventaire + Entretien semi-directif + Observation directe"),
            ("Outil de collecte",   "Grille d'inventaire du SIH + Guide d'entretien + Questionnaire utilisateurs"),
            ("Population cible",    "Responsable informatique du CHU + Infirmier chef du BOC"),
        ],
        "method_text": "L'identification et l'analyse du SIH du BOC ont été conduites par un inventaire des composantes du système en place (logiciels, matériel, modules disponibles, interfaces) réalisé avec le responsable informatique. Des entretiens ont évalué l'utilisation effective de chaque module et identifié les lacunes fonctionnelles. Une observation directe lors de la préparation du programme opératoire a complété la démarche. Les résultats ont été synthétisés sous forme d'une cartographie fonctionnelle du SIH.",
    },
    {
        "title": "Activité 2 — Analyse des rapports d'activités et des bilans de la structure/unité de soins (rédaction, emplacement, circuit, accès, traitement, exploitation)",
        "intro_paras": [
            "Les rapports d'activités constituent des instruments fondamentaux de la redevabilité et du pilotage. Ils documentent les réalisations du BOC en termes d'activité chirurgicale, de ressources mobilisées et de résultats obtenus. La théorie de la gestion axée sur les résultats (GAR) place les données probantes et les indicateurs au cœur du cycle de gestion des organisations [1].",
            "La qualité des rapports est conditionnée par la fiabilité du système de collecte à la source : enregistrement en temps réel dans les registres opératoires, codification correcte des actes et exhaustivité des données [2]. Les lacunes dans ce recueil conduisent à des conclusions erronées et des décisions managériales inappropriées (AbouZahr & Boerma, 2005) [3]. Les rapports suivent un circuit défini : rédaction par l'infirmier chef, validation par le chef de service, transmission à la direction et exploitation par le bureau des statistiques [4].",
        ],
        "refs": [
            "[1] Programme des Nations Unies pour le Développement. (2009). Handbook on Planning, Monitoring and Evaluating for Development Results. UNDP.",
            "[2] Ministère de la Santé du Maroc. (2018). Guide d'élaboration des rapports d'activité des établissements de santé. DPRF.",
            "[3] AbouZahr, C., & Boerma, T. (2005). Health information systems: The foundations of public health. Bulletin of the World Health Organization, 83(8), 578–583.",
            "[4] Patton, M. Q. (2008). Utilization-Focused Evaluation (4e éd.). SAGE Publications.",
        ],
        "method": [
            ("Lieu de l'étude",     "BOC + Bureau des statistiques sanitaires — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Investigation documentaire + Entretien semi-directif + Observation du circuit"),
            ("Outil de collecte",   "Grille d'analyse des rapports + Guide d'entretien + Diagramme de flux documentaire"),
            ("Population cible",    "Infirmier chef du BOC + Technicien en statistiques sanitaires du CHU"),
        ],
        "method_text": "L'analyse des rapports d'activité du BOC a été conduite par une revue systématique des rapports mensuels, trimestriels et annuels des deux dernières années. Une grille standardisée a évalué la qualité selon quatre critères : complétude, exactitude, ponctualité et exploitation effective. Des entretiens avec l'infirmier chef et le technicien en statistiques sanitaires ont permis de comprendre le circuit documentaire complet et d'identifier les points de fragilité. Un diagramme de flux du circuit a été construit.",
    },
    {
        "title": "Activité 3 — Identification du recueil et du traitement informatisé des données pertinentes pour la gestion de l'unité",
        "intro_paras": [
            "La disponibilité de données fiables et accessibles est un prérequis de la prise de décision fondée sur des données probantes (EIDM). Au BOC, les données pertinentes incluent : activité chirurgicale, ressources humaines, consommation pharmaceutique et indicateurs qualité/sécurité [1].",
            "La stratégie e-santé de l'OMS (2021) promeut la transformation numérique des systèmes de santé comme levier d'amélioration de la qualité [2]. La conformité à la loi n° 09-08 impose des exigences de qualité (exactitude, complétude), de sécurité et de traçabilité des données de santé numériques [3]. La normalisation des processus de saisie est indispensable pour automatiser la production des tableaux de bord [4].",
        ],
        "refs": [
            "[1] Lavis, J. N., et al. (2009). SUPPORT tools for evidence-informed health policymaking. Health Research Policy and Systems, 7(Suppl 1), I1.",
            "[2] Organisation mondiale de la Santé. (2021). Global Strategy on Digital Health 2020–2025. WHO.",
            "[3] Loi n° 09-08 relative à la protection des données à caractère personnel. (2009). Bulletin Officiel du Royaume du Maroc.",
            "[4] Oachs, P. K., & Watters, A. L. (Eds.). (2016). Health Information Management (5e éd.). AHIMA Press.",
        ],
        "method": [
            ("Lieu de l'étude",     "BOC + Service informatique — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Inventaire + Entretien semi-directif + Audit qualité des données"),
            ("Outil de collecte",   "Grille d'inventaire des données informatisées + Guide d'entretien + Grille d'audit"),
            ("Population cible",    "Responsable informatique + Infirmier chef + Technicien en statistiques sanitaires"),
        ],
        "method_text": "L'identification du recueil informatisé des données au BOC a été conduite par un inventaire des données effectivement saisies dans les systèmes disponibles. Un audit de qualité sur un échantillon de dossiers a évalué la complétude et l'exactitude des informations. Des entretiens avec le responsable informatique et le technicien en statistiques ont permis de comprendre les flux et interfaces entre systèmes. Les résultats ont été synthétisés sous forme d'une cartographie des flux d'information numérique, avec identification des lacunes.",
    },
    {
        "title": "Activité 4 — Identification du système d'organisation et de conservation des archives au sein d'une structure/unité de soins",
        "intro_paras": [
            "L'archivage des documents de santé (registres opératoires, comptes rendus d'anesthésie, consentements éclairés) constitue une obligation légale et éthique garantissant la continuité des soins, la redevabilité médicale et la protection juridique des patients et des professionnels [1].",
            "La circulaire du Ministère de la Santé impose des délais de conservation minimaux et des modalités spécifiques aux établissements publics [2]. La norme ISO 15489:2016 fournit un cadre pour la gestion des archives physiques et numériques [3]. La numérisation progressive des dossiers dans le cadre de la stratégie e-santé soulève des questions de cybersécurité, de confidentialité et de durabilité numérique [4].",
        ],
        "refs": [
            "[1] Ministère de la Santé du Maroc. (2015). Circulaire relative à la gestion et l'archivage des dossiers médicaux. DHSA.",
            "[2] Organisation internationale de normalisation. (2016). ISO 15489-1:2016 — Records management. ISO.",
            "[3] Shepherd, E., & Yeo, G. (2003). Managing Records: A Handbook of Principles and Practice. Facet Publishing.",
            "[4] Sprehe, J. T. (2005). The positive benefits of electronic records management. Government Information Quarterly, 22(2), 297–303.",
        ],
        "method": [
            ("Lieu de l'étude",     "BOC + Service des archives — CHU Mohammed VI d'Oujda"),
            ("Méthode de collecte", "Audit physique + Entretien semi-directif + Analyse documentaire"),
            ("Outil de collecte",   "Grille d'audit des archives (conformité réglementaire) + Guide d'entretien + Procédures d'archivage"),
            ("Population cible",    "Infirmier chef du BOC + Responsable des archives du CHU"),
        ],
        "method_text": "L'identification du système d'archivage du BOC a été conduite par un audit physique des espaces de stockage, évaluant les conditions de conservation, l'organisation du classement et la conformité aux délais réglementaires sur un échantillon de dossiers. Des entretiens avec l'infirmier chef et le responsable des archives ont permis de comprendre les procédures formelles et le niveau d'avancement de la numérisation. Les résultats ont été présentés sous forme d'un plan d'amélioration avec mesures prioritaires à court et moyen terme.",
    },
]

for idx, act in enumerate(activities_obj5, 1):
    add_heading(act["title"], 2)
    add_heading(f"{idx}.1 Introduction", 3)
    for para in act["intro_paras"]:
        add_body(para)
    add_body("Références :", bold=True, after=2)
    for ref in act["refs"]:
        add_ref(ref)
    add_heading(f"{idx}.2 Méthodologie", 3)
    add_table_method(act["method"])
    add_body(act["method_text"])

# ── Save
doc.save("/home/user/stage/rapport_stage_MQSS.docx")
print("Saved: rapport_stage_MQSS.docx")
