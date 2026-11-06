"""Génère public/data/budget_1980.xlsx et public/data/budget_2026.xlsx.

ATTENTION : relancer ce script écrase les corrections faites à la main dans les
fichiers Excel. Il ne sert qu'à documenter la construction initiale des données.

Méthode (détaillée dans la feuille « Méthode » de chaque fichier) :
- Périmètre : ensemble des administrations publiques (APU = État + ODAC +
  collectivités locales + Sécurité sociale), comptabilité nationale Insee.
- 2026 : structure par fonction COFOG 2024 (Insee T_3301), totaux 2026 issus
  de la LFI/LFSS 2026 (dépenses 56,6 % du PIB, déficit 5,0 % du PIB).
- 1980 : totaux Insee (T_3201, base 2020) ; ventilation reconstituée par
  indexation des catégories COFOG 2024 sur des indicateurs disponibles en 1980
  (DREES, DEPP, Insee) ; le reste est réparti selon la structure COFOG 1995.
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUT = Path(__file__).resolve().parent.parent / "public" / "data"
FRANC = 6.55957

# --- Insee T_3301 : dépenses APU par fonction COFOG (Md€ courants) -----------
COFOG = {  # code: (1995, 2024)
    "01": (101.098, 181.103), "0102": (1.398, 4.881), "0104": (9.668, 19.881),
    "0107": (43.924, 58.870),
    "02": (29.851, 54.199), "0203": (None, 2.007), "0204": (None, 1.682),
    "03": (18.027, 52.113),
    "04": (77.799, 166.073), "0401": (15.374, 45.257), "0402": (None, 6.199),
    "0403": (None, 15.461), "0404": (None, 1.625), "0405": (None, 60.689),
    "0406": (None, 1.621), "0407": (None, 11.183), "0408": (None, 21.429),
    "0409": (None, 2.609),
    "05": (6.824, 30.285), "0505": (None, 1.760),
    "06": (12.848, 42.129), "0601": (3.458, 8.415),
    "07": (83.748, 261.156), "0705": (None, 6.708),
    "08": (13.198, 43.068),
    "09": (69.915, 148.640), "0904": (6.803, 12.308),
    "1001": (None, 83.899), "1002": (None, 391.954), "1003": (None, 40.608),
    "1004": (None, 66.711), "1005": (None, 48.521), "1006": (None, 21.147),
    "1007": (None, 36.461), "1009": (None, 3.728),
}
COFOG_TOTAL_2024 = 1671.794


def c24(*codes):
    return sum(COFOG[c][1] for c in codes)


# Catégories de dépenses : id, libellé, icône, description, valeur COFOG 2024
DEP = [
    ("retraites", "Retraites et vieillesse", "👵",
     "Pensions de retraite (base et complémentaires), réversion, minimum vieillesse, dépendance des personnes âgées (APA).",
     c24("1002", "1003")),
    ("sante", "Santé", "🏥",
     "Hôpitaux publics, remboursements de soins de ville et de médicaments, prévention, santé publique.",
     c24("07") - c24("0705")),
    ("famille", "Famille et enfance", "👨‍👩‍👧",
     "Allocations familiales, aides à la garde d'enfants, congés maternité, aide sociale à l'enfance.",
     c24("1004")),
    ("emploi", "Chômage et emploi", "💼",
     "Allocations chômage, France Travail, aides à l'embauche et à l'apprentissage, formation professionnelle.",
     c24("1005", "0401")),
    ("logement", "Logement", "🏠",
     "Aides personnelles au logement (APL, ALS, ALF) et aides à la construction de logements sociaux.",
     c24("1006", "0601")),
    ("pauvrete", "Pauvreté et exclusion", "🤝",
     "RSA, prime d'activité, hébergement d'urgence, action sociale des communes.",
     c24("1007", "1009")),
    ("handicap", "Invalidité et handicap", "♿",
     "Indemnités journalières maladie, pensions d'invalidité, AAH, établissements pour personnes handicapées.",
     c24("1001")),
    ("scolaire", "Écoles, collèges, lycées", "🏫",
     "Enseignement primaire et secondaire : enseignants, bâtiments scolaires, cantines, transport scolaire.",
     c24("09") - c24("0904")),
    ("superieur", "Enseignement supérieur et recherche", "🎓",
     "Universités, grandes écoles publiques, organismes de recherche (CNRS, CEA, CNES…), recherche appliquée.",
     c24("0904", "0104", "0204", "0408", "0505", "0705")),
    ("defense", "Défense", "🛡️",
     "Armées, équipements militaires, dissuasion nucléaire, opérations extérieures.",
     c24("02") - c24("0204", "0203")),
    ("securite", "Sécurité et justice", "⚖️",
     "Police, gendarmerie, pompiers, tribunaux, prisons.",
     c24("03")),
    ("dette", "Intérêts de la dette", "💸",
     "Intérêts versés aux créanciers de l'État, de la Sécurité sociale et des collectivités.",
     c24("0107")),
    ("transports", "Transports et infrastructures", "🚆",
     "Routes, ferroviaire, transports en commun, ports, aéroports, télécommunications.",
     c24("0405", "0406")),
    ("economie", "Agriculture, industrie et commerce", "🏭",
     "Soutien à l'agriculture, à l'industrie, au commerce, au tourisme et aux entreprises.",
     c24("0402", "0404", "0407", "0409")),
    ("environnement", "Environnement et énergie", "🌱",
     "Déchets, eaux usées, lutte contre la pollution, biodiversité, soutien aux énergies.",
     c24("05", "0403") - c24("0505")),
    ("culture", "Culture, sport et loisirs", "🎭",
     "Musées, patrimoine, audiovisuel public, équipements sportifs, associations.",
     c24("08")),
    ("territoires", "Aménagement et services locaux", "🏘️",
     "Urbanisme, voirie communale, eau potable, éclairage public, équipements collectifs.",
     c24("06") - c24("0601")),
    ("administration", "Administration et action extérieure", "🏛️",
     "Fonctionnement des ministères, du Parlement et des collectivités, services fiscaux, diplomatie, aide au développement, contribution à l'UE.",
     c24("01") - c24("0104", "0107") + c24("0203")),
]
assert abs(sum(d[4] for d in DEP) - COFOG_TOTAL_2024) < 0.01, sum(d[4] for d in DEP)

# --- Paramètres macro ------------------------------------------------------
PIB_1980 = 447.4          # Md€ courants, Insee T_3216 (PO / taux de PO)
PIB_2024 = 2935.2
PIB_2026 = 3060.0         # 2025 Insee 2991,1 Md€ + ~2,3 % de croissance nominale
DEP_APU_1980 = 214.9636   # Insee T_3201
REC_APU_1980 = 213.0273
DEP_APU_2026 = round(0.566 * PIB_2026, 1)        # LFI 2026 : 56,6 % du PIB
REC_APU_2026 = round(DEP_APU_2026 - 0.050 * PIB_2026, 1)  # déficit 5,0 % du PIB
INTERETS_1980 = 5.5151    # Insee T_3201, D41
INTERETS_2026 = 72.0      # ~2,35 % du PIB (charge d'intérêts APU en hausse, 65,7 en 2025)
DEFENSE_2026_DELTA = 57.15 - 47.2  # mission Défense LFI 2026 vs 2024 (hors pensions)
CPI_1980_TO_2026 = round(0.5035 * FRANC * 1.012, 3)  # 100 F 1980 = 50,35 € 2025 (Insee) ; inflation 2026 ~1,2 %

# --- 2026 : COFOG 2024 recalé sur le total 2026 ------------------------------
fixed26 = {"dette": INTERETS_2026,
           "defense": dict((d[0], d[4]) for d in DEP)["defense"] + DEFENSE_2026_DELTA}
rest24 = sum(d[4] for d in DEP if d[0] not in fixed26)
k26 = (DEP_APU_2026 - sum(fixed26.values())) / rest24
dep26 = {d[0]: round(fixed26.get(d[0], d[4] * k26), 2) for d in DEP}

# --- 1980 : indexation sur indicateurs ---------------------------------------
# DREES, prestations de protection sociale tous secteurs (Md€) : (1980, 2024)
DREES = {"vieillesse": (44.296, 426.699), "maladie": (29.010, 278.117),
         "inval_atmp": (4.793 + 4.579, 54.102 + 7.220), "famille": (14.274, 65.477),
         "emploi": (7.581, 51.301), "logement": (1.880, 16.059),
         "pauvrete": (0.443, 34.186)}


def idx(key, base):
    a, b = DREES[key]
    return base * a / b


v24 = dict((d[0], d[4]) for d in DEP)
DIE_1980, DIE_2024 = 29.4, 196.0  # DEPP, dépense intérieure d'éducation
educ_1980 = (c24("09")) * DIE_1980 / DIE_2024
dep80 = {
    "retraites": (idx("vieillesse", v24["retraites"]), "moyenne",
                  "COFOG 2024 indexé sur les prestations vieillesse-survie DREES (1980/2024)."),
    "sante": (idx("maladie", v24["sante"]), "moyenne",
              "COFOG 2024 indexé sur les prestations maladie DREES (1980/2024)."),
    "famille": (idx("famille", v24["famille"]), "moyenne",
                "COFOG 2024 indexé sur les prestations famille DREES (1980/2024)."),
    "emploi": (idx("emploi", v24["emploi"]), "faible",
               "COFOG 2024 indexé sur les prestations emploi DREES ; les politiques actives de l'emploi étaient plus réduites en 1980."),
    "logement": (idx("logement", COFOG["1006"][1]) + 0.005 * PIB_1980, "faible",
                 "Allocations logement DREES + aides à la pierre estimées à 0,5 % du PIB (aides à la construction importantes en 1980)."),
    "pauvrete": (idx("pauvrete", v24["pauvrete"]), "moyenne",
                 "COFOG 2024 indexé sur les prestations pauvreté DREES. Le RMI (ancêtre du RSA) n'existait pas avant 1988."),
    "handicap": (idx("inval_atmp", v24["handicap"]), "moyenne",
                 "COFOG 2024 indexé sur les prestations invalidité + accidents du travail DREES."),
    "scolaire": (educ_1980 * 0.91, "moyenne",
                 "Dépense intérieure d'éducation 1980 (DEPP, 29,4 Md€) × part publique COFOG 2024 ; 91 % pour le scolaire."),
    "superieur": (educ_1980 * 0.09 + 0.010 * PIB_1980, "faible",
                  "9 % de la dépense publique d'éducation + recherche publique estimée à 1,0 % du PIB."),
    "defense": (0.029 * PIB_1980, "moyenne",
                "Budget de la défense 1980 ≈ 3,1 % du PIB (loi de programmation 1977-1982 : 87,3 MdF), hors gendarmerie ≈ 2,9 %."),
    "dette": (INTERETS_1980, "élevée", "Intérêts versés par les APU (Insee, comptes nationaux, D41)."),
}
# Catégories restantes : structure 1995 (% du PIB), recalée pour retrouver le total Insee 1980
PROXY95 = {
    "securite": COFOG["03"][0],
    "transports": None, "economie": None, "environnement": None,
    "culture": COFOG["08"][0],
    "territoires": COFOG["06"][0] - COFOG["0601"][0],
    "administration": COFOG["01"][0] - COFOG["0104"][0] - COFOG["0107"][0],
}
# 1995 : détail des affaires économiques non publié dans notre extrait -> on applique
# la structure 2024 au total « 04 hors 0401 » de 1995 (hors R&D, comptée en recherche).
aff95 = COFOG["04"][0] - COFOG["0401"][0]
aff24 = {k: v24[k] for k in ("transports", "economie")}
aff24["environnement_energie"] = COFOG["0403"][1]
tot_aff24 = sum(aff24.values()) + COFOG["0408"][1]
PROXY95["transports"] = aff95 * aff24["transports"] / tot_aff24
PROXY95["economie"] = aff95 * aff24["economie"] / tot_aff24
PROXY95["environnement"] = COFOG["05"][0] + aff95 * aff24["environnement_energie"] / tot_aff24

residual = DEP_APU_1980 - sum(v[0] for v in dep80.values())
k80 = residual / sum(PROXY95.values())
for k, v in PROXY95.items():
    dep80[k] = (v * k80, "faible",
                "Pas de ventilation par fonction avant 1995 : part estimée d'après la structure COFOG 1995, "
                "recalée pour que le total égale les dépenses APU 1980 de l'Insee.")
assert abs(sum(v[0] for v in dep80.values()) - DEP_APU_1980) < 0.01

# --- Recettes ---------------------------------------------------------------
REC = [
    ("cotisations", "Cotisations sociales", "🧾",
     "Cotisations patronales et salariales finançant retraites, maladie, chômage, famille."),
    ("tva", "TVA et taxes sur la consommation", "🛒",
     "TVA, droits sur le tabac et l'alcool, taxes sur les assurances, droits de douane, autres taxes sur les produits."),
    ("revenu", "Impôt sur le revenu et CSG", "💶",
     "Impôt sur le revenu, CSG, CRDS, prélèvements sociaux sur les revenus du capital."),
    ("societes", "Impôt sur les sociétés", "🏢",
     "Impôt sur les bénéfices des entreprises et contributions associées."),
    ("locaux", "Impôts locaux et sur la production", "🏙️",
     "Taxe foncière, taxe d'habitation, taxe professionnelle / CFE / CVAE, taxe sur les salaires, versement mobilité."),
    ("energie", "Taxes sur l'énergie", "⛽",
     "Taxe sur les carburants (TICPE, ex-TIPP), accises sur l'électricité et le gaz."),
    ("patrimoine", "Patrimoine et successions", "🏰",
     "Droits de succession et de donation, droits de mutation immobilière, impôt sur la fortune."),
    ("autres", "Autres recettes", "📦",
     "Ventes de services publics, dividendes et revenus du patrimoine public, transferts, ajustements."),
]
# 2025 (Insee T_3217 / T_3201, Md€)
r25 = {
    "cotisations": 498.736,
    "tva": 208.82 + 3.58 + (128.06 - 30.46 - 8.27 - 3.40 - 0.97 - 20.21),
    "revenu": 156.58 + 9.27 + 14.41 + 5.31 + 103.63 - 7.17 - 2.14 + 0.37 + 11.59,
    "societes": 69.47 - 3.93 - 6.20 + 1.15 + 7.52,
    "locaux": 60.08 + 72.72 + 2.90,
    "energie": 30.46 + 8.27 + 3.40 + 0.97,
    "patrimoine": 20.21 + 2.71 + 21.68,
}
REC_APU_2025 = 1562.5197
r25["autres"] = REC_APU_2025 - sum(r25.values())
k_r26 = REC_APU_2026 / REC_APU_2025
rec26 = {k: round(v * k_r26, 2) for k, v in r25.items()}

# 1980 (Insee T_3201) : D2 67,3974 ; D5 33,0894 ; D91 0,9664 ; D61 84,9964 ; D995 -1,3778
D2_80, D5_80 = 67.3974, 33.0894
# Structure de D2 en 1995 (T_3217) : TVA 89,58 ; énergie 23,45 ; DMTO 4,83 ;
# autres taxes sur produits/importations 25,29 ; impôts sur la production 50,23
d2_95 = {"tva": 89.58 + 1.78 + (51.79 - 23.45 - 4.83), "energie": 23.45,
         "patrimoine": 4.83, "locaux": 14.42 + 35.81}
s95 = sum(d2_95.values())
IS_80 = 0.021 * PIB_1980                    # IS = 2,1 % du PIB en 1980 (Sénat)
TH_80 = 0.005 * PIB_1980                    # taxe d'habitation ≈ 0,5 % du PIB
AUTRES_D5_80 = 0.004 * PIB_1980             # vignette, redevance, divers
rec80 = {
    "cotisations": (84.9964, "élevée", "Cotisations sociales nettes (Insee, D61)."),
    "tva": (D2_80 * d2_95["tva"] / s95, "moyenne",
            "Impôts sur la production et les importations 1980 (Insee) ventilés selon leur structure de 1995."),
    "energie": (D2_80 * d2_95["energie"] / s95, "moyenne",
                "Impôts sur la production et les importations 1980 (Insee) ventilés selon leur structure de 1995."),
    "locaux": (D2_80 * d2_95["locaux"] / s95 + TH_80, "moyenne",
               "Impôts sur la production (structure 1995) + taxe d'habitation estimée à 0,5 % du PIB."),
    "patrimoine": (D2_80 * d2_95["patrimoine"] / s95 + 0.9664, "moyenne",
                   "Droits de mutation (structure 1995) + droits de succession et donation (Insee, D91)."),
    "societes": (IS_80, "moyenne", "IS = 2,1 % du PIB en 1980 (rapport du Sénat sur le PLF 1997)."),
    "revenu": (D5_80 - IS_80 - TH_80 - AUTRES_D5_80, "moyenne",
               "Impôts courants sur le revenu (Insee, D5) moins IS, taxe d'habitation et taxes diverses. La CSG n'existait pas (créée en 1991)."),
}
rec80["autres"] = (REC_APU_1980 - sum(v[0] for v in rec80.values()), "élevée",
                   "Recettes non fiscales (ventes, revenus de la propriété, transferts) et ajustements (Insee).")

# Indexation de chaque dépense pour le scénario « 2026 à la sauce 1980 » :
# Tout en € constants. francais = même montant par Français qu'en 1980 ; une clé de la feuille Indicateurs
# = même montant par bénéficiaire ; retraites = même effort par cotisant (la pension suit la richesse par
# cotisant et le nombre de cotisants par retraité, le surplus étant financé par les cotisations) ;
# reel = montant réel 2026 conservé (intérêts de la dette : contrainte héritée).
INDEX_SAUCE = {
    "retraites": "retraites", "famille": "moins20", "emploi": "chomeurs", "scolaire": "eleves",
    "superieur": "etudiants", "dette": "reel",
}

# --- Écriture Excel -----------------------------------------------------------
HEAD = Font(bold=True, color="FFFFFF")
FILL = PatternFill("solid", fgColor="1F3A5F")
INPUT = PatternFill("solid", fgColor="FFF4CE")


def sheet_table(ws, headers, rows, widths):
    ws.append(headers)
    for c in ws[1]:
        c.font, c.fill = HEAD, FILL
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for r in rows:
        ws.append(r)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C2"
    ws.row_dimensions[1].height = 32


def write(year, dep_rows, rec_rows, params, sources, methode, indicators):
    wb = Workbook()
    francs = year == 1980
    in_col = "Montant (milliards de francs courants)" if francs else "Montant (milliards d'€ courants)"
    for title, cats, values in (("Dépenses", DEP, dep_rows), ("Recettes", REC, rec_rows)):
        ws = wb.active if title == "Dépenses" else wb.create_sheet(title)
        ws.title = title
        heads = ["id", "Catégorie", "Icône", "Description", in_col]
        if francs:
            heads.append("Montant (milliards d'€ courants)")
        heads += ["Montant (milliards d'€ constants 2026)", "Fiabilité", "Source / méthode"]
        if year == 2026 and title == "Dépenses":
            heads.append("Indexation « 2026 à la sauce 1980 »")
        rows = []
        for i, cat in enumerate(cats, 2):
            val, fiab, note = values[cat[0]]
            amount = round(val * FRANC, 2) if francs else round(val, 2)
            row = [cat[0], cat[1], cat[2], cat[3], amount]
            if francs:
                row.append(f"=E{i}/Paramètres!$B$3")
                row.append(f"=F{i}*Paramètres!$B$4")
            else:
                row.append(f"=E{i}*Paramètres!$B$4")
            rows += [row + [fiab, note] + ([INDEX_SAUCE.get(cat[0], "francais")] if year == 2026 and title == "Dépenses" else [])]
        n = len(cats) + 1
        total = ["TOTAL", "Total", "", "", f"=SUM(E2:E{n})"]
        if francs:
            total += [f"=SUM(F2:F{n})", f"=SUM(G2:G{n})"]
        else:
            total += [f"=SUM(F2:F{n})"]
        rows.append(total + ["", ""])
        sheet_table(ws, heads, rows, [14, 34, 7, 60, 18] + ([18] if francs else []) + [20, 11, 80])
        for r in range(2, n + 1):
            ws.cell(r, 5).fill = INPUT
            for c in range(5, 8 if francs else 7):
                ws.cell(r, c).number_format = "#,##0.00"
        for c in ws[n + 1]:
            c.font = Font(bold=True)

    ws = wb.create_sheet("Paramètres")
    sheet_table(ws, ["clé", "Valeur", "Paramètre", "Unité", "Source"], params, [16, 14, 44, 22, 90])
    for r in range(2, len(params) + 2):
        ws.cell(r, 2).fill = INPUT

    ws = wb.create_sheet("Indicateurs")
    sheet_table(ws, ["clé", "Indicateur", "Valeur", "Unité", "Commentaire (affiché sur le site)", "Fiabilité", "Source"],
                indicators, [18, 44, 12, 18, 70, 11, 90])
    for r in range(2, len(indicators) + 2):
        ws.cell(r, 3).fill = INPUT

    ws = wb.create_sheet("Sources")
    sheet_table(ws, ["Donnée", "Source", "Lien"], sources, [36, 70, 90])
    ws = wb.create_sheet("Méthode")
    ws.column_dimensions["A"].width = 140
    for line in methode:
        ws.append([line])
        ws.cell(ws.max_row, 1).alignment = Alignment(wrap_text=True)
    OUT.mkdir(parents=True, exist_ok=True)
    wb.save(OUT / f"budget_{year}.xlsx")



# --- Indicateurs transversaux (feuille « Indicateurs ») ------------------------
# clé, indicateur, valeur, unité, commentaire affiché, fiabilité, source
IND_LABELS = {
    "agents_fp": ("Agents des trois fonctions publiques", "millions"),
    "agents_secu": ("Salariés des organismes de Sécurité sociale", "millions"),
    "masse_salariale": ("Rémunérations des agents publics", "% du PIB"),
    "retraites_nb": ("Retraités", "millions"),
    "cotisants_nb": ("Cotisants aux régimes de retraite", "millions"),
    "moins20": ("Jeunes de moins de 20 ans", "millions"),
    "eleves": ("Élèves (écoles, collèges, lycées)", "millions"),
    "etudiants": ("Étudiants", "millions"),
    "chomeurs": ("Chômeurs (au sens du BIT)", "millions"),
    "immigres": ("Immigrés (nés étrangers à l'étranger)", "millions"),
    "chomage_etrangers": ("Taux de chômage des étrangers", "% des actifs"),
}


def ind(key, val, comment, fiab, src):
    label, unit = IND_LABELS[key]
    return (key, label, val, unit, comment, fiab, src)


IND80 = [
    ind("agents_fp", 4.6, "Dont environ 0,45 M d'agents des PTT, sortis de la fonction publique en 1991 : hors PTT, 4,15 M.",
        "moyenne", "DGAFP-Insee, séries longues 1980 : FPE ≈ 2,83 M (PTT inclus), FPT 1,075 M, FPH 0,70 M (via iFRAP, 2021)"),
    ind("agents_secu", 0.18, "Estimation : le régime général comptait encore 166 000 ETP en 2004.",
        "faible", "Estimation à partir de l'annexe 10 du PLFSS (166 443 ETP en 2004) et des autres régimes"),
    ind("masse_salariale", 12.6, "Salaires et cotisations employeurs des trois fonctions publiques.",
        "moyenne", "Insee, comptes nationaux (via iFRAP, 2021)"),
    ind("retraites_nb", 9.0, "Estimation tous régimes (le seul régime général en comptait 4,99 M).",
        "faible", "Insee, Tableaux de l'économie française : régime général 4,989 M en 1980 ; population de 60 ans et plus ≈ 9,4 M"),
    ind("cotisants_nb", 22.1, "Estimation tous régimes : valeur choisie entre 21,6 et 22,7 M pour équilibrer le scénario « sauce 1980 » (le régime général seul : 13,35 M).",
        "faible", "Pas de total tous régimes publié pour 1980. Bornes : 21,6 M (estimation initiale), 22,7 M (Insee, emploi BIT, métropole). Valeur calée pour un solde proche de 0. Insee, TEF : régime général 13,354 M de cotisants pour 4,989 M de retraités"),
    ind("moins20", 16.8, "30,5 % de la population.", "moyenne", "Insee, pyramide des âges 1980"),
    ind("eleves", 12.9, "Collèges et lycées : 5,29 M ; écoles : ≈ 7,4 M ; outre-mer inclus.",
        "moyenne", "DEPP / Assemblée nationale (5 287 100 élèves du second degré en 1981-82) ; Insee"),
    ind("etudiants", 1.18, "Inscriptions dans l'enseignement supérieur en 1980-81.", "élevée", "MESR-SIES, séries longues"),
    ind("chomeurs", 1.28, "Taux de chômage d'environ 5,5 %.", "élevée", "Insee, séries longues du chômage BIT (1 276 000 en métropole)"),
    ind("immigres", 4.0, "7,4 % de la population (recensement de 1982).", "moyenne", "Insee, recensement 1982 (≈ 4,04 M d'immigrés)"),
    ind("chomage_etrangers", 14.0, "Recensement de 1982 : 220 000 chômeurs étrangers pour 1,35 M d'étrangers en emploi (taux national 1982 : 7,1 %).",
        "moyenne", "Insee, recensement 1982, cité par le Gisti (Les étrangers dans le monde du travail)"),
]

IND26 = [
    ind("agents_fp", 5.85, "État 2,58 M, territoriale 2,02 M, hospitalière 1,25 M (fin 2024, dernière donnée), contrats aidés exclus.",
        "élevée", "DGAFP, Stats rapides mai 2026 : effectifs de la fonction publique fin 2024"),
    ind("agents_secu", 0.16, "Régime général (~145 000 salariés) et autres régimes.",
        "moyenne", "Ucanss, carnet de l'emploi 2026"),
    ind("masse_salariale", 12.3, "Dernière donnée comparable (2019).",
        "moyenne", "Insee, comptes nationaux (via iFRAP, 2021)"),
    ind("retraites_nb", 17.4, "Retraités de droit direct ou de réversion, tous régimes (17,2 M fin 2023).",
        "moyenne", "Insee / DREES : 17,2 M de retraités et 30,4 M de cotisants en 2023, prolongé à 2026"),
    ind("cotisants_nb", 30.6, "30,4 M de cotisants pour 17,2 M de retraités en 2023.",
        "élevée", "Insee, cotisants, retraités et rapport démographique tous régimes"),
    ind("moins20", 15.9, "23 % de la population.", "élevée", "Insee, bilan démographique (23,5 % en 2023)"),
    ind("eleves", 11.9, "Premier degré ≈ 6,2 M, second degré ≈ 5,7 M.", "moyenne", "DEPP, Repères et références statistiques"),
    ind("etudiants", 3.01, "Inscriptions 2024-2025 (3 012 800).", "élevée", "MESR-SIES, note d'information 2025"),
    ind("chomeurs", 2.7, "Taux de chômage de 8,3 % au 2e trimestre 2026.", "élevée", "Insee, Informations rapides n° 192, août 2026"),
    ind("immigres", 7.7, "11,3 % de la population ; 2,6 M ont acquis la nationalité française (2024).",
        "élevée", "Insee Première n° 2076, Immigrés et étrangers en France en 2024"),
    ind("chomage_etrangers", 15.7, "Dernière donnée par nationalité (2019), contre 7,8 % pour les Français. Immigrés en 2024 : 13,8 % (hors UE : 15,8 %).",
        "moyenne", "Insee, enquête Emploi 2019, cité par l'Assemblée nationale (QE n° 1402) ; DSED 2024 pour les immigrés"),
]

COMMON_SOURCES = [
    ("Dépenses et recettes APU", "Insee, comptes nationaux base 2020, tableau 3.201",
     "https://www.insee.fr/fr/statistiques/8988845"),
    ("Impôts par catégorie / prélèvements obligatoires", "Insee, tableaux 3.216 et 3.217",
     "https://www.insee.fr/fr/statistiques/8988845"),
    ("Dépenses par fonction (COFOG)", "Insee, tableau 3.301 (1995-2024)",
     "https://www.insee.fr/fr/statistiques/8574707"),
    ("Prestations sociales par risque 1959-2024", "DREES, comptes de la protection sociale",
     "https://www.data.gouv.fr/datasets/les-comptes-de-la-protection-sociale/"),
]

write(1980,
      {k: dep80[k] for k in v24},
      rec80,
      [
          ("annee", 1980, "Année", "", ""),
          ("franc", FRANC, "Taux de conversion franc → euro", "F pour 1 €", "Taux officiel fixé le 31/12/1998"),
          ("coef", CPI_1980_TO_2026, "Coefficient € 1980 → € constants 2026", "multiplicateur",
           "Insee, convertisseur franc-euro (100 F de 1980 = 50,35 € de 2025) × inflation 2026 estimée à 1,2 %"),
          ("pib", round(PIB_1980 * FRANC, 1), "PIB", "milliards de F courants", "Insee, comptes nationaux base 2020 (447,4 Md€)"),
          ("pop", 55.0, "Population totale", "millions", "Insee : 53,7 M en métropole au 1/1/1980 + ~1,3 M outre-mer"),
          ("etrangers", 3.6, "Population étrangère", "millions", "Insee, recensement 1982 : 3,7 M d'étrangers (6,8 %)"),
          ("solde_etat", -30.3, "Solde budgétaire de l'État", "milliards de F", "Journal de l'année 1981 (déficit de l'État seul)"),
          ("dette_pib", 20.7, "Dette publique", "% du PIB", "Insee, dette au sens de Maastricht"),
      ],
      COMMON_SOURCES + [
          ("Dépense intérieure d'éducation 1980", "DEPP / Insee : 29,4 Md€, 6,5 % du PIB",
           "https://www.insee.fr/fr/statistiques/4238411"),
          ("Loi de finances pour 1980", "Loi n° 80-30 du 18 janvier 1980 ; rapports du Sénat 1979-1980",
           "https://www.senat.fr/rap/1979-1980/i1979_1980_0050_01.pdf"),
          ("Impôt sur les sociétés 1980", "Sénat, rapport PLF 1997 : IS = 2,1 % du PIB en 1980",
           "https://www.senat.fr/rap/l96-086-1-1/l96-086-1-144.html"),
      ],
      [
          "Budget 1980 de l'ensemble des administrations publiques (État, Sécurité sociale, collectivités locales).",
          "Les totaux (dépenses 215,0 Md€, recettes 213,0 Md€, déficit 1,9 Md€ soit 0,4 % du PIB) sont ceux de l'Insee.",
          "La ventilation par fonction n'existe pas avant 1995. Elle est reconstituée :",
          " - prestations sociales : catégories COFOG 2024 indexées sur l'évolution 1980→2024 des prestations par risque (DREES) ;",
          " - éducation : dépense intérieure d'éducation 1980 (DEPP) ; défense : budget 1980 en % du PIB ; intérêts : Insee ;",
          " - autres fonctions : structure COFOG 1995 recalée pour que le total égale les dépenses Insee 1980.",
          "Les montants sont saisis en milliards de francs courants (colonne jaune). Les autres colonnes sont calculées.",
          "Fiabilité : élevée = donnée publiée ; moyenne = reconstitution à partir d'indicateurs publiés ; faible = estimation.",
      ],
      IND80)

write(2026,
      {k: (dep26[k], "moyenne",
           "Structure COFOG 2024 (Insee) recalée sur le total 2026 de la LFI (56,6 % du PIB)."
           if k not in fixed26 else
           ("Intérêts APU 2026 estimés (~2,35 % du PIB ; 65,7 Md€ en 2025)." if k == "dette" else
            "COFOG 2024 + hausse de la mission Défense entre la LFI 2024 et la LFI 2026 (+9,95 Md€)."))
       for k in v24},
      {k: (rec26[k], "moyenne", "Structure 2025 (Insee, tableau 3.217) recalée sur les recettes 2026 prévues.")
       for k in r25},
      [
          ("annee", 2026, "Année", "", ""),
          ("franc", None, "Taux de conversion franc → euro", "", "Sans objet (montants en euros)"),
          ("coef", 1, "Coefficient € 2026 → € constants 2026", "multiplicateur", "Année de référence"),
          ("pib", PIB_2026, "PIB", "milliards d'€ courants", "Insee 2025 : 2 991,1 Md€ ; croissance nominale 2026 estimée à 2,3 %"),
          ("pop", 69.1, "Population totale", "millions", "Insee, bilan démographique : 69,1 M au 1/1/2026"),
          ("etrangers", 6.35, "Population étrangère", "millions", "Insee : 6,3 M d'étrangers en 2025 (9,1 %)"),
          ("solde_etat", -134.6, "Solde budgétaire de l'État", "milliards d'€", "LFI 2026 (loi n° 2026-103 du 19/02/2026)"),
          ("dette_pib", 118.0, "Dette publique", "% du PIB", "Prévision fin 2026 (115,7 % fin 2025 selon l'Insee ; 118,5 % selon la Cour des comptes)"),
          ("cas_education", 25.16, "Contribution de l'État aux pensions, part de l'Éducation nationale", "milliards d'€", "PLF 2026, mission Enseignement scolaire (Sénat)"),
          ("taux_cas", 82.28, "Taux de contribution de l'État aux pensions des civils", "%", "Taux en vigueur au 1er janvier 2026"),
          ("taux_cas_1980", 14, "Taux de cotisation retraite de référence (1980)", "%", "Hypothèse retenue ; à sourcer. La surcotisation est la part de la contribution au-dessus de ce taux"),
          ("ecole_inclusive", 4.74, "École inclusive, dont AESH (ajoutée au scolaire de la sauce 1980)", "milliards d'€", "PLF 2026 : 4,74 Md€, dont 3,16 Md€ pour les AESH (Sénat)"),
      ],
      COMMON_SOURCES + [
          ("Loi de finances 2026", "Loi n° 2026-103 du 19 février 2026",
           "https://www.legifrance.gouv.fr/jorf/id/JORFTEXT000053508155"),
          ("Loi de financement de la Sécurité sociale 2026", "LFSS 2026, déficit ROBSS 19,4 Md€", ""),
          ("Exécution 2026", "DGFiP, situation mensuelle budgétaire : -145,9 Md€ fin juillet 2026", ""),
      ],
      [
          "Budget 2026 de l'ensemble des administrations publiques (État, Sécurité sociale, collectivités locales).",
          f"Total des dépenses : 56,6 % du PIB ({DEP_APU_2026} Md€) ; déficit public visé : 5,0 % du PIB.",
          "Ventilation : structure COFOG 2024 de l'Insee, recalée sur ce total (intérêts et défense traités à part).",
          "Les comptes 2026 définitifs seront publiés par l'Insee en 2027 : ce sont donc des prévisions.",
          "Montants saisis en milliards d'euros courants (colonne jaune).",
      ],
      IND26)

# La cellule B3 du fichier 1980 est le taux franc/euro ; B4 le coefficient € constants.
print("dépenses 1980", round(DEP_APU_1980, 1), "2026", DEP_APU_2026)
print("recettes 1980", round(REC_APU_1980, 1), "2026", REC_APU_2026)
print("coef", CPI_1980_TO_2026)
for k in v24:
    print(f"{k:15s} {dep80[k][0]:8.2f} {dep26[k]:9.2f}")
for k in r25:
    print(f"{k:15s} {rec80[k][0]:8.2f} {rec26[k]:9.2f}")
