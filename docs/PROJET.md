# Projet Budget 1980 vs 2026

Mise à jour : 31 octobre 2026.

Voir aussi : [TASKS.md](TASKS.md) pour les questions et tâches en cours, [DEMANDES.md](DEMANDES.md) pour l'historique des demandes, [README.md](../README.md) pour lancer le site et modifier les données.

## Objectif

Site web en français, très simple d'usage, qui compare le budget public de la France en 1980 et en 2026, en euros constants 2026 par Français. Il sert de point de repère pour la présidentielle de 2027 : un budget 1980 quasi équilibré face au déficit actuel.

Public visé : grand public, sur ordinateur, tablette et téléphone. La qualité de l'interface est la priorité.

## Contenu du site

1. **Trois budgets** : 1980, 2026, et 1980 transposé sur la France de 2026.
2. **Où va l'argent ?** : 18 catégories de dépenses et 8 de recettes, avec détail et sources.
3. **Et si 2026 suivait les règles de 1980 ?** : le budget « sauce 1980 », poste par poste.
4. **À vous de jouer** : curseurs par poste, résultat en barres verticales, rapport PDF et Excel.
5. **Pour remettre en perspective** et **Indicateurs transversaux** : PIB, dette, agents publics, retraités, étrangers, etc.
6. **Sources et méthode**.

## Technique

- HTML, CSS et JavaScript sans framework ; SheetJS lit et écrit les fichiers Excel.
- Fichiers : `public/index.html`, `public/app.js`, `public/styles.css`, `public/data/budget_1980.xlsx`, `public/data/budget_2026.xlsx`.
- `server.py` : serveur de test local, port 8080, sans cache.
- `tools/build_data.py` : a généré les Excel. Le relancer écrase les corrections faites à la main.
- Dépôt : https://github.com/OS-SaaS-Apps/simu_spending_1980_2026 (public, licence MIT, auteur OS-SaaS-Apps).
- En ligne : https://os-saas-apps.github.io/simu_spending_1980_2026/. Chaque push sur `main` republie `public/` (`.github/workflows/pages.yml`).
- Couleurs : 1980 bleu, 2026 orange, transposé vert d'eau, sauce 1980 jaune, votre budget violet.

## Décisions de périmètre

- Périmètre : toutes les administrations publiques (État, agences, collectivités locales, Sécurité sociale).
- Unité : euros constants 2026 par personne de **nationalité française**, pas par habitant.
- Catégories fonctionnelles, jamais les ministères, dont les noms changent trop souvent.
- Priorité aux dépenses, recettes présentées aussi.
- 18 catégories de dépenses et 8 de recettes.
- Montants courants et constants dans les Excel ; 1980 saisi en francs.
- Excel hébergés sur le serveur et lus par URL (`DATA_URLS` dans `public/app.js`).
- Le déficit 2026 affiché (−153 Md€) est le déficit prévu sur l'année, toutes administrations confondues. Les −145,9 Md€ de fin juillet concernent l'État seul et ne s'extrapolent pas linéairement.

## Chiffres clés

| | 1980 (€ de 2026) | 2026 réel | Sauce 1980 |
|---|---|---|---|
| Dépenses | 718 Md€ environ | 1 732 Md€ | 1 092 Md€ |
| Recettes | 712 Md€ environ | 1 579 Md€ | 1 085 Md€ |
| Solde | −0,4 % du PIB | −153 Md€ (−5,0 % du PIB) | −6,5 Md€ (−0,2 % du PIB) |

Paramètres : coefficient € 1980 → € 2026 = 3,342 ; PIB 2026 estimé à 3 060 Md€ ; 55,0 M d'habitants et 3,6 M d'étrangers en 1980 ; 69,1 M et 6,35 M en 2026.

## Règles du budget « sauce 1980 »

Le budget de 1980, en euros constants, est appliqué à la France de 2026. La règle de chaque poste figure dans la colonne « Indexation » de la feuille Dépenses de `budget_2026.xlsx`.

- **Par défaut** : même montant par Français. C'est la population, et non le PIB, qui sert de référence, puisque tout est déjà en euros constants.
- **Moins de 20 ans, chômeurs, élèves, étudiants** : même montant par bénéficiaire.
- **Retraites** : même effort par cotisant. La pension de 1980 suit la richesse par cotisant et le nombre de cotisants par retraité. Cela garde 10,0 % du PIB, soit 307 Md€ et 1 470 € par mois au lieu de 2 120 €.
- **Intérêts de la dette** : montant réel de 2026 (72 Md€), car c'est une contrainte héritée et non un choix politique.
- **Recettes** : même montant par personne en emploi. Le surplus de pensions par rapport à cette règle est ajouté aux cotisations, si bien que les retraites ne creusent pas le solde.
- **Cotisants 1980** : fixés à 22,1 M, valeur calibrée pour approcher un solde nul. Ce n'est pas un chiffre mesuré. Bornes : 21,6 M (première estimation) et 22,7 M (emploi Insee BIT). À 22,7 M, le solde serait d'environ −22 Md€.

### École

- La sauce ajoute 4,74 Md€ pour l'école inclusive et les AESH, qui n'existaient pas en 1980.
- La sauce supprime l'éducation prioritaire : les ZEP sont créées en 1981, et les REP en découlent. C'est indiqué sous la ligne école.
- Dans le tableau sauce 1980, la barre 2026 est affichée hors 20,9 Md€ de surcotisation retraite, sur la base d'un taux de référence de 14 % (contre 82,28 % de taux CAS Pensions en 2026 pour un CAS éducation de 25,16 Md€). Le total du budget reste à 1 732 Md€.
- Résultat : 1 070 € contre 1 890 € par Français, soit 5 650 € contre 9 970 € par élève.

### Choix écartés ou retirés

- Tuiles de transparence sur les retraites et la santé : refusées.
- Contribution budgétaire nette des immigrés (OCDE) : retirée, car il n'y a pas de valeur 1980.
- Part du personnel hospitalier non soignant (34 %) : retirée, car il n'y a pas de valeur 1980 et le chiffre est contesté.
- Recalcul des aides sociales selon la nationalité : non fait, aucune source officielle ne ventile ces dépenses par nationalité.

## Le jeu « À vous de jouer »

- Les curseurs **partent de la sauce 1980**, pour les dépenses comme pour les recettes.
- Chaque curseur couvre au moins ±50 % et permet d'atteindre le montant réel de 2026. Des repères marquent la sauce 1980 et 2026 réel.
- Écarts calculés par rapport à la sauce 1980.
- Bandeau fixe : dépenses, recettes, déficit, en % du PIB aussi.
- Barres verticales : 2026 réel → votre budget, en Md€, avec le déficit hachuré.
- Boutons : « Repartir de la sauce 1980 » et « Appliquer 2026 réel ».
- Les choix sont gardés dans le navigateur.
- Le montant par retraité, élève ou chômeur se recalcule en direct.
- Rapport PDF par la fenêtre d'impression du navigateur (3 pages) : résumé, barres verticales, mes choix, détail complet.
- Rapport Excel `mon_budget.xlsx` : feuilles Résumé, Mes choix, Dépenses, Recettes, sans graphique.
- Dans le jeu, la barre 2026 de l'école est le montant réel, surcotisation comprise, pour que le total reste exact.

## Affichage

- Écarts du tableau sauce 1980 : sauce 1980 moins 2026 réel.
- Barres horizontales : unité explicite, par exemple « 4 890 € / Français » ou « 307 Md€ ».
- École (écoles, collèges, lycées) : en mode Par Français, les barres et l'écart sont en € / élève dans les trois tableaux, car le nombre d'élèves baisse quand la population augmente. La ligne a sa propre échelle. Les totaux et le mode Total ne changent pas.
- Sous 860 px, les lignes des tableaux s'empilent.

## Analyses de fond

### École : pourquoi environ +100 % par Français ?

- Site : 1 070 € sauce 1980 contre 1 890 € hors surcotisation.
- Le chiffre 2026 du site vient de la comptabilité nationale (11 730 €/élève). La DEPP donne environ 10 500 € pour 2024-25, tous financeurs, donc le site semble trop haut d'environ 1 270 €/élève.
- DEPP hors surcotisation : environ 8 700 €/élève, soit 1 650 €/Français, contre 1 395 € en 1980 par Français avec l'école inclusive, soit +18 %.
- Les salaires réels des enseignants ont baissé, ils n'expliquent donc pas l'écart.
- Enseignants, sur la base des chiffres 1980 fournis par l'utilisateur et non vérifiés : 700 000 → 890 000. Pour 1 000 élèves, 54 → 75 (+38 %). Pour 1 000 Français, 13,6 → 14,2 (+4 %). Élèves pour 100 Français : 25,1 → 19,0 (−24 %).
- À ratio de 1980 constant, il y a environ 244 000 enseignants de plus. Les sources n'en localisent qu'environ un quart : dédoublement en REP (16 000 ETP), classes plus petites en primaire (environ 45 000 au total, dédoublement compris). Le reste est sans doute dans le second degré et le temps partiel, sans chiffre 1980.

### Autres catégories

- Transports : montant 1980 plausible (environ 106 MdF en 1982, État et collectivités).
- Handicap : plausible, sans source directe pour 1980.
- Environnement : le montant 1980 (340 €/Français, 1,2 % du PIB) vient de la structure de 1995 et paraît trop élevé. Le soutien aux énergies renouvelables pèse 8,2 Md€ sur 45 Md€ en 2026.
- Santé : la part des frais de gestion a baissé (7,5 % en 1979 au régime général, environ 3 % en 2022), soit environ 130 € par Français aux deux dates. Administration hospitalière non comprise.
- Environ 24 % des dépenses de 1980 sont réparties avec la structure de 1995.

## Sources principales

- Insee, comptes nationaux base 2020 : tableaux 3.201, 3.216, 3.217 et COFOG 3.301 (1995-2024).
- DREES, comptes de la protection sociale 1959-2024.
- DEPP, dépense intérieure d'éducation et RERS.
- LFI 2026 (loi n° 2026-103 du 19 février 2026) et LFSS 2026.
- Cour des comptes, COR, DGAFP, Ucanss.

La liste complète et les liens figurent dans les feuilles « Sources » des Excel et en bas du site.

## Méthode de travail

- Réponses courtes et directes, sans `AskUserQuestion` : les questions se posent en texte.
- Lire directement les Excel plutôt que de bâtir des commandes compliquées.
