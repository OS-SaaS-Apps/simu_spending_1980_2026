# Budget 1980 vs 2026

Site web qui compare le budget public de la France (État, Sécurité sociale, collectivités locales) en 1980 et en 2026, en euros constants 2026 par Français. Il transpose aussi le budget de 1980 sur la France de 2026 (« sauce 1980 ») et propose un jeu pour composer son propre budget, avec un rapport PDF ou Excel.

Point de repère pour la présidentielle de 2027 : un budget 1980 quasi équilibré face au déficit actuel.

**En ligne : https://os-saas-apps.github.io/simu_spending_1980_2026/**

## Lancer en local

Il faut Python 3.

```
python server.py        # puis ouvrir http://localhost:8080
```

Le site est statique : n'importe quel serveur web qui sert le dossier `public/` suffit. Chaque push sur `main` le publie sur GitHub Pages (`.github/workflows/pages.yml`). Ouvrir `index.html` directement depuis le disque ne marche pas, car le navigateur bloque la lecture des Excel.

## Arborescence

```
public/              site (servi tel quel)
  index.html
  app.js             calculs et affichage
  styles.css
  data/              budget_1980.xlsx, budget_2026.xlsx
tools/
  build_data.py      génération initiale des Excel
server.py            serveur de test local
docs/                documentation du projet
```

## Modifier les données

Ouvrir `public/data/budget_1980.xlsx` ou `public/data/budget_2026.xlsx` :

- **Dépenses** / **Recettes** : seule la colonne jaune (montant) est à saisir. 1980 se saisit en milliards de francs, 2026 en milliards d'euros. Les autres colonnes sont calculées.
- **Paramètres** : coefficient d'inflation, PIB, population totale, population étrangère, dette.
- **Sources** et **Méthode** : affichées telles quelles sur le site.

Enregistrer, puis recharger la page : les changements apparaissent immédiatement.

L'adresse des fichiers Excel se règle dans `DATA_URLS`, en haut de `public/app.js`.

`tools/build_data.py` a généré les fichiers à l'origine (`pip install -r requirements.txt`, puis `python tools/build_data.py`). Le relancer écrase les corrections faites à la main.

## Documentation

- [docs/PROJET.md](docs/PROJET.md) : objectif, décisions, règles de calcul, analyses.
- [docs/TASKS.md](docs/TASKS.md) : questions et tâches en cours.
- [docs/DEMANDES.md](docs/DEMANDES.md) : historique des demandes.

## Sources

Insee (comptes nationaux), DREES, DEPP, lois de finances 2026, Cour des comptes, COR. La liste complète figure dans les feuilles « Sources » des Excel et en bas du site. Les données brutes sont publiques (Licence Ouverte).

## Licence

[MIT](LICENSE) © 2026 OS-SaaS-Apps.
