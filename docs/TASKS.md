# Questions et tâches en cours

Mise à jour : 31 octobre 2026. Contexte dans [PROJET.md](PROJET.md).

## Questions à trancher

- [ ] **École, chiffre 2026** : aligner la ligne 2026 sur la DEPP (environ 10 500 €/élève) au lieu de la comptabilité nationale (11 730 €/élève) ?
- [ ] **École, règle sauce 1980** : garder le même montant par élève, ou passer au même montant par Français ? Le passage ajouterait environ 20 Md€ et mettrait le solde sauce 1980 à environ −27 Md€.
- [ ] **École, cohérence des tableaux** : « Où va l'argent ? » et le jeu affichent 11 730 €/élève en 2026, surcotisation comprise, et le tableau sauce 1980 affiche 9 970 €/élève hors surcotisation. Faut-il unifier ?
- [ ] **Enseignement supérieur et autres postes** : passer aussi en € / étudiant, ou selon le bénéficiaire pour d'autres postes ?
- [ ] **Retraites, base de calcul** : garder la richesse par cotisant pour les retraites seules, ou harmoniser avec les autres postes (euros de 1980 partout, ou richesse par cotisant partout) ?
- [ ] **Jeu, intérêts de la dette** : bloquer ce curseur, puisque ce n'est pas un choix politique ?
- [ ] **Jeu, rapport** : permettre de recharger ses choix sur le site à partir d'un fichier ? Aujourd'hui, le rapport se lit mais ne se recharge pas.

## Données à vérifier

- [ ] Taux de cotisation retraite de 14 % en 1980 dans le public : chiffre donné par l'utilisateur, sans source.
- [ ] Enseignants 1980 (700 000) et non-enseignants (230 000 à 250 000) : chiffres donnés par l'utilisateur, à vérifier et à ramener en temps plein.
- [ ] Rétropolation DEPP de 1980 : détail des cotisations imputées non contrôlé.
- [ ] Environnement 1980 : le montant vient de la structure de 1995 et paraît trop élevé.
- [ ] Dépenses de 1980 réparties avec la structure de 1995 (environ 24 % du total) : chercher des sources directes.
- [ ] Nombre de cotisants 1980 : 22,1 M est une valeur calibrée, à remplacer si un total tous régimes est trouvé.
- [ ] Élèves et moins de 20 ans : valeurs estimées dans la feuille Indicateurs.

## À faire

- [ ] Mettre à jour 2026 avec les comptes COFOG définitifs de l'Insee, attendus en 2027.
- [ ] Mettre le site en ligne et choisir l'hébergement ; régler `DATA_URLS` si les Excel changent d'adresse.
- [ ] Reporter dans `tools/build_data.py` toute correction manuelle des Excel avant de le relancer.
- [ ] Tester le rapport PDF dans une vraie fenêtre d'impression, sur Chrome, Firefox, Safari et sur téléphone. Il n'a été testé qu'en impression simulée sous Chromium.
