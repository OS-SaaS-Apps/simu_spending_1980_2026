# Historique des demandes

Demandes de l'utilisateur, dans l'ordre, telles qu'écrites. Les décisions qui en découlent sont dans [PROJET.md](PROJET.md).


## 2026-10-28

- I want to build a web application with a really easy interface that shows the french budget (state+social) of 1980, its equivalent in constant €/french people, the budget of 2026, its equivalent in constant €/french people and the transposed budget of 1980 onto 2026 and its equivalent in constant €/french people. We should find around 15 to 20 categories to split the budget into (ministries, cross topics, ...). In a second phase we will let users adjust some cursors to adapt the transposed budget like a game. Put maximal focus on a really good, intuitive and usable ui design. Use french as a language. The goal is to give a reference point for the next presidential election by comparing a quasi balanced budget (1980) to our current unsustainable budget (145 billions in deficit in the first 6 months). The budget data should be available in one excel file per year (1980 and 2026) so I can make corrections easily if needed. Look on the web for this data and write the source for both budget in a memory markdown file. Challenge my requirement, ask questions before implementing.
- 1. everything related to the state incl local, 2. constant /french people (not inhabitant which counts foreigners), 3. not real ministries as they change name often, 4. focus on spending, but include revenue as well but how to categorize it?, 5. no the excel is on the server but the site can read it via url, 6. general public on dektops, tablets and smartphones, 7. let's discuss this later
- 1. nationals only, 2. 19+8?, 3. ok, 4. ok and put the effective € values and the constant € values, 5. i do not care for now - for testing start your own server serving the html and the excels

## 2026-10-29

- pour 2026 tu as mis −153 Md€ de deficit mis ce chiffre n'est que pour le debut de l'annee pas l'annee complete. tu as une explication pour la difference du montant retraite et vieillesse? Les retraites ne gagnent pas 2x plus qu'en 1980 si? et rajoute quelques indicateurs transverseaux du budget 1980/2026 comme le nombre de fonctionnaires (total etat+social+territorial+agences) et si possible le taux d'administratif (2026 environ 34%)
- heu si on est a -145.9Mrds fin juillet on devrait etre a -300Mrds fin decembre non?
- in me manque le budget 2026 a la sauce 1980: par exemple pour les retraites on regarde le defcit de 1980 des regimes de retraites et on fixe le deficit 2026 des retraites a ce chiffre (toujours en euros constant) et de même on peut extrapoler les pensions en suivant la reduction du nombre de cotisant par retraité -> quel est donc le budget retraite dans cette configuration?
- tu rajoute un panel comme pour "Ou va l'argent" et tu fais les calculs pour toutes les categories - si tu as besoin de valeurs de references (comme le nombre de cotisant par retraite pour les retraites cherche en ligne). Et verifie en ligne les chiffres de 1980
- je pense qu'il faut integrer quelque part le nombre d'etrangers/immigres sur le sol francais en 1980 et 2026 (puisqu'ils auraient du entre autre compenser la baisse de la natalite. Si le nombre d'etranger (soit de non francais) est plus important aujourd'hui on pourrait envisager d'utiliser ca pour recalculer le budget necessaire aux aides sociales en prenant en compte le ratio d'etranger qui ne travaillent pas
- oui
- sans donnee de 1980 inutile de montrer cette contribution des immigres
- pour ces deux derniers essaye de trouver des chiffres en ligne (ca peut etre des chiffres de 1978 ou 1979 aussi)
- pour le budget 2026 sauce 1980 tu utilises souvent "Même part de la richesse produite qu'en 1980" mais comme nous sommes en euros constant la bonne comparaison c'est l'evolution de la population francaise et pas du PIB
- pour les interets de la dette il faut garder les chiffres reels de 2026 - ce n'est pas vraiment un choix politique mais plutot une contrainte externe

## 2026-10-30

- comment on est passe de -6Mrds a -123Mrds de deficit pour le budget 2026 sauce 1980
- mais par exemple pour les retraites on a scaler ca en baissant les pensions moyennes donc je ne comprends pas
- qu'est ce que tu raconte? hier tu me disais que la pension moyenne en 2026 sauce 1980 etait de +6% par rapport a 1980 a 1490€
- reprends la premiere version
- y a un moyen d'augmenter les recettes "sauce 1980" pour avoir un deficit 0? ATTENTION: il faut une regle logique
- ok calcule les recettes par personne en emploi plutôt que par Français
- ok ajoute des cotisants en 1980 pour arriver a environ 0 de solde
- pour les retraites, dans le bar graph rajoute le montant moyen de retraite pour la sauce 1980 - y a t il d'autres categories qui devraient avoir une telle indication? Education?
- pour l'education le budget actuel est grevé s'une surcotisation retraite pour les agents du public. ca fausse les € par etudiants - qu'en est il du budget sauce 1980? comment rendre ceci transparent? Pareil pour la sante avec une partie du budget consacre a l'administratif en 2026 (CPAM, ARS, ...) plus importante qu'en 1980 ce qui traffique les €/francais
- non
- dans le tableau "Et si 2026 suivait les règles de 1980 ?" tu donnes la difference dans la colonne de droite mais tu calcul 2026 - "sauce 1980" alors que ca devrait etre l'opposé
- quelques comparaisons m'etonnent comme l'ecole, l'invalidite/handicap et l'infrastructure - par contre4 pour l'envoronnement je trouve la baisse faible au vu des investissements pharaoniqesx que l'iont fait pouir les enr et par l'ademe en 2026
- mais pour l'ecole ca voudrait dire que le budget sauce 1980 divise par deux le budget par eleve!!! comment c'est possible?
- ton cout par eleve tient il compte des investissements des regions qui ont la responsabilite des colleges et lycees depuis la decentralisation?
- ok j'ai identifié quelques raisons pour le +100% entre 1980 et 2026: les sur-cotisation retraite, la plus longue duree d'etude surtout dans le supérieur plus cher, la multiplication des normes, le transfer aux region de l'infrastructure (avec le numerique genre tablette / laptop pour des eleves) - ca explique les +100% ou j'oublie encore quelque chose?
- la cotisation retraite standard dans le publique est de 14% (donc en 1980) recalcule avec ce chiffre. En gros il faut retirer l'effet sur-cotisation des €/eleve du budget 2026 et ajouter les aesh/aide aux eleves en difficultees au budget sauce 1980
- pour le budget 2026 pour les aesh / handicap voici les chiffres (en 1980: 0€): Le budget consacré à l'école inclusive et aux AESH (accompagnants d'élèves en situation de handicap) dans le cadre de la loi de finances est le suivant : • Budget global de l'école inclusive : Il atteint 4,74 milliards d'euros dans le Projet de Loi de Finances (PLF) pour 2026, soit une hausse de 100 millions d'euros par rapport à 2025 (qui s'élevait à environ 4,6 milliards d'euros). • Coût des AESH : À lui seul, le coût de l'accompagnement par les AESH représente une part majeure de ce budget (évalué à environ 3 milliards d'euros par an).
- le chiffre par eleve sauce 1980 n'a pas bougé
- c'est quoi cette commande de fou furieux ? tu peux pas juste lire le excel?
- moi je retgarde le bar graph: 1070€ sauce 1980 et 2220€ en 2026
- oui retire la ... mais on reste a +100% sur les chiffres par francais!!! malgré les 4 raisons évoquées
- ok mais les salaires des enseignants n'ont pas fait +76% et les batiments non plus et les enfants par classe n'ont pas l'air d'avoir diminuer (en tout cas pas chez mes enfants!)
- en 1980: • Enseignants (Public + Privé sous contrat) : ~700 000 • Personnels Non Enseignants (Administratifs, Techniques, Santé, Vie Scolaire) : ~230 000 à 250 000
- mais l'augmentation du nombre d'enseignant devraitr aller de pair avec l'augmentation du nombre d'eleve et de la population donc ce ne devrait pas etre un grand changement dans les chiffres €/francais
- c'est enorme!!! ils sont ou tous ces "profs"? C'est lie au ZEP?
- ok indique quelque part que le budget sauce 1980 supprime les rep
- ok on passe au "jeu" ajoute une nouvelle section avec a la base une copie de la derniere table mais en rajoutant des sliders par categorie qui augmentent ou diminuent le budget de cette categorie, il faut quelque part montrer le budget total inclus le nouveau deficit. On fait ca pour les depenses et on verra les recettes apres

## 2026-10-31

- ok maintenant on fait un tableau recettes sauce 1980
- currently the sliders align with the 2026 values - let them start at the sauce 1980 value. Et il faudrait pouvoir exporter un rapport qui permette a un utolisateur de repartir avec la liste de ses choix
- sous le "jeu" affiche le résultat avec des bar graph verticaux: Chiffres totaux pour le Budget Recettes/Depense 2026 -> le Budget Recettes/Depenses Sause 1980 ajustées par l'utilisateur - inclus ce resume dans le rapport
- dans les bar graph horizontaux l'unite n'est pas claire. Indique clairement 1450 € / francais ou selon la valeur affichee
- dans "mes choix" dans le pdf il manque une colonne 2026
- ok sauvegarde mes demandes, nos decisions, les details du projet dans un ou plusieurs md files. note les questions et taches en cours dans tasks.md

## 2026-11-01

- pour la categorie education nationale je pense a passer les bar graph en € / eleve a cause de l'augmentation de la population francaise mais la diminution du nombre d'eleves
- on est d'accord que tu as integrer le budget pour le aesh dans la sauce 1980?

## 2026-11-06

- I plan to add this project to github - look at the currrent folder and organize the files according to best practices. Add a readme and a .gitignore
- look for api keys and passwords - report any finding
- this is a one shot project - it is probably not reusable - which licence should I use? (choix : MIT, au nom de OS-SaaS-Apps)
- ok the target is https://github.com/OS-SaaS-Apps/simu_spending_1980_2026.git
- change to OS-SaaS-Apps (auteur des commits, adresse noreply GitHub)
- is there a way to "host" this project on github? like with pages? → oui
- make sure the task list is updated and write in the project memory md file what we did last
