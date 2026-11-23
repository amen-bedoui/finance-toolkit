# Validation sur NVIDIA (US GAAP, FY2024-FY2026, M$)

**Pourquoi Nvidia.** Profil opposé à Wavestone et LVMH : US GAAP, exercice clos fin janvier, hypercroissance (CA +65 % en FY2026), marges d'EBITDA > 60 %, trésorerie nette, rachats d'actions massifs, rémunération en actions importante, investissements financiers lourds. Les états sont publiés par la société (communiqués de résultats / 10-K, SEC).

**Sources.** États financiers : communiqués du 25/02/2026 (FY2026, FY2025) et du 21/02/2024 (FY2024). Mise à jour v1.2 : communiqué T3 FY27 (8-K du 17/11/2026), cours de clôture Nasdaq du 20/11/2026 (125,53 $, −14,0 % ; NVIDIA Investor Relations, recoupé Yahoo Finance), 10 ans américain 5,07 % (18/11/2026), prime de risque Damodaran 4,14 % (01/09/2026), bêta sectoriel Damodaran semi-conducteurs 1,49 (janvier 2026), consensus LSEG relayé par CNBC (CA +90 % en FY27, +70 % en FY28), cible moyenne TipRanks 229,33 $ (44 analystes, 18/11/2026). Pairs : agrégateurs, non datés pour la plupart.

## Rapprochement avec les chiffres publiés (tests `tests/validate_nvidia.py`, tous verts)

| Grandeur | Publié | Classeur |
|---|---|---|
| Revenue FY2024 / FY2025 / FY2026 | 60 922 / 130 497 / 215 938 | identique |
| Operating income | 32 972 / 81 453 / 130 387 | identique |
| Net income | 29 760 / 72 880 / 120 067 | identique |
| Total assets | 65 728 / 111 601 / 206 803 | identique |
| Shareholders' equity FY2026 | 157 293 | identique |
| EBITDA FY2025 / FY2026 (operating income + D&A) | 83 317 / 133 230 | identique (recoupé par un agrégateur : 83,32 Md / 133,23 Md) |
| Trésorerie nette FY2026 (liquidités + titres − dette) | 54 088 | identique |

## Choix de mapping
- **Charges** : coût des ventes → « achats » ; R&D + frais commerciaux et généraux − D&A → « charges externes » ; D&A → dotations. La rémunération en actions reste dans les charges (US GAAP) ; elle est saisie en mémo pour calculer l'EBITDA ajusté (139,6 Md$ en FY2026) et le ratio SBC / CA (3,0 %).
- **Produits financiers** : intérêts + « other income, net » (9,0 Md$ en FY2026, dont l'essentiel de gains latents sur titres : 8,9 Md$ retraités du flux).
- **Bilan** : trésorerie = liquidités + titres négociables ; participations non cotées → actifs financiers (et « actifs non opérationnels » de la valorisation). Les dettes locatives opérationnelles sont incluses dans les dettes financières et déclarées en mémo ; **sous US GAAP le loyer opérationnel est déjà en charges** : loyers IFRS 16 = 0, traitement 1 (dette nette hors dettes locatives).

## Ce que le test a révélé (et corrigé)
1. **Rapprochement de trésorerie** : sans retraitement, l'écart FY2026 était de −43 Md$ (20 % du CA). Il manquait la rémunération en actions (non décaissée), les gains latents (non encaissés), les impôts payés sur actions gratuites (7,9 Md$), les achats de titres non cotés (17,5 Md$) et la licence/acquisition Groq (13 Md$). Ces lignes sont désormais saisissables ; l'écart tombe à −2,6 Md$ (1,2 % du CA). Le contrôle n'est pas contourné : il reste visible.
2. **EBITDA ajusté de la rémunération en actions** et **multiples implicites du marché** ajoutés à la Synthèse.
3. **Comparables par PER** ajoutés (le VE/EBITDA seul était inutilisable faute de données de pairs).
4. Le classeur **signale** qu'un panel est insuffisant (« < 3 pairs ») au lieu de produire une fourchette trompeuse ; le football field affiche « n/d » pour la méthode concernée.

## Résultats (hypothèses illustratives)
- Au cours de 125,53 $ (20/11/2026), la capitalisation boursière est ≈ 3 036 Md$ pour 24,19 Md d'actions ; ces multiples de marché sont recalculés dans l'onglet Synthèse du classeur.
- DCF (WACC ≈ 13,2 % avec un bêta de 1,5, entre le bêta brut 5 ans ≈ 2,2 et un niveau sectoriel) : 60 $ (Bear), 129 $ (Base), 188 $ (Bull). Le cours se situe entre le Base et le Bull : le marché prend des hypothèses plus proches du Bull ou un taux d'actualisation plus bas.
- PER de comparables : 146 à 232 $ (3 pairs seulement ; Micron, cyclique, tire le bas, Broadcom et Analog Devices le haut).
- LBO / M&A : exercices mécaniques sans portée (aucun sponsor ni acquéreur à cette taille).

## Limites que ce test met en évidence
- **Horizon de 5 ans + valeur terminale** : la croissance de l'année 1 (90 %) et de l'année 2 (70 %) est alignée sur le consensus ; les années 3 à 5 (25 %, 12 %, 7 %) et les marges sont des hypothèses de l'auteur. Valeur Base à 5 ans ≈ 269 $ ; l'onglet **DCF_10ans** (croissance convergeant vers 3 %, marge vers 50 %, multiple de sortie 14x) donne ≈ 233 $, soit −13 %, avec un poids de la valeur terminale de 49 % (Gordon) à 57 % (multiple). Le résultat dépend fortement de la marge et du multiple de long terme.
- **DCF inversé** : au cours de 125,53 $, le rendement annuel implicite du scénario Base est ≈ 20,8 %, contre un WACC de 11,2 % : le cours est très inférieur à la valeur du Base ; soit le marché doute du Base (enquête SEC, gouvernance, Chine, cycle d'investissement), soit il actualise à un taux plus élevé.
- **Note de marché** : cours cible pondéré ≈ 137 $ (−13 %) ; l'appréciation mécanique ressort « défavorable » sous ces seuls hypothèses illustratives. Ce n'est pas un avis sur la valeur Nvidia.
- Le bêta (1,5) est un jugement ; la valeur DCF y est très sensible (voir la table de sensibilité au WACC).
- Les données de pairs sont gratuites et partiellement datées : à remplacer par Bloomberg, FactSet ou CapIQ pour un usage sérieux. Aucun résultat trimestriel postérieur au 10-K FY2026 n'est intégré.
- Cours à ± quelques % selon la source et le jour (plusieurs agrégateurs divergent) : à re-vérifier avant publication.

## Recoupement avec les états SEC (jeu Hugging Face)
Les états 10-K extraits du jeu `DenyTranDFW/SEC_10K_FSNoNDS_Zip` donnent pour FY2024 les mêmes montants que les entrées du modèle (CA, résultat opérationnel, résultat net, D&A, total actif, capitaux propres, SBC, liquidités). Ils ne couvrent pas FY2025 ni FY2026.

**Repères de marge issus du panel de pairs** (`data/peers/semis_fundamentals_FY2023-24.csv`, marge d'EBITDA du dernier exercice) : Nvidia FY2024 56,6 %, Broadcom 54,6 %, Texas Instruments 48,6 %, Microchip 48,8 %, NXP 35,9 %, Analog Devices 33,8 %. La marge cible de 50 % retenue pour l'année 10 du DCF se situe dans la zone des meilleurs pairs établis ; elle reste un choix, pas une prévision. Le panel ne contient ni cours ni capitalisations : les multiples de comparables restent à documenter avec des cours datés.

## Registre des sources (onglets Sources, Marche, Qualitatif)
Sur l'exemple Nvidia, seules **13 % des hypothèses clés sont « documentées »** (source citée, datée, de moins de 180 jours) : taux sans risque et cours. Les autres relèvent du jugement de l'auteur (prime de risque, bêta, croissance, multiples) ou de sources anciennes ou non datées (panel de marges de FY2023-24, nombre d'actions). Le cours date de 6 jours (alerte au-delà de 5). Le consensus des analystes n'est pas renseigné faute de source professionnelle. Le profil de risque qualitatif ressort « élevé » (restrictions d'export, dépendance à la fabrication et à la mémoire, cycle d'investissement en IA : appréciations de l'auteur à étayer avec les facteurs de risque du 10-K). Niveau de conviction indicatif : **moyen** (4 contrôles réussis sur 8). C'est le message voulu : le modèle montre ce qui est établi et ce qui ne l'est pas.


## Mise à niveau analyste (v1.2, 23/11/2026)

- L'ancien exemple (croissance de 55 % en année 1, cours de 158,09 $ non recoupé) était périmé : Nvidia avait déjà publié 9 mois de l'exercice FY27 et relevé sa guidance. Le Base est maintenant calé sur le consensus pour les deux premières années ; l'écart modèle/consensus sur le CA est donc faible **par construction**, ce qui n'est pas une validation.
- Valorisation calée sur le bilan du 25/10/2026 (paramètre « part de l'année écoulée » = 0,75) : 25 % seulement du FCFF de FY27 est actualisé, car le reste est déjà dans la trésorerie. Les actifs non opérationnels valent 76 800 M$ = titres non cotés 65 528 (8-K T3 FY27) + trésorerie nette générée depuis la clôture FY26 de 11 272 (97 731 − 32 371 − 54 088).
- Cours cible du modèle ≈ 262 $ (pondération DCF 5 et 10 ans et comparables sur base LTM) contre 229 $ pour le consensus TipRanks (écart ≈ +14 %, donc au-dessus des analystes) ; les agrégateurs divergent (247 $ Benzinga, 264 $ MarketBeat) et les cibles sont en cours de révision.
- Registre des sources : 6 hypothèses documentées sur 15, 5 de jugement, 4 sur source ancienne (bêta de janvier 2026, panel de pairs FY2023-24, actif non opérationnel de février 2026). Conviction indicative : moyenne.
- Qualitatif mis à jour : démission du PDG le 12/11/2026 (intérim Colette Kress), perspective de crédit Moody's ramenée de positive à stable le 18/11, chute du SOX de 12,25 % le 05/11, engagements d'approvisionnement (≈ 390 Md$ selon la presse, à vérifier dans le 10-Q).
- **Transactions (v1.3)** : 5 opérations réelles avec source : Marvell/Lumentum (nov. 2026, VE 115 Md$, prime 23,9 %), Avago/Broadcom (2015), Microchip/Microsemi (2018), Analog Devices/Hittite (2014), ON/Fairchild (2015). Trois ont un EBITDA calculé à partir des communiqués : Lumentum (EBITDA ajusté LTM 1 476 M$, VE/EBITDA 77,9x), Microsemi (résultat d'exploitation non-GAAP 578 M$, 17,6x), Fairchild (EBITDA GAAP 159 M$, 15,1x). Médiane 17,6x, Q1–Q3 16,3x–47,7x ; valeur de contrôle de NVIDIA ≈ 95–268 $ (médiane 102 $). Les définitions d'EBITDA diffèrent (ajusté, non-GAAP, GAAP) et l'échantillon est dominé par des cibles très différentes de NVIDIA (Fairchild à faible marge, Lumentum en hypercroissance) : la fourchette est donc indicative et très large. Broadcom et Hittite n'ont pas d'EBITDA retenu (aucun chiffre vérifié).
- **Base LTM (v1.3)** : les multiples de pairs sont des LTM ; ils sont désormais appliqués à l'EBITDA LTM de NVIDIA au 25/10/2026 : 133 230 + (194 791 + 3 378) − (86 088 + 2 031) = 243 280 M$ (8-K T3 FY27). Le résultat net LTM publié (≈ 226 257 M$) contient des gains sur titres ; le PER est appliqué à un résultat net normalisé de 198 923 M$ (EBIT LTM 239 090 × (1 − 16,8 %)). Valeur par comparables : VE/EBITDA médiane ≈ 239 $, PER médian ≈ 364 $.
- **Pairs (v1.3)** : Broadcom (23,2x, 28/10), AMD (81,1x), TSMC (18,0x), Marvell (52,5x, exclu : cours perturbé par l'acquisition de Lumentum), multiples Finviz du 11/11/2026. Médiane ≈ 23x, très dispersée (18x à 81x). Le PER des pairs vient d'agrégateurs non datés (Broadcom 44x, Analog Devices 49x, Micron 14x) : il est la partie la moins fiable du panel.
- Limites : cours en forte baisse après le 18/11 (cotation intrajournalière ≈ 127 $ le 20/11 selon Investing.com, non retenue faute de clôture) ; pairs et multiples toujours issus d'agrégateurs ; aucun consensus d'EBITDA ; pas de transactions comparables réelles ; ce n'est pas un conseil en investissement.

## Mise à jour v1.4 (23/11/2026)

- **Cours** : clôture du 20/11/2026 à 125,53 $ (−14,0 % le jour de l'information sur une enquête de la SEC, Reuters). Le cours est âgé de 3 jours à la date d'analyse ; le marché a rouvert le 23/11, à revérifier avant tout usage.
- **Décote juridique de 10 %** (onglet Marche, B25) appliquée à la cible : c'est un jugement de l'auteur, non calibré, car aucun montant de sanction n'est connu. Cible avant décote 243,97 $, après décote 219,57 $, contre 229,33 $ pour le consensus (−4,3 %).
- **Onglet Pont_Consensus** : prix implicites (cours, consensus, modèle) en VE/EBITDA et PER, sensibilité de 0 à 40 % de décote, et décote nécessaire pour justifier le cours actuel (≈ 49 %).
- **Règle de valeur aberrante** (Reco_Marche, B17) : une méthode qui s'écarte de plus de 40 % de la médiane des méthodes est exclue de la moyenne. La méthode PER (364 $, +47 %) est donc écartée et sa valeur reste affichée. Le seuil de 40 % est un choix de l'auteur. Un graphique compare les méthodes.
- **Sensibilité aux multiples des pairs** (Pont_Consensus, section 4) : les multiples du panel datent du 28/10 (Broadcom) et du 11/11 (autres), et le secteur a baissé depuis. Un repli de 30 % des multiples ramène la valeur par comparables de 238,67 $ à 168,66 $ et la cible du modèle de 219,57 $ à 204,75 $, soit toujours +63 % au-dessus du cours. Je n'ai pas pu rafraîchir les cours des pairs à une date commune avec des sources gratuites fiables : la table borne l'effet, elle ne le mesure pas.
- **Précision** : les multiples de pairs et celui de NVIDIA sont tous des LTM ; la base est cohérente. Le risque porte sur la date et sur la qualité des agrégateurs, pas sur la méthode.
