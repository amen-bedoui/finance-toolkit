# Changelog

## 1.4.0
- Cours NVIDIA rebasé sur la clôture du 20/11/2026 (125,53 $).
- Nouvel onglet `Pont_Consensus` (écart modèle / consensus / cours, multiples implicites, sensibilité à la décote).
- Décote juridique paramétrable (`marche.decote`, `marche.decote_source`) appliquée à la cible de `Reco_Marche`.
- Règle automatique de valeur aberrante sur les méthodes de la cible (seuil paramétrable) et graphique par méthode.
- Sensibilité de la cible au repli des multiples des pairs (Pont_Consensus).
- 79 tests.

## 1.3.0 — 2026-11-23
- DCF calé sur un bilan intermédiaire : la part écoulée du FCFF N+1 est retirée, avec période d'actualisation explicite ; exemple NVIDIA recalé sur le bilan du 25/10/2026 (8-K T3 FY27).
- Onglet Transactions alimenté depuis le JSON (colonne source) ; 5 transactions semi-conducteurs réelles pour NVIDIA ; statistiques déclenchées à partir de 3 données exploitables par colonne.
- Panel de pairs NVIDIA daté (Broadcom, AMD, TSMC ; Marvell exclu).
- Base LTM optionnelle dans Comparables : multiples de marché appliqués à l'EBITDA et au résultat net des 12 derniers mois.
- Transactions NVIDIA : EBITDA calculé pour 3 opérations (Lumentum, Microsemi, Fairchild), VE/EBITDA médian 17,6x ; 76 tests.

## 1.2.0 — 2026-11-23
- Exemple NVIDIA rebasé sur des données datées : T3 FY27 (8-K 17/11/2026), consensus LSEG, clôture Nasdaq du 18/11/2026, prime de risque et bêta Damodaran, taux 10 ans du 18/11. Hypothèses de jugement signalées comme telles.
- Nouveau paramètre « fraction d'année écoulée » (valorisation à une date postérieure à la clôture), appliqué aux DCF 5 et 10 ans et au DCF inversé.
- Qualitatif mis à jour (gouvernance, Chine, approvisionnement, secteur) ; 74 tests.

## 1.1.0 — 2026-11-23
- Onglets Sources (registre des hypothèses), Marche (cours daté, consensus) et Qualitatif (forces concurrentielles, risques, catalyseurs).
- Niveau de conviction dans la note de marché ; 73 tests.

## 1.0.0 — 2026-11-23
- Classeur Excel à formules natives : SIG, CAF, BFR, flux, ratios, scores, DuPont, variations N/N-1.
- Valorisation : DCF 3 scénarios, DCF à 10 ans avec convergence, DCF inversé, comparables (VE/EBITDA et PER), transactions, LBO, accrétion/dilution, football field.
- Note de marché : cours cible pondéré, risque/rendement, notation de crédit synthétique indicative.
- Import balance/FEC avec mapping PCG éditable ; rapport PDF ; 67 tests.
- Validation sur Wavestone, LVMH et NVIDIA (US GAAP) ; panel de pairs semi-conducteurs (SEC).
