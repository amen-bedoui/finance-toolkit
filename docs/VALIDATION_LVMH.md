# Validation sur LVMH (comptes consolidés 2023-2025, M€)

**Sources.** Comptes consolidés au 31 décembre 2025 (compte de résultat, bilan, flux, capitaux propres ; traduction anglaise publiée le 12/02/2026 et chiffres clés du site lvmh.com) et communiqué de résultats du 27/01/2026. Cours : 382,40 € relevé le 20/11/2026 sur un widget de cotation (à re-vérifier le jour de la publication). Nombre d'actions : capital 149 306 082 € / nominal 0,30 € = 497 686 940 (non retraité des actions auto-détenues).

## Rapprochement avec les chiffres publiés (tests `tests/validate_lvmh.py`, tous verts)

| Grandeur | Publié | Classeur |
|---|---|---|
| Ventes 2023 / 2024 / 2025 | 86 153 / 84 683 / 80 807 | identique |
| Résultat opérationnel courant | 22 802 / 19 571 / 17 755 | identique |
| Résultat net avant minoritaires | 15 952 / 12 958 / 11 222 | identique |
| Total bilan | 143 694 / 149 190 / 142 037 | identique |
| Capitaux propres (avec minoritaires) 2025 | 68 949 | identique |
| EBITDA 2025 (ROC + dotations, dont droits d'utilisation) | 25 756 | identique |

Ces égalités valident le **mapping** (états par fonction → rubriques du modèle) et la cohérence interne, pas les modèles de valorisation.

## Choix de mapping (à connaître)
- Les charges étant présentées **par fonction**, « achats » = coût des ventes ; « charges externes » = solde des charges d'exploitation hors dotations ; « personnel » = 0 (inclus dans les fonctions). L'EBITDA et le ROC sont exacts, mais la marge brute du modèle (≈ 66 %) n'a de sens que parce que le coût des ventes est repris tel quel.
- Dotations = amortissements et provisions nets + amortissement des droits d'utilisation (IFRS 16). Quote-part des sociétés mises en équivalence → « autres produits ».
- Autres produits et charges opérationnels (−656 en 2025) → éléments exceptionnels, donc **sous** l'EBIT du modèle (= ROC publié).
- Dettes locatives incluses dans les dettes financières (mémo IFRS 16 saisi) ; loyers = remboursement + intérêts des dettes locatives (3 527 en 2025).
- Engagements d'achat de titres de minoritaires (6 331) classés en autres dettes ; ils sont aussi repris dans les hypothèses comme dette assimilée (avec les minoritaires comptables, 1 477).

## Écarts de définition
- **Dette financière nette.** Publiée : 6 857 (2025), 9 228 (2024), 10 746 (2023). Le modèle donne 11 549 / 13 311 / 14 133 hors dettes locatives, car il part des seuls emprunts moins la trésorerie du bilan ; la définition du groupe déduit aussi des placements financiers courants et des dérivés. L'écart 2025 (4 692) est ajouté aux « actifs non opérationnels » de la valorisation pour éviter de pénaliser la valeur des capitaux propres ; un test le verrouille.
- **Flux.** Le FCFF du modèle (12 577 en 2025) est un flux économique théorique avant loyers ; le « cash flow disponible d'exploitation » publié (11 333) est après loyers, impôt et intérêts décaissés. Définitions non comparables.
- **Minoritaires** pris à leur valeur comptable, alors que leur valeur de marché est plus élevée.

## Résultats des méthodes (hypothèses illustratives, 23/11/2026)
- DCF : 306 € (Bear) / 412 € (Base) / 468 € (Bull) par action contre 382,40 €. Ce n'est pas une estimation de la « vraie » valeur : croissance, marges, WACC (≈ 9,7 %) et multiple de sortie sont des hypothèses d'école.
- Comparables : 436 à 980 € (Q1-Q3). L'écart énorme vient du panel : Hermès (21x) et Richemont (20x) valent très au-dessus de Kering (11x), Moncler et Prada (7-10x). Le choix des pairs détermine le résultat ; ces multiples proviennent d'agrégateurs ou d'extraits non primaires et ne sont pas datés. **À remplacer par Bloomberg, FactSet ou CapIQ avant toute diffusion sérieuse.**
- Bêtas : seuls deux bêtas (Kering, Prada) sont disponibles, d'où un bêta saisi (mode 1) et non une médiane de pairs.
- LBO et accrétion/dilution : exercices **mécaniques** uniquement. Aucun sponsor ne finance un LBO à plus de 200 Md€ ; l'acquéreur du M&A est fictif.

## Ce que ce test ne prouve pas
Qu'un modèle simplifié (CAF additive, pas de retraitement fiscal fin, minoritaires comptables) reproduit les valorisations d'analystes. Il prouve que les états d'un grand groupe cotés se saisissent proprement et se rapprochent des chiffres publiés.
