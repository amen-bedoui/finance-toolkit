# Validation sur société réelle — Wavestone SA

**Objectif.** Vérifier que le classeur reproduit les chiffres publiés d'une société cotée (états consolidés IFRS, M€, exercices clos au 31 mars 2024, 2025 et 2026), au-delà de l'exemple fictif.

**Source des données.** Communiqués et rapports financiers publiés par la société (Euronext / AMF). Les lignes IFRS ont été regroupées dans les rubriques simplifiées de l'onglet Saisie (`data/examples/wavestone_public.json`). Cours, nombre d'actions et multiples de pairs viennent d'agrégateurs gratuits (dates variables) : ils servent à tester la mécanique, pas à fonder une décision.

## Résultats (tests `tests/validate_wavestone.py`, tous verts)

| Grandeur | Publié | Classeur | Écart |
|---|---|---|---|
| CA 2025/26 | 954,347 | 954,347 | 0 |
| Résultat net 2025/26 | 82,100 | 82,100 | 0 |
| Total bilan 2025/26 | 996,894 | 996,894 | 0 |
| Capitaux propres 2025/26 | 702,465 | 702,465 | 0 |
| EBITDA (RO courant + dotations nettes) | 132,758 | 132,758 | 0 |
| Trésorerie nette hors dettes locatives | 121,4 | 121,421 | < 0,1 |

Ces égalités testent surtout la qualité du **mapping** et l'équilibre des contrôles internes (bilan, résultat). Elles ne valident pas, à elles seules, les modèles de valorisation.

## Écarts de définition (volontaires, à connaître)

- **CAF.** Le classeur calcule *résultat net + dotations* (102,2 en 2025/26). La CAF publiée par la société (139,4) est avant intérêts et impôt et inclut les rémunérations en actions : définitions non comparables.
- **Free cash-flow.** Le FCFF du classeur est un flux économique théorique (NOPAT + dotations − capex − ΔBFR), avant loyers IFRS 16 et avant M&A. Le FCF publié est après loyers et avant acquisitions.
- **EBIT.** Le classeur tient compte de tous les produits et charges d'exploitation saisis ; il diffère du « résultat opérationnel courant » publié (119,9) par les éléments non courants.
- **Capex estimé.** Quand le capex n'est pas saisi, il est estimé par Δ immobilisations + dotations. Avec des acquisitions de filiales (65,9 / 44,0 / 10,6 M€ sur les trois exercices) cette estimation devient fausse : le capex réel doit être saisi. C'est pourquoi le modèle le demande pour les années où il y a des acquisitions.
- **Rapprochement de trésorerie.** Écart résiduel 2025/26 : 1,8 M€ (0,2 % du CA). Sur 2024/25 l'écart est plus élevé (≈ 39 M€) : variations de capital, de provisions et éléments hors exploitation non saisis ; l'onglet Cashflow le montre explicitement plutôt que de le masquer.

## Défauts trouvés grâce à ce test (et corrigés)

1. Capex estimé erroné en présence d'acquisitions → ajout de lignes *acquisitions* et *rachats* dans le rapprochement, capex réel en entrée.
2. Dette nette gonflée par les dettes locatives IFRS 16 pour un cabinet de conseil → commutateur IFRS 16 (loyers en charge, dette nette hors dettes locatives).
3. DSO/DPO faussés par la TVA (clients TTC / CA HT) → correction par le taux de TVA saisi.
4. Mapping PCG : le préfixe `10` captait les réserves (`106`) comme du capital → trouvé par le test d'aller-retour de l'import, corrigé.

## Ce que cette validation ne prouve pas

- Les hypothèses DCF (croissance, marge, WACC) sont **illustratives**. La valeur par action obtenue n'est pas une estimation de la valeur de Wavestone.
- Un panel de 3 à 4 pairs est trop étroit pour un bêta ou un multiple robustes.
- Aucune comparaison avec un modèle d'analyste ou un rapport d'évaluation indépendant n'a été faite.
