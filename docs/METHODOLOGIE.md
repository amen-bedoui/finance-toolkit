# Méthodologie

## Architecture
`src/build_workbook.py` génère un classeur Excel (formules vivantes, sans macro) à partir d'un fichier JSON (`data/examples/*.json`) ou vierge (`--blank`). `src/analyse.py` lit le classeur calculé et produit le rapport PDF. `src/import_balance.py` remplit l'onglet Saisie depuis une balance générale ou un FEC grâce à `data/pcg_mapping.csv` (éditable).

## Calculs
- **SIG / CAF** : marge brute, valeur ajoutée, EBE/EBITDA, EBIT, résultat. CAF = résultat net + dotations (méthode additive simplifiée).
- **Bilan fonctionnel** : FRNG, BFR (exploitation / hors exploitation), trésorerie nette, dette nette ; dette nette hors dettes locatives (définition « covenant »).
- **Ratios** : rentabilité, croissance, liquidité, solvabilité, activité (DSO/DPO corrigés de la TVA), cash. Lecture automatique selon des seuils par profil (onglet Seuils).
- **Scores** : Altman Z'' (non industriel), Piotroski F, DuPont à 5 facteurs. Outils de screening.
- **DCF** : FCFF actualisés au WACC (CAPM), valeur terminale Gordon-Shapiro et multiple de sortie pondérées, convention mid-year optionnelle, 3 scénarios pondérés, sensibilités en formule fermée (SUMPRODUCT).
- **IFRS 16** : commutateur. Traitement 1 : loyers déduits du FCFF et de l'EBITDA de sortie, dette nette hors dettes locatives. Traitement 0 : dettes locatives dans la dette.
- **WACC** : bêta saisi, ou médiane des bêtas désendettés des pairs ré-endettée (Hamada), éventuellement ajustée (Blume).
- **Comparables** : quartiles du VE/EBITDA et du PER du panel appliqués à l'EBITDA et au résultat net N ; football field ; alerte si moins de 3 pairs ; multiples implicites au cours de marché.
- **DCF à 10 ans** : années 1-5 = scénario Base ; années 6-10, la croissance converge linéairement vers g et la marge vers une cible saisie (vide = marge de l'année 5) ; capex, D&A et BFR restent aux ratios de l'année 5. Réduit le poids de la valeur terminale pour les profils d'hypercroissance.
- **DCF inversé** : grille de taux d'actualisation (g + 1 % à g + 21 %) ; le taux pour lequel la valeur par action égale le cours, par interpolation linéaire, est le rendement annuel attendu si le scénario Base se réalise. Comparé au WACC, il mesure la marge de sécurité.
- **Transactions comparables** : multiples VE/EBITDA, VE/CA et prime de contrôle sur des transactions saisies ; résultat « n/d » sous 3 transactions. Valeur de contrôle : ne se compare pas à un cours minoritaire sans retrancher la prime.
- **Note de marché** : cours cible = moyenne pondérée (poids saisis, renormalisés sur les méthodes disponibles) du DCF 5 ans, DCF 10 ans, comparables VE/EBITDA et PER ; appréciation mécanique selon des seuils de potentiel ; ratio gain/perte Bull/Bear ; notation synthétique par couverture des intérêts (table pédagogique de type Damodaran, non une notation d'agence).
- **Registre des hypothèses (Sources)** : chaque hypothèse clé porte une source, une date, un niveau de confiance et une justification ; statut automatique « documentée » (source citée hors jugement, datée, < 180 jours), « jugement », « non datée », « source ancienne » ou « non sourcée ».
- **Marché** : date de référence, ancienneté du cours (alerte au-delà de 5 jours), consensus (CA, EBITDA, cours cible) comparé au scénario Base ; un écart > 10 % sur N+1 doit être expliqué.
- **Qualitatif** : cinq forces concurrentielles (1-5), registre des risques (probabilité × impact, 1-25), catalyseurs ; profil « élevé / modéré / faible » selon le score maximal et un seuil réglable.
- **Conviction** : nombre de contrôles réussis (méthodes disponibles, poids de la valeur terminale, dispersion, panel, hypothèses documentées, fraîcheur du cours, risque qualitatif, cours cible dans le football field) ; ≥ 75 % élevée, ≥ 50 % moyenne. Un modèle sans sources ne peut donc pas afficher une conviction élevée.
- **Rapprochement de trésorerie** : CAF + ΔBFR − capex + Δ dettes − dividendes − acquisitions − rachats − autres investissements et décaissements + rémunération en actions − gains latents, comparé à la variation constatée au bilan ; l'écart résiduel est affiché, jamais masqué.
- **Rémunération en actions (SBC)** : conservée dans les charges (traitement prudent) ; EBITDA ajusté et ratio SBC/CA fournis en information.
- **LBO** : dette = multiple × EBITDA, flux du scénario Base, cash sweep ; TRI et prix maximal pour un TRI cible (la dette ne dépend pas du prix, donc formule fermée).
- **M&A** : accrétion / dilution du BPA avec financement mixte, synergies phasées, amortissement des incorporels issus de l'écart d'acquisition.

## Limites assumées
Modèles simplifiés ; pas de convexité de la dette, de covenants, d'effets de timing, de retraitements fiscaux fins ni de normalisation de l'EBITDA. Les seuils sectoriels sont des jugements d'analyste. Outil d'analyse et de formation : aucun conseil en investissement.
