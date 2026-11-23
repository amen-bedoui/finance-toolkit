# Panel de pairs (semi-conducteurs) — fondamentaux FY2023/FY2024

`semis_fundamentals_FY2023-24.csv` : croissance, marges (opérationnelle, EBITDA = résultat opérationnel + D&A, nette, FCF = flux d'exploitation − capex), R&D et SBC / CA, dette nette / EBITDA, ROE.

**Source** : états financiers 10-K issus de la base SEC « Financial Statement and Notes », via le jeu Hugging Face `DenyTranDFW/SEC_10K_FSNoNDS_Zip` (dossier `OFFICE_OF_MANUFACTURING`, fichier `SEMICONDUCTORS_AND_RELATED_DEVICES.rar`). Recréer le CSV : extraire l'archive puis `python tools/build_peer_panel.py <dossier> <sortie.pkl>`.

**Contrôle** : NVIDIA FY2024 (clos le 31/01/2024) recoupé avec les montants du modèle : CA 60 922, résultat opérationnel 32 972, résultat net 29 760, D&A 1 508, total actif 65 728, capitaux propres 42 978, SBC 3 549, liquidités + titres 25 984 (M$) : identiques.

**Limites**
- Dernier exercice disponible : FY2023 pour la plupart des pairs, FY2024 pour Nvidia et Marvell. Les exercices plus récents ne sont pas dans le jeu.
- Pas de cours ni de capitalisation : ce panel donne des fondamentaux, pas de multiples de valorisation.
- Extraction par balises XBRL : une balise absente donne une cellule vide (par ex. R&D d'ON Semiconductor). La dette nette / EBITDA est indicative (couverture des balises de dette variable).
- Qualcomm est absent du jeu. Le jeu est sous licence GPL : le CSV ne contient que des ratios calculés à partir de données SEC publiques ; vérifier la compatibilité avant de redistribuer les parquet bruts.
