# Packaging Engine MMS – macro VBA de suggestion de caisses

Livrable principal : **`packaging_engine_mms_v2.xlsm`** (classeur d'origine + 2 nouvelles feuilles + macro VBA + boutons « Remplir »).

## Ouverture dans Excel

1. Si Windows affiche « Les macros ont été bloquées » : clic droit sur le fichier → Propriétés → cocher **Débloquer** → OK, puis rouvrir.
2. Cliquer sur **Activer les macros** / **Activer le contenu**.
3. Si les boutons « Remplir » n'apparaissent pas : `Alt+F8` → `Packaging_Setup` → Exécuter (recrée les boutons).

## Feuilles

| Feuille | Rôle |
|---|---|
| `Produits` | Inchangée. Chaque produit = jusqu'à 3 sous-articles (`sub_Item_n`, `quantite_n`, dimensions). Colonne `Par Saison?` : `Oui` = emballage Summer en été. |
| `Caisses` | Inchangée. `Saison` = `All seasons` ou `Summer season`. |
| `Packaging Calculator` | `B1` Marge vide All Seasons (5 %), `B2` Marge vide Summer (20 %), modifiables. Table `Item / Quantity / Suggested Package All Seasons / Suggested Package Summer / Info` à partir de la ligne 4. Bouton **Remplir**. |
| `Mixed Pack Helper` | Une commande par ligne : `Produit1 / Quantité1 / … / Produit5 / Quantité5` (A:J), résultats `Suggested Package All Seasons / Suggested Package Summer / Info` (K:M). Quantité vide = 1. Bouton **Remplir**. Les marges sont celles de `Packaging Calculator`. |

`Item` / `ProduitN` = SKU (ex. `9004`) ou, à défaut, le nom exact du produit (un produit composé de la liste est accepté comme n'importe quel autre).

## Règles implémentées

- Hauteur utilisable d'une caisse = `Hauteur × (1 − marge)` (la marge vide est en haut).
- Placement 3D sans débordement (heuristique « espaces maximaux », 6 orientations possibles par article, plusieurs stratégies essayées, la meilleure retenue). Constante `ALLOW_TIP_OVER` dans le module pour interdire de coucher les articles.
- Priorité au **minimum de caisses** : si tout tient dans une seule caisse, on prend la plus petite qui convient ; sinon on remplit la caisse qui prend le plus de volume, et on recommence avec le reste.
- All Seasons : uniquement les caisses `All seasons`. Summer : uniquement les caisses `Summer season`, et seulement si le produit (ou au moins un produit de la commande) est `Par Saison? = Oui` ; sinon la colonne Summer reprend le résultat All Seasons.
- Format du résultat : `1:Caisse Super Small_2:Caisse V2_1:Caisse V3`.
- Erreurs écrites dans les cellules résultat : produit inconnu, dimensions manquantes, article trop grand pour toutes les caisses.

## Fichiers

- `packaging_engine_mms_v2.xlsm` – le classeur livré.
- `vba/modPackaging.bas` – le code VBA (importable dans n'importe quel classeur : VBE → Fichier → Importer).
- `tools/build_xlsx.py` – ajoute les feuilles au classeur d'origine (openpyxl).
- `tools/make_vbaproject.py` – génère `vbaProject.bin` (format MS-OVBA) à partir du `.bas`, sans Excel.
- `tools/finalize_xlsm.py` – assemble le `.xlsm` (VBA + boutons de formulaire).
- `tools/packer_prototype.py` – prototype Python de l'algorithme (même logique que le VBA), utile pour tester : `python3 tools/packer_prototype.py source/packaging_engine_mms_v1.xlsx`.
- `source/packaging_engine_mms_v1.xlsx` – classeur d'origine.

Regénérer le classeur :

```bash
python3 tools/build_xlsx.py source/packaging_engine_mms_v1.xlsx /tmp/v2.xlsx
python3 tools/make_vbaproject.py /tmp/vbaProject.bin vba/modPackaging.bas
python3 tools/finalize_xlsm.py /tmp/v2.xlsx /tmp/vbaProject.bin packaging_engine_mms_v2.xlsm
```
