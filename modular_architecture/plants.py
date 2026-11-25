# plants.py — Brazil plant registry
#
# One PlantSpec per template file; one Job per line block in that template.
#   Job(sheet, cons, header_text, source, db_line, ...)
#     sheet       template sheet ("PC", "TC", "Extruded", "Other", "Stax", ...)
#     cons        block index on that sheet (0 = first block, columns B..I)
#     header_text line name exactly as in the template dropdown
#     source      downtime export: "APP", "ITA" or "MES"
#     db_line     'Linha' as it appears in the exports
#     type_value  line type (row 7) for Default-layout sheets
#     te_target   TE % target cell
#     fryer       PC only: (wash water temp, moisture aim, surface water)
#     skus        product table, top to bottom. Each SKU lists the
#                 (Corte, Sabor) pairs from the production export that feed it.
#
# Bootstrapped from the OCT 2026 goldens by analysis/build_plants_from_goldens.py.
# From here on, edit this file by hand: add a SKU row when a new product
# appears (the app warns "SKU ... is not in the SKU master").
from __future__ import annotations

import os as _os
import unicodedata as _ud
from typing import Dict

from engine import Job, PlantSpec, SKU


# ---------------------------------------------------------------------------
# AMACOCO_PETROLINA  (Snk_TE_Amacoco Petrolina (TE)_P11 2026 OCT(P10 UK).xlsm)
# ---------------------------------------------------------------------------
AMACOCO_PETROLINA = PlantSpec(
    key='AMACOCO_PETROLINA',
    jobs=[
        # TBA19 (APP)
        Job('Other', 0, '1 - TBA 19', 'APP', 'TBA19', type_value='Other', te_target=0.67, skus=(
            SKU('AMACOCO 330 ML', ptype='Dairy', seasoning=0, rated=2020, db_keys=()),
        )),
        # TBA21 (APP)
        Job('Other', 1, '2 - TBA 21', 'APP', 'TBA21', type_value='Other', te_target=0.65, skus=(
            SKU('AGUA DE COCO KEROCOCO 12x1000 ML', ptype='Dairy', seasoning=0, rated=7140, db_keys=()),
        )),
        # TBA22 (APP)
        Job('Other', 2, '3 - TBA 22', 'APP', 'TBA22', type_value='Other', te_target=0.68, skus=(
            SKU('AGUA DE COCO KEROCOCO 200 ML', ptype='Dairy', seasoning=0, rated=4080, db_keys=()),
        )),
    ],
)

# ---------------------------------------------------------------------------
# CIC  (Snk_TE_CIC_P11 2026 OCT(P10 UK).xlsm)
# ---------------------------------------------------------------------------
CIC = PlantSpec(
    key='CIC',
    jobs=[
        # PC60 (APP)
        Job('PC', 0, '1 - PC60', 'APP', 'PC60', te_target=0.669444, fryer=(20, 0.015, 0.08), skus=(
            SKU('RUFFLES SAL', ptype='Wavy', seasoning=0.0121, oil=0, col_f=0, db_keys=(('CORTE ACANALADO', 'SAL'),)),
            SKU('RUFFLES CEBOLA E SALSA', ptype='Wavy', seasoning=0.048, oil=0, col_f=0, db_keys=(('CORTE ACANALADO', 'CEBOLA E SALSA'),)),
            SKU('SENSAÇÕES FRANGO GRELHADO', ptype='Flats', seasoning=0.05, oil=0, col_f=0, db_keys=(('SENSACOES', 'FRANGO GRELHADO'), ('SENSACOES II', 'FRANGO GRELHADO'))),
            SKU('LAYS SOUR CREAM', ptype='Flats', seasoning=0.06, oil=0, col_f=0, db_keys=(('LAYS LISA', 'CREMA AGRIA'),)),
            SKU('LAYS CLASSICA', ptype='Flats', seasoning=0.0121, oil=0, col_f=0, db_keys=(('LAYS LISA', 'SAL'),)),
            SKU('NA MESA', ptype='Shoe String', seasoning=0.009, oil=0, col_f=2830.296, db_keys=(('PALHA', 'SAL'),)),
            SKU('NA MESA EXTRA FINA', ptype='Shoe String', seasoning=0.009, oil=0, col_f=2830.296, db_keys=(('PALHA EXTRA FINA', 'SAL'),)),
            SKU('RUFFLES CHURRASCO', ptype='Wavy', seasoning=0.0513, oil=0, col_f=0, db_keys=(('CORTE ACANALADO', 'CHURRASCO'),)),
            SKU('SENSAÇÕES PEITO DE PERU', ptype='Flats', seasoning=0.045, oil=0, col_f=0, db_keys=(('SENSACOES II', 'PEITO DE PERU'), ('SENSACOES', 'PEITO DE PERU'))),
            SKU('LAYS VINEGAR', ptype='Flats', seasoning=0.049, oil=0, col_f=0, db_keys=(('LAYS LISA', 'VINAGRE'),)),
        )),
        # DORITOS_TC2 (APP)
        Job('TC', 0, '2 - TC', 'APP', 'DORITOS_TC2', te_target=0.809444, skus=(
            SKU('DORITOS NACHO', ptype='Doritos - 17.6 gr', seasoning=0.093, oil=0.048, col_f=0.048, db_keys=(('DORITOS', 'QUESO'),)),
            SKU('TOSTITOS SAL', ptype='Other', seasoning=0.004, oil=0, col_f=1447.188, db_keys=(('TOSTITOS REDONDO', 'SAL'),)),
            SKU('TOSTITOS PIMENTA MEXICANA', ptype='Other', seasoning=0.06, oil=0.048, col_f=1447.188, db_keys=(('TOSTITOS REDONDO', 'PIMIENTA/PIMENTA'),)),
        )),
        # CLEXTRAL III (APP)
        Job('Extruded', 0, '4 - Clextral', 'APP', 'CLEXTRAL III', type_value='Baked', te_target=0.785444, skus=(
            SKU('CEBOLITOS', ptype='MD', seasoning=0.3331, rated=960.336, db_keys=(('CEBOLITOS', 'CEBOLLA'),)),
            SKU('CHEETOS BOLA', ptype='HD', seasoning=0.3, rated=770, db_keys=(('CHEETOS BOLA', 'QUESO'),)),
            SKU('FANDANGOS PRESUNTO', ptype='HD', seasoning=0.1837, rated=1175.472, db_keys=(('FANDANGOS', 'PRESUNTO'),)),
            SKU('CHEETOS ONDA', ptype='MD', seasoning=0.2798, rated=1130.714, db_keys=(('CHEETOS ONDA', 'REQUEIJAO'),)),
            SKU('CHEETOS LUA', ptype='MD', seasoning=0.2994, rated=1029.882, db_keys=(('CHEETOS LUA', 'QUESO PARMESANO'),)),
            SKU('FANDANGOS QUEIJO', ptype='HD', seasoning=0.1805, rated=1180.08, db_keys=(('FANDANGOS', 'QUESO'),)),
            SKU('FANDANGOS CHURRASCO', ptype='HD', seasoning=0.1738, rated=1189.728, db_keys=(('FANDANGOS', 'CHURRASCO'),)),
        )),
        # CLEXTRAL VII (APP)
        Job('Extruded', 1, '8 - Clextral 7', 'APP', 'CLEXTRAL VII', type_value='Baked', te_target=0.840444, skus=(
            SKU('CEBOLITOS', ptype='MD', seasoning=0.3331, rated=960.336, db_keys=(('CEBOLITOS', 'CEBOLLA'),)),
            SKU('FANDANGOS QUEIJO', ptype='MD', seasoning=0.1805, rated=1180.08, db_keys=(('FANDANGOS', 'QUESO'),)),
            SKU('FANDANGOS PRESUNTO', ptype='HD', seasoning=0.1837, rated=1175.472, db_keys=(('FANDANGOS', 'PRESUNTO'),)),
            SKU('CHEETOS LUA', ptype='MD', seasoning=0.2994, rated=1030, db_keys=(('CHEETOS LUA', 'QUESO PARMESANO'),)),
            SKU('CHEETOS ONDA', ptype='MD', seasoning=0.2798, rated=1130.714, db_keys=(('CHEETOS ONDA', 'REQUEIJAO'),)),
            SKU('FANDANGOS CHURRASCO', ptype='HD', seasoning=0.1738, rated=1189.728, db_keys=(('FANDANGOS', 'CHURRASCO'),)),
            SKU('CHEETOS BOLA', ptype='HD', seasoning=0.3, rated=770, db_keys=(('CHEETOS BOLA', 'QUESO'),)),
        )),
        # STIKSY (APP)
        Job('Other', 0, '3 - Stiksy', 'APP', 'STIKSY', type_value='Fried Wheat', te_target=0.847444, skus=(
            SKU('STIKSY', ptype='PRETZEL', seasoning=0.041, rated=604.17, db_keys=(('STIKSY', 'SAL'),)),
        )),
        # PINGO CIC (APP)
        Job('Other', 1, '5 - Pingo', 'APP', 'PINGO CIC', type_value='Fried Wheat', te_target=0.835444, skus=(
            SKU('PINGO BACON', ptype='HD', seasoning=0.0298, rated=630.63, db_keys=(('PINGO REDONDO', 'BACON/TOCINO'),)),
            SKU('PINGO PICANHA', ptype='HD', seasoning=0.04, rated=528, db_keys=(('PINGO QUADRADO', 'PICANHA'),)),
        )),
        # BACON CIC (APP)
        Job('Other', 2, '6 - Bacon', 'APP', 'BACON CIC', type_value='Pellets', te_target=0.810444, skus=(
            SKU('BACONZITOS', ptype='PELLETS', seasoning=0.06, rated=564, db_keys=(('BACONZITOS', 'BACON/TOCINO'),)),
        )),
        # POPCORNERS (APP)
        Job('Other', 3, '7 - Popcorners', 'APP', 'POPCORNERS', type_value='Popcorn', te_target=0.660444, skus=(
            SKU('POPCORNERS CHEDDAR', ptype='POPCORNERS', seasoning=0.1485, col_f=0.1556, rated=104.5, db_keys=(('POPCORN', 'QUESO'),)),
            SKU('POPCORNERS KETTLE', ptype='POPCORNERS', seasoning=0.1098, col_f=0.1445, rated=104.5, db_keys=(('POPCORN', 'KETTLE CORN'),)),
            SKU('POPCORNERS SAL', ptype='POPCORNERS', seasoning=0.013, col_f=0.08, rated=118, db_keys=(('POPCORN', 'SAL'),)),
        )),
    ],
)

# ---------------------------------------------------------------------------
# FEIRA_DE_SANTANA  (Snk_TE_Feira de Santana_P11 2026 OCT(P10 UK).xlsm)
# ---------------------------------------------------------------------------
FEIRA_DE_SANTANA = PlantSpec(
    key='FEIRA_DE_SANTANA',
    jobs=[
        # BOSCH (APP)
        Job('Other', 0, '3 - BOSH', 'APP', 'BOSCH', type_value='Other', te_target=0.74, skus=(
            SKU('TODDY SACHET 700G', ptype='Dairy', seasoning=0, rated=1680, db_keys=(('CHOCOLATE 700', 'CHOCOLATE'),)),
            SKU('TODDY SACHET 300G', ptype='Dairy', seasoning=0, rated=990, db_keys=(('CHOCOLATE 300', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 1,02KGX12', ptype='Dairy', seasoning=0, rated=2203, db_keys=(('CHOCOLATE 1020', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 560GRX16', ptype='Dairy', seasoning=0, rated=1152, db_keys=(('CHOCOLATE 560', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 480GRX20', ptype='Dairy', seasoning=0, rated=1152, db_keys=(('CHOCOLATE 480', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 620GRX12', ptype='Dairy', seasoning=0, rated=1488, db_keys=(('CHOCOLATE 620', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 175GRX42', ptype='Dairy', seasoning=0, rated=630, db_keys=(('CHOCOLATE 175', 'CHOCOLATE'),)),
            SKU('TODDY LIGHT 280G', ptype='Dairy', seasoning=0, rated=924, db_keys=(('CHOCOLATE 280', 'CHOCOLATE LIGHT'),)),
        )),
        # PLF (APP)
        Job('Other', 1, '1 - PLF', 'APP', 'PLF', type_value='Other', te_target=0.753, skus=(
            SKU('TODDY LIGHT 24X380', ptype='Dairy', seasoning=0, rated=3876, db_keys=(('TODDY LIGHT', 'CHOCOLATE 380'),)),
            SKU('TODDY ORIGINAL 750GRX12', ptype='Dairy', seasoning=0, rated=6750, db_keys=(('CHOCOLATE 800', 'CHOCOLATE'),)),
            SKU('TODDY ORIG POTE 370GX24', ptype='Dairy', seasoning=0, rated=3774, db_keys=(('CHOCOLATE 400', 'CHOCOLATE'),)),
        )),
        # MASIPACK (APP)
        Job('Other', 2, '2 - MASIPACK 1', 'APP', 'MASIPACK', type_value='Other', te_target=0.75, skus=(
            SKU('TODDY SACHET 700G', ptype='Dairy', seasoning=0, rated=1680, db_keys=(('CHOCOLATE 700', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 1,02KGX12', ptype='Dairy', seasoning=0, rated=1530, db_keys=(('CHOCOLATE 1020', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 620GRX12', ptype='Dairy', seasoning=0, rated=1488, db_keys=(('CHOCOLATE 620', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 480GRX20', ptype='Dairy', seasoning=0, rated=1152, db_keys=(('CHOCOLATE 480', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 300GRX24', ptype='Dairy', seasoning=0, rated=990, db_keys=(('CHOCOLATE 300', 'CHOCOLATE'),)),
            SKU('TODDY LIGHT 280G', ptype='Dairy', seasoning=0, rated=924, db_keys=(('CHOCOLATE 280', 'CHOCOLATE LIGHT'),)),
            SKU('TODDY ORIGINAL SACHET 560GRX16', ptype='Dairy', seasoning=0, rated=1152, db_keys=(('CHOCOLATE 560', 'CHOCOLATE'),)),
        )),
        # SIG 1 (APP)
        Job('Other', 3, '4 - SIG 1', 'APP', 'SIG 1', type_value='Other', te_target=0.655, skus=(
            SKU('TODDYNHO 9X3X200', ptype='Dairy', seasoning=0, rated=5040, db_keys=(('CARTONADO', '3 PACK CHOCOLATE'),)),
            SKU('TODDYNHO LEVINHO 200ML', ptype='Dairy', seasoning=0, rated=4992, db_keys=(('CARTONADO', '3 PACK LIGHT'),)),
        )),
        # SIG 2 (APP)
        Job('Other', 4, '5 - SIG 2', 'APP', 'SIG 2', type_value='Other', te_target=0.666, skus=(
            SKU('TODDYNHO 9X3X200', ptype='Dairy', seasoning=0, rated=5040, db_keys=(('CARTONADO', '3 PACK CHOCOLATE'),)),
            SKU('TODDYNHO LEVINHO 200ML', ptype='Dairy', seasoning=0, rated=4992, db_keys=(('CARTONADO', '3 PACK LIGHT'),)),
        )),
    ],
)

# ---------------------------------------------------------------------------
# ITU  (Snk_TE_ITU_P11 2026 OCT(P10 UK).xlsm)
# ---------------------------------------------------------------------------
ITU = PlantSpec(
    key='ITU',
    jobs=[
        # PC50 (MES)
        Job('PC', 0, '1 - PC 50', 'MES', 'PC50', te_target=0.713, fryer=(20, 0.015, 0.08), skus=(
            SKU('RUFFLES SAL', ptype='Wavy', seasoning=0.0121, oil=0, col_f=0, db_keys=(('CORTE ACANALADO', 'SAL'),)),
            SKU('RUFFLES PIMENTA MEXICANA', ptype='Wavy', seasoning=0.07, oil=0, col_f=0, db_keys=(('CORTE ACANALADO', 'PIMIENTA/PIMENTA'),)),
            SKU('LAYS SOUR CREAM', ptype='Flats', seasoning=0.06, oil=0, col_f=0, db_keys=(('LAYS LISA', 'CREMA AGRIA'),)),
            SKU('LAYS CLASSICA', ptype='Flats', seasoning=0.0121, oil=0, col_f=0, db_keys=(('LAYS LISA', 'SAL'),)),
            SKU('NA MESA', ptype='Shoe String', seasoning=0.009, oil=0, col_f=2636.06, db_keys=(('PALHA', 'SAL'),)),
            SKU('SENSAÇÕES PEITO DE PERU', ptype='Flats', seasoning=0.045, oil=0, col_f=0, db_keys=(('SENSACOES', 'PEITO DE PERU'), ('SENSACOES II', 'PEITO DE PERU'))),
            SKU('LAYS VINEGAR', ptype='Flats', seasoning=0.049, oil=0, col_f=0, db_keys=(('LAYS LISA', 'VINAGRE'),)),
            SKU('LAYS PICANHA', ptype='Flats', seasoning=0.07, oil=0, col_f=0, db_keys=(('LAYS LISA', 'PICANHA'),)),
            SKU('LAYS TACO MEXICANO', ptype='Flats', seasoning=0.07, oil=0, col_f=0, db_keys=()),
            SKU('LAYS QUEIJO', ptype='Flats', seasoning=0.07, oil=0, col_f=0, db_keys=(('LAYS LISA', 'QUESO'),)),
            SKU('RUFFLES CEBOLA E SALSA', ptype='Wavy', seasoning=0.048, oil=0, col_f=0, db_keys=(('CORTE ACANALADO', 'CEBOLA E SALSA'),)),
        )),
        # KETTLE (APP)
        Job('PC', 1, '8 - KETTLE ITU', 'APP', 'KETTLE', te_target=0.8, skus=(
            SKU('LAYS RUSTICA SAL', ptype='Other', seasoning=0.011, oil=0, col_f=158.24, db_keys=()),
            SKU('LAYS RUSTICA CREAM CHEESE', ptype='Other', seasoning=0.06, oil=0, col_f=150.4, db_keys=()),
        )),
        # DORITOS_TC1 (MES)
        Job('TC', 0, '2 - TC', 'MES', 'DORITOS_TC1', te_target=0.799, skus=(
            SKU('DORITOS NACHO', ptype='Doritos - 17.6 gr', seasoning=0.093, oil=0.048, col_f=0.048, db_keys=(('DORITOS', 'QUESO'),)),
            SKU('DORITOS PIZZA', ptype='Doritos - 17.6 gr', seasoning=0.087, oil=0.048, col_f=0.048, db_keys=(('DORITOS', 'PIZZA'),)),
            SKU('DORITOS KETCHUP', ptype='Doritos - 17.6 gr', seasoning=0.085, oil=0.048, col_f=0.048, db_keys=(('DORITOS', 'KETCHUP'),)),
            SKU('DORITOS SPICY NACHO', ptype='Doritos - 17.6 gr', seasoning=0.08, oil=0.048, col_f=0.048, db_keys=(('DORITOS', 'SPICY'),)),
        )),
        # CLEXTRAL I (MES)
        Job('Extruded', 0, '3 - Clextral 1', 'MES', 'CLEXTRAL I', type_value='Baked', te_target=0.849, skus=(
            SKU('CEBOLITOS', ptype='MD', seasoning=0.3331, rated=960.336, db_keys=(('CEBOLITOS', 'CEBOLLA'),)),
            SKU('CHEETOS BOLA', ptype='HD', seasoning=0.3, rated=770, db_keys=(('CHEETOS BOLA', 'QUESO'),)),
            SKU('CHEETOS LUA', ptype='HD', seasoning=0.2994, rated=1029.882, db_keys=(('CHEETOS LUA', 'QUESO PARMESANO'),)),
            SKU('INSUMOS LUA', ptype='HD', seasoning=0.2994, rated=945.81, db_keys=()),
            SKU('CHEETOS ONDA', ptype='HD', seasoning=0.2798, rated=1130.714, db_keys=(('CHEETOS ONDA', 'REQUEIJAO'),)),
            SKU('FANDANGOS PRESUNTO', ptype='HD', seasoning=0.1837, rated=1175.472, db_keys=(('FANDANGOS', 'PRESUNTO'),)),
            SKU('FANDANGOS QUEIJO', ptype='HD', seasoning=0.1805, rated=1180.08, db_keys=(('FANDANGOS', 'QUESO'),)),
            SKU('INSUMOS PATINHA', ptype='HD', seasoning=0.37, rated=850.5, db_keys=(('CHEETOS PATINHA', 'QUESO'),)),
            SKU('INSUMOS TRIANGULO', ptype='HD', seasoning=0.37, rated=850.5, db_keys=(('CHEETOS TRIANGULO', 'QUESO'),)),
        )),
        # CLEXTRAL II (MES)
        Job('Extruded', 1, '4 - Clextral 2', 'MES', 'CLEXTRAL II', type_value='Baked', te_target=0.785, skus=(
            SKU('CEBOLITOS', ptype='MD', seasoning=0.3331, rated=960.336, db_keys=(('CEBOLITOS', 'CEBOLLA'),)),
            SKU('CHEETOS BOLA', ptype='HD', seasoning=0.3, rated=770, db_keys=(('CHEETOS BOLA', 'QUESO'),)),
            SKU('CHEETOS LUA', ptype='HD', seasoning=0.2994, rated=1029.882, db_keys=(('CHEETOS LUA', 'QUESO PARMESANO'),)),
            SKU('CHEETOS ONDA', ptype='HD', seasoning=0.2798, rated=1130.714, db_keys=(('CHEETOS ONDA', 'REQUEIJAO'),)),
            SKU('FANDANGOS PRESUNTO', ptype='HD', seasoning=0.1837, rated=1175.472, db_keys=(('FANDANGOS', 'PRESUNTO'),)),
            SKU('FANDANGOS QUEIJO', ptype='HD', seasoning=0.1805, rated=1180.08, db_keys=(('FANDANGOS', 'QUESO'),)),
            SKU('INSUMOS LUA', ptype='HD', seasoning=0.2994, rated=945.81, db_keys=()),
            SKU('INSUMOS PATINHA', ptype='HD', seasoning=0.37, rated=850.5, db_keys=(('CHEETOS PATINHA', 'QUESO'),)),
            SKU('INSUMOS TRIANGULO', ptype='HD', seasoning=0.37, rated=850.5, db_keys=(('CHEETOS TRIANGULO', 'QUESO'),)),
        )),
        # CRUNCHY ITU (APP)
        Job('Other', 0, '9 - Crunchy ITU', 'APP', 'CRUNCHY ITU', type_value='Fritos', te_target=0.81, skus=(
            SKU('CRUNCHY WHITE CHEESE', ptype='HD', seasoning=0.21, rated=592.5, db_keys=(('CHEETOS KRUNCHY', 'QUESO'),)),
            SKU('CRUNCHY MELLOW CHEDDAR', ptype='HD', seasoning=0.228, rated=579, db_keys=()),
            SKU('CRUNCHY PIMENTA MEXICANA', ptype='HD', seasoning=0.251, rated=561.75, db_keys=(('CHEETOS KRUNCHY', 'PIMIENTA/PIMENTA'),)),
            SKU('BACONZITOS', ptype='HD', seasoning=0.06, rated=517, db_keys=(('BACONZITOS', 'BACON/TOCINO'),)),
        )),
        # PINGO ITU (APP)
        Job('Other', 1, '5 - Pingo', 'APP', 'PINGO ITU', type_value='Fried Wheat', te_target=0.86, skus=(
            SKU('PINGO BACON', ptype='HD', seasoning=0.0198, rated=539.11, db_keys=(('PINGO REDONDO', 'BACON/TOCINO'),)),
            SKU('PINGO PICANHA', ptype='HD', seasoning=0.04, rated=432, db_keys=(('PINGO QUADRADO', 'PICANHA'),)),
        )),
    ],
)

# ---------------------------------------------------------------------------
# LUCKY  (Snk_TE_Lucky_P11 2026 OCT(P10 UK).xlsm)
# ---------------------------------------------------------------------------
LUCKY = PlantSpec(
    key='LUCKY',
    jobs=[
        # FOFURA (ITA)
        Job('Extruded', 0, '3 - FOFURA', 'ITA', 'FOFURA', type_value='Baked', te_target=0.857, skus=(
            SKU('FOFURA CEBOLA', ptype='MD', rated=2160, db_keys=(('ANEL', 'CEBOLLA'),)),
            SKU('FOFURA CHURRASCO', ptype='MD', rated=2400, db_keys=(('ANEL', 'CHURRASCO'),)),
            SKU('FOFURA REQUEIJÃO', ptype='MD', rated=2400, db_keys=(('ANEL', 'REQUEIJAO'),)),
            SKU('FOFURA PRESUNTO', ptype='MD', rated=2400, db_keys=(('ANEL', 'PRESUNTO'),)),
            SKU('FOFURA QUEIJO', ptype='MD', rated=2160, db_keys=(('ANEL', 'QUESO'),)),
            SKU('SKINY NATURAL', ptype='MD', rated=2100, db_keys=(('SKYNI', 'NATURAL'),)),
            SKU('SKINY QUEIJO', ptype='MD', rated=2100, db_keys=(('SKYNI', 'QUESO'),)),
            SKU('SKINY PRESUNTO', ptype='MD', rated=2100, db_keys=(('SKYNI', 'PRESUNTO'),)),
        )),
        # TORCIDA (ITA)
        Job('Other', 0, '2 - TORCIDA', 'ITA', 'TORCIDA', type_value='Fried Wheat', te_target=0.85, skus=(
            SKU('CEBOLA', ptype='HD', seasoning=0, rated=4650, db_keys=(('TORCIDA TOR', 'CEBOLLA'),)),
            SKU('BACON', ptype='HD', seasoning=0, rated=4200, db_keys=(('TORCIDA TOR', 'BACON/TOCINO'),)),
            SKU('PIMENTA MEXICANA', ptype='HD', seasoning=0, rated=4650, db_keys=(('TORCIDA TOR', 'PIMIENTA/PIMENTA'),)),
            SKU('QUEIJO', ptype='HD', seasoning=0, rated=4650, db_keys=(('TORCIDA TOR', 'QUESO'),)),
            SKU('COSTELINHA', ptype='HD', seasoning=0, rated=4650, db_keys=(('TORCIDA TOR', 'COSTILLAR'),)),
            SKU('CHURRASCO', ptype='HD', seasoning=0, rated=4650, db_keys=(('TORCIDA TOR', 'CHURRASCO'),)),
            SKU('CALABRESA', ptype='HD', seasoning=0, rated=4650, db_keys=()),
            SKU('TORCIDA CAMARÃO', ptype='HD', seasoning=0, rated=4650, db_keys=(('TORCIDA TOR', 'CAMARON PIMIENTA'),)),
            SKU('PAO DE ALHO', ptype='HD', seasoning=0, rated=4650, db_keys=(('TORCIDA TOR', 'AJO'),)),
            SKU('VINAGRETE', ptype='HD', seasoning=0, rated=4650, db_keys=(('TORCIDA TOR', 'VINAGRETE'),)),
            SKU('TORCIDA MIX', ptype='HD', seasoning=0, rated=4650, db_keys=()),
        )),
    ],
)

# ---------------------------------------------------------------------------
# NE_RECIFE  (Snk_TE_NE Recife_P11 2026 OCT(P10 UK).xlsm)
# ---------------------------------------------------------------------------
NE_RECIFE = PlantSpec(
    key='NE_RECIFE',
    jobs=[
        # CLEXTRAL VI (APP)
        Job('Extruded', 0, '2 - Clextral', 'APP', 'CLEXTRAL VI', type_value='Baked', te_target=0.79, skus=(
            SKU('CEBOLITOS', ptype='MD', seasoning=0.3331, rated=960.336, db_keys=(('CEBOLITOS', 'CEBOLLA'),)),
            SKU('CHEETOS BOLA', ptype='HD', seasoning=0.3, rated=770, db_keys=(('CHEETOS BOLA', 'QUESO'),)),
            SKU('CHEETOS LUA', ptype='MD', seasoning=0.2994, rated=1029.882, db_keys=(('CHEETOS LUA', 'QUESO PARMESANO'),)),
            SKU('CHEETOS ONDA', ptype='MD', seasoning=0.2798, rated=1130.714, db_keys=(('CHEETOS ONDA', 'REQUEIJAO'),)),
            SKU('FANDANGOS PRESUNTO', ptype='MD', seasoning=0.1837, rated=1175.472, db_keys=(('FANDANGOS', 'PRESUNTO'),)),
            SKU('FANDANGOS QUEIJO', ptype='MD', seasoning=0.1805, rated=1180.08, db_keys=(('FANDANGOS', 'QUESO'),)),
            SKU('FANDANGOS CHURRASCO', ptype='MD', seasoning=0.1738, rated=1189.728, db_keys=(('FANDANGOS', 'CHURRASCO'),)),
            SKU('FANDANGOS BACON', ptype='HD', seasoning=0.1963, rated=1157.328, db_keys=(('FANDANGOS', 'BACON/TOCINO'),)),
        )),
        # SCHAAF_1 (APP)
        Job('Extruded', 1, '1 - SCHAAF', 'APP', 'SCHAAF_1', type_value='Baked', te_target=0.8, skus=(
            SKU('CEBOLITOS', ptype='MD', seasoning=0.3331, rated=500.175, db_keys=(('CEBOLITOS', 'CEBOLLA'),)),
            SKU('CHEETOS LUA', ptype='HD', seasoning=0.2994, rated=490.42, db_keys=(('CHEETOS LUA', 'QUESO PARMESANO'),)),
            SKU('CHEETOS ONDA', ptype='HD', seasoning=0.2798, rated=540.15, db_keys=(('CHEETOS ONDA', 'REQUEIJAO'),)),
            SKU('FANDANGOS PRESUNTO', ptype='MD', seasoning=0.1837, rated=612.225, db_keys=(('FANDANGOS', 'PRESUNTO'),)),
            SKU('FANDANGOS QUEIJO', ptype='MD', seasoning=0.1805, rated=614.625, db_keys=(('FANDANGOS', 'QUESO'),)),
            SKU('FANDANGOS BACON', ptype='HD', seasoning=0.1963, rated=602.775, db_keys=(('FANDANGOS', 'BACON/TOCINO'),)),
            SKU('FANDANGOS CHURRASCO', ptype='MD', seasoning=0.1738, rated=619.65, db_keys=(('FANDANGOS', 'CHURRASCO'),)),
        )),
        # SCHAAF_2 (APP)
        Job('Extruded', 2, '4 - SCHAAF 2', 'APP', 'SCHAAF_2', type_value='Baked', te_target=0.83, skus=(
            SKU('CEBOLITOS', ptype='MD', seasoning=0.3331, rated=500.175, db_keys=(('CEBOLITOS', 'CEBOLLA'),)),
            SKU('CHEETOS LUA', ptype='MD', seasoning=0.2994, rated=490.42, db_keys=(('CHEETOS LUA', 'QUESO PARMESANO'),)),
            SKU('CHEETOS ONDA', ptype='MD', seasoning=0.3331, rated=500.175, db_keys=(('CHEETOS ONDA', 'REQUEIJAO'),)),
            SKU('FANDANGOS PRESUNTO', ptype='MD', seasoning=0.1838, rated=612.15, db_keys=(('FANDANGOS', 'PRESUNTO'),)),
            SKU('FANDANGOS QUEIJO', ptype='MD', seasoning=0.1806, rated=614.55, db_keys=(('FANDANGOS', 'QUESO'),)),
            SKU('FANDANGOS BACON', ptype='MD', seasoning=0.1963, rated=602.775, db_keys=(('FANDANGOS', 'BACON/TOCINO'),)),
            SKU('FANDANGOS CHURRASCO', ptype='HD', seasoning=0.1738, rated=619.65, db_keys=(('FANDANGOS', 'CHURRASCO'),)),
            SKU('CHEETOS BOLA', ptype='MD', seasoning=0.3, rated=385, db_keys=(('CHEETOS BOLA', 'QUESO'),)),
        )),
        # TORCIDA REC (APP)
        Job('Other', 0, '3 - Torcida Recife', 'APP', 'TORCIDA REC', type_value='Fried Wheat', te_target=0.85, skus=(
            SKU('CHURRASCO', ptype='HD', seasoning=0.015, rated=689.5, db_keys=(('TORCIDA TOR', 'CHURRASCO'),)),
            SKU('BACON', ptype='HD', seasoning=0.015, rated=689.5, db_keys=(('TORCIDA TOR', 'BACON/TOCINO'),)),
            SKU('CALABRESA', ptype='HD', seasoning=0.016, rated=688.8, db_keys=()),
            SKU('QUEIJO', ptype='HD', seasoning=0.013, rated=690.9, db_keys=(('TORCIDA TOR', 'QUESO'),)),
            SKU('PIMENTA MEXICANA', ptype='HD', seasoning=0.05, rated=665, db_keys=(('TORCIDA TOR', 'PIMIENTA/PIMENTA'),)),
            SKU('CEBOLA', ptype='HD', seasoning=0.013, rated=690.9, db_keys=(('TORCIDA TOR', 'CEBOLLA'),)),
            SKU('TORCIDA CAMARÃO', ptype='HD', seasoning=0.0275, rated=680.75, db_keys=(('TORCIDA TOR', 'CAMARON PIMIENTA'),)),
            SKU('COSTELINHA', ptype='HD', seasoning=0.0285, rated=680.75, db_keys=(('TORCIDA TOR', 'COSTILLAR'),)),
        )),
    ],
)

# ---------------------------------------------------------------------------
# SETE_LAGOAS  (Snk_TE_Sete Lagoas_P11 2026 OCT(P10 UK).xlsm)
# ---------------------------------------------------------------------------
SETE_LAGOAS = PlantSpec(
    key='SETE_LAGOAS',
    jobs=[
        # PC37 (MES)
        Job('PC', 0, '1 - PC 37', 'MES', 'PC37', te_target=0.707, fryer=(20, 0.015, 0.08), skus=(
            SKU('RUFFLES CEBOLA E SALSA', ptype='Wavy', seasoning=0.048, db_keys=(('CORTE ACANALADO', 'CEBOLA E SALSA'),)),
            SKU('RUFFLES CHURRASCO', ptype='Wavy', seasoning=0.0513, db_keys=(('CORTE ACANALADO', 'CHURRASCO'),)),
            SKU('RUFFLES SAL', ptype='Wavy', seasoning=0.0121, db_keys=(('CORTE ACANALADO', 'SAL'),)),
            SKU('SENSACOES PERU', ptype='Flats', seasoning=0.045, db_keys=(('SENSACOES', 'PEITO DE PERU'), ('SENSACOES II', 'PEITO DE PERU'))),
            SKU('SENSACOES FRANGO', ptype='Flats', seasoning=0.05, db_keys=(('SENSACOES', 'FRANGO GRELHADO'), ('SENSACOES II', 'FRANGO GRELHADO'))),
            SKU('LAYS SOUR CREAM', ptype='Flats', seasoning=0.06, db_keys=(('LAYS LISA', 'CREMA AGRIA'),)),
            SKU('LAYS VINEGAR', ptype='Flats', seasoning=0.049, db_keys=(('LAYS LISA', 'VINAGRE'),)),
            SKU('LAYS CLASSICA', ptype='Flats', seasoning=0.0121, db_keys=(('LAYS LISA', 'SAL'),)),
            SKU('PALHA', ptype='Shoe String', seasoning=0.009, col_f=1763.98, db_keys=(('PALHA', 'SAL'),)),
            SKU('PALHA EXTRA FINA', ptype='Shoe String', seasoning=0.009, col_f=1763.98, db_keys=(('PALHA EXTRA FINA', 'SAL'),)),
        )),
        # Doritos 3 (MES)
        Job('TC', 0, '2 - TC', 'MES', 'Doritos 3', te_target=0.74, skus=(
            SKU('DORITOS NACHO', ptype='Doritos - 17.6 gr', seasoning=0.093, oil=0.048, col_f=0.048, db_keys=(('DORITOS', 'QUESO'),)),
            SKU('DINAMITAS FLAMIN HOT', ptype='Other', seasoning=0.095, oil=0.048, col_f=942, db_keys=(('DORITOS DINAMITA', 'FLAMIN HOT'),)),
            SKU('DINAMITAS PIMENTA MEXICANA', ptype='Other', seasoning=0.08, oil=0.048, col_f=959, db_keys=(('DORITOS DINAMITA', 'PIMIENTA/PIMENTA'),)),
            SKU('DORITOS SWEET CHILI', ptype='Doritos - 17.6 gr', seasoning=0.095, oil=0.048, col_f=0, db_keys=(('DORITOS', 'SWEET CHILI'),)),
        )),
        # Doritos 4 (MES)
        Job('TC', 1, '5 - TC-04', 'MES', 'Doritos 4', te_target=0.817, skus=(
            SKU('DORITOS NACHO', ptype='Doritos - 17.6 gr', seasoning=0.093, oil=0.048, col_f=0.048, db_keys=(('DORITOS', 'QUESO'),)),
            SKU('DORITOS SWEET CHILI', ptype='Doritos - 17.6 gr', seasoning=0.095, oil=0.048, col_f=0.048, db_keys=(('DORITOS', 'SWEET CHILI'),)),
        )),
        # CLX IV (MES)
        Job('Extruded', 0, '3 - Clextral', 'MES', 'CLX IV', type_value='Baked', te_target=0.847, skus=(
            SKU('CHEETOS ONDA', ptype='MD', seasoning=0.2798, rated=1130.714, db_keys=(('CHEETOS ONDA', 'REQUEIJAO'),)),
            SKU('CHEETOS LUA', ptype='MD', seasoning=0.2994, rated=1029.882, db_keys=(('CHEETOS LUA', 'QUESO PARMESANO'),)),
            SKU('SKINY NATURAL', ptype='HD', seasoning=0.1369, rated=1078.875, db_keys=(('SKYNI', 'NATURAL'),)),
            SKU('SKINY PRESUNTO', ptype='HD', seasoning=0.1355, rated=1080.625, db_keys=(('SKYNI', 'PRESUNTO'),)),
            SKU('SKINY QUEIJO', ptype='HD', seasoning=0.2955, rated=880.625, db_keys=(('SKYNI', 'QUESO'),)),
            SKU('CEBOLITOS', ptype='MD', seasoning=0.3331, rated=960, db_keys=(('CEBOLITOS', 'CEBOLLA'),)),
            SKU('FANDANGOS PRESUNTO', ptype='HD', seasoning=0.1837, rated=1175.472, db_keys=(('FANDANGOS', 'PRESUNTO'),)),
            SKU('FANDANGOS QUEIJO', ptype='HD', seasoning=0.1805, rated=1180.08, db_keys=(('FANDANGOS', 'QUESO'),)),
        )),
        # CLX V (MES)
        Job('Extruded', 1, '4 - Clextral 5', 'MES', 'CLX V', type_value='Baked', te_target=0.89, skus=(
            SKU('CHEETOS ONDA', ptype='MD', seasoning=0.2798, rated=1131, db_keys=(('CHEETOS ONDA', 'REQUEIJAO'),)),
            SKU('CHEETOS LUA', ptype='MD', seasoning=0.2994, rated=1030, db_keys=(('CHEETOS LUA', 'QUESO PARMESANO'),)),
            SKU('SKINY QUEIJO', ptype='HD', seasoning=0.2955, rated=881, db_keys=(('SKYNI', 'QUESO'),)),
            SKU('SKINY NATURAL', ptype='HD', seasoning=0.1369, rated=1079, db_keys=(('SKYNI', 'NATURAL'),)),
            SKU('CEBOLITOS', ptype='MD', seasoning=0.3331, rated=960, db_keys=(('CEBOLITOS', 'CEBOLLA'),)),
            SKU('FANDANGOS PRESUNTO', ptype='HD', seasoning=0.1837, rated=1175, db_keys=(('FANDANGOS', 'PRESUNTO'),)),
        )),
    ],
)

# ---------------------------------------------------------------------------
# SOROCABA  (Snk_TE_Sorocaba_P11 2026 OCT(P10 UK).xlsm)
# ---------------------------------------------------------------------------
SOROCABA = PlantSpec(
    key='SOROCABA',
    jobs=[
        # SCHAAF_3 (APP)
        Job('Extruded', 0, '3 - SCHAAF', 'APP', 'SCHAAF_3', type_value='Baked', te_target=0.8, skus=(
            SKU('CHEETOS LUA', ptype='MD', seasoning=0.2994, rated=805.69, db_keys=(('CHEETOS LUA', 'QUESO PARMESANO'),)),
            SKU('CHEETOS ONDA', ptype='MD', seasoning=0.2798, rated=900.25, db_keys=(('CHEETOS ONDA', 'REQUEIJAO'),)),
            SKU('FANDANGOS PRESUNTO', ptype='HD', seasoning=0.1837, rated=1126.494, db_keys=(('FANDANGOS', 'PRESUNTO'),)),
            SKU('FANDANGOS QUEIJO', ptype='HD', seasoning=0.1805, rated=1130.91, db_keys=(('FANDANGOS', 'QUESO'),)),
        )),
        # TODDY_2 (APP)
        Job('Other', 0, '10 - OPTIMA 1 SOR', 'APP', 'TODDY_2', type_value='Other', te_target=0.68, skus=(
            SKU('TODDY ORIGINAL 12X750', ptype='MM- POTE', seasoning=0, rated=6750, db_keys=()),
            SKU('TODDY ORIGINAL 24X370', ptype='MM- POTE', seasoning=0, rated=5772, db_keys=()),
            SKU('TODDY ORIGINAL 24X200', ptype='MM- POTE', seasoning=0, rated=2880, db_keys=()),
            SKU('TODDY ORIGINAL EXP 200', ptype='MM- POTE', seasoning=0, rated=2880, db_keys=(('TODDY 200', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL EXP 370', ptype='MM- POTE', seasoning=0, rated=5772, db_keys=(('CHOCOLATE 400', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL EXP 750', ptype='MM- POTE', seasoning=0, rated=6750, db_keys=(('CHOCOLATE 800', 'CHOCOLATE'),)),
        )),
        # TODDY_1 (APP)
        Job('Other', 1, '11 - OPTIMA 2 SOR', 'APP', 'TODDY_1', type_value='Other', te_target=0.65, skus=(
            SKU('TODDY ORIGINAL 12X750', ptype='MM- POTE', seasoning=0, rated=6750, db_keys=()),
            SKU('TODDY ORIGINAL 24X370', ptype='MM- POTE', seasoning=0, rated=5772, db_keys=()),
            SKU('TODDY ORIGINAL 24X200', ptype='MM- POTE', seasoning=0, rated=2880, db_keys=()),
            SKU('TODDY ORIGINAL EXP 200', ptype='MM- POTE', seasoning=0, rated=2880, db_keys=(('TODDY 200', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL EXP 370', ptype='MM- POTE', seasoning=0, rated=5772, db_keys=(('CHOCOLATE 400', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL EXP 750', ptype='MM- POTE', seasoning=0, rated=6750, db_keys=(('CHOCOLATE 800', 'CHOCOLATE'),)),
        )),
        # TODDY_3 (APP)
        Job('Other', 2, '12 - BOSCH SOR', 'APP', 'TODDY_3', type_value='Other', te_target=0.81, skus=(
            SKU('TODDY SACHET 700G', ptype='MM- SACHET', seasoning=0, rated=1680, db_keys=(('CHOCOLATE 700', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL 6X1800', ptype='MM- SACHET', seasoning=0, rated=3024, db_keys=(('CHOCOLATE 2000', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 1,02KGX12', ptype='MM- SACHET', seasoning=0, rated=2203, db_keys=(('CHOCOLATE 1020', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 480GRX20', ptype='MM- SACHET', seasoning=0, rated=1152, db_keys=(('CHOCOLATE 480', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 300GRX24', ptype='MM- SACHET', seasoning=0, rated=720, db_keys=(('CHOCOLATE 300', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 560GRX24', ptype='MM- SACHET', seasoning=0, rated=1152, db_keys=(('CHOCOLATE 560', 'CHOCOLATE'),)),
        )),
        # TODDY_4 (APP)
        Job('Other', 3, '13 - SYNTHEGON 9 SOR', 'APP', 'TODDY_4', type_value='Other', te_target=0.78, skus=(
            SKU('TODDY ORIGINAL 6X1800', ptype='MM- SACHET', seasoning=0, rated=4320, db_keys=(('CHOCOLATE 2000', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 480GRX20', ptype='MM- SACHET', seasoning=0, rated=1728, db_keys=(('CHOCOLATE 480', 'CHOCOLATE'),)),
            SKU('TODDY SACHET 700G', ptype='MM- SACHET', seasoning=0, rated=2520, db_keys=(('CHOCOLATE 700', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 1,02KGX12', ptype='MM- SACHET', seasoning=0, rated=3060, db_keys=(('CHOCOLATE 1020', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 300GRX24', ptype='MM- SACHET', seasoning=0, rated=1260, db_keys=(('CHOCOLATE 300', 'CHOCOLATE'),)),
            SKU('TODDY ORIGINAL SACHET 560GRX24', ptype='MM- SACHET', seasoning=0, rated=1722.93443, db_keys=(('CHOCOLATE 560', 'CHOCOLATE'),)),
        )),
        # SPEED_2 (APP)
        Job('Other', 4, '7 - L02', 'APP', 'SPEED_2', type_value='Other', te_target=0.6, skus=(
            SKU('TODDYNHO 9X3X200', ptype='Dairy', seasoning=0, rated=5040, db_keys=(('CARTONADO', '3 PACK CHOCOLATE'),)),
            SKU('TODDYNHO LEVINHO 200ML', ptype='Dairy', seasoning=0, rated=4992, db_keys=(('CARTONADO', '3 PACK LIGHT'),)),
        )),
        # SPEED_3 (APP)
        Job('Other', 5, '8 - L03', 'APP', 'SPEED_3', type_value='Other', te_target=0.59, skus=(
            SKU('TODDYNHO 9X3X200', ptype='Dairy', seasoning=0, rated=5040, db_keys=(('CARTONADO', '3 PACK CHOCOLATE'),)),
            SKU('TODDYNHO LEVINHO 200ML', ptype='Dairy', seasoning=0, rated=4992, db_keys=(('CARTONADO', '3 PACK LIGHT'),)),
        )),
        # SPEED_4 (APP)
        Job('Other', 6, '9 - L04', 'APP', 'SPEED_4', type_value='Other', te_target=0.61, skus=(
            SKU('TODDYNHO 9X3X200', ptype='Dairy', seasoning=0, rated=5040, db_keys=(('CARTONADO', '3 PACK CHOCOLATE'),)),
            SKU('TODDYNHO LEVINHO 200ML', ptype='Dairy', seasoning=0, rated=4992, db_keys=(('CARTONADO', '3 PACK LIGHT'),)),
        )),
        # CANISTER_1 (APP)
        Job('Stax', 0, '5 - Canister', 'APP', 'CANISTER_1', te_target=0.66, skus=(
            SKU('RUFFLES TUBO SAL', ptype='CANISTER', seasoning=0, db_keys=(('STAX ACANALADO', 'SAL'),)),
            SKU('RUFFLES TUBO SOUR CREAM', ptype='CANISTER', seasoning=0, db_keys=(('STAX ACANALADO', 'CEBOLLA'),)),
        )),
    ],
)


# -------------------------
# Filename matching
# -------------------------
# Keyword (lower-case, accents stripped) -> REGISTRY key. Brazilian templates
# are named like "Snk_TE_<Plant>_P11 2026 OCT(P10 UK).xlsm".
_FNAME_KEYWORDS = [
    ('amacoco', 'AMACOCO_PETROLINA'),
    ('_cic_', 'CIC'),
    ('feira de santana', 'FEIRA_DE_SANTANA'),
    ('_itu_', 'ITU'),
    ('_lucky_', 'LUCKY'),
    ('ne recife', 'NE_RECIFE'),
    ('sete lagoas', 'SETE_LAGOAS'),
    ('sorocaba', 'SOROCABA'),
]


def _norm_fname(s: str) -> str:
    """Strip diacritics and lowercase a filename string for matching."""
    s = "".join(ch for ch in _ud.normalize("NFKD", s) if not _ud.combining(ch))
    return s.lower()


def match_filename(fname: str) -> "str | None":
    """Return the REGISTRY key for a template filename, or None if unrecognised."""
    n = _norm_fname(_os.path.basename(fname))
    for kw, key in _FNAME_KEYWORDS:
        if kw in n:
            return key
    return None


def find_templates(folder: str) -> "list[tuple[str, str]]":
    """Scan *folder* for .xlsm templates whose name matches a known plant.
    Returns (plant_key, absolute_filepath) pairs sorted by filename.
    Completed/filled copies ("_complete", "_Filled") are skipped."""
    results = []
    try:
        for fname in sorted(_os.listdir(folder)):
            low = fname.lower()
            if not low.endswith(".xlsm") or low.startswith("~$"):
                continue
            if "_complete" in low or "_filled" in low:
                continue
            key = match_filename(fname)
            if key:
                results.append((key, _os.path.join(folder, fname)))
    except OSError:
        pass
    return results


# -------------------------
# Registry
# -------------------------
REGISTRY: Dict[str, PlantSpec] = {
    AMACOCO_PETROLINA.key: AMACOCO_PETROLINA,
    CIC.key: CIC,
    FEIRA_DE_SANTANA.key: FEIRA_DE_SANTANA,
    ITU.key: ITU,
    LUCKY.key: LUCKY,
    NE_RECIFE.key: NE_RECIFE,
    SETE_LAGOAS.key: SETE_LAGOAS,
    SOROCABA.key: SOROCABA,
}

