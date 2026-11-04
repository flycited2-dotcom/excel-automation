import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from transform import apply_pricing, load_markups


def test_homeline_uses_group_markup_not_fixed_price():
    """Фикс-цены HOMELINE отключены: цена считается наценкой от закупки."""
    config = {'markups_file': 'наценки по группам товара.xlsx'}
    group_markups = load_markups(config)
    # Раздел фикс-цен (после строки «Артикул») не должен попадать в наценки.
    assert 'UT-00244222' not in group_markups
    df = pd.DataFrame([{'article': 'UT-00244222', 'brand': 'HOMELINE',
                        'group': 'Духовки', 'price_in': 5000.0}])
    markup = group_markups.get('Духовки', 10)
    price = apply_pricing(df, group_markups, 10).price_out.iloc[0]
    assert price != 6390
    assert price == round(5000 * (1 + markup / 100))
