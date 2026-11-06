import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from transform import apply_pricing, drop_supplier_only_brands, load_markups, round_to_ending

RULES = {'HOMELINE': {'supplier_only': True, 'round_to': 95,
                      'markups': {'лари': 4, 'стиральн': 4, 'холодильн': 4,
                                  'духовк': 5, 'мини-печ': 5}}}


def item(group, price_in, brand='HOMELINE', **kw):
    return dict(dict(article='A', brand=brand, group=group, name=f'{brand} X', price_in=price_in), **kw)


def test_fixed_price_block_is_not_read_as_group_markups():
    group_markups = load_markups({'markups_file': 'наценки по группам товара.xlsx'})
    assert 'UT-00244222' not in group_markups


@pytest.mark.parametrize('price, expected', [
    (21015, 20995), (21044, 20995), (21046, 21095), (21095, 21095), (130, 95), (6628.4, 6595),
])
def test_round_to_nearest_95(price, expected):
    assert round_to_ending(price, 95) == expected


@pytest.mark.parametrize('group, markup', [
    ('Морозильные лари', 4), ('Стиральные машины с фронтальной загрузкой', 4),
    ('Холодильники с верхней морозильной камерой', 4), ('Духовки настольные', 5),
])
def test_homeline_brand_markup_by_category(group, markup):
    df = pd.DataFrame([item(group, 20000)])
    price = apply_pricing(df, {group: 2}, 10, RULES).price_out.iloc[0]
    assert price == round_to_ending(20000 * (1 + markup / 100), 95)


def test_homeline_unknown_category_uses_group_markup_and_rounding():
    df = pd.DataFrame([item('Чайники', 1000)])
    assert apply_pricing(df, {'Чайники': 10}, 10, RULES).price_out.iloc[0] == 1095


def test_other_brands_unchanged():
    df = pd.DataFrame([item('Морозильные лари', 20000, brand='Bosch')])
    assert apply_pricing(df, {'Морозильные лари': 2}, 10, RULES).price_out.iloc[0] == 20400


def test_tlt_homeline_rows_are_ignored():
    df = pd.DataFrame([item('Морозильные лари', 100, source='tlt'),
                       item('Морозильные лари', 100, brand='', name='Ларь HOMELINE HCF-1', source='tlt'),
                       item('Морозильные лари', 100, brand='Vitek', source='tlt'),
                       item('Морозильные лари', 100)])
    kept = drop_supplier_only_brands(df, RULES)
    assert kept.brand.tolist() == ['Vitek', 'HOMELINE']
    assert kept.get('source').fillna('').tolist() == ['tlt', '']


def test_supplier_rows_without_group_do_not_break_build(tmp_path):
    """06.11.2026: у Технолидера 4 товара пришли без группы — сборка упала."""
    import openpyxl
    from transform import read_supplier_price, MISSING_GROUP
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.append(['00-1', 'Холодильники', 'Brand', 'Холодильник A', 1000, '*', None, None])
    sheet.append(['00-2', None, 'Hotpoint', 'Встраиваемый холодильник B', 2000, None, '*', None])
    path = tmp_path / 'supplier.xlsx'
    book.save(path)
    df = read_supplier_price(path, {'read_options': {'skiprows': 0}})
    assert df.group.tolist() == ['Холодильники', MISSING_GROUP]
    priced = apply_pricing(df, {'Холодильники': 5}, 10)
    assert priced.price_out.tolist() == [1050, 2200]
    # и даже если пустая группа пришла из другого источника — не падаем
    raw = pd.DataFrame([item(float('nan'), 100, brand='X')])
    assert apply_pricing(raw, {}, 10).price_out.iloc[0] == 110
