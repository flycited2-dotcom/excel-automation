import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from splithub_catalog import catalog_frame, merge_products, parse_products
from transform import apply_pricing, build_pricelist


def product(**kwargs):
    return dict(dict(id='1', sku='test', model='TEST', brand='Brand', price=100,
                     stock='in_stock', group='inv', type='split'), **kwargs)


def test_admin_stock_price_visibility_and_custom_items():
    base = [product(), product(id='2', sku='second'), product(id='3', sku='hidden')]
    admin = {'ok': True, 'overrides': {'test': {'stock': 'out'},
             'second': {'price': 125}, 'hidden': {'active': '0'}},
             'custom_products': [product(id='4', sku='new')]}
    df = catalog_frame(merge_products(base, admin))
    assert df.article.tolist() == ['2', '4']
    assert apply_pricing(df, {}, {}, 50).price_out.tolist() == [125, 100]


@pytest.mark.parametrize('stock', ['out', 'order_7', 'days_1_2', 'wait', None])
def test_only_confirmed_stock(stock):
    assert catalog_frame([product(stock=stock)]).empty


def test_parser_does_not_execute_javascript():
    assert parse_products('var PRODUCTS = [{"id":"1"}]; alert(1)') == [{'id': '1'}]
    with pytest.raises(ValueError):
        parse_products('<html>failure</html>')
    with pytest.raises(ValueError):
        merge_products([], {'ok': False})


def test_workbook_hierarchy_and_price():
    items = [product(id=str(i), sku=str(i), group=group, type=kind)
             for i, (group, kind) in enumerate([
                 ('rashod', 'sol'), ('pac_onoff', 'duct'), ('pac_inv', 'cassette'),
                 ('inv', 'split'), ('pac_onoff', 'cassette'), ('onoff', 'split')])]
    df = catalog_frame(items)
    wb = openpyxl.Workbook()
    build_pricelist(df, wb, {'group_positions': {
        'Классические сплит-системы': 1, 'Инверторные сплит-системы': 1.3}}, {})
    headers = [row[0].value for row in wb.active if isinstance(row[0].value, str)
               and row[0].value not in ('№', 'Brand')]
    assert headers == ['Классические сплит-системы', 'Инверторные сплит-системы',
        'Полупромышленные — Кассетные — Инверторные',
        'Полупромышленные — Кассетные — ON/OFF',
        'Полупромышленные — Канальные — ON/OFF', 'Расходные материалы — Brand']
    assert all(row[4].value == 100 for row in wb.active if isinstance(row[0].value, int))


def test_unknown_available_equipment_is_not_silently_misclassified():
    with pytest.raises(ValueError):
        catalog_frame([product(group='pac_inv', type='unknown')])


def test_brand_series_headers_and_continuous_numbering():
    df = catalog_frame([product(id='1', brand='Beta', series='O2'),
                        product(id='2', brand='Alpha', series='iKON'),
                        product(id='3', brand='Alpha', series='Axiom'),
                        product(id='4', brand='Alpha', series='iKON')])
    wb = openpyxl.Workbook()
    build_pricelist(df, wb, {}, {})
    rows = list(wb.active.values)
    labels = [r[0] for r in rows if isinstance(r[0], str)]
    assert labels == ['№', 'Инверторные сплит-системы', 'Alpha', 'Серия: Axiom',
                      'Серия: iKON', 'Beta', 'Серия: O2']
    assert [r[0] for r in rows if isinstance(r[0], int)] == [1, 2, 3, 4]
    assert [r[1] for r in rows if isinstance(r[0], int)] == ['3', '2', '4', '1']
    product_rows = [r for r in wb.active if isinstance(r[0].value, int)]
    assert product_rows[0][0].fill.fgColor.rgb != product_rows[-1][0].fill.fgColor.rgb


def test_capacity_order_uses_btu_instead_of_model_code_or_price():
    sizes = ['24', '07', '60', '09', '18', '12', '55', '25', '50', '65', '20', '23']
    df = catalog_frame([product(id=str(i), series='ICE', btu=size,
                        model='ICE' + str(100-i), price=1000-i)
                        for i, size in enumerate(sizes)])
    wb = openpyxl.Workbook()
    build_pricelist(df, wb, {}, {})
    ordered = [int(r[1].value) for r in wb.active if isinstance(r[0].value, int)]
    assert [int(sizes[i]) for i in ordered] == sorted(map(int, sizes))
