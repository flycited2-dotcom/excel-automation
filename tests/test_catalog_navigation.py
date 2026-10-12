from collections import Counter
from pathlib import Path
import sys
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import openpyxl
import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from catalog_navigation import SECTIONS, build_navigation, link_navigation_images, section_for
from transform import build_pricelist
from ready_price import capture


def test_category_menu_preserves_goods_prices_and_outbox(tmp_path):
    groups = ['Классические сплит-системы', 'Расходные материалы — Медь',
              'Холодильники', 'Электрочайники', 'Телевизоры 46" и более',
              'Инверторные генераторы', 'Стабилизаторы напряжения',
              'Мотоблоки и садовая техника', 'Водонагреватели накопительные',
              'Пылесосы вертикальные', 'Фены', 'Сковороды', 'Маршрутизаторы', 'Неизвестная группа']
    df = pd.DataFrame([dict(article=str(i), group=g, brand='Brand', name=f'Item {i}',
                           price_in=100+i, price_out=100+i, source='supplier')
                       for i, g in enumerate(groups)])
    wb = openpyxl.Workbook()
    build_pricelist(df, wb, {}, {})
    links = build_navigation(df, wb, {}, {}, {})
    path = tmp_path / 'БытТехОпт_20261012.xlsx'
    wb.save(path)
    link_navigation_images(path, links)
    result = openpyxl.load_workbook(path)
    assert result.active.title == 'Главная'
    assert len(result.active._images) == len(SECTIONS)
    def goods(sheet):
        return [tuple(r[:6]) for r in sheet.values if isinstance(r[0], int)]
    combined = [r[1:] for title, _, _ in SECTIONS for r in goods(result[title])]
    assert Counter(combined) == Counter(r[1:] for r in goods(result['Прайс клиента']))
    for title, _, _ in SECTIONS:
        assert len(goods(result[title])) == 1
        assert result[title]['H2'].hyperlink.target == "#'Главная'!A1"
    with ZipFile(path) as z:
        rels = ET.fromstring(z.read('xl/drawings/_rels/drawing1.xml.rels'))
        targets = [r.attrib['Target'] for r in rels if r.attrib['Type'].endswith('/hyperlink')]
        assert targets == links
        drawing = ET.fromstring(z.read('xl/drawings/drawing1.xml'))
        assert len(drawing.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}hlinkClick')) == len(SECTIONS)
    source = tmp_path / 'supplier.xls'
    source.write_bytes(b'supplier input')
    manifest = capture(path, input_path=source, outbox=tmp_path / 'outbox')
    assert (manifest.parent / 'price.xlsx').read_bytes() == path.read_bytes()


@pytest.mark.parametrize('group, expected', [
    ('Инверторные генераторы', 'Генераторы'),
    ('Телевизоры 27"- 32"', 'Телевизоры и ТВ'),
    ('Подставки, крепления для ТВ', 'Телевизоры и ТВ'),
    ('Сушилки для овощей и фруктов', 'Мелкая бытовая техника'),
    ('Инверторные стиральные машины', 'Крупная бытовая техника'),
    ('Аксессуары к климатической технике', 'Расходники'),
    ('Аксессуары к товарам красота и здоровье', 'Красота и здоровье'),
    ('Чайники и заварочные чайники', 'Посуда и товары для дома'),
    ('Электрочайники', 'Мелкая бытовая техника'),
    ('Полупромышленные — Кассетные — Инверторные', 'Кондиционеры'),
    ('Новая неизвестная группа', 'Прочие товары'),
])
def test_category_classification(group, expected):
    assert section_for(group) == expected


def test_empty_sections_are_not_shown():
    df = pd.DataFrame([dict(article='1', group='Классические сплит-системы',
                           brand='Test', name='AC 07', price_out=100)])
    wb = openpyxl.Workbook()
    build_pricelist(df, wb, {}, {})
    links = build_navigation(df, wb, {}, {}, {})
    assert wb.sheetnames == ['Главная', 'Прайс клиента', 'Кондиционеры']
    assert links == ["#'Кондиционеры'!A1"]
