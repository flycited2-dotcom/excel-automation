from collections import Counter
from pathlib import Path
import sys
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import openpyxl
import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from catalog_navigation import SECTIONS, build_navigation, section_for
from workbook_media import optimize_workbook_images
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
    optimize_workbook_images(path)
    result = openpyxl.load_workbook(path)
    assert result.active.title == 'Главная'
    assert len(result.active._images) == len(SECTIONS)
    assert result.active.freeze_panes is None  # шапка не съедает экран телефона
    def goods(sheet):
        return [tuple(r[:6]) for r in sheet.values if isinstance(r[0], int)]
    combined = [r[1:] for title, _, _ in SECTIONS for r in goods(result[title])]
    assert Counter(combined) == Counter(r[1:] for r in goods(result['Прайс клиента']))
    for title, _, _ in SECTIONS:
        assert len(goods(result[title])) == 1
        assert result[title]['H2'].hyperlink.location == "'Главная'!A1"
    assert links == [f"'{title}'!A1" for title, _, _ in SECTIONS]
    home_locations = {c.hyperlink.location for row in result['Главная'].iter_rows()
                      for c in row if c.hyperlink is not None}
    assert set(links) <= home_locations and "'Прайс клиента'!A1" in home_locations
    hint = result['Главная']['B6']
    assert hint.fill.fgColor.rgb.endswith('FFFF00') and hint.font.b  # подсказку не пропустить
    assert_mobile_friendly(path)
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
    assert links == ["'Кондиционеры'!A1"]


def assert_mobile_friendly(path):
    """Мобильные просмотрщики: только внутренние ссылки и стандартный DrawingML."""
    with ZipFile(path) as z:
        names = z.namelist()
        for name in names:
            if not name.endswith(('.xml', '.rels')):
                continue
            text = z.read(name).decode('utf-8')
            # ElementTree без register_namespace пишет ns0:/ns1: — их не понимают
            # упрощённые парсеры (Polaris/Samsung и др.).
            assert '<ns0:' not in text and ':ns0=' not in text, name
            assert 'hlinkClick' not in text and 'alphaModFix' not in text, name
            if name.endswith('.rels'):
                # Ни одной «внешней» ссылки на лист вида #'Лист'!A1.
                assert 'Target="#' not in text, name
        for name in names:
            if name.startswith('xl/drawings/') and name.endswith('.xml'):
                text = z.read(name).decode('utf-8')
                assert 'oneCellAnchor' not in text, name


def test_section_buttons_are_big_colored_and_linked():
    """Клик по картинкам на телефонах не работает — нажимают на кнопку под ней."""
    from catalog_navigation import BUTTON_COLORS
    df = pd.DataFrame([dict(article='1', group='Классические сплит-системы', brand='T',
                            name='AC', price_out=100),
                       dict(article='2', group='Сковороды', brand='T', name='Pan', price_out=50)])
    wb = openpyxl.Workbook()
    build_pricelist(df, wb, {}, {})
    build_navigation(df, wb, {}, {}, {})
    home = wb['Главная']
    for anchor, title, span in [('B17', 'Кондиционеры', 'B17:F19'),
                                ('H17', 'Посуда и товары для дома', 'H17:L19')]:
        button = home[anchor]
        assert span in {str(r) for r in home.merged_cells.ranges}  # во всю ширину картинки
        assert title in button.value and 'Открыть →' in button.value
        assert button.font.b and button.font.color.rgb.endswith('FFFFFF') and button.font.sz >= 14
        assert button.fill.fgColor.rgb[-6:] in BUTTON_COLORS
        assert button.alignment.horizontal == 'center'
        assert button.hyperlink.location == f"'{title}'!A1"
        # описание под кнопкой тоже ведёт в раздел — больше площадь нажатия
        below = home[anchor[0] + '20']
        assert below.hyperlink.location == f"'{title}'!A1"
