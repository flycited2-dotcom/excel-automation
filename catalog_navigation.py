"""Category sheets and a photo menu, generated on every price run."""
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import os
import xml.etree.ElementTree as ET

from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU

BASE = Path(__file__).parent
TITLE_COLORS = ('DBEAF6', 'DEF0E6', 'F5E4D9', 'E9E0F3', 'F3DFE8',
                'E2EDD8', 'DDEEF0', 'F3EBD6', 'E1E6F4', 'EDE3D9',
                'EEE1ED', 'E5EEDC', 'DFE8F1', 'E9E5F0')
SECTIONS = [
    ('Кондиционеры', 'climate', 'Бытовые и полупромышленные сплит-системы'),
    ('Расходники', 'supplies', 'Медь, изоляция, крепёж и материалы для монтажа'),
    ('Крупная бытовая техника', 'large', 'Холодильники, стиральные машины, плиты'),
    ('Мелкая бытовая техника', 'small', 'Техника для кухни, уборки и ухода'),
    ('Телевизоры и ТВ', 'tv', 'Телевизоры, антенны, пульты и крепления'),
    ('Генераторы', 'generators', 'Инверторные и бензиновые генераторы'),
    ('Стабилизаторы', 'stabilizers', 'Стабилизаторы напряжения'),
    ('Садовая техника', 'garden', 'Мотоблоки, садовая техника и аксессуары'),
    ('Отопление и горячая вода', 'heating', 'Обогреватели, водонагреватели и котлы'),
    ('Уборка и уход за одеждой', 'cleaning', 'Пылесосы, утюги, отпариватели'),
    ('Красота и здоровье', 'beauty', 'Фены, триммеры, приборы для ухода'),
    ('Посуда и товары для дома', 'home', 'Посуда, кухонные принадлежности и товары для дома'),
    ('Электроника', 'electronics', 'Компьютеры, аудио, сеть и аксессуары'),
    ('Прочие товары', 'other', 'Инструменты и остальные товары'),
]


def section_for(group):
    g = str(group).casefold()
    if 'сплит-систем' in g or g.startswith('полупромышлен') or 'кондиционер' in g:
        return 'Кондиционеры'
    if g.startswith('расходные материалы'):
        return 'Расходники'
    if 'к климатической технике' in g:
        return 'Расходники'
    if any(k in g for k in ('к вбт', 'к кбт')):
        return 'Крупная бытовая техника'
    if 'к кухонной технике' in g:
        return 'Мелкая бытовая техника'
    if g in ('для напитков и бара', 'для чая и кофе', 'чайники и заварочные чайники'):
        return 'Посуда и товары для дома'
    if any(k in g for k in ('телевизор', 'для тв', 'тв-', 'пульты ду', 'медиаплеер', 'dvb')):
        return 'Телевизоры и ТВ'
    if 'генератор' in g:
        return 'Генераторы'
    if 'стабилизатор' in g:
        return 'Стабилизаторы'
    if any(k in g for k in ('садов', 'мотоблок')):
        return 'Садовая техника'
    if any(k in g for k in ('водонагрев', 'котлы', 'конвектор', 'электрокамин',
                            'обогрев', 'радиатор', 'тепловентил', 'теплый пол', 'плоские эмаль')):
        return 'Отопление и горячая вода'
    if any(k in g for k in ('пылесос', 'утюг', 'отпар', 'пароочист', 'уходу за домом и одеждой')):
        return 'Уборка и уход за одеждой'
    if any(k in g for k in ('фены', 'стрижк', 'волос', 'маникюр', 'педикюр',
                            'весы наполь', 'красот', 'здоров')):
        return 'Красота и здоровье'
    if any(k in g for k in ('электроник', 'мфу', 'ноутбук', 'аудио', 'наушник',
                            'маршрутизатор', 'кабель и переходник', 'защитные', 'чехлы',
                            'мыши', 'видеорегистратор', 'intel', 'аккумулятор', 'зарядные', 'элементы питания')):
        return 'Электроника'
    if any(k in g for k in ('холодильник', 'морозиль', 'стиральн', 'сушильн',
                            'плиты ', 'электрические плиты', 'духовые шкаф', 'поверхност',
                            'посудомо', 'вытяжк', 'водонагрев', 'плоские эмаль', 'котлы')):
        return 'Крупная бытовая техника'
    if any(k in g for k in ('посуд', 'сковород', 'ножи', 'принадлежност', 'приборы стол',
                            'столовые прибор', 'термос', 'термопосуд', 'мангал', 'барбекю',
                            'хранени', 'вилки', 'ложки', 'мойки', 'смесител', 'пищевых отходов',
                            'средства для', 'средства по', 'садовая мебель')):
        return 'Посуда и товары для дома'
    if any(k in g for k in ('микроволн', 'мультивар', 'аэрогрил', 'грил', 'напитк',
                            'чая', 'коф', 'чайник', 'пылесос', 'утюг', 'отпар', 'фены',
                            'стрижк', 'волос', 'маникюр', 'весы', 'красот', 'блендер',
                            'миксер', 'измельчител', 'мясоруб', 'соковыжим', 'тостер',
                            'бутерброд', 'блинниц', 'сушилки для', 'плитки', 'духовки настоль',
                            'конвектор', 'обогрев', 'радиатор', 'тепловентил', 'теплый пол',
                            'мелкая техник', 'вентилятор', 'фритюр')):
        return 'Мелкая бытовая техника'
    return 'Прочие товары'


def _label(ws, cell_range, text, size=12, color='193C57', fill=None, link=None):
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(':')[0]]
    cell.value = text
    cell.font = Font(name='Arial', size=size, bold=True, color=color)
    cell.alignment = Alignment(vertical='center', wrap_text=True, indent=1)
    if fill:
        for row in ws[cell_range]:
            for c in row:
                c.fill = PatternFill('solid', fgColor=fill)
    if link:
        cell.hyperlink = link
    return cell


def build_navigation(df, wb, config, group_order, group_after):
    from transform import build_pricelist
    home = wb.create_sheet('Главная', 0)
    home.sheet_view.showGridLines = False
    home.sheet_view.zoomScale = 85
    home.sheet_properties.tabColor = '758ADB'
    home.freeze_panes = 'A9'
    for col in 'ABCDEFGHIJKLM':
        home.column_dimensions[col].width = 9
    destinations = df['group'].map(section_for)
    sections = [section for section in SECTIONS if (destinations == section[0]).any()]
    end_row = 9 + ((len(sections)+1)//2)*14 + 4
    for row in range(1, end_row+1):
        home.row_dimensions[row].height = 21
        for cell in home[row][:13]:
            cell.fill = PatternFill('solid', fgColor='F1F3F8')
    _label(home, 'B2:L3', 'БЫТ ТЕХ ОПТ  /  КАТАЛОГ', 22, '26334D')
    _label(home, 'B4:L4', f'{datetime.now():%d.%m.%Y}  ·  {len(df):,} товаров  ·  {len(sections)} разделов'.replace(',', ' '), 11, '68758C')
    _label(home, 'B6:L7', 'Нажмите на карточку или название — откроется нужный раздел', 12,
           '405275', 'E5EAF5')
    image_links = []
    for index, (title, image_name, subtitle) in enumerate(sections):
        row = 9 + (index//2)*14
        left, right = ('B','F') if index%2 == 0 else ('H','L')
        subset = df[destinations == title]
        sheet = wb.create_sheet(title)
        build_pricelist(subset, wb, config, group_order, group_after, worksheet=sheet)
        sheet.sheet_view.showGridLines = False
        sheet.sheet_properties.tabColor = '758ADB'
        sheet.column_dimensions['H'].width = 24
        sheet['H2'] = 'На главную'
        sheet['H2'].hyperlink = "#'Главная'!A1"
        sheet['H2'].font = Font(name='Arial', bold=True, size=12, color='546BB5', underline='single')
        target = f"#'{title}'!A1"
        path = BASE / 'assets' / 'catalog' / 'glass' / f'{image_name}.png'
        if not path.is_file():
            raise FileNotFoundError(f'Не найдена картинка каталога: {path.name}')
        image = Image(str(path))
        image.anchor = OneCellAnchor(
            _from=AnchorMarker(col=ord(left)-65, row=row-1),
            ext=XDRPositiveSize2D(pixels_to_EMU(340), pixels_to_EMU(227)))
        home.add_image(image)
        image_links.append(target)
        _label(home, f'{left}{row+9}:{right}{row+10}', title, 14,
               '26334D', TITLE_COLORS[index % len(TITLE_COLORS)], target)
        _label(home, f'{left}{row+11}:{right}{row+12}',
               f'{subtitle}\nПозиций: {len(subset)}', 10, '64728A')
    footer = 9 + ((len(sections)+1)//2)*14
    _label(home, f'B{footer}:L{footer+1}', 'ОТКРЫТЬ ВЕСЬ ПРАЙС  →', 13, 'FFFFFF', '657DBB', "#'Прайс клиента'!A1")
    full = wb['Прайс клиента']
    full['H2'] = 'На главную'
    full['H2'].hyperlink = "#'Главная'!A1"
    full['H2'].font = Font(name='Arial', bold=True, color='546BB5', underline='single')
    full.column_dimensions['H'].width = 24
    wb.active = 0
    home.print_area = f'A1:M{footer+2}'
    home.page_setup.orientation = 'portrait'
    home.page_setup.paperSize = home.PAPERSIZE_A4
    home.page_setup.fitToWidth = 1
    home.page_setup.fitToHeight = 0
    home.sheet_properties.pageSetUpPr.fitToPage = True
    return image_links


def link_navigation_images(filepath, targets):
    """Add DrawingML click relationships unsupported by openpyxl's image API."""
    if not targets:
        return
    drawing = 'xl/drawings/drawing1.xml'
    rel_path = 'xl/drawings/_rels/drawing1.xml.rels'
    a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    xdr = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    rel = 'http://schemas.openxmlformats.org/package/2006/relationships'
    office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    tmp = Path(filepath).with_suffix('.nav.tmp')
    try:
        with ZipFile(filepath) as src, ZipFile(tmp, 'w', ZIP_DEFLATED) as dst:
            tree = ET.fromstring(src.read(drawing))
            relationships = ET.fromstring(src.read(rel_path))
            pictures = tree.findall(f'.//{{{xdr}}}cNvPr')
            if len(pictures) != len(targets):
                raise ValueError('Количество картинок меню не совпадает со ссылками')
            # Excel picture transparency softens the glow against the pale sheet
            # without modifying or regenerating the original illustrations.
            for blip in tree.findall(f'.//{{{a}}}blip'):
                ET.SubElement(blip, f'{{{a}}}alphaModFix', {'amt': '70000'})
            for i, (picture, target) in enumerate(zip(pictures, targets), 1):
                rid = f'navLink{i}'
                ET.SubElement(picture, f'{{{a}}}hlinkClick', {f'{{{office}}}id': rid,
                              'tooltip': 'Открыть раздел каталога'})
                ET.SubElement(relationships, f'{{{rel}}}Relationship', {
                    'Id': rid, 'Type': office + '/hyperlink', 'Target': target, 'TargetMode': 'External'})
            for info in src.infolist():
                data = src.read(info.filename)
                if info.filename == drawing:
                    data = ET.tostring(tree, encoding='utf-8', xml_declaration=True)
                elif info.filename == rel_path:
                    data = ET.tostring(relationships, encoding='utf-8', xml_declaration=True)
                dst.writestr(info, data)
        os.replace(tmp, filepath)
    finally:
        tmp.unlink(missing_ok=True)
