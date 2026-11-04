"""Category sheets and a photo menu, generated on every price run."""
from datetime import datetime
from io import BytesIO
from pathlib import Path

from PIL import Image as PILImage
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.worksheet.hyperlink import Hyperlink

BASE = Path(__file__).parent
# Насыщенные цвета кнопок «Открыть →» (белый текст читается на всех).
BUTTON_COLORS = ('2F6FD6', '1F9D6B', 'D9682B', '7A4FC9', 'C9406A',
                 '4E9A2F', '178A9E', 'B7871B', '3F5BC2', '9C5B2E',
                 'A8459A', '5F8F2A', '2D7FB8', '6A55B0')
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


PAGE_BG = 'F1F3F8'
CARD_W, CARD_H = 340, 227          # px; 5 колонок по 9 ≈ 68 px, 8 строк по 21 pt = 28 px
CARD_OPACITY = 0.7                 # осветление «запекается» в картинку, без эффектов DrawingML


def link_to_sheet(cell, sheet_title):
    """Внутренняя ссылка через location (как пишет сам Excel).

    Строка вида "#'Лист'!A1" превращается в openpyxl во ВНЕШНЮЮ связь — мобильные
    просмотрщики пытаются открыть её как веб-адрес, и переход не работает.
    """
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sheet_title}'!A1",
                               display=str(cell.value or sheet_title))


def card_image(path):
    """PNG карточки, смешанный с фоном страницы: тот же вид, что прежний alphaModFix."""
    with PILImage.open(path) as source:
        picture = source.convert('RGBA')
    background = PILImage.new('RGBA', picture.size, '#' + PAGE_BG)
    blended = PILImage.alpha_composite(background, picture)
    result = PILImage.blend(background, blended, CARD_OPACITY).convert('RGB')
    stream = BytesIO()
    result.save(stream, format='PNG')
    stream.seek(0)
    return Image(stream)


def _label(ws, cell_range, text, size=12, color='193C57', fill=None, link=None, center=False):
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(':')[0]]
    cell.value = text
    cell.font = Font(name='Arial', size=size, bold=True, color=color)
    cell.alignment = (Alignment(horizontal='center', vertical='center', wrap_text=True) if center
                      else Alignment(vertical='center', wrap_text=True, indent=1))
    if fill:
        for row in ws[cell_range]:
            for c in row:
                c.fill = PatternFill('solid', fgColor=fill)
    if link:
        link_to_sheet(cell, link)
    return cell


def build_navigation(df, wb, config, group_order, group_after):
    from transform import build_pricelist
    home = wb.create_sheet('Главная', 0)
    home.sheet_view.showGridLines = False
    home.sheet_view.zoomScale = 85
    home.sheet_properties.tabColor = '758ADB'
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
    _label(home, 'B6:L7', 'Нажимайте на цветную кнопку «Открыть →» под картинкой. Если не открывается — '
           'разделы есть на вкладках листов внизу экрана', 12, '1F2937', 'FFFF00')
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
        link_to_sheet(sheet['H2'], 'Главная')
        sheet['H2'].font = Font(name='Arial', bold=True, size=12, color='546BB5', underline='single')
        target = f"'{title}'!A1"
        path = BASE / 'assets' / 'catalog' / 'glass' / f'{image_name}.png'
        if not path.is_file():
            raise FileNotFoundError(f'Не найдена картинка каталога: {path.name}')
        image = card_image(path)
        col = ord(left) - 65
        # twoCellAnchor editAs="oneCell" — так картинки пишет сам Excel;
        # oneCellAnchor многие сторонние просмотрщики не отображают.
        image.anchor = TwoCellAnchor(
            editAs='oneCell',
            _from=AnchorMarker(col=col, row=row-1),
            to=AnchorMarker(col=col+5, colOff=0, row=row-1+8, rowOff=pixels_to_EMU(3)))
        home.add_image(image)
        image_links.append(target)
        # Клик по картинке на телефонах не работает, поэтому прямо под ней —
        # крупная цветная кнопка во всю ширину карточки (3 строки).
        _label(home, f'{left}{row+8}:{right}{row+10}', f'{title}\nОткрыть →', 15, 'FFFFFF',
               BUTTON_COLORS[index % len(BUTTON_COLORS)], title, center=True)
        _label(home, f'{left}{row+11}:{right}{row+12}',
               f'{subtitle}\nПозиций: {len(subset)}', 10, '64728A', link=title)
    footer = 9 + ((len(sections)+1)//2)*14
    _label(home, f'B{footer}:L{footer+1}', 'ОТКРЫТЬ ВЕСЬ ПРАЙС  →', 13, 'FFFFFF', '657DBB', 'Прайс клиента')
    full = wb['Прайс клиента']
    full['H2'] = 'На главную'
    link_to_sheet(full['H2'], 'Главная')
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

