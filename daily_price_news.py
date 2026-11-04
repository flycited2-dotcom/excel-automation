"""Short factual Telegram highlights; baseline advances only after delivery."""
from collections import Counter
from datetime import datetime
from hashlib import sha256
from html import escape
import json
import math
from pathlib import Path
import re

from catalog_navigation import section_for

BASE = Path(__file__).parent


def write_snapshot(df, output_path, *, supplier_fallback=False):
    path = Path(output_path)
    date = re.search(r'_(\d{8})(?:_\d{6})?\.xlsx$', path.name)
    if not date:
        raise ValueError('Missing price issue date')
    products = {}
    ambiguous = set()
    for row in df.to_dict('records'):
        price = float(row.get('price_out') or 0)
        if not math.isfinite(price) or price <= 0:
            continue
        def clean(key):
            value = row.get(key)
            return '' if value is None or (isinstance(value, float) and math.isnan(value)) else str(value).strip()
        source, article, brand, name = (clean(k) for k in ('source', 'article', 'brand', 'name'))
        key = json.dumps([source, article, brand, name], ensure_ascii=False)
        product = {'name': name, 'price': round(price, 2), 'section': section_for(clean('group'))}
        if key in products and products[key] != product:
            ambiguous.add(key)
        products[key] = product
    for key in ambiguous:
        products.pop(key, None)
    payload = {'schema': 1, 'date': datetime.strptime(date[1], '%Y%m%d').date().isoformat(),
               'sha256': sha256(path.read_bytes()).hexdigest(),
               'supplier_fallback': bool(supplier_fallback), 'products': products}
    from ready_price import write_json
    write_json(path.with_suffix('.news.json'), payload)


def state_path(recipient, state_dir=None):
    folder = Path(state_dir) if state_dir is not None else BASE / 'output' / 'caption-state'
    return folder / (sha256(str(recipient).encode()).hexdigest() + '.json')


def prepare_note(filepath, recipient, *, state_dir=None):
    path = Path(filepath)
    sidecar = path.with_suffix('.news.json')
    if not sidecar.is_file():
        return '', None
    current = json.loads(sidecar.read_text(encoding='utf-8'))
    if current.get('schema') != 1 or current.get('sha256') != sha256(path.read_bytes()).hexdigest():
        raise ValueError('Price news snapshot does not match workbook')
    baseline = state_path(recipient, state_dir)
    try:
        previous = json.loads(baseline.read_text(encoding='utf-8')) if baseline.is_file() else None
    except (ValueError, OSError):
        previous = None
    if not isinstance(previous, dict) or previous.get('schema') != 1:
        previous = None
    if previous and previous.get('sha256') == current['sha256']:
        return previous['note'], previous
    day = datetime.strptime(current['date'], '%Y-%m-%d').date().toordinal()
    products = current['products']
    lines = []
    if (previous and previous.get('schema') == 1 and previous['date'] <= current['date']
            and not current['supplier_fallback'] and not previous.get('supplier_fallback')):
        old = previous['products']
        cheaper = [(key, item) for key, item in products.items()
                   if key in old and item['price'] < old[key]['price'] - 0.005]
        new = [item for key, item in products.items() if key not in old]
        if cheaper or new:
            compared = datetime.strptime(previous['date'], '%Y-%m-%d').strftime('%d.%m.%Y')
            lines.append(f'По сравнению с прайсом от {compared}:')
        if cheaper:
            key, item = max(cheaper, key=lambda pair: (old[pair[0]]['price'] - pair[1]['price']) / old[pair[0]]['price'])
            def money(value):
                return f'{value:,.2f}'.rstrip('0').rstrip('.').replace(',', ' ').replace('.', ',')
            name = escape(item['name'][:75] + ('…' if len(item['name']) > 75 else ''))
            lines.append(f'📉 Снизились цены на {len(cheaper)} поз.: {name} — '
                         f'{money(old[key]["price"])} → {money(item["price"])} ₽.')
        if new:
            section, count = Counter(item['section'] for item in new).most_common(1)[0]
            lines.append(f'🆕 В разделе «{escape(section)}» появились в прайсе {count} поз.')
    if not lines and products:
        sections = sorted({item['section'] for item in products.values()})
        section = sections[day % len(sections)]
        templates = ['💡 Обратите внимание на раздел «{}» — подборка в приложенном прайсе.',
                     '🔎 Сегодня загляните в раздел «{}»: модели и цены — внутри файла.',
                     '📌 Подбираете технику? Откройте раздел «{}» на главной странице прайса.']
        lines.append(templates[day % len(templates)].format(escape(section)))
    if current['supplier_fallback']:
        lines.insert(0, 'ℹ️ Часть цен поставщика взята из предыдущего сохранённого файла.')
    current['note'] = '\n'.join(lines)
    return current['note'], current


def record_success(payload, recipient, *, state_dir=None):
    if payload is not None:
        from ready_price import write_json
        write_json(state_path(recipient, state_dir), payload)
