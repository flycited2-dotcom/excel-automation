"""Fresh public website catalog, including admin edits; never use stale stock."""
import json
import math
import re
from pathlib import Path

import pandas as pd
import requests

TYPES = {
    'cassette': 'Кассетные', 'floor-ceiling': 'Напольно-потолочные',
    'duct': 'Канальные', 'column': 'Колонные', 'console': 'Консольные',
    'split': 'Настенные', 'multi': 'Мультисплит-системы',
}


def parse_products(script):
    match = re.search(r'\b(?:var|let|const)\s+PRODUCTS\s*=\s*', script)
    if not match:
        raise ValueError('Не найден PRODUCTS в каталоге СплитХаб')
    products, _ = json.JSONDecoder().raw_decode(script[match.end():])
    if not isinstance(products, list) or not products:
        raise ValueError('Некорректный базовый каталог СплитХаб')
    return products


def merge_products(products, admin):
    if admin.get('ok') is not True:
        raise ValueError('Не получены актуальные изменения каталога СплитХаб')
    overrides = admin.get('overrides', {})
    custom = admin.get('custom_products', [])
    if not isinstance(overrides, dict) or not isinstance(custom, list):
        raise ValueError('Некорректный ответ каталога СплитХаб')
    merged, seen = [], set()
    for product in products:
        item = dict(product)
        # sku/id are stable on the website; edits cannot change identity.
        item.update({k: v for k, v in overrides.get(item.get('sku'), {}).items()
                     if k not in ('sku', 'id')})
        merged.append(item)
        seen.add(item.get('sku'))
    for item in custom:
        if item.get('sku') and item['sku'] not in seen:
            merged.append(dict(item))
            seen.add(item['sku'])
    return merged


def product_group(p):
    group, kind = p.get('group'), p.get('type')
    if group == 'truba':
        return 'Расходные материалы — Медная труба'
    if group == 'rashod':
        text = f'{p.get("brand", "")} {p.get("model", "")}'.casefold()
        categories = [
            ('Дренажные насосы', ('насос', 'помпа')),
            ('Дренаж', ('дренаж', 'ruvinil', 'металлопласт', 'гибкий')),
            ('Фитинги для медной трубы', ('фитинг', 'рефнет', 'соединитель', 'латун', 'муфта')),
            ('Изоляция', ('термаэко', 'изоляция', 'утеплитель')),
            ('Кронштейны', ('кронштейн',)),
            ('Ленты и изолента', ('лента', 'скотч')),
            ('Кабель и провод', ('пвс', 'провод', 'кабель')),
            ('Крепёж', ('болт', 'гайка', 'шайба', 'дюбель', 'глухарь', 'анкер')),
            ('Фреон', ('фреон',)),
        ]
        for label, words in categories:
            if any(word in text for word in words):
                return 'Расходные материалы — ' + label
        return 'Расходные материалы — ' + str(p.get('brand') or 'Прочее')
    if group in ('pac_inv', 'pac_onoff'):
        if kind not in TYPES:
            raise ValueError(f'Неизвестный тип полупрома: {kind}')
        drive = 'Инверторные' if group == 'pac_inv' else 'ON/OFF'
        return f'Полупромышленные — {TYPES[kind]} — {drive}'
    if group == 'inv' and kind == 'split':
        return 'Инверторные сплит-системы'
    if group == 'onoff' and kind == 'split':
        return 'Классические сплит-системы'
    raise ValueError(f'Неизвестная категория СплитХаб: {group}/{kind}')


def group_position(group):
    if group.startswith('Полупромышленные — '):
        for index, label in enumerate(TYPES.values()):
            if group.startswith(f'Полупромышленные — {label} — '):
                return 1.6 + index * .04 + (0 if group.endswith('Инверторные') else .02)
    if group.startswith('Расходные материалы — '):
        return 1.95 if group.endswith('Медная труба') else 2.0
    return None


def catalog_frame(products):
    records = []
    for p in products:
        if p.get('active') in (0, '0') or p.get('stock') != 'in_stock':
            continue
        price = float(p.get('price') or 0)
        if not math.isfinite(price) or price <= 0:
            raise ValueError(f'Некорректная цена СплитХаб: {p.get("sku")}')
        if not p.get('model') or not (p.get('id') or p.get('sku')):
            raise ValueError('Товар СплитХаб без модели или артикула')
        name = ' '.join(str(p.get(k) or '').strip() for k in ('model', 'descShort')).strip()
        records.append(dict(article=str(p.get('id') or p['sku']),
                            group=product_group(p), brand=str(p.get('brand') or ''),
                            series=str(p.get('series') or '').strip(),
                            capacity_btu=p.get('btu'),
                            name=name, price_in=price, price_out=price,
                            source='splithub_site'))
    return pd.DataFrame(records, columns=['article', 'group', 'brand', 'name',
                                         'price_in', 'price_out', 'source', 'series', 'capacity_btu'])


def load_catalog(config):
    base = config['splithub_site_url'].rstrip('/') + '/'
    with requests.Session() as session:
        session.headers.update({'Cache-Control': 'no-cache'})
        response = session.get(base + 'products.js', timeout=45)
        response.raise_for_status()
        products = parse_products(response.content.decode('utf-8-sig'))
        response = session.get(base + 'api/admin.php',
                               params={'action': 'products_overrides_public'}, timeout=45)
        response.raise_for_status()
        products = merge_products(products, response.json())
    frame = catalog_frame(products)
    # Keep exact inputs for provenance and troubleshooting, not as a stock fallback.
    from datetime import datetime
    folder = Path(__file__).parent / 'input' / 'splithub'
    folder.mkdir(parents=True, exist_ok=True)
    snapshot = folder / f'catalog_{datetime.now():%Y%m%d_%H%M%S_%f}.json'
    snapshot.write_text(json.dumps({'url': base, 'products': products}, ensure_ascii=False), encoding='utf-8')
    return frame, snapshot
