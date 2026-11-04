"""Compress embedded Excel pictures without resaving cells or drawings."""
from hashlib import sha256
from io import BytesIO
from pathlib import Path
import os
import posixpath
import xml.etree.ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

from PIL import Image, ImageOps


def _compress(data):
    with Image.open(BytesIO(data)) as original:
        image = ImageOps.exif_transpose(original)
        limit = 1600 if image.width / image.height >= 3 else 680
        if original.format == 'JPEG' and max(image.size) <= limit:
            return data, 'jpg'
        image.thumbnail((limit, limit), Image.Resampling.LANCZOS)
        transparent = image.mode in ('RGBA', 'LA') and image.getextrema()[-1][0] < 255
        transparent = transparent or (image.mode == 'P' and 'transparency' in image.info)
        output = BytesIO()
        if transparent:
            image.save(output, format='PNG', optimize=True)
            return output.getvalue(), 'png'
        image.convert('RGB').save(output, format='JPEG', quality=84, subsampling=0, optimize=True)
        return output.getvalue(), 'jpg'


REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'


def _tostring(tree, namespace):
    """XML с пространством имён по умолчанию, как пишет Excel (без префикса ns0:)."""
    ET.register_namespace('', namespace)
    return ET.tostring(tree, encoding='utf-8', xml_declaration=True)


def optimize_workbook_images(filepath):
    """Keep all worksheet data and hyperlinks; compress and share duplicate images."""
    path = Path(filepath)
    tmp = path.with_suffix('.media.tmp')
    try:
        with ZipFile(path) as source:
            replacements, media, cache = {}, {}, {}
            for name in source.namelist():
                if not name.startswith('xl/media/') or not name.lower().endswith(('.png', '.jpg', '.jpeg')):
                    continue
                data = source.read(name)
                digest = sha256(data).hexdigest()
                if digest not in cache:
                    compressed, extension = _compress(data)
                    target = f'xl/media/compact_{sha256(compressed).hexdigest()}.{extension}'
                    media[target] = compressed
                    cache[digest] = target
                replacements[name] = cache[digest]
            if not replacements:
                return
            with ZipFile(tmp, 'w', ZIP_DEFLATED, compresslevel=9) as dest:
                for info in source.infolist():
                    if info.filename in replacements:
                        continue
                    data = source.read(info.filename)
                    if info.filename.endswith('.rels'):
                        tree = ET.fromstring(data)
                        base = posixpath.dirname(posixpath.dirname(info.filename))
                        changed = False
                        for relationship in tree:
                            if not relationship.attrib.get('Type', '').endswith('/image'):
                                continue
                            target = relationship.attrib['Target']
                            absolute = (target.lstrip('/') if target.startswith('/') else
                                        posixpath.normpath(posixpath.join(base, target)))
                            if absolute in replacements:
                                relationship.set('Target', '/' + replacements[absolute])
                                changed = True
                        if changed:
                            # default_namespace: без префиксов ns0: — их не понимают
                            # упрощённые мобильные просмотрщики.
                            data = _tostring(tree, REL_NS)
                    elif info.filename == '[Content_Types].xml':
                        tree = ET.fromstring(data)
                        ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
                        for entry in list(tree):
                            if entry.attrib.get('PartName', '').lstrip('/') in replacements:
                                tree.remove(entry)
                        defaults = {entry.attrib.get('Extension') for entry in tree}
                        for extension, mime in [('jpg', 'image/jpeg'), ('png', 'image/png')]:
                            if extension not in defaults:
                                ET.SubElement(tree, f'{{{ns}}}Default', Extension=extension, ContentType=mime)
                        data = _tostring(tree, ns)
                    dest.writestr(info.filename, data)
                for name, data in media.items():
                    dest.writestr(name, data)
        os.replace(tmp, path)
    finally:
        tmp.unlink(missing_ok=True)
