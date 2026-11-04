from io import BytesIO
from pathlib import Path
import sys
from zipfile import ZipFile

import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from workbook_media import optimize_workbook_images


def test_compression_shares_images_and_preserves_workbook(tmp_path):
    source = tmp_path / 'picture.png'
    Image.new('RGB', (1536, 1024), '#8394c5').save(source)
    original = source.read_bytes()
    book = openpyxl.Workbook()
    sheet = book.active
    sheet['A1'] = 'Catalog'
    sheet['A1'].hyperlink = "#'Second'!A1"
    sheet['B2'] = 12345
    sheet.add_image(ExcelImage(source), 'D4')
    second = book.create_sheet('Second')
    second.add_image(ExcelImage(source), 'B2')
    path = tmp_path / 'book.xlsx'
    book.save(path)
    with ZipFile(path) as zipfile:
        xml = {n: zipfile.read(n) for n in zipfile.namelist()
               if n.startswith(('xl/worksheets/', 'xl/drawings/')) and n.endswith('.xml')}
    optimize_workbook_images(path)
    with ZipFile(path) as zipfile:
        assert all(zipfile.read(n) == data for n, data in xml.items())
        media = {n: zipfile.read(n) for n in zipfile.namelist() if n.startswith('xl/media/')}
        assert len(media) == 1
        image = Image.open(BytesIO(next(iter(media.values()))))
        assert image.size == (680, 453)
    loaded = openpyxl.load_workbook(path)
    assert loaded.active['B2'].value == 12345
    assert loaded.active['A1'].hyperlink.target == "#'Second'!A1"
    assert len(loaded.active._images) == len(loaded['Second']._images) == 1
    optimize_workbook_images(path)
    with ZipFile(path) as zipfile:
        assert {n: zipfile.read(n) for n in zipfile.namelist() if n.startswith('xl/media/')} == media
    assert source.read_bytes() == original
