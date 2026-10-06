"""Create labeled contact sheets from the rendered final PDF pages."""
from pathlib import Path
from PIL import Image, ImageDraw

folder = Path('/private/tmp/agrovector-unified')
for prefix, columns, count in [('deck', 3, 9), ('doc', 4, 12)]:
    pages = sorted(folder.glob(prefix + '-*.png'))
    for start in range(0, len(pages), count):
        batch = pages[start:start + count]
        sheet = Image.new('RGB', (columns * 400, ((len(batch) + columns - 1) // columns) * 590), '#cccccc')
        draw = ImageDraw.Draw(sheet)
        for i, page in enumerate(batch):
            im = Image.open(page)
            im.thumbnail((390, 555))
            x, y = (i % columns) * 400, (i // columns) * 590
            sheet.paste(im, (x, y + 25))
            draw.text((x + 5, y + 5), page.stem, fill='black')
        sheet.save(folder / f'{prefix}-contact-{start // count + 1}.jpg')
