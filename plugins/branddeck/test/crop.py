"""Crop regions of a rendered slide for close inspection."""
import sys
from PIL import Image

src, dst, box = sys.argv[1], sys.argv[2], sys.argv[3]
left, top, right, bottom = (int(v) for v in box.split(','))
image = Image.open(src)
crop = image.crop((left, top, right, bottom))
scale = 2
crop = crop.resize((crop.width * scale, crop.height * scale), Image.LANCZOS)
crop.save(dst)
print(f"{dst} {crop.width}x{crop.height}")
