"""Optional reference recipe; prints JSON, never installs or writes fixtures.

Run with Pillow12.3.0/LittleCMS2.19 to reproduce the checked reference tuples.
The mandatory PowerShell suites use precomputed REFERENCE.json and do not need
Python or Pillow. Inputs are the CC0 profiles in this directory and synthetic
channel values only.
"""
from pathlib import Path
import json, hashlib
from PIL import Image, ImageCms, __version__ as pillow_version

root = Path(__file__).resolve().parent
reference = json.loads((root / 'REFERENCE.json').read_text(encoding='utf-8'))
manifest = json.loads((root / 'MANIFEST.json').read_text(encoding='utf-8'))
for profile in manifest['profiles']:
    path = root / Path(profile['source_path']).name
    assert len(path.read_bytes()) == profile['bytes']
    assert hashlib.sha256(path.read_bytes()).hexdigest() == profile['sha256']
target = ImageCms.ImageCmsProfile(str(root / 'sRGB-v4.icc'))
wide = ImageCms.ImageCmsProfile(str(root / 'AdobeCompat-v2.icc'))
cmyk = ImageCms.ImageCmsProfile(str(root / 'CGATS001Compat-v2-micro.icc'))

def transform(mode, values, source):
    image = Image.new(mode, (len(values),1))
    image.putdata([tuple(value) for value in values])
    return ImageCms.profileToProfile(image, source, target, renderingIntent=1,
                                     outputMode='RGB', flags=0)

rgb = transform('RGB', reference['rgb_source'], wide)
ink = transform('CMYK', reference['cmyk_source'], cmyk)
alpha_rgb = transform('RGB', [value[:3] for value in reference['alpha_source']], wide)
alpha = Image.new('RGBA', (4,1))
alpha.putdata([tuple(value) for value in reference['alpha_source']])
alpha_rgb.putalpha(alpha.getchannel('A'))
white = Image.new('RGBA', (4,1), (255,255,255,255))
out = {
    'Pillow':pillow_version, 'LittleCMS':ImageCms.core.littlecms_version,
    'intent':1, 'flags':0, 'BPC':False,
    'wide_rgb':list(rgb.getdata()), 'tagged_cmyk':list(ink.getdata()),
    'tagged_alpha_white':list(Image.alpha_composite(white,alpha_rgb).convert('RGB').getdata()),
    'untagged_alpha_white':list(Image.alpha_composite(white,alpha).convert('RGB').getdata())
}
print(json.dumps(out, indent=2))
