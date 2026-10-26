"""화면안 이미지에서 휴대폰 화면·B.O 영역을 잘라 PNG로 저장한다.
- box는 '보이는 이미지' 기준 좌표, scale = 원본폭 / 보이는폭 (view 도구로 볼 때 축소된 경우)
- clean: 잡티(붙여넣기 아이콘 등) 영역을 주변 배경색으로 덮음
사용 예:
  crop('mail.png', (117,104,241,368), 'n_list.png', scale=1393/876)
  clean('q1.png', (941,340,1006,370), sample=(1100,350))"""
from PIL import Image, ImageDraw
def crop(src, box, out, scale=1.0):
    im = Image.open(src).convert('RGB')
    c = im.crop(tuple(int(v * scale) for v in box)); c.save(out); return c.size
def clean(src, rect, sample, out=None):
    im = Image.open(src).convert('RGB')
    ImageDraw.Draw(im).rectangle(rect, fill=im.getpixel(sample)); im.save(out or src)
