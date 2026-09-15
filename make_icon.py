# make_icon.py
from PIL import Image, ImageDraw, ImageFont

CORAL = (242, 105, 75, 255)
CORAL_D = (217, 80, 58, 255)
WHITE = (255, 255, 255, 255)
FOLD = (228, 226, 222, 255)

def draw_icon(s):
    img = Image.new('RGBA', (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([0, 0, s - 1, s - 1], radius=int(s * 0.22), fill=CORAL)
    px0, py0 = int(s * 0.26), int(s * 0.16)
    px1, py1 = int(s * 0.74), int(s * 0.84)
    f = int(s * 0.16)
    d.polygon([(px0, py0), (px1 - f, py0), (px1, py0 + f), (px1, py1), (px0, py1)], fill=WHITE)
    d.polygon([(px1 - f, py0), (px1, py0 + f), (px1 - f, py0 + f)], fill=FOLD)
    lh = max(2, int(s * 0.045))
    for i, yy in enumerate((0.30, 0.38, 0.46)):
        y = int(s * yy)
        w = int(s * (0.34 if i != 1 else 0.26))
        d.rounded_rectangle([int(s * 0.34), y, int(s * 0.34) + w, y + lh], radius=lh // 2, fill=(242, 105, 75, 150))
    by0, by1 = int(s * 0.60), int(s * 0.78)
    d.rounded_rectangle([int(s * 0.30), by0, int(s * 0.70), by1], radius=int(s * 0.05), fill=CORAL_D)
    if s >= 48:
        try:
            font = ImageFont.truetype('C:/Windows/Fonts/arialbd.ttf', int(s * 0.13))
        except Exception:
            font = ImageFont.load_default()
        bb = d.textbbox((0, 0), 'PDF', font=font)
        tw, th = bb[2] - bb[0], bb[3] - bb[1]
        d.text(((s - tw) / 2 - bb[0], (by0 + by1) / 2 - th / 2 - bb[1]), 'PDF', font=font, fill=WHITE)
    return img

if __name__ == '__main__':
    base = draw_icon(256)
    base.save('pdf.ico', format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
    base.save('icon.png')
    print('OK: pdf.ico + icon.png created')