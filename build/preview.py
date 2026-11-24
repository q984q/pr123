"""Рендер четырёх вариантов дизайна приложения — по 4 слайда в сетке 2x2."""
import sys, os, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pymupdf as fitz
from PIL import Image, ImageOps
import appendix2 as A2

PPI = 144.0                # 13.333 x 7.5 -> 1920 x 1080
SW, SH = 13.333, 7.5
W, H = int(SW * PPI), int(SH * PPI)

REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
SERIF = "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf"
SERIFB = "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"

_fonts = {}
def font(path):
    if path not in _fonts:
        _fonts[path] = fitz.Font(fontfile=path)
    return _fonts[path]

def tw(text, size, path=REG):
    return font(path).text_length(text, fontsize=size)

def wrap(text, size, maxw, path=REG):
    out = []
    for para in text.split("\n"):
        words, line = para.split(" "), ""
        for w in words:
            t = (line + " " + w).strip()
            if tw(t, size, path) <= maxw or not line:
                line = t
            else:
                out.append(line); line = w
        out.append(line)
    return out

def text(page, x, y, s, size, color, path=REG, maxw=None, lh=1.24, align="l", track=0):
    """y — верх строки. Возвращает y после последней строки."""
    lines = wrap(s, size, maxw, path) if maxw else s.split("\n")
    for ln in lines:
        w = tw(ln, size, path)
        px = x
        if align == "r": px = x - w
        elif align == "c": px = x - w / 2
        page.insert_text((px, y + size * 0.86), ln, fontsize=size,
                         fontname="f" + str(abs(hash(path)) % 999), color=color)
        y += size * lh
    return y

def rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))

def bg(page, color):
    page.draw_rect(fitz.Rect(0, 0, W, H), color=None, fill=rgb(color))

def box(page, x, y, w, h, fill=None, stroke=None, lw=1, r=0):
    o = rgb(stroke) if stroke else None
    f = rgb(fill) if fill else None
    page.draw_rect(fitz.Rect(x, y, x + w, y + h), color=o, fill=f, width=lw, radius=r if r else None)

def line(page, x1, y1, x2, y2, color, lw=1):
    page.draw_line(fitz.Point(x1, y1), fitz.Point(x2, y2), color=rgb(color), width=lw)

_imgs = {}
def iu(v): return v * PPI

def photo(page, path, x, y, w, h, crop=True):
    key = (path, round(w), round(h))
    if key not in _imgs:
        im = Image.open(path).convert("RGB")
        im = ImageOps.fit(im, (int(w), int(h)), method=Image.LANCZOS, centering=(0.5, 0.5))
        tmp = tempfile.mktemp(suffix=".png")
        im.save(tmp)
        _imgs[key] = tmp
    page.insert_image(fitz.Rect(x, y, x + w, y + h), filename=_imgs[key])

def scrim(page, rect, c, a0=0.85, a1=0.0, steps=90, horizontal=True):
    x0, y0, x1, y1 = rect
    for i in range(steps):
        t = i / steps
        amid = fitz.Identity
        a = a0 + (a1 - a0) * t
        if horizontal:
            xx0 = x0 + (x1 - x0) * t; xx1 = x0 + (x1 - x0) * (t + 1 / steps)
            page.draw_rect(fitz.Rect(xx0, y0, xx1 + 0.6, y1), color=None, fill=c, fill_opacity=a)
        else:
            yy0 = y0 + (y1 - y0) * t; yy1 = y0 + (y1 - y0) * (t + 1 / steps)
            page.draw_rect(fitz.Rect(x0, yy0, x1, yy1 + 0.6), color=None, fill=c, fill_opacity=a)

Z = {z["n"]: z for z in A2.ZONES}
z = Z["06"]

def newpage(doc):
    p = doc.new_page(width=W, height=H)
    for path in (REG, BOLD, SERIF, SERIFB):
        p.insert_font(fontname="f" + str(abs(hash(path)) % 999), fontfile=path)
    return p

# ---------------------------------------------------------------- ВАРИАНТ A
def variant_a():
    """Тихий светлый: бумага, тонкая линейка, серифный заголовок."""
    PAP = "#F4F1EA"; INK = "#171614"; ACC = "#A75B24"; MUT = "#6F6A61"; RULE = "#D8D2C6"; CARD = "#FFFFFF"
    d = fitz.open(); M = iu(0.95)
    # 1 обложка
    p = newpage(d); bg(p, PAP)
    text(p, M, iu(0.72), "SOKOL INDUSTRIAL PARK · ПРИЛОЖЕНИЕ", 15, rgb(ACC), BOLD)
    text(p, M, iu(1.25), "Двадцать четыре", 96, rgb(INK), SERIFB, maxw=W - 2 * M)
    text(p, M, iu(2.35), "зоны по отдельности", 96, rgb(INK), SERIFB, maxw=W - 2 * M)
    text(p, M, iu(3.5), "Состав, площади, технология, инженерия и границы применимости", 22, rgb(MUT), REG, maxw=iu(6.4))
    line(p, M, iu(4.55), W - M, iu(4.55), RULE, 1.5)
    text(p, M, iu(4.72), "103 слайда · 4 слайда на зону · фото по всем 24 зонам", 16, rgb(MUT))
    text(p, W - M, iu(4.72), "612,36 га · участки № 2–5", 16, rgb(MUT), REG, align="r")
    photo(p, "assets3/x_aerial.jpg", 0, iu(5.25), W, iu(2.25))
    # 2 обложка зоны
    p = newpage(d); bg(p, PAP)
    text(p, M, iu(0.9), "ЗОНА 06 ИЗ 24 · ПРОИЗВОДСТВО", 15, rgb(ACC), BOLD)
    y = text(p, M, iu(1.5), z["name"], 80, rgb(INK), SERIFB, maxw=iu(6.0))
    line(p, M, y + iu(0.35), M + iu(1.1), y + iu(0.35), ACC, 2.5)
    text(p, M, y + iu(0.62), z["purpose"], 20, rgb("#3D3A35"), REG, maxw=iu(5.7), lh=1.42)
    box(p, M, iu(5.75), iu(5.4), iu(1.05), fill=CARD, stroke=RULE)
    text(p, M + iu(0.3), iu(5.95), "ПЛОЩАДЬ", 12, rgb(MUT), BOLD)
    text(p, M + iu(0.3), iu(6.16), z["area"], 34, rgb(INK), SERIFB)
    text(p, M + iu(2.0), iu(6.2), z["plots"], 14, rgb(MUT), lh=1.3)
    photo(p, A2.PHOTOS["06"], iu(7.35), iu(0.72), iu(5.15), iu(6.06))
    # 3 разделитель
    p = newpage(d); bg(p, PAP)
    text(p, M, iu(1.05), "РАЗДЕЛ ПЕРВЫЙ", 15, rgb(ACC), BOLD)
    text(p, M, iu(1.55), "Производство", 88, rgb(INK), SERIFB)
    text(p, M, iu(2.85), "Шесть кластеров, 221 га. Каждый — отдельный продукт, отдельный заказчик и отдельные сети.", 22, rgb("#3D3A35"), REG, maxw=iu(7.2), lh=1.4)
    line(p, M, iu(4.35), W - M, iu(4.35), RULE, 1.5)
    rows = [("01", "Дата-центр и ИИ", "32 га"), ("02", "Робототехника", "32 га"),
            ("03", "Авиация и БПЛА", "44 га"), ("04", "Биомедицина", "22 га"),
            ("05", "Электроника", "27 га"), ("06", "Автопром и шины", "64 га")]
    for i, (n, nm, a) in enumerate(rows):
        yy = iu(4.6) + i * iu(0.44)
        text(p, M, yy, n, 15, rgb(ACC), BOLD)
        text(p, M + iu(0.7), yy - iu(0.02), nm, 19, rgb(INK))
        text(p, M + iu(5.2), yy - iu(0.02), a, 19, rgb(MUT), REG, align="r")
    box(p, iu(7.4), iu(4.5), iu(5.1), iu(1.75), fill=None, stroke=RULE)
    text(p, iu(7.7), iu(4.72), "ВСЕГО", 12, rgb(MUT), BOLD)
    text(p, iu(7.7), iu(4.98), "221 га", 44, rgb(INK), SERIFB)
    text(p, iu(7.7), iu(5.72), "из 612,36 га всех участков", 14, rgb(MUT))
    # 4 цифры
    p = newpage(d); bg(p, PAP)
    text(p, M, iu(0.62), "06 · ЦИФРЫ И ОГРАНИЧЕНИЯ", 13, rgb(ACC), BOLD)
    text(p, M, iu(0.95), "Что известно точно, а что нет", 46, rgb(INK), SERIFB)
    line(p, M, iu(1.82), W - M, iu(1.82), RULE, 1.5)
    for i, (big, cap) in enumerate(z["tech"]):
        x = M + i * iu(2.95)
        text(p, x, iu(2.05), big, 40, rgb(INK), SERIFB)
        text(p, x, iu(2.62), cap, 14, rgb(MUT), REG, maxw=iu(2.5), lh=1.3)
    line(p, M, iu(3.35), W - M, iu(3.35), RULE)
    text(p, M, iu(3.52), "ИНЖЕНЕРИЯ", 12, rgb(ACC), BOLD)
    text(p, iu(6.85), iu(3.52), "ЧЕГО НЕТ", 12, rgb(ACC), BOLD)
    yy = iu(3.85)
    for e in z["eng"]:
        yy = text(p, M, yy, "—  " + e, 15, rgb("#3D3A35"), REG, maxw=iu(5.5), lh=1.45) + iu(0.1)
    yy = iu(3.85)
    for e in z["limits"]:
        yy = text(p, iu(6.85), yy, "—  " + e, 15, rgb("#3D3A35"), REG, maxw=iu(5.5), lh=1.45) + iu(0.1)
    box(p, M, iu(6.42), W - 2 * M, iu(0.62), fill="#EDE8DC")
    box(p, M, iu(6.42), iu(0.05), iu(0.62), fill=ACC)
    text(p, M + iu(0.28), iu(6.61), "СТАТУС: " + z["status"].upper(), 14, rgb(INK), BOLD)
    return d

# ---------------------------------------------------------------- ВАРИАНТ B
def variant_b():
    """Швейцарская сетка: белый лист, жёсткие правила, синий акцент, крупные поля."""
    BG = "#FFFFFF"; INK = "#0D0D0D"; ACC = "#0047FF"; MUT = "#7A7A7A"; RULE = "#0D0D0D"; SOFT = "#EDEDED"
    d = fitz.open(); M = iu(0.62)
    p = newpage(d); bg(p, BG)
    text(p, M, iu(0.5), "SOKOL / ПРИЛОЖЕНИЕ / 2026", 13, rgb(INK), BOLD)
    line(p, M, iu(0.78), W - M, iu(0.78), RULE, 2)
    text(p, M, iu(0.95), "24 зоны", 128, rgb(INK), BOLD)
    text(p, M, iu(2.65), "по отдельности", 62, rgb(ACC), BOLD)
    photo(p, "assets3/x_aerial.jpg", iu(7.2), iu(0.95), W - M - iu(7.2), iu(4.3))
    line(p, M, iu(5.55), W - M, iu(5.55), RULE, 1)
    for i, (a, b) in enumerate([("103", "слайда"), ("4", "на зону"), ("24", "зоны с фото"), ("612,36", "га, 4 участка")]):
        x = M + i * iu(3.05)
        text(p, x, iu(5.7), a, 40, rgb(INK), BOLD)
        text(p, x, iu(6.35), b, 15, rgb(MUT))
    # зона
    p = newpage(d); bg(p, BG)
    line(p, M, iu(0.5), W - M, iu(0.5), RULE, 2)
    text(p, M, iu(0.62), "06 / 24", 13, rgb(ACC), BOLD)
    text(p, W - M, iu(0.62), "ПРОИЗВОДСТВО", 13, rgb(MUT), BOLD, align="r")
    text(p, M, iu(1.15), "Автопром\nи шины", 96, rgb(INK), BOLD, lh=1.02)
    text(p, M, iu(3.55), z["purpose"], 18, rgb("#333333"), REG, maxw=iu(4.7), lh=1.5)
    photo(p, A2.PHOTOS["06"], iu(6.0), iu(0.9), iu(6.75), iu(4.1))
    line(p, iu(6.0), iu(5.2), W - M, iu(5.2), RULE, 1)
    text(p, iu(6.0), iu(5.35), "ПЛОЩАДЬ", 12, rgb(MUT), BOLD)
    text(p, iu(6.0), iu(5.6), "64 га", 48, rgb(INK), BOLD)
    text(p, iu(8.6), iu(5.38), "УЧАСТКИ", 12, rgb(MUT), BOLD)
    text(p, iu(8.6), iu(5.64), z["plots"], 16, rgb(INK), REG, maxw=iu(3.9), lh=1.35)
    line(p, M, iu(6.85), W - M, iu(6.85), RULE, 2)
    text(p, M, iu(6.98), "SOKOL INDUSTRIAL PARK", 12, rgb(MUT))
    text(p, W - M, iu(6.98), "06", 12, rgb(MUT), REG, align="r")
    # разделитель: синяя плашка на всю высоту слева
    p = newpage(d); bg(p, BG)
    box(p, 0, 0, iu(4.35), H, fill=ACC)
    text(p, iu(0.62), iu(0.62), "РАЗДЕЛ 01", 14, rgb("#FFFFFF"), BOLD)
    text(p, iu(0.62), iu(1.05), "Произ-\nводство", 84, rgb("#FFFFFF"), BOLD, lh=1.02)
    text(p, iu(5.0), iu(0.62), "Шесть кластеров · 221 га", 15, rgb(MUT), BOLD)
    line(p, iu(5.0), iu(0.95), W - M, iu(0.95), RULE, 2)
    rows = [("01", "Дата-центр и ИИ", "32"), ("02", "Робототехника", "32"),
            ("03", "Авиация и БПЛА", "44"), ("04", "Биомедицина", "22"),
            ("05", "Электроника", "27"), ("06", "Автопром и шины", "64")]
    for i, (n, nm, a) in enumerate(rows):
        yy = iu(1.25) + i * iu(0.86)
        text(p, iu(5.0), yy, n, 16, rgb(ACC), BOLD)
        text(p, iu(5.85), yy - iu(0.03), nm, 27, rgb(INK), BOLD)
        text(p, W - M, yy - iu(0.03), a + " га", 27, rgb(MUT), REG, align="r")
        line(p, iu(5.0), yy + iu(0.62), W - M, yy + iu(0.62), "#D8D8D8", 1)
    text(p, iu(5.0), iu(6.6), "Итого 221 га из 612,36 га всех участков", 15, rgb(MUT), REG, maxw=iu(7.6))
    # цифры
    p = newpage(d); bg(p, BG)
    line(p, M, iu(0.5), W - M, iu(0.5), RULE, 2)
    text(p, M, iu(0.62), "06 / ЦИФРЫ И ОГРАНИЧЕНИЯ", 13, rgb(ACC), BOLD)
    text(p, M, iu(1.05), "Что известно точно, а что нет", 52, rgb(INK), BOLD)
    for i, (big, cap) in enumerate(z["tech"]):
        x = M + i * iu(3.05)
        text(p, x, iu(2.1), big, 52, rgb(INK), BOLD)
        text(p, x, iu(2.95), cap, 14, rgb(MUT), REG, maxw=iu(2.6), lh=1.35)
    line(p, M, iu(3.55), W - M, iu(3.55), RULE, 1)
    box(p, M, iu(3.7), iu(5.9), iu(0.3), fill=ACC)
    text(p, M, iu(3.76), "ИНЖЕНЕРИЯ", 13, rgb("#FFFFFF"), BOLD)
    text(p, iu(7.0), iu(3.76), "ЧЕГО НЕТ", 13, rgb(ACC), BOLD)
    yy = iu(4.25)
    for e in z["eng"]:
        yy = text(p, M, yy, e, 15, rgb("#222222"), REG, maxw=iu(5.6), lh=1.45) + iu(0.16)
    yy = iu(4.25)
    for e in z["limits"]:
        yy = text(p, iu(7.0), yy, e, 15, rgb("#222222"), REG, maxw=iu(5.5), lh=1.45) + iu(0.16)
    line(p, M, iu(6.45), W - M, iu(6.45), RULE, 2)
    text(p, M, iu(6.6), "СТАТУС", 12, rgb(MUT), BOLD)
    text(p, iu(2.0), iu(6.58), z["status"], 18, rgb(INK), BOLD, maxw=iu(7.2))
    text(p, W - M, iu(6.6), "СВЯЗИ: " + z["links"].upper(), 12, rgb(MUT), REG, align="r", maxw=iu(3.0))
    return d

# ---------------------------------------------------------------- ВАРИАНТ C
def variant_c():
    """Годовой отчёт: глубокий тёмно-синий, кремовый текст, золото, воздух."""
    BG = "#0B1526"; INK = "#F3EDE1"; ACC = "#D9A94C"; MUT = "#8C99AE"; CARD = "#132035"
    d = fitz.open(); M = iu(1.0)
    p = newpage(d); bg(p, BG)
    photo(p, "assets3/x_aerial.jpg", 0, 0, W, H)
    scrim(p, (0, 0, W, H), rgb("#08101C"), a0=0.93, a1=0.55)
    text(p, M, iu(1.0), "ПРИЛОЖЕНИЕ К КОЛОДЕ SOKOL", 15, rgb(ACC), BOLD)
    line(p, M, iu(1.42), M + iu(1.5), iu(1.42), ACC, 2)
    text(p, M, iu(1.75), "Двадцать четыре зоны", 76, rgb(INK), SERIF, maxw=iu(9.0))
    text(p, M, iu(3.05), "по отдельности", 76, rgb(INK), SERIF)
    text(p, M, iu(4.6), "Состав · площади · технология · инженерия · границы применимости", 20, rgb(MUT))
    box(p, M, iu(5.6), iu(11.3), iu(0.02), fill=ACC)
    for i, (a, b) in enumerate([("103", "слайда"), ("4", "слайда на зону"), ("24", "зоны с фото"), ("612,36", "га, четыре участка")]):
        x = M + i * iu(2.85)
        text(p, x, iu(5.9), a, 38, rgb(ACC), SERIFB)
        text(p, x, iu(6.5), b, 14, rgb(MUT))
    # зона
    p = newpage(d); bg(p, BG)
    photo(p, A2.PHOTOS["06"], iu(7.5), 0, iu(5.83), H)
    scrim(p, (iu(7.5), 0, iu(8.6), H), rgb("#0B1526"), a0=1.0, a1=0.0, steps=50)
    text(p, M, iu(1.0), "ЗОНА 06 · ПРОИЗВОДСТВО", 14, rgb(ACC), BOLD)
    y = text(p, M, iu(1.55), z["name"], 68, rgb(INK), SERIF, maxw=iu(6.0))
    line(p, M, y + iu(0.4), M + iu(1.3), y + iu(0.4), ACC, 2)
    text(p, M, y + iu(0.75), z["purpose"], 18, rgb("#C9D1DE"), REG, maxw=iu(5.7), lh=1.5)
    text(p, M, iu(5.05), "ПЛОЩАДЬ", 11, rgb(MUT), BOLD)
    text(p, M, iu(5.28), z["area"], 42, rgb(ACC), SERIFB)
    box(p, M + iu(2.35), iu(5.32), iu(0.02), iu(0.55), fill="#2A3A52")
    text(p, M + iu(2.7), iu(5.3), z["plots"], 14, rgb(MUT), lh=1.4)
    text(p, M, iu(6.75), "SOKOL INDUSTRIAL PARK · 2026", 11, rgb("#5A6880"))
    # разделитель
    p = newpage(d); bg(p, BG)
    box(p, 0, 0, W, iu(0.28), fill=ACC)
    text(p, M, iu(1.15), "РАЗДЕЛ ПЕРВЫЙ", 14, rgb(ACC), BOLD)
    text(p, M, iu(1.6), "Производство", 92, rgb(INK), SERIF)
    text(p, M, iu(3.2), "Шесть кластеров, 221 га. Каждый — отдельный продукт,\nотдельный заказчик и отдельные сети.", 21, rgb("#C9D1DE"), REG, lh=1.55)
    rows = [("01", "Дата-центр и ИИ", "32 га"), ("02", "Робототехника", "32 га"),
            ("03", "Авиация и БПЛА", "44 га"), ("04", "Биомедицина", "22 га"),
            ("05", "Электроника", "27 га"), ("06", "Автопром и шины", "64 га")]
    for i, (n, nm, a) in enumerate(rows[:3]):
        yy = iu(4.55) + i * iu(0.6)
        text(p, M, yy, n, 13, rgb(ACC), BOLD)
        text(p, M + iu(0.65), yy - iu(0.05), nm, 20, rgb(INK), SERIF)
        text(p, iu(6.6), yy - iu(0.04), a, 20, rgb(MUT), SERIF, align="r")
    for i, (n, nm, a) in enumerate(rows[3:]):
        yy = iu(4.55) + i * iu(0.6)
        text(p, iu(7.3), yy, n, 13, rgb(ACC), BOLD)
        text(p, iu(7.95), yy - iu(0.05), nm, 20, rgb(INK), SERIF)
        text(p, W - M, yy - iu(0.04), a, 20, rgb(MUT), SERIF, align="r")
    box(p, M, iu(6.45), W - 2 * M, iu(0.02), fill="#2A3A52")
    text(p, M, iu(6.6), "ВСЕГО 221 га", 16, rgb(ACC), BOLD)
    text(p, W - M, iu(6.6), "из 612,36 га всех участков", 14, rgb(MUT), REG, align="r")
    # цифры
    p = newpage(d); bg(p, BG)
    text(p, M, iu(0.72), "06 · ЦИФРЫ И ОГРАНИЧЕНИЯ", 13, rgb(ACC), BOLD)
    text(p, M, iu(1.05), "Что известно точно, а что нет", 44, rgb(INK), SERIF)
    box(p, M, iu(1.85), W - 2 * M, iu(0.02), fill="#2A3A52")
    for i, (big, cap) in enumerate(z["tech"]):
        x = M + i * iu(2.9)
        text(p, x, iu(2.1), big, 40, rgb(ACC), SERIFB)
        text(p, x, iu(2.65), cap, 13, rgb(MUT), REG, maxw=iu(2.4), lh=1.35)
    box(p, M, iu(3.5), iu(5.5), iu(2.55), fill=CARD)
    box(p, iu(6.9), iu(3.5), iu(5.4), iu(2.55), fill=CARD)
    text(p, M + iu(0.4), iu(3.72), "ИНЖЕНЕРИЯ", 11, rgb(ACC), BOLD)
    text(p, iu(7.3), iu(3.72), "ЧЕГО НЕТ", 11, rgb(ACC), BOLD)
    yy = iu(4.05)
    for e in z["eng"]:
        yy = text(p, M + iu(0.4), yy, e, 13.5, rgb("#CBD3E0"), REG, maxw=iu(4.7), lh=1.45) + iu(0.12)
    yy = iu(4.05)
    for e in z["limits"]:
        yy = text(p, iu(7.3), yy, e, 13.5, rgb("#CBD3E0"), REG, maxw=iu(4.6), lh=1.45) + iu(0.12)
    text(p, M, iu(6.35), "СТАТУС", 11, rgb(MUT), BOLD)
    text(p, M, iu(6.58), z["status"], 19, rgb(INK), SERIF, maxw=iu(9.0))
    text(p, W - M, iu(6.62), "СВЯЗИ · " + z["links"].upper(), 10.5, rgb(MUT), REG, align="r", maxw=iu(3.2))
    return d

# ---------------------------------------------------------------- ВАРИАНТ D
def variant_d():
    """Индустриальный контраст: почти чёрный, оранжевый сигнал, огромные цифры, фото во всю ширину."""
    BG = "#111111"; INK = "#FFFFFF"; ACC = "#FF5A1F"; MUT = "#8E8E8E"; CARD = "#1B1B1B"
    d = fitz.open(); M = iu(0.7)
    p = newpage(d); bg(p, BG)
    photo(p, "assets3/x_aerial.jpg", 0, 0, W, iu(3.9))
    scrim(p, (0, iu(2.3), W, iu(3.9)), rgb("#111111"), a0=0.0, a1=1.0, horizontal=False, steps=60)
    text(p, M, iu(0.6), "SOKOL · ПРИЛОЖЕНИЕ", 14, rgb(ACC), BOLD)
    text(p, M, iu(4.15), "24 ЗОНЫ", 104, rgb(INK), BOLD)
    text(p, M, iu(5.55), "подробно по каждой", 34, rgb(ACC), BOLD)
    box(p, M, iu(6.35), iu(0.6), iu(0.09), fill=ACC)
    text(p, M, iu(6.62), "103 слайда · 4 на зону · фото по всем 24 · 612,36 га", 16, rgb(MUT))
    # зона
    p = newpage(d); bg(p, BG)
    photo(p, A2.PHOTOS["06"], iu(7.0), 0, iu(6.33), H)
    scrim(p, (iu(6.2), 0, iu(8.2), H), rgb("#111111"), a0=1.0, a1=0.0, steps=70)
    text(p, M, iu(0.62), "ЗОНА 06 / 24", 14, rgb(ACC), BOLD)
    text(p, iu(2.2), iu(0.62), "ПРОИЗВОДСТВО", 14, rgb("#555555"), BOLD)
    text(p, M, iu(1.15), "АВТОПРОМ\nИ ШИНЫ", 62, rgb(INK), BOLD, lh=1.06)
    text(p, M, iu(3.25), z["purpose"], 16, rgb("#B9B9B9"), REG, maxw=iu(5.4), lh=1.5)
    text(p, M, iu(5.35), "64", 82, rgb(ACC), BOLD)
    text(p, M + iu(1.45), iu(5.75), "га", 26, rgb(INK), BOLD)
    text(p, M, iu(6.55), z["plots"].replace("\n", " · "), 13, rgb(MUT), maxw=iu(5.2))
    # разделитель: фото на всю ширину + крупная цифра раздела
    p = newpage(d); bg(p, BG)
    photo(p, "assets3/z02_robotics.jpg", 0, 0, W, H)
    scrim(p, (0, 0, iu(7.4), H), rgb("#111111"), a0=0.97, a1=0.15, steps=80)
    text(p, M, iu(0.85), "РАЗДЕЛ 01", 14, rgb(ACC), BOLD)
    text(p, M, iu(1.2), "ПРОИЗВОДСТВО", 76, rgb(INK), BOLD)
    text(p, M, iu(2.75), "6 кластеров · 221 га", 22, rgb(ACC), BOLD)
    text(p, M, iu(3.35), "Дата-центр и ИИ 32 · Робототехника 32 · Авиация и БПЛА 44\nБиомедицина 22 · Электроника 27 · Автопром и шины 64", 15, rgb("#C4C4C4"), REG, lh=1.6)
    box(p, M, iu(4.75), iu(5.9), iu(0.02), fill=ACC)
    text(p, M, iu(5.0), "Этот раздел — про заводы. Дальше идут логистика,\nгород и инженерия: у них другие заказчики и другие сети.", 14, rgb("#8E8E8E"), REG, lh=1.6)
    # цифры
    p = newpage(d); bg(p, BG)
    text(p, M, iu(0.6), "01", 20, rgb(ACC), BOLD)
    text(p, iu(1.35), iu(0.62), "ЦИФРЫ И ОГРАНИЧЕНИЯ · ЗОНА 06", 13, rgb(MUT), BOLD)
    text(p, M, iu(1.0), "ЧТО ТОЧНО, А ЧТО НЕТ", 52, rgb(INK), BOLD)
    for i, (big, cap) in enumerate(z["tech"]):
        x = M + i * iu(3.05)
        box(p, x, iu(2.0), iu(2.75), iu(1.42), fill=CARD)
        text(p, x + iu(0.28), iu(2.18), big, 34, rgb(ACC), BOLD)
        text(p, x + iu(0.28), iu(2.72), cap, 12, rgb(MUT), REG, maxw=iu(2.2), lh=1.3)
    text(p, M, iu(3.72), "ИНЖЕНЕРИЯ", 12, rgb(ACC), BOLD)
    text(p, iu(6.9), iu(3.72), "ЧЕГО НЕТ", 12, rgb(ACC), BOLD)
    yy = iu(4.02)
    for e in z["eng"]:
        yy = text(p, M, yy, e, 14, rgb("#D2D2D2"), REG, maxw=iu(5.7), lh=1.45) + iu(0.14)
    yy = iu(4.02)
    for e in z["limits"]:
        yy = text(p, iu(6.9), yy, e, 14, rgb("#D2D2D2"), REG, maxw=iu(5.6), lh=1.45) + iu(0.14)
    box(p, M, iu(6.45), W - 2 * M, iu(0.72), fill=ACC)
    text(p, M + iu(0.3), iu(6.62), "СТАТУС: " + z["status"].upper(), 15, rgb("#111111"), BOLD, maxw=iu(11.0))
    return d


def sheet(doc, name):
    """4 слайда в сетке 2x2 + подписи."""
    CW, CH = 2, 2
    pad, cap = 34, 74
    tw_, th_ = (W / CW) * 1.0, (H / CH) * 1.0
    k = 0.62
    cw, ch = tw_ * k, th_ * k
    out = fitz.open()
    PW = int(pad + CW * (cw + pad))
    PH = int(pad + CH * (ch + cap + pad) + 90)
    pg = out.new_page(width=PW, height=PH)
    pg.draw_rect(fitz.Rect(0, 0, PW, PH), color=None, fill=rgb("#F0F0F0"))
    pg.insert_font(fontname="f" + str(abs(hash(BOLD)) % 999), fontfile=BOLD)
    for i in range(min(4, doc.page_count)):
        r, c = divmod(i, CW)
        x = pad + c * (cw + pad); y = 100 + r * (ch + cap + pad)
        pix = doc[i].get_pixmap(dpi=72)
        tmp = tempfile.mktemp(suffix=".png"); pix.save(tmp)
        pg.insert_image(fitz.Rect(x, y, x + cw, y + ch), filename=tmp)
        pg.draw_rect(fitz.Rect(x, y, x + cw, y + ch), color=rgb("#C6C6C6"), width=1)
    return out, PW, PH


if __name__ == "__main__":
    os.makedirs("design", exist_ok=True)
    specs = [("A", "paper", "Бумага: светлый фон, серифный заголовок, тонкие линейки"),
             ("B", "swiss", "Швейцарская сетка: белый лист, жирные правила, синий акцент"),
             ("C", "report", "Годовой отчёт: тёмно-синий, кремовый сериф, золото"),
             ("D", "signal", "Сигнал: чёрный, оранжевый, огромные цифры, фото во всю ширину")]
    for key, slug, desc in specs:
        doc = {"A": variant_a, "B": variant_b, "C": variant_c, "D": variant_d}[key]()
        out = fitz.open()
        pad, cap = 34, 26
        cw, ch = W * 0.62, H * 0.62
        PW = int(pad + 2 * (cw + pad)); PH = int(pad + 2 * (ch + pad) + 56)
        pg = out.new_page(width=PW, height=PH)
        pg.draw_rect(fitz.Rect(0, 0, PW, PH), color=None, fill=rgb("#E8E8E8"))
        pg.insert_font(fontname="f" + str(abs(hash(BOLD)) % 999), fontfile=BOLD)
        for i in range(4):
            r, c = divmod(i, 2)
            x = pad + c * (cw + pad); y = pad + r * (ch + pad)
            tmp = tempfile.mktemp(suffix=".png"); doc[i].get_pixmap(dpi=100).save(tmp)
            pg.insert_image(fitz.Rect(x, y, x + cw, y + ch), filename=tmp)
            pg.draw_rect(fitz.Rect(x, y, x + cw, y + ch), color=rgb("#BDBDBD"), width=1)
        pg.insert_text((pad, PH - 22), f"ВАРИАНТ {key} · {desc}", fontsize=15,
                       fontname="f" + str(abs(hash(BOLD)) % 999), color=rgb("#222222"))
        pg.get_pixmap(dpi=72).save(f"design/{key}_{slug}.png")
        print("design/%s_%s.png" % (key, slug))
