"""Единая презентация: основная колода (19) + приложение (103) в одном файле."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import deck2 as D                      # строит 19 слайдов основной колоды
import appendix as A                   # палитра, хелперы, ZONES 01–06, SECTIONS
import appendix2 as A2                 # зоны 07–24, фото, builders приложения

MAIN = D.prs
A.prs = MAIN                           # хелперы appendix теперь рисуют в общем файле
A2.prs = MAIN

sld, rect, txt, hair = A.sld, A.rect, A.txt, A.hair
M, CW = A.M, A.CW
W, MUT, ACC, DIM, SUB, CARD, LINE = A.W, A.MUT, A.ACC, A.DIM, A.SUB, A.CARD, A.LINE


def divider_into_appendix():
    """Шов между колодой и приложением."""
    s = sld()
    txt(s, M, 0.95, CW, 0.28, "Дальше — приложение", 8.5, True, ACC, caps=True)
    txt(s, M, 1.35, 10.2, 1.3, "Двадцать четыре зоны\nпо отдельности", 34, True, W, sp=1.04)
    hair(s, 3.30)
    txt(s, M, 3.56, 6.9, 1.4,
        "Всё, что в колоде сжато до строки, разобрано здесь по одной зоне: состав, площади, "
        "участки, технология, инженерия и границы применимости.", 10.5, False, MUT, sp=1.38)
    txt(s, M, 5.05, 6.9, 0.9,
        "Четыре слайда на зону: обложка → что входит → как это работает → цифры и ограничения.",
        9.4, False, SUB, sp=1.34)
    # правая карточка — четыре раздела
    rect(s, 8.05, 1.55, 4.56, 4.55, CARD)
    txt(s, 8.34, 1.78, 4.0, 0.26, "СОСТАВ ПРИЛОЖЕНИЯ", 7.4, True, ACC, caps=True)
    rows = [("Производство", "6 зон", "221 га"),
            ("Логистика", "5 зон", "127 га"),
            ("Город и наука", "7 зон", "119 га"),
            ("Инженерия и среда", "6 зон", "145,36 га")]
    for i, (nm, cnt, ar) in enumerate(rows):
        y = 2.16 + i * 0.86
        txt(s, 8.34, y, 2.5, 0.3, nm, 10.2, True, W)
        txt(s, 8.34, y + 0.30, 1.6, 0.26, cnt, 8.4, False, MUT)
        txt(s, 10.30, y + 0.02, 2.0, 0.3, ar, 11.4, True, SUB, align=A.PP_ALIGN.RIGHT)
        hair(s, y + 0.62, 8.34, 3.98, LINE)
    txt(s, 8.34, 5.72, 3.98, 0.3, "612,36 га · участки № 2–5", 8.2, False, DIM)
    txt(s, M, 6.82, CW, 0.3, "103 слайда · фото по всем 24 зонам · ноябрь 2026", 8, False, DIM)


def divider_sections_mark():
    pass


divider_into_appendix()

A.contents()
for sec in A.SECTIONS:
    ids = list(sec["ids"])
    A2.sec_div(sec, ids, A2.ZONES)
    for i in ids:
        z = A2.ZONES[i]
        A2.z_cover(z, i, len(A2.ZONES))
        A2.z_comp(z, i, len(A2.ZONES))
        A2.z_flow(z, i, len(A2.ZONES))
        A2.z_facts(z, i, len(A2.ZONES))
A.ending()

MAIN.save("Sokol_единая_презентация.pptx")
n = len(MAIN.slides._sldIdLst)
print("slides:", n, "| колода:", n - 104, "| приложение: 103 + шов")
