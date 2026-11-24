# -*- coding: utf-8 -*-
"""Build all 6 deliverable decks + registries from the 47 source parts."""
import sys, os, re, json
sys.path.insert(0, "/tmp")
from lxml import etree
from pptx_engine import DeckAssembler, NS, replace_text_in_slide, make_textbox_sp, next_sp_id
import slide_builders as SB

SRC = "/tmp/src_pptx"
EXTRACTED = "/tmp/extracted"
OUT = "/home/user/pr123/deliverables"
os.makedirs(OUT, exist_ok=True)
os.makedirs(os.path.join(OUT, "registry"), exist_ok=True)

# ----------------------------------------------------------------------------
# 1. Source part tables: deck -> list of (start_page, end_page, filename)
# ----------------------------------------------------------------------------
def part_name(deck, i, a, b):
    return f"{deck}__часть_{i:02d}__страницы_{a:03d}-{b:03d}.pptx"

DECKS = {
    "Alatau":   ("Алатау___презентация_на_русском", [(1,6),(7,12),(13,14)]),
    "V8":       ("Sokol_V8", [(1,33),(34,60),(61,72)]),
    "Animated": ("Sokol_8_1_Animated_RU", [(1,7),(8,18),(19,24),(25,34),(35,41),(42,48),(49,54),(55,61),(62,69),(70,73),(74,74)]),
    "Hub100":   ("Sokol_8_1_Hub100_RU", [(1,9),(10,21),(22,29),(30,36),(37,41),(42,47),(48,50)]),
    "Investor": ("Sokol_8_1_Investor_RU", [(1,9),(10,19),(20,27),(28,34),(35,39),(40,45),(46,48)]),
    "Park":     ("Sokol_8_1_Park_RU", [(1,3),(4,5),(6,8),(9,12),(13,17),(18,21),(22,26),(27,29),(30,34),(35,36),(37,38),(39,40),(41,42),(43,44),(45,47),(48,50)]),
}

def src_ref(deck, page):
    """Return (abs_path, local_idx, part_file) for a 1-based global page."""
    prefix, parts = DECKS[deck]
    for i, (a, b) in enumerate(parts, start=1):
        if a <= page <= b:
            fn = part_name(prefix, i, a, b)
            return os.path.join(SRC, fn), page - a, fn
    raise KeyError(f"{deck} page {page} out of range")

# ----------------------------------------------------------------------------
# 2. Original speaker notes from /tmp/extracted/*.txt
# ----------------------------------------------------------------------------
def load_notes():
    table = {}
    file_map = {
        "Alatau": "Алатау___презентация_на_русском.txt",
        "V8": "Sokol_V8.txt",
        "Animated": "Sokol_8_1_Animated_RU.txt",
        "Hub100": "Sokol_8_1_Hub100_RU.txt",
        "Investor": "Sokol_8_1_Investor_RU.txt",
        "Park": "Sokol_8_1_Park_RU.txt",
    }
    for deck, fn in file_map.items():
        with open(os.path.join(EXTRACTED, fn)) as f:
            blocks = f.read().split("--- Slide ")[1:]
        for b in blocks:
            page = int(b[:3])
            notes = ""
            if "[NOTES]:" in b:
                notes = b.split("[NOTES]:", 1)[1].strip()
            table[(deck, page)] = notes
    return table

ORIG_NOTES = load_notes()

# ----------------------------------------------------------------------------
# 3. Patch / mutation helpers
# ----------------------------------------------------------------------------
def para_text(p_el):
    return "".join((t.text or "") for t in p_el.findall(".//a:t", NS))

def rewrite_paragraph(p_el, new_text):
    """Set paragraph's first run to new_text, drop other runs and <a:br/>."""
    runs = p_el.findall("a:r", NS)
    if not runs:
        return
    first_t = runs[0].find("a:t", NS)
    if first_t is None:
        return
    first_t.text = new_text
    for r in runs[1:]:
        p_el.remove(r)
    for br in p_el.findall("a:br", NS):
        p_el.remove(br)

def patch_para_map(mapping):
    """mapping: {exact_stripped_para_text: new_text}. Returns mutate_fn."""
    def _fn(root):
        hits = []
        for p_el in root.findall(".//a:p", NS):
            key = para_text(p_el).strip()
            if key in mapping:
                rewrite_paragraph(p_el, mapping[key])
                hits.append(key[:40])
        return hits
    return _fn

# --- SEZ patches ---
PATCH_SEZ_INCENTIVES = patch_para_map({
    "По данным проекта, площадка Sokol": "Площадка Sokol планируется",
    "входит в СЭЗ. Льготы применяются": "к получению статуса СЭЗ (предварительно). Льготы применятся",
})
PATCH_SEZ_PARK025 = patch_para_map({
    "По данным проекта, площадка Sokol": "Площадка Sokol планируется",
    "входит в её границы.": "к включению в её границы (предварительно).",
})
PATCH_PARK005 = patch_para_map({
    "к февралю 2024 года, по данным прокуратуры.": "к февралю 2024 года, по публикациям СМИ [предварительно].",
})

# --- Roadmap date patches (exact standalone labels) ---
ROADMAP_LEGEND = ("* Горизонты этапов предварительные [TBD], уточняются по решениям о земле, "
                  "резидентам и техусловиям: этапы 1–2 — 2027–2030, этапы 3–4 — 2030–2035, этап 5 — после 2035.")
DATES_123 = ["2027–2030*", "2027–2030*", "2030–2035*", "2030–2035*", "после 2035*"]

def patch_roadmap_dates(labels):
    """labels: list of 5 exact standalone label texts, in stage order."""
    date_map = {l: d for l, d in zip(labels, DATES_123)}
    def _fn(root):
        hits = []
        for t in root.findall(".//a:t", NS):
            if t.text and t.text.strip() in date_map:
                lbl = t.text.strip()
                t.text = f"{lbl} · {date_map[lbl]}"
                hits.append(lbl)
        # footnote legend
        sptree = root.find("p:cSld/p:spTree", NS)
        sid = next_sp_id(root)
        sptree.append(make_textbox_sp(
            sid, "RoadmapLegend", 428625, 6565000, 9400000, 220000,
            [{"text": ROADMAP_LEGEND, "sz": 900, "bold": False, "color": "B8CACC", "align": "l"}],
            pad=(0, 0, 0, 0)
        ))
        return hits
    return _fn

PATCH_V8_070      = patch_roadmap_dates(["Подготовка", "Первые заводы", "Расширение", "Жилой город", "Интеграция"])
PATCH_STAGES_CITY = patch_roadmap_dates(["Земля и сети", "Первые заводы", "Терминал и кампус", "Жильё и школы", "Городская инфраструктура"])
PATCH_PARK_002    = patch_roadmap_dates(["Партнёры и продукт", "Площадки и сети", "Первые резиденты", "Поставки и сервис", "Рост парка"])
PATCH_STAGES_INV  = patch_roadmap_dates(["Земля и сети", "Первые заводы", "Терминал и кампус", "Жильё и школы", "Расширение на участки № 6–9"])
PATCH_STAGES_HUB  = patch_roadmap_dates(["Земля и сети", "Первые заводы", "Хаб, станция и кампус", "Жильё и школы", "Расширение авиации на № 6"])

# --- Alatau EN line patches ---
PATCH_ALATAU_TITLE = patch_para_map({
    "МЕЖДУНАРОДНАЯ ИНВЕСТИЦИОННАЯ ПРЕЗЕНТАЦИЯ  |  КОНЦЕПЦИЯ ПРОЕКТА":
    "МЕЖДУНАРОДНАЯ ИНВЕСТИЦИОННАЯ ПРЕЗЕНТАЦИЯ  |  КОНЦЕПЦИЯ ПРОЕКТА  ·  EN: INTERNATIONAL INVESTMENT PRESENTATION | PROJECT CONCEPT",
})
def patch_alatau_sources_attr(root):
    sptree = root.find("p:cSld/p:spTree", NS)
    sid = next_sp_id(root)
    sptree.append(make_textbox_sp(
        sid, "AttributionLine", 502920, 6000000, 10500000, 250000,
        [{"text": "Изображения: проектный комплект «Алатау». Картографические материалы при использовании — © OpenStreetMap contributors, Natural Earth.",
          "sz": 1000, "bold": False, "color": "94A3B8", "align": "l"}],
        pad=(0, 0, 0, 0)
    ))

PATCH_ALATAU_CTA = patch_para_map({
    "ИНВЕСТИРОВАТЬ · СТРОИТЬ · ВНЕДРЯТЬ ИННОВАЦИИ · РАСТИ":
    "ИНВЕСТИРОВАТЬ · СТРОИТЬ · ВНЕДРЯТЬ ИННОВАЦИИ · РАСТИ  ·  EN: INVEST · BUILD · INNOVATE · GROW",
})

# --- Caption overlays for textless render/map slides ---
CAPTIONS = {
    ("Animated", 10): ("КАРТА · ГЕОГРАФИЯ", "Евразия: Казахстан и площадка Sokol",
                       "Алматинская область связывает площадку с рынками Китая и Центральной Азии · © OpenStreetMap, Natural Earth"),
    ("Animated", 11): ("КАРТА · КОРИДОРЫ", "Алтынколь — Хоргос и Транскаспийский маршрут",
                       "Железнодорожный переход KZ–CN и морская переправа через Каспий · © OpenStreetMap, Natural Earth"),
    ("Animated", 12): ("КАРТА · ДОСТУПНОСТЬ", "Трасса Алматы — Астана и дальние коридоры",
                       "Грузовой въезд на площадку требует отдельного проекта · © OpenStreetMap, Natural Earth, Esri"),
    ("Animated", 16): ("КАРТА · ЖЕЛЕЗНАЯ ДОРОГА", "Ж/д коридор в районе станции Сорбулак",
                       "Точка примыкания терминала согласуется с КТЖ [требует согласования] · © OpenStreetMap"),
    ("Animated", 25): ("КАРТА · ВНЕШНИЕ РЫНКИ", "Выход грузов в Центральную Азию, ЕАЭС и Европу",
                       "Маршрут зависит от коридора и действующих ограничений · © OpenStreetMap, Natural Earth"),
    ("Hub100", 11): ("КАРТА · ГЕОГРАФИЯ", "Алматинская область: направление Алматы — Астана",
                     "Площадка связывает рынки Китая и Центральной Азии · © OpenStreetMap, Natural Earth, Esri"),
    ("Hub100", 12): ("КАРТА · КОРИДОРЫ", "Алтынколь — Хоргос и Транскаспийский маршрут",
                     "Железнодорожный переход KZ–CN и морская переправа через Каспий · © OpenStreetMap, Natural Earth"),
    ("Hub100", 13): ("КАРТА · ЛОКАЦИЯ", "76-й км трассы Алматы — Астана",
                     "Автодром Sokol указывает адрес; грузовой въезд в парк предстоит спроектировать · © OpenStreetMap, Natural Earth, Esri"),
    ("Hub100", 16): ("КАРТА · ЖЕЛЕЗНАЯ ДОРОГА", "Ж/д коридор в районе станции Сорбулак",
                     "Точка примыкания терминала согласуется с КТЖ [требует согласования] · © OpenStreetMap"),
    ("Hub100", 23): ("КАРТА · ТРАНСКАСПИЙ", "Перегрузка груза в порту для отправки за Каспий",
                     "Маршрут зависит от товара, покупателя и условий транзита · © OpenStreetMap, Natural Earth"),
    ("Investor", 11): ("КАРТА · ГЕОГРАФИЯ", "Алматинская область: направление Алматы — Астана",
                       "Площадка связывает рынки Китая и Центральной Азии · © OpenStreetMap, Natural Earth, Esri"),
    ("Investor", 12): ("КАРТА · КОРИДОРЫ", "Алтынколь — Хоргос и Транскаспийский маршрут",
                       "Железнодорожный переход KZ–CN и морская переправа через Каспий · © OpenStreetMap, Natural Earth"),
    ("Investor", 13): ("КАРТА · ЛОКАЦИЯ", "76-й км трассы Алматы — Астана",
                       "Автодром Sokol указывает адрес; грузовой въезд в парк предстоит спроектировать · © OpenStreetMap, Natural Earth, Esri"),
    ("Investor", 16): ("КАРТА · ЖЕЛЕЗНАЯ ДОРОГА", "Ж/д коридор в районе станции Сорбулак",
                       "Точка примыкания терминала согласуется с КТЖ [требует согласования] · © OpenStreetMap"),
    ("Investor", 21): ("КАРТА · ТРАНСКАСПИЙ", "Перегрузка груза в порту для отправки за Каспий",
                       "Маршрут зависит от товара, покупателя и условий транзита · © OpenStreetMap, Natural Earth"),
}
# Notes-only alt descriptions (slides already carry some text)
NOTES_ALT = {
    ("Park", 3):  "Описание изображения: обзорная карта Алматинской области с положением площадки у автодрома Sokol (Илийский район, 76-й км трассы Алматы — Астана). © OpenStreetMap, Natural Earth, Esri.",
    ("Park", 32): "Описание изображения: схема международных грузовых коридоров — ж/д выход в Китай через Алтынколь и Достык, Транскаспийский маршрут с морской переправой. © OpenStreetMap, Natural Earth.",
    ("Park", 33): "Описание изображения: автомобильные маршруты к Sokol по трассе Алматы — Астана (76-й км). Грузовой въезд в парк предстоит спроектировать. © OpenStreetMap, Natural Earth, Esri.",
}

# Scrubbed notes for legally sensitive slides
CUSTOM_NOTES = {
    ("Park", 5): ("Хронология площадки автодрома Sokol: 2011 — презентация проекта в акимате; 2012 — начало работ на площадке; "
                  "2020–2022 — строительство большого кольца по портфолио проектировщика Tilke (длина 4 495 м, 13 поворотов). "
                  "Накопленные вложения ТОО «СТК Sokol» в спортивно-технический комплекс — свыше 20 млрд тенге к февралю 2024 года "
                  "по публикациям СМИ [предварительно; к бюджету индустриального парка не относится].\n"
                  "Источники: inform.kz (a2352588) · tilke.de, Portfolio 2024, стр. 19 · tengrinews.kz (217315)."),
    ("Investor", 9): ("Статус СЭЗ для площадки Sokol — ПЛАНИРУЕМЫЙ [предварительно]: включение территории, перечень приоритетных "
                      "видов деятельности и сроки определяются правовым решением уполномоченных органов РК. Льготы участников "
                      "(КПН −100% по приоритетной деятельности; земельный налог, налог на имущество и плата за землю −100%; "
                      "сроки по категориям проектов A/B/C — до 7/15/25 лет) применяются ПРИ ПОЛУЧЕНИИ СТАТУСА СЭЗ и соблюдении условий режима.\n"
                      "Налоговый кодекс РК, ст. 734–735; Закон РК о СЭЗ, ст. 51."),
    ("Hub100", 9): ("Статус СЭЗ для площадки Sokol — ПЛАНИРУЕМЫЙ [предварительно]: включение территории, перечень приоритетных "
                    "видов деятельности и сроки определяются правовым решением уполномоченных органов РК. Льготы участников "
                    "(КПН −100% по приоритетной деятельности; земельный налог, налог на имущество и плата за землю −100%; "
                    "сроки по категориям проектов A/B/C — до 7/15/25 лет) применяются ПРИ ПОЛУЧЕНИИ СТАТУСА СЭЗ и соблюдении условий режима.\n"
                    "Налоговый кодекс РК, ст. 734–735; Закон РК о СЭЗ, ст. 51."),
}

ROADMAP_NOTE = ("Диапазоны дат этапов добавлены редакцией как предварительные [TBD]: этапы 1–2 — 2027–2030, "
                "этапы 3–4 — 2030–2035, этап 5 — после 2035. Уточняются по решениям о земле, резидентах и техническим условиям.")

# ----------------------------------------------------------------------------
# 4. Plan entry constructors
# ----------------------------------------------------------------------------
F = "факт"; P = "предварительно"; S = "требует согласования"

def clone(deck, page, status=P, mut=None, block="", note_extra=""):
    return {"kind": "clone", "deck": deck, "page": page, "status": status,
            "mut": mut or [], "block": block, "note_extra": note_extra}

def custom(builder, bg, title, status=P, block="", note_extra=""):
    return {"kind": "custom", "builder": builder, "bg": bg, "title": title,
            "status": status, "block": block, "note_extra": note_extra}

SOKOL_BG = "0B202B"
ALATAU_BG = "0F1C2E"

def sokol_summary(mode):
    return custom(lambda r: SB.build_sokol_unified_summary(r, scenario_mode=mode), SOKOL_BG,
                  "Суть проекта в одном слайде (Executive Summary)", P, "02 · Суть проекта",
                  "Новый сводный слайд, собран редакцией: ключевые показатели 612,36 га / 500 га / 952 га и три сценария "
                  "компоновки (Park / Investor / Hub100) сведены в одну таблицу для прозрачного согласования. [предварительно]")

def sokol_econ():
    return custom(SB.build_sokol_economics, SOKOL_BG,
                  "Экономика проекта и инвестиционные параметры", P, "10 · Экономика локализации",
                  "Новый слайд редакции. CAPEX, рабочие места, население и грузооборот — структурированные плейсхолдеры "
                  "[Уточняется инициатором / TBD]; бенчмарк Duisburg Gateway Terminal — по данным DGT и duisport.")

def sokol_sources():
    return custom(SB.build_sokol_sources_legal, SOKOL_BG,
                  "Источники и правовые оговорки", F, "13 · Источники и оговорки",
                  "Новый слайд редакции: статус СЭЗ (планируемый), земля/генплан (не ПДП), согласования с КТЖ и сетевыми "
                  "организациями, атрибуции DGT/duisport, OpenStreetMap/Natural Earth/OSRM/Esri, AutoFlight "
                  "(© Tengrinews.kz / Болат Айтмолда; съёмка 24KZ; видео производителя), отказ от подтверждения брендов.")

def sokol_contacts():
    return custom(SB.build_sokol_contacts, SOKOL_BG,
                  "Контакты и приглашение к партнёрству (RU/EN/KZ)", P, "14 · Контакты",
                  "Новый слайд редакции. Контакты дирекции — плейсхолдеры [Уточняется инициатором / TBD]. "
                  "Дубль ключевого оффера на EN и KZ.")

# ----------------------------------------------------------------------------
# 5. Deck plans
# ----------------------------------------------------------------------------
def plan_master():
    B = lambda s: f"{s}"
    plan = [
        clone("V8", 1, P, block="01 · Титул"),
        sokol_summary("master"),
        clone("V8", 2, P, block="02 · Суть проекта"),
        clone("V8", 3, P, block="02 · Суть проекта"),
        clone("V8", 4, P, block="03 · География и транспорт"),
        clone("V8", 7, F, block="03 · География и транспорт"),
        clone("V8", 8, F, block="03 · География и транспорт"),
        clone("V8", 9, F, block="03 · География и транспорт"),
        clone("Park", 5, F, mut=[PATCH_PARK005], block="03 · География и транспорт",
              note_extra="РЕДАКЦИОННАЯ ПРАВКА: ссылка на госорган в строке об инвестициях заменена на «по публикациям СМИ [предварительно]»; соответствующий URL удалён из заметок."),
        clone("V8", 11, P, block="03 · География и транспорт"),
        clone("V8", 12, P, block="04 · Земля и генплан"),
        clone("Park", 6, P, block="04 · Земля и генплан"),
        clone("Park", 7, P, block="04 · Земля и генплан"),
        clone("Park", 8, P, block="04 · Земля и генплан"),
        clone("Park", 9, P, block="04 · Земля и генплан"),
        clone("V8", 13, P, block="04 · Земля и генплан"),
        clone("V8", 24, P, block="05 · Кластеры"),
        clone("V8", 25, P, block="05 · Кластеры"),
        clone("V8", 26, P, block="05 · Кластеры"),
        clone("V8", 27, P, block="05 · Кластеры"),
        clone("Animated", 28, P, block="05 · Кластеры",
              note_extra="Иллюстративный сценарий ИИ: окно разгрузки 13:00–15:30 при задержке поезда. Числа условные."),
        clone("V8", 28, P, block="05 · Кластеры"),
        clone("V8", 29, P, block="05 · Кластеры"),
        clone("V8", 30, P, block="05 · Кластеры"),
        clone("V8", 31, F, block="05 · Кластеры",
              note_extra="Сохранена атрибуция: © Tengrinews.kz / Болат Айтмолда; съёмка 24KZ; видео производителя."),
        clone("V8", 32, P, block="05 · Кластеры"),
        clone("V8", 33, P, block="05 · Кластеры"),
        clone("V8", 34, F, block="05 · Кластеры"),
        clone("V8", 35, P, block="05 · Кластеры"),
        clone("V8", 36, P, block="05 · Кластеры"),
        clone("V8", 37, P, block="05 · Кластеры"),
        clone("V8", 38, P, block="05 · Кластеры"),
        clone("V8", 39, P, block="05 · Кластеры"),
        clone("V8", 40, P, block="05 · Кластеры"),
        clone("V8", 41, P, block="05 · Кластеры"),
        clone("V8", 42, S, block="05 · Кластеры"),
        clone("Park", 23, P, block="05 · Кластеры"),
        clone("Park", 24, S, block="05 · Кластеры"),
        clone("V8", 43, P, block="05 · Кластеры"),
        clone("V8", 14, S, block="06 · Логистика"),
        clone("V8", 15, P, block="06 · Логистика"),
        clone("V8", 16, P, block="06 · Логистика"),
        clone("Hub100", 19, S, block="06 · Логистика"),
        clone("V8", 17, P, block="06 · Логистика"),
        clone("V8", 18, P, block="06 · Логистика"),
        clone("V8", 19, P, block="06 · Логистика"),
        clone("V8", 20, P, block="06 · Логистика"),
        clone("V8", 21, P, block="06 · Логистика"),
        clone("V8", 22, P, block="06 · Логистика"),
        clone("V8", 23, F, block="06 · Логистика"),
        clone("Hub100", 17, F, block="06 · Логистика",
              note_extra="Бенчмарк Duisburg Gateway Terminal: до 850 тыс. TEU/год, ≈23,5 га, 6→12 путей, >730 м — по данным DGT и duisport."),
        clone("Hub100", 18, P, block="06 · Логистика"),
        clone("Hub100", 21, P, block="06 · Логистика"),
        clone("Park", 26, S, block="06 · Логистика"),
        clone("Park", 27, P, block="06 · Логистика"),
        clone("Park", 29, P, block="06 · Логистика"),
        clone("Park", 30, P, block="06 · Логистика"),
        clone("V8", 44, P, block="07 · Город и образование"),
        clone("V8", 45, P, block="07 · Город и образование"),
        clone("V8", 46, P, block="07 · Город и образование"),
        clone("V8", 47, P, block="07 · Город и образование"),
        clone("V8", 48, P, block="07 · Город и образование"),
        clone("V8", 49, P, block="07 · Город и образование"),
        clone("V8", 50, P, block="07 · Город и образование"),
        clone("V8", 51, P, block="07 · Город и образование"),
        clone("V8", 52, P, block="07 · Город и образование"),
        clone("V8", 53, P, block="07 · Город и образование"),
        clone("V8", 54, P, block="07 · Город и образование"),
        clone("V8", 55, P, block="07 · Город и образование"),
        clone("V8", 56, P, block="07 · Город и образование"),
        clone("V8", 57, P, block="07 · Город и образование"),
        clone("Park", 34, P, block="07 · Город и образование"),
        clone("Park", 35, P, block="07 · Город и образование"),
        clone("V8", 58, S, block="08 · Инженерия"),
        clone("V8", 59, P, block="08 · Инженерия"),
        clone("V8", 60, P, block="08 · Инженерия"),
        clone("V8", 61, S, block="08 · Инженерия"),
        clone("V8", 62, P, block="08 · Инженерия"),
        clone("V8", 63, P, block="08 · Инженерия"),
        clone("V8", 64, P, block="08 · Инженерия"),
        clone("V8", 65, P, block="08 · Инженерия"),
        clone("V8", 66, P, block="08 · Инженерия"),
        clone("Park", 47, P, block="08 · Инженерия"),
        clone("V8", 5, P, block="09 · СЭЗ и преференции"),
        clone("V8", 6, P, block="09 · СЭЗ и преференции"),
        clone("Investor", 9, P, mut=[PATCH_SEZ_INCENTIVES], block="09 · СЭЗ и преференции",
              note_extra="РЕДАКЦИОННАЯ ПРАВКА: утверждение о фактическом включении площадки в СЭЗ заменено на «планируется к получению статуса СЭЗ (предварительно)»."),
        clone("Investor", 10, P, block="09 · СЭЗ и преференции",
              note_extra="НДС 0% и СТЗ применяются при получении статуса СЭЗ и соблюдении условий таможенной процедуры."),
        clone("V8", 67, P, block="10 · Экономика локализации"),
        clone("V8", 68, P, block="10 · Экономика локализации"),
        clone("Animated", 69, P, block="10 · Экономика локализации",
              note_extra="Маршрут заказа на партию шин: приёмка → образец → лаборатория → заезды на автодроме → отгрузка."),
        sokol_econ(),
        clone("V8", 70, P, mut=[PATCH_V8_070], block="11 · Дорожная карта", note_extra=ROADMAP_NOTE),
        clone("V8", 69, P, block="12 · Предложение партнёрам",
              note_extra="Xiaomi Auto, MG, VOYAH — только иллюстративные примеры; участие в проекте не подтверждено."),
        clone("V8", 71, P, block="12 · Предложение партнёрам"),
        sokol_sources(),
        sokol_contacts(),
        clone("V8", 72, P, block="14 · Контакты",
              note_extra="Англоязычный дубль титульного оффера (EN title): SOKOL INDUSTRIAL PARK · 612.36 HECTARES · KAZAKHSTAN × CHINA."),
    ]
    return plan

def plan_variant_cut(deck, roadmap_page, roadmap_patch, sez_pages, extra_note_pages=None):
    """Full variant deck in original order (checklist numbering preserved) + appended custom blocks."""
    total = {"Investor": 48, "Hub100": 50, "Park": 50, "Animated": 74}[deck]
    mut_by_page = {}
    if roadmap_page:
        mut_by_page.setdefault(roadmap_page, []).append(roadmap_patch)
    if deck in ("Investor", "Hub100"):
        mut_by_page.setdefault(9, []).append(PATCH_SEZ_INCENTIVES)
    if deck == "Park":
        mut_by_page.setdefault(5, []).append(PATCH_PARK005)
        mut_by_page.setdefault(25, []).append(PATCH_SEZ_PARK025)
    plan = []
    for pg in range(1, total + 1):
        status = P
        ne = ""
        if roadmap_page and pg == roadmap_page:
            ne = ROADMAP_NOTE
        if (deck, pg) in CUSTOM_NOTES:
            ne = (ne + "\n" if ne else "") + "РЕДАКЦИОННАЯ ПРАВКА: см. корректировку формулировок о статусе СЭЗ / атрибуции."
        if (deck, pg) in CAPTIONS:
            ne = (ne + "\n" if ne else "") + "Слайд-визуал без текста: добавлена видимая подпись (caption bar) с описанием и атрибуцией картографии."
        if (deck, pg) in NOTES_ALT:
            ne = (ne + "\n" if ne else "") + NOTES_ALT[(deck, pg)]
        if deck == "Animated" and pg == 71:
            ne = (ne + "\n" if ne else "") + "Xiaomi Auto, MG, VOYAH — только иллюстративные примеры; участие в проекте не подтверждено."
        if deck == "Hub100" and pg == 17:
            status = F
            ne = (ne + "\n" if ne else "") + "Бенчмарк Duisburg Gateway Terminal — по данным DGT и duisport."
        if (deck, pg) in (("Animated", 33),):
            ne = (ne + "\n" if ne else "") + "Сохранена атрибуция: © Tengrinews.kz / Болат Айтмолда; съёмка 24KZ; видео производителя."
        plan.append(clone(deck, pg, status, mut=mut_by_page.get(pg, []), block="", note_extra=ne))
    plan.append(sokol_econ())
    plan.append(sokol_sources())
    plan.append(sokol_contacts())
    return plan

def plan_alatau():
    plan = [
        clone("Alatau", 1, P, mut=[PATCH_ALATAU_TITLE], block="01 · Титул",
              note_extra="Добавлена английская строка-дубль заголовка. Номер форума и спикер — [TBD], см. слайд «Источники»."),
        clone("Alatau", 2, P, block="02 · Суть проекта"),
        clone("Alatau", 3, P, block="04 · Земля и активы",
              note_extra="Пересечение прав 550 га и плана 500 га подлежит проверке по кадастровым данным."),
        clone("Alatau", 4, F, block="03 · География и контекст"),
        clone("Alatau", 5, P, block="04 · Земля и генплан",
              note_extra="Схема планировки 500 га предварительная, статуса утверждённого ПДП не имеет."),
        clone("Alatau", 6, P, block="05 · Кластеры",
              note_extra="270 га: ИИ 75 · робототехника 65 · БПЛА 50 · биомед/фарма 45 · электрооборудование 35."),
        clone("Alatau", 7, S, block="07 · Город и соцобъекты",
              note_extra="БЛОК ERG: публикация параметров участия ERG (больница 15 га, школа 22 га, экспо 5 га, муниципальные работы) "
                         "только после официального письменного одобрения ERG и заключения договоров. Оговорка сохранена на слайде."),
        clone("Alatau", 8, P, block="12 · Предложение партнёрам"),
        clone("Alatau", 9, F, block="09 · Правовой режим",
              note_extra="Конституционный закон № 286-VIII от 08.05.2026; «одно окно» — Alatau City Authority. Льготы подтверждаются отдельно."),
        custom(SB.build_alatau_economics, ALATAU_BG, "Экономика проекта: финансы и рабочие места", P, "10 · Экономика и занятость",
               "Новый слайд редакции: CAPEX, рабочие места, население и налоговый вклад — плейсхолдеры [Уточняется инициатором / TBD]."),
        clone("Alatau", 10, P, block="11 · Дорожная карта",
              note_extra="Этапы: 2027–2030 (пилот) · 2030–2035 (расширение) · после 2035 (полная интеграция). График уточняется."),
        clone("Alatau", 11, P, block="12 · Предложение партнёрам"),
        custom(SB.build_alatau_team_partners, ALATAU_BG, "Команда, институты развития и партнёры", S, "12 · Команда и партнёры",
               "Новый слайд редакции. Состав партнёров и спикер форума — [TBD]; участие ERG — только после одобрения ERG."),
        clone("Alatau", 12, P, block="02 · Видение"),
        clone("Alatau", 13, P, mut=[PATCH_ALATAU_CTA], block="12 · Призыв к действию",
              note_extra="Добавлена английская строка-дубль оффера (EN)."),
        clone("Alatau", 14, F, mut=[patch_alatau_sources_attr], block="13 · Источники и оговорки",
              note_extra="Ссылки: adilet.zan.kz (Z2600000286), alataucity.kz. Заглушки форума/спикера помечены как уточняемые."),
        custom(SB.build_alatau_contacts, ALATAU_BG, "Контакты (RU/EN)", P, "14 · Контакты",
               "Новый слайд редакции. Контакты проектного офиса — плейсхолдеры [Уточняется инициатором / TBD]; дубль на EN."),
    ]
    return plan

# ----------------------------------------------------------------------------
# 6. Assembly
# ----------------------------------------------------------------------------
def compose_notes(entry, part_file=None, local_idx=None):
    lines = []
    if entry["kind"] == "clone":
        deck, page = entry["deck"], entry["page"]
        if (deck, page) in CUSTOM_NOTES:
            orig = CUSTOM_NOTES[(deck, page)]
        else:
            orig = ORIG_NOTES.get((deck, page), "")
        prefix, _ = DECKS[deck]
        lines = [
            f"ИСТОЧНИК: {prefix} — страница {page:03d} (файл части: {os.path.basename(part_file)}, слайд {local_idx + 1} в части).",
            f"СТАТУС: {entry['status']}.",
        ]
        if entry.get("note_extra"):
            lines.append(entry["note_extra"])
        if orig:
            lines.append("--- Исходные заметки ---")
            lines.append(orig)
        return "\n".join(lines)
    else:
        lines = [
            f"ИСТОЧНИК: новый слайд, собран редакцией из сведённых данных проекта (см. реестр слайдов).",
            f"СТАТУС: {entry['status']}.",
        ]
        if entry.get("note_extra"):
            lines.append(entry["note_extra"])
        return "\n".join(lines)

def build_deck(name, plan, template_deck, strip_video, badge_style, deck_tag, out_path,
               idx_offset=0, total_override=None):
    template_path, _, _ = src_ref(template_deck, 1)
    asm = DeckAssembler(template_path,
                        default_dark=(ALATAU_BG if badge_style == "alatau" else SOKOL_BG),
                        default_light="F7F5F0",
                        strip_video=strip_video)
    total = total_override if total_override else len(plan)
    hit_log = []
    for i, entry in enumerate(plan, start=1):
        i = i + idx_offset
        if entry["kind"] == "clone":
            path, local_idx, part_file = src_ref(entry["deck"], entry["page"])
            notes = compose_notes(entry, part_file, local_idx)
            patches = list(entry["mut"])
            caption = CAPTIONS.get((entry["deck"], entry["page"]))
            is_cover = (i == 1) or (entry["deck"], entry["page"]) in (("V8", 72), ("Animated", 74),
                                                                      ("Hub100", 50), ("Investor", 48), ("Park", 50))
            def make_mut(patches=patches, caption=caption, i=i, total=total, is_cover=is_cover,
                         badge_style=badge_style, deck_tag=deck_tag, key=(entry["deck"], entry["page"])):
                def _mut(root):
                    for p in patches:
                        h = p(root)
                        if h:
                            hit_log.append((key, h))
                    if badge_style == "sokol":
                        SB.update_sokol_header_and_pagenum(root, i, total, deck_tag=deck_tag, is_cover_or_end=is_cover)
                    elif badge_style == "alatau":
                        SB.update_alatau_footer_and_pagenum(root, i, total, is_cover=(i == 1))
                    if caption:
                        SB.add_render_caption_overlay(root, caption[0], caption[1], caption[2], i, total, deck_tag=deck_tag)
                return _mut
            asm.add_cloned_slide(path, local_idx, notes_text=notes, mutate_fn=make_mut())
        else:
            notes = compose_notes(entry)
            builder = entry["builder"]
            def make_build(builder=builder, i=i, total=total, badge_style=badge_style, deck_tag=deck_tag):
                def _b(root):
                    builder(root)
                    if badge_style == "sokol":
                        SB.update_sokol_header_and_pagenum(root, i, total, deck_tag=deck_tag, is_cover_or_end=False)
                    elif badge_style == "alatau":
                        SB.update_alatau_footer_and_pagenum(root, i, total, is_cover=False)
                return _b
            asm.add_custom_slide(entry["bg"], make_build(), notes_text=notes)
    n, size = asm.save(out_path)
    return n, size, asm.stripped_videos, hit_log

def write_registry(reg_path, deck_title, plan, sizes_line):
    rows = []
    rows.append(f"# Реестр слайдов — {deck_title}")
    rows.append("")
    rows.append(f"Собрано автоматически 24.11.2026 · {sizes_line}")
    rows.append("")
    rows.append("| № | Блок структуры | Источник (файл, страница) | Статус | Примечание |")
    rows.append("|--:|----------------|---------------------------|--------|------------|")
    for i, entry in enumerate(plan, start=1):
        if entry["kind"] == "clone":
            prefix, _ = DECKS[entry["deck"]]
            src = f"{prefix}, стр. {entry['page']:03d}"
        else:
            src = "новый слайд редакции"
        block = entry.get("block", "") or "—"
        note = (entry.get("note_extra") or "").split("\n")[0][:110]
        note = note.replace("|", "/")
        rows.append(f"| {i:02d} | {block} | {src} | {entry['status']} | {note} |")
    rows.append("")
    with open(reg_path, "w") as f:
        f.write("\n".join(rows))

# ----------------------------------------------------------------------------
# 7. Build all decks
# ----------------------------------------------------------------------------
RESULTS = []

def run_deck(key, out_name, reg_name, title, plan, template_deck, strip_video, badge_style, deck_tag):
    out_path = os.path.join(OUT, out_name)
    print(f"\n=== Building {key} -> {out_name} ({len(plan)} slides) ===")
    n, size, stripped, hits = build_deck(key, plan, template_deck, strip_video, badge_style, deck_tag, out_path)
    print(f"    slides={n} size={size/1e6:.1f}MB videos_stripped={stripped} patch_hits={len(hits)}")
    for k, h in hits:
        print(f"    patch hit {k}: {h}")
    write_registry(os.path.join(OUT, "registry", reg_name), title, plan,
                   f"{n} слайдов · файл {out_name} · {size/1e6:.1f} МБ")
    RESULTS.append((key, out_name, n, size))
    return out_path

if __name__ == "__main__":
    which = sys.argv[1] if len(sys.argv) > 1 else "all"

    if which in ("all", "alatau"):
        run_deck("Alatau", "Alatau_SmartCity_RU.pptx", "REESTR_Alatau.md",
                 "Алатау — умный город 500 га (обновлённая редакция)", plan_alatau(), "Alatau", False, "alatau", "ALATAU")

    if which in ("all", "park"):
        run_deck("Park", "SOKOL_Park_RU.pptx", "REESTR_SOKOL_Park.md",
                 "SOKOL · Генплан парка (вырезка Park)", plan_variant_cut("Park", 2, PATCH_PARK_002, [25]),
                 "Park", True, "sokol", "SOKOL · ГЕНПЛАН")

    if which in ("all", "investor"):
        run_deck("Investor", "SOKOL_Investor_RU.pptx", "REESTR_SOKOL_Investor.md",
                 "SOKOL · Инвесторский оффер (вырезка Investor)", plan_variant_cut("Investor", 46, PATCH_STAGES_INV, [9]),
                 "Investor", True, "sokol", "SOKOL · ИНВЕСТОРУ")

    if which in ("all", "hub100"):
        run_deck("Hub100", "SOKOL_Hub100_RU.pptx", "REESTR_SOKOL_Hub100.md",
                 "SOKOL · Логистический хаб 100 га (вырезка Hub100)", plan_variant_cut("Hub100", 48, PATCH_STAGES_HUB, [9]),
                 "Hub100", True, "sokol", "SOKOL · ХАБ 100 ГА")

    if which in ("all", "master"):
        run_deck("Master", "SOKOL_Master_RU.pptx", "REESTR_SOKOL_Master.md",
                 "SOKOL INDUSTRIAL PARK — мастер-презентация", plan_master(), "V8", True, "sokol", "SOKOL INDUSTRIAL PARK")

    if which in ("all", "animated"):
        run_deck("Animated", "SOKOL_Animated_RU.pptx", "REESTR_SOKOL_Animated.md",
                 "SOKOL · Витринная анимированная версия", plan_variant_cut("Animated", 72, PATCH_STAGES_CITY, []),
                 "Animated", False, "sokol", "SOKOL INDUSTRIAL PARK")

    print("\n=== SUMMARY ===")
    for key, out_name, n, size in RESULTS:
        print(f"  {key:10s} {out_name:30s} {n:3d} slides {size/1e6:7.1f} MB")
