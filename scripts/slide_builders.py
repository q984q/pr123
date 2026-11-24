from lxml import etree
from pptx_engine import NS, make_textbox_sp, next_sp_id, replace_text_in_slide

# Palette constants
# Alatau: Dark Slate Navy #0F1C2E, Card #172D48, Gold #E8B66C, Teal #00A899, Text #FFFFFF, Muted #AABBCB
# Sokol:  Deep Petrol #0B202B, Card #132F3D, Gold #E9C28E, Teal #3CC1BF, Text #F8F4ED, Muted #B8CACC

def add_rect(sptree, sp_id, name, x, y, cx, cy, fill_hex=None, border_hex=None, border_w=12700, radius=False):
    geom = 'roundRect' if radius else 'rect'
    av = '<a:gd name="adj" fmla="val 5000"/>' if radius else ''
    fill_xml = f'<a:solidFill><a:srgbClr val="{fill_hex}"/></a:solidFill>' if fill_hex else '<a:noFill/>'
    ln_xml = f'<a:ln w="{border_w}"><a:solidFill><a:srgbClr val="{border_hex}"/></a:solidFill></a:ln>' if border_hex else '<a:ln><a:noFill/></a:ln>'
    el = etree.fromstring(
        f'<p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">'
        f'<p:nvSpPr><p:cNvPr id="{sp_id}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="{geom}"><a:avLst>{av}</a:avLst></a:prstGeom>{fill_xml}{ln_xml}</p:spPr>'
        f'</p:sp>'
    )
    sptree.append(el)

def update_alatau_footer_and_pagenum(slide_root, page_num, total_pages, is_cover=False):
    sptree = slide_root.find("p:cSld/p:spTree", NS)
    # Remove verbose 2-line footer boxes and old page number box
    to_remove = []
    for sp in sptree.findall("p:sp", NS):
        txt = "".join(t.text or "" for t in sp.findall(".//a:t", NS)).strip()
        xfrm = sp.find("p:spPr/a:xfrm", NS)
        y = int(xfrm.find("a:off", NS).get("y", "0")) if (xfrm is not None and xfrm.find("a:off", NS) is not None) else 0
        if "Китайско-казахстанский демонстрационный парк «Алатау»" in txt or "Сотрудничество в сфере зелёной цифровой промышленности" in txt:
            to_remove.append(sp)
        elif y > 6350000 and txt.isdigit() and len(txt) <= 3:
            to_remove.append(sp)
    for sp in to_remove:
        sptree.remove(sp)

    if not is_cover:
        sid = next_sp_id(slide_root)
        # Compact logo-badge in bottom-left
        badge = make_textbox_sp(
            sid, "AlatauLogoBadge", 502920, 6465000, 3200000, 240000,
            [{"text": "ALATAU  ·  SMART CITY 500 HA", "sz": 900, "bold": True, "color": "E8B66C", "align": "l"}],
            fill_hex="172D48", border_hex="00A899", border_w=9525, radius=True, anchor="ctr", pad=(91440, 20000, 91440, 20000)
        )
        sptree.append(badge)
        # Page number bottom-right
        pbox = make_textbox_sp(
            sid + 1, "PageNumber", 10850000, 6465000, 840000, 240000,
            [{"text": f"{page_num:02d} / {total_pages:02d}", "sz": 950, "bold": True, "color": "AABBCB", "align": "r"}],
            anchor="ctr", pad=(0, 0, 0, 0)
        )
        sptree.append(pbox)

def update_sokol_header_and_pagenum(slide_root, page_num, total_pages, deck_tag="SOKOL INDUSTRIAL PARK", is_cover_or_end=False):
    if is_cover_or_end:
        return
    sptree = slide_root.find("p:cSld/p:spTree", NS)
    # Determine if light or dark slide by checking <p:bg> or top text colors
    bg_clr_el = slide_root.find("p:cSld/p:bg//a:srgbClr", NS)
    is_light_slide = (bg_clr_el is not None and bg_clr_el.get("val", "").upper() == "F7F5F0")
    accent_col = "8A683C" if is_light_slide else "E9C28E"
    muted_col  = "48616B" if is_light_slide else "B8CACC"
    badge_bg   = "E9E4DA" if is_light_slide else "132F3D"

    # Replace old "SOKOL / V8" or "XX / 72" or "SOKOL  /  XX" counters
    removed_old_num = False
    for sp in list(sptree.findall("p:sp", NS)):
        txt = "".join(t.text or "" for t in sp.findall(".//a:t", NS)).strip()
        xfrm = sp.find("p:spPr/a:xfrm", NS)
        x = int(xfrm.find("a:off", NS).get("x", "0")) if (xfrm is not None and xfrm.find("a:off", NS) is not None) else 0
        y = int(xfrm.find("a:off", NS).get("y", "0")) if (xfrm is not None and xfrm.find("a:off", NS) is not None) else 0
        if txt in ("SOKOL / V8", "SOKOL INDUSTRIAL PARK") and x > 7000000 and y < 500000:
            sptree.remove(sp)
        elif y < 500000 and x > 9000000 and "SOKOL" in txt and "/" in txt:
            sptree.remove(sp)
        elif y > 6150000 and x > 9800000 and "/" in txt and len(txt) <= 10:
            sptree.remove(sp)
            removed_old_num = True

    sid = next_sp_id(slide_root)
    # Compact top-right brand & slide counter badge
    top_badge = make_textbox_sp(
        sid, "SokolBrandBadge", 9250000, 215000, 2520000, 245000,
        [{"text": f"{deck_tag}  ·  {page_num:02d}/{total_pages:02d}", "sz": 850, "bold": True, "color": accent_col, "align": "r"}],
        fill_hex=None, anchor="ctr", pad=(0, 0, 0, 0)
    )
    sptree.append(top_badge)

def add_render_caption_overlay(slide_root, category_ru, title_ru, subtitle_ru, page_num, total_pages, deck_tag="SOKOL INDUSTRIAL PARK"):
    """Adds a sleek semi-solid bottom editorial caption bar to full-bleed render/video slides that previously had no text."""
    sptree = slide_root.find("p:cSld/p:spTree", NS)
    sid = next_sp_id(slide_root)
    # Bottom dark bar y=5480000..6858000
    add_rect(sptree, sid, "CaptionBackdrop", 0, 5520000, 12192000, 1338000, fill_hex="0B202B", border_hex="3CC1BF", border_w=9525)
    sptree.append(make_textbox_sp(
        sid + 1, "CaptionCategory", 438150, 5620000, 8000000, 220000,
        [{"text": category_ru, "sz": 1000, "bold": True, "color": "E9C28E", "align": "l"}],
        pad=(0, 0, 0, 0)
    ))
    sptree.append(make_textbox_sp(
        sid + 2, "CaptionTitle", 438150, 5860000, 10200000, 420000,
        [{"text": title_ru, "sz": 2000, "bold": True, "color": "F8F4ED", "align": "l"}],
        pad=(0, 0, 0, 0)
    ))
    sptree.append(make_textbox_sp(
        sid + 3, "CaptionSub", 438150, 6310000, 10200000, 380000,
        [{"text": subtitle_ru, "sz": 1150, "bold": False, "color": "B8CACC", "align": "l"}],
        pad=(0, 0, 0, 0)
    ))
    sptree.append(make_textbox_sp(
        sid + 4, "CaptionPage", 10700000, 5620000, 1070000, 240000,
        [{"text": f"{page_num:02d} / {total_pages:02d}", "sz": 1000, "bold": True, "color": "E9C28E", "align": "r"}],
        pad=(0, 0, 0, 0)
    ))

# ==============================================================================
# CUSTOM SLIDES FOR ALATAU
# ==============================================================================

def build_alatau_team_partners(slide_root):
    sptree = slide_root.find("p:cSld/p:spTree", NS)
    sid = 10
    sptree.append(make_textbox_sp(
        sid, "Title", 502920, 340000, 11155680, 440000,
        [{"text": "Команда проекта, институты развития и партнёры (KZ × CN)", "sz": 2400, "bold": True, "color": "FFFFFF"}],
        pad=(0,0,0,0)
    ))
    sptree.append(make_textbox_sp(
        sid+1, "Subtitle", 502920, 790000, 11155680, 280000,
        [{"text": "Координация в режиме «одного окна» · Статус коммерческих соглашений уточняется по каждому участнику [предварительно]", "sz": 1200, "bold": False, "color": "AABBCB"}],
        pad=(0,0,0,0)
    ))
    add_rect(sptree, sid+2, "Divider", 502920, 1180000, 11155680, 12700, fill_hex="35506A")

    cards = [
        ("ГОСУДАРСТВЕННЫЙ КОНТУР РК", "Alatau City Authority и МНЭ РК",
         "• Координация инвестпроектов в режиме «одного окна»\n• Спецправовой режим: Конституционный закон № 286-VIII (08.05.2026)\n• Согласование земли, градостроительства и преференций", "00A899"),
        ("ИНИЦИАТОР И ДЕВЕЛОПМЕНТ", "Проектная компания (SPV) парка «Алатау»",
         "• Мастер-девелопмент территории 500 га (270 / 150 / 80 га)\n• Привлечение соинвесторов по 5 промышленным кластерам\n• Руководитель / спикер форума: [ФИО, должность — TBD]", "E8B66C"),
        ("СОЦИАЛЬНАЯ ИНФРАСТРУКТУРА", "Потенциальное участие ERG [требует одобрения]",
         "• Больница (15 га), международная школа (22 га), экспоцентр (5 га)\n• Муниципальные сети и въездная группа\n• Внимание: параметры публикуются только после письменного одобрения ERG", "E8B66C"),
        ("МЕЖДУНАРОДНЫЕ ПАРТНЁРЫ (КНР)", "Индустриальные и финансовые партнёры",
         "• Технологические компании КНР: ИИ, роботы, БПЛА, биомед, энерго\n• EPC-подрядчики, операторы цифрового города и профильные фонды\n• Статус переговоров и перечень якорных резидентов: [уточняется / TBD]", "00A899"),
    ]
    coords = [
        (502920, 1420000), (6200000, 1420000),
        (502920, 3820000), (6200000, 3820000)
    ]
    for idx, ((tag, heading, body, acc), (x, y)) in enumerate(zip(cards, coords)):
        base_id = sid + 10 + idx * 4
        add_rect(sptree, base_id, f"Card_{idx}", x, y, 5460000, 2180000, fill_hex="172D48", border_hex=acc, border_w=15875, radius=True)
        sptree.append(make_textbox_sp(
            base_id+1, f"CardText_{idx}", x+180000, y+140000, 5100000, 1920000,
            [
                {"text": tag, "sz": 1000, "bold": True, "color": acc, "spcAft": 400},
                {"text": heading, "sz": 1650, "bold": True, "color": "FFFFFF", "spcAft": 600},
                {"text": body, "sz": 1200, "bold": False, "color": "CBD5E1", "lnSpc": 125000},
            ],
            pad=(0,0,0,0)
        ))

    sptree.append(make_textbox_sp(
        sid+40, "Footnote", 502920, 6120000, 11155680, 240000,
        [{"text": "Примечание: наименования партнёров и состав делегации на форуме «Китай — Центральная Азия» фиксируются по официальным соглашениям.", "sz": 1000, "bold": False, "color": "94A3B8"}],
        pad=(0,0,0,0)
    ))

def build_alatau_economics(slide_root):
    sptree = slide_root.find("p:cSld/p:spTree", NS)
    sid = 10
    sptree.append(make_textbox_sp(
        sid, "Title", 502920, 340000, 11155680, 440000,
        [{"text": "Экономика проекта 500 га: инвестиции, занятость и целевые показатели", "sz": 2300, "bold": True, "color": "FFFFFF"}],
        pad=(0,0,0,0)
    ))
    sptree.append(make_textbox_sp(
        sid+1, "Subtitle", 502920, 790000, 11155680, 280000,
        [{"text": "Финансово-экономическая модель формируется по очередям 2027–2035+ · Показатели со статусом [TBD] уточняются инициатором", "sz": 1200, "bold": False, "color": "AABBCB"}],
        pad=(0,0,0,0)
    ))
    add_rect(sptree, sid+2, "Divider", 502920, 1180000, 11155680, 12700, fill_hex="35506A")

    kpis = [
        ("500 га", "Территория мастер-плана", "270 га кластеры · 150 га город · 80 га экозона (права на 550 га уточняются)", "E8B66C"),
        ("[TBD] млн $", "Совокупный CAPEX проекта", "Инфраструктура парка, производственные корпуса и социальные объекты", "00A899"),
        ("[TBD] тыс.", "Новых рабочих мест", "Высокотехнологичная занятость в 5 кластерах ИИ, роботов, БПЛА и биомеда", "E8B66C"),
        ("[TBD] тыс.", "Жителей и специалистов", "Население умного города в городе, учащиеся школы (22 га) и медперсонал", "00A899"),
        ("[TBD] млн $/год", "Налоговый и экспортный вклад", "Прогноз поступлений и несырьевого экспорта после выхода на полную мощность", "E8B66C"),
        ("2027–2035+", "Горизонт реализации (3 этапа)", "Пилот 2027–2030 · Масштабирование 2030–2035 · Полная интеграция после 2035", "00A899"),
    ]
    for i, (big, label, desc, col) in enumerate(kpis):
        col_i = i % 3
        row_i = i // 3
        x = 502920 + col_i * 3780000
        y = 1400000 + row_i * 2220000
        bid = sid + 10 + i * 3
        add_rect(sptree, bid, f"KPI_{i}", x, y, 3590000, 2020000, fill_hex="172D48", border_hex=col, border_w=14000, radius=True)
        sptree.append(make_textbox_sp(
            bid+1, f"KPITxt_{i}", x+160000, y+150000, 3270000, 1750000,
            [
                {"text": big, "sz": 2600, "bold": True, "color": col, "spcAft": 400},
                {"text": label, "sz": 1350, "bold": True, "color": "FFFFFF", "spcAft": 400},
                {"text": desc, "sz": 1050, "bold": False, "color": "CBD5E1", "lnSpc": 120000},
            ],
            pad=(0,0,0,0)
        ))

    sptree.append(make_textbox_sp(
        sid+40, "SummaryNote", 502920, 5920000, 11155680, 420000,
        [{"text": "Формат участия инвесторов: SPV по кластерам, совместные предприятия KZ×CN, прямое размещение производств и девелопмент городской среды. Расчёт CAPEX, IRR и сроков окупаемости выполняется индивидуально под ТЗ резидента.", "sz": 1050, "bold": False, "color": "AABBCB"}],
        pad=(0,0,0,0)
    ))

def build_alatau_contacts(slide_root):
    sptree = slide_root.find("p:cSld/p:spTree", NS)
    sid = 10
    sptree.append(make_textbox_sp(
        sid, "Title", 502920, 340000, 11155680, 440000,
        [{"text": "Контакты для инвесторов и партнёров · Contact Information", "sz": 2400, "bold": True, "color": "FFFFFF"}],
        pad=(0,0,0,0)
    ))
    sptree.append(make_textbox_sp(
        sid+1, "Subtitle", 502920, 790000, 11155680, 280000,
        [{"text": "Китайско-казахстанский демонстрационный парк «Алатау» · China–Kazakhstan Alatau Smart City Demonstration Park", "sz": 1200, "bold": False, "color": "E8B66C"}],
        pad=(0,0,0,0)
    ))
    add_rect(sptree, sid+2, "Divider", 502920, 1180000, 11155680, 12700, fill_hex="35506A")

    # Left card: RU/EN Project Office & One-Stop Shop
    add_rect(sptree, sid+10, "LeftCard", 502920, 1450000, 5450000, 4550000, fill_hex="172D48", border_hex="00A899", border_w=15875, radius=True)
    sptree.append(make_textbox_sp(
        sid+11, "LeftCardText", 720000, 1650000, 5020000, 4200000,
        [
            {"text": "КООРДИНАЦИЯ ИНВЕСТПРОЕКТОВ («ОДНО ОКНО»)", "sz": 1050, "bold": True, "color": "00A899", "spcAft": 500},
            {"text": "Alatau City Authority / Проектный офис «Алатау»", "sz": 1750, "bold": True, "color": "FFFFFF", "spcAft": 800},
            {"text": "Адрес / Location:\nРеспублика Казахстан, Алматинская область, г. Алатау\n[Точный адрес офиса уточняется / TBD]", "sz": 1250, "bold": False, "color": "CBD5E1", "spcAft": 600},
            {"text": "Контактное лицо / Contact Person:\n[ФИО, должность руководителя дирекции — TBD]", "sz": 1250, "bold": False, "color": "CBD5E1", "spcAft": 600},
            {"text": "Телефон / Phone:  [+7 (727) XXX-XX-XX — TBD]\nE-mail:  [invest@alatau-project.kz — TBD]\nОфициальный портал города:  https://alataucity.kz", "sz": 1250, "bold": True, "color": "E8B66C"},
        ],
        pad=(0,0,0,0)
    ))

    # Right card: Next steps EN + RU + QR placeholder
    add_rect(sptree, sid+20, "RightCard", 6210000, 1450000, 5450000, 4550000, fill_hex="172D48", border_hex="E8B66C", border_w=15875, radius=True)
    sptree.append(make_textbox_sp(
        sid+21, "RightCardText", 6430000, 1650000, 5020000, 2750000,
        [
            {"text": "INVESTMENT INQUIRIES & NEXT STEPS (EN / RU)", "sz": 1050, "bold": True, "color": "E8B66C", "spcAft": 500},
            {"text": "How to start cooperation / Порядок старта:", "sz": 1650, "bold": True, "color": "FFFFFF", "spcAft": 600},
            {"text": "1. Запрос параметров участка и техусловий (Industry & Plot Brief)\n2. Совместная проверка в Alatau City Authority (One-Stop Review)\n3. Согласование формата SPV / СП и дорожной карты 2027–2030", "sz": 1200, "bold": False, "color": "CBD5E1", "lnSpc": 125000},
        ],
        pad=(0,0,0,0)
    ))
    # QR Box placeholder
    add_rect(sptree, sid+22, "QRBox", 6430000, 4450000, 1300000, 1300000, fill_hex="0F1C2E", border_hex="00A899", border_w=12700, radius=True)
    sptree.append(make_textbox_sp(
        sid+23, "QRText", 6430000, 4450000, 1300000, 1300000,
        [{"text": "[QR-КОД]\nМатериалы\nпроекта\n(TBD)", "sz": 1050, "bold": True, "color": "00A899", "align": "ctr"}],
        anchor="ctr"
    ))
    sptree.append(make_textbox_sp(
        sid+24, "ENOfferSummary", 7920000, 4500000, 3520000, 1200000,
        [
            {"text": "ALATAU SMART CITY · 500 HA", "sz": 1200, "bold": True, "color": "FFFFFF", "spcAft": 300},
            {"text": "5 AI-Driven Industrial Clusters (270 ha) · Urban Living & Social Core (150 ha) · Eco-Zone (80 ha) · Constitutional Law No. 286-VIII.", "sz": 1050, "bold": False, "color": "AABBCB"},
        ],
        pad=(0,0,0,0)
    ))

# ==============================================================================
# CUSTOM SLIDES FOR SOKOL
# ==============================================================================

def build_sokol_unified_summary(slide_root, scenario_mode="master"):
    """
    Slide 2 'Суть проекта в одном слайде' / Unified Key Metrics & Scenarios
    scenario_mode: 'master' | 'investor' | 'hub100' | 'park' | 'animated'
    """
    sptree = slide_root.find("p:cSld/p:spTree", NS)
    sid = 10
    sptree.append(make_textbox_sp(
        sid, "Section", 428625, 230000, 8000000, 220000,
        [{"text": "СУТЬ ПРОЕКТА В ОДНОМ СЛАЙДЕ  ·  EXECUTIVE SUMMARY", "sz": 1100, "bold": True, "color": "E9C28E"}],
        pad=(0,0,0,0)
    ))
    title_map = {
        "master":   "Индустриальный город Sokol: 612,36 га первой очереди и 952 га резерва",
        "investor": "Индустриальный город Sokol: 500 га для 6 кластеров и компактный хаб 15 га",
        "hub100":   "Индустриальный город Sokol: международный сухой порт 100 га и 6 кластеров",
        "park":     "Генеральный план Sokol: 612,36 га (участки № 2–5), 6 кластеров и резерв 952 га",
        "animated": "Индустриальный город Sokol: 612,36 га — заводы, мультимодальный хаб и жилая среда",
    }
    sptree.append(make_textbox_sp(
        sid+1, "Title", 428625, 520000, 11334750, 520000,
        [{"text": title_map.get(scenario_mode, title_map["master"]), "sz": 2400, "bold": True, "color": "F8F4ED"}],
        pad=(0,0,0,0)
    ))
    sptree.append(make_textbox_sp(
        sid+2, "Subtitle", 428625, 1040000, 11334750, 280000,
        [{"text": "76-й км трассы Алматы — Астана (Илийский р-н) · Партнёрство KZ × CN · Специальный экономический режим — планируемый [предварительно]", "sz": 1150, "bold": False, "color": "B8CACC"}],
        pad=(0,0,0,0)
    ))

    # 4 top metric pills
    pills = [
        ("612,36 га", "Первая очередь (уч. № 2–5)", "Уч. № 2 (62) · № 3 (100) · № 4 (162,36) · № 5 (288 га); автодром № 1 — 205 га отдельно"),
        ("500 га", "6 промышленных кластеров", "Базовый баланс: ИИ и ЦОД 100 · Авиация 80 · Роботы 100 · Электроника 90 · Биомед 60 · Автопром 70 га"),
        ("952 га", "Резерв развития (уч. № 6–9)", "Уч. № 6 (500) · № 7 (93,2) · № 8 (116,8) · № 9 (242 га) — всего 1 564,36 га (уч. № 2–9)"),
        ("32 га + 14 га", "Жилая среда и образование", "Жилой квартал (32 га, буфер ~73 м от ЦОД), технологический университет, школы и клиника"),
    ]
    for i, (num, head, sub) in enumerate(pills):
        x = 428625 + i * 2880000
        y = 1440000
        bid = sid + 10 + i * 3
        add_rect(sptree, bid, f"Pill_{i}", x, y, 2730000, 2050000, fill_hex="132F3D", border_hex="E9C28E" if i % 2 == 0 else "3CC1BF", border_w=14000, radius=True)
        sptree.append(make_textbox_sp(
            bid+1, f"PillTxt_{i}", x+140000, y+130000, 2450000, 1820000,
            [
                {"text": num, "sz": 2300, "bold": True, "color": "E9C28E" if i % 2 == 0 else "3CC1BF", "spcAft": 300},
                {"text": head, "sz": 1250, "bold": True, "color": "F8F4ED", "spcAft": 350},
                {"text": sub, "sz": 980, "bold": False, "color": "B8CACC", "lnSpc": 118000},
            ],
            pad=(0,0,0,0)
        ))

    # Scenario comparison matrix at bottom (resolves §3 contradictions transparently!)
    add_rect(sptree, sid+30, "ScenarioBox", 428625, 3700000, 11334750, 2480000, fill_hex="102835", border_hex="3CC1BF", border_w=12700, radius=True)
    sptree.append(make_textbox_sp(
        sid+31, "ScenarioTitle", 620000, 3830000, 10950000, 320000,
        [{"text": "СЦЕНАРИИ КОМПОНОВКИ ЛОГИСТИКИ И КЛАСТЕРОВ (ПРЕДВАРИТЕЛЬНО, УТОЧНЯЮТСЯ С КТЖ И ЯКОРНЫМИ РЕЗИДЕНТАМИ)", "sz": 1050, "bold": True, "color": "3CC1BF"}],
        pad=(0,0,0,0)
    ))
    cols_txt = [
        ("Сценарий 1 · Базовый генплан (Park)",
         "• Промышленность: 473 га (уч. 2–5) + 27 га инж. сервис (уч. 1) = 500 га\n• Логистика и таможня: 30 га + ж/д зона 20 га (17 груз. + 3 пасс., 12 путей, 1,86 км)\n• Автопром 70 га (43+27) · Авиация 80 га"),
        ("Сценарий 2 · Компактный оффер (Investor)",
         "• Промышленность внутри уч. № 2–5: 500 га\n• Компактный логистический хаб: 15 га (весовая, склады, ж/д фронт 12 путей)\n• Акцент на автопром: 145 га (+30 га испытаний на уч. № 1) · Авиация 100 га"),
        ("Сценарий 3 · Международный хаб (Hub 100)",
         "• Мультимодальный сухой порт: 100 га (ж/д терминал 25 га, склады кл. А 120–150 тыс. м², 4–6 кранов, пути 850–1 050 м) + станция 5 га\n• Кластеры: 418 га (очередь 1) + 82 га (расширение авиации на уч. № 6) = 500 га"),
    ]
    for j, (ch, cb) in enumerate(cols_txt):
        cx_pos = 620000 + j * 3680000
        sptree.append(make_textbox_sp(
            sid+35+j, f"ScenCol_{j}", cx_pos, 4220000, 3520000, 1850000,
            [
                {"text": ch, "sz": 1200, "bold": True, "color": "E9C28E", "spcAft": 350},
                {"text": cb, "sz": 1000, "bold": False, "color": "F8F4ED", "lnSpc": 120000},
            ],
            pad=(0,0,0,0)
        ))

    sptree.append(make_textbox_sp(
        sid+45, "ENFooter", 428625, 6320000, 11334750, 250000,
        [{"text": "SOKOL INDUSTRIAL PARK (KZ × CN): 612.36 ha Phase 1 + 952 ha Expansion Reserve · 6 High-Tech Clusters · 12-Track Rail Hub · Planned SEZ Regime", "sz": 950, "bold": False, "color": "B8CACC"}],
        pad=(0,0,0,0)
    ))

def build_sokol_economics(slide_root):
    sptree = slide_root.find("p:cSld/p:spTree", NS)
    sid = 10
    sptree.append(make_textbox_sp(
        sid, "Section", 428625, 230000, 8000000, 220000,
        [{"text": "ЭКОНОМИКА ПРОЕКТА И ИНВЕСТИЦИОННЫЕ ПАРАМЕТРЫ", "sz": 1100, "bold": True, "color": "E9C28E"}],
        pad=(0,0,0,0)
    ))
    sptree.append(make_textbox_sp(
        sid+1, "Title", 428625, 520000, 11334750, 520000,
        [{"text": "Целевые показатели инвестиций, рабочих мест и грузооборота [предварительно]", "sz": 2300, "bold": True, "color": "F8F4ED"}],
        pad=(0,0,0,0)
    ))
    sptree.append(make_textbox_sp(
        sid+2, "Subtitle", 428625, 1040000, 11334750, 280000,
        [{"text": "Финансовая модель рассчитывается по очередям ввода (2027–2035+) и спецификации якорных резидентов · Значения [TBD] заполняются инициатором", "sz": 1150, "bold": False, "color": "B8CACC"}],
        pad=(0,0,0,0)
    ))

    kpis = [
        ("[TBD] млн $", "CAPEX инфраструктуры 1-й очереди", "Подготовка 612,36 га, ж/д примыкание у ст. Сорбулак, подстанции, водоподготовка и дороги", "E9C28E"),
        ("[TBD] млрд $", "Прямые инвестиции резидентов", "Производственные линии 6 кластеров (500 га), ЦОД, складские комплексы класса А (120–150 тыс. м²)", "3CC1BF"),
        ("[TBD] тыс.", "Новых рабочих мест в парке", "Инженеры, операторы автоматизированных линий, специалисты ЦОД, логистики и R&D-лабораторий", "E9C28E"),
        ("[TBD] тыс.", "Жителей и студентов кампуса", "Население жилых кварталов (32 га) и контингент Международного технологического университета (14 га)", "3CC1BF"),
        ("[TBD] тыс. TEU/год", "Целевой грузооборот терминала", "Ориентир международного бенчмарка: Duisburg Gateway Terminal — до 850 тыс. TEU/год на ≈23,5 га (по данным DGT/duisport)", "E9C28E"),
        ("до 25 лет", "Горизонт преференций (план)", "При получении статуса СЭЗ: КПН −100%, земля и имущество −100% (кат. A/B/C — до 7/15/25 лет), СТЗ и НДС 0%", "3CC1BF"),
    ]
    for i, (big, head, desc, col) in enumerate(kpis):
        c = i % 3
        r = i // 3
        x = 428625 + c * 3830000
        y = 1450000 + r * 2220000
        bid = sid + 10 + i * 3
        add_rect(sptree, bid, f"EconCard_{i}", x, y, 3670000, 2020000, fill_hex="132F3D", border_hex=col, border_w=14000, radius=True)
        sptree.append(make_textbox_sp(
            bid+1, f"EconTxt_{i}", x+150000, y+140000, 3370000, 1750000,
            [
                {"text": big, "sz": 2300, "bold": True, "color": col, "spcAft": 350},
                {"text": head, "sz": 1300, "bold": True, "color": "F8F4ED", "spcAft": 350},
                {"text": desc, "sz": 1000, "bold": False, "color": "B8CACC", "lnSpc": 118000},
            ],
            pad=(0,0,0,0)
        ))

    sptree.append(make_textbox_sp(
        sid+40, "AskBar", 428625, 5940000, 11334750, 560000,
        [{"text": "ПРЕДЛОЖЕНИЕ ИНВЕСТОРУ (INVESTOR ASK): выбор участка от [TBD] га под завод / склад / городской объект · совместное предприятие (KZ×CN) или 100% иностранное участие · фиксация мощностей сетей (МВт, м³/сут) и графика запуска на этапе предпроекта.", "sz": 1100, "bold": True, "color": "F8F4ED"}],
        fill_hex="102835", border_hex="E9C28E", border_w=12700, radius=True, anchor="ctr", pad=(140000, 60000, 140000, 60000)
    ))

def build_sokol_sources_legal(slide_root):
    """Formats Sokol 'Источники и правовые оговорки' following the exact standard of Alatau p.014."""
    sptree = slide_root.find("p:cSld/p:spTree", NS)
    sid = 10
    sptree.append(make_textbox_sp(
        sid, "Section", 428625, 230000, 8000000, 220000,
        [{"text": "ИСТОЧНИКИ И ПРАВОВЫЕ ОГОВОРКИ  ·  SOURCES & LEGAL DISCLAIMERS", "sz": 1100, "bold": True, "color": "E9C28E"}],
        pad=(0,0,0,0)
    ))
    sptree.append(make_textbox_sp(
        sid+1, "Title", 428625, 520000, 11334750, 500000,
        [{"text": "Источники данных, картографии и обязательные правовые примечания", "sz": 2300, "bold": True, "color": "F8F4ED"}],
        pad=(0,0,0,0)
    ))
    add_rect(sptree, sid+2, "Divider", 428625, 1100000, 11334750, 12700, fill_hex="27495A")

    items = [
        ("Статус СЭЗ и преференций (Предварительно):",
         "Специальный экономический режим для территории Sokol рассматривается инициатором (планируемый режим). Границы, виды приоритетной деятельности и сроки (до 25 лет, категории A/B/C) подлежат утверждению правовым решением уполномоченных органов РК."),
        ("Земельные участки и генеральный план:",
         "Контуры и площади участков № 2–5 (612,36 га), резерва № 6–9 (952 га) и соседнего участка автодрома № 1 (205 га, отдельно) приведены предварительно по схеме «Sokol Земля» и не являются утверждённым ПДП. Программы испытаний согласуются с оператором автодрома (кольцо по проекту бюро Tilke — 4 495 м)."),
        ("Транспорт, ж/д примыкание и инженерные сети:",
         "Варианты примыкания к линии Жетыген — Казыбек бек (район ст. Сорбулак, 12 путей) требуют инженерного расчёта и согласования с АО «НК «КТЖ» (справочно учтены открытые материалы ESIA / AIIB обхода Алматы). Объёмы воды, газа и электроэнергии подтверждаются у владельцев сетей."),
        ("Отраслевые ориентиры и внешние упоминания:",
         "Показатели Duisburg Gateway Terminal (до 850 тыс. TEU/год, ≈23,5 га, 6→12 путей, >730 м) приведены по открытым данным DGT и duisport. Марки Xiaomi Auto, MG и VOYAH указаны исключительно как иллюстративные примеры для проработки — их участие в проекте не подтверждено. Авиапредприятия региона (АРЗ № 405, BURUNDAIAVIA) являются самостоятельными организациями."),
        ("Картография и медиаматериалы:",
         "Карты и маршруты: © OpenStreetMap contributors, Natural Earth, OSRM, Esri World Imagery / World Shaded Relief. Демонстрационный полёт аэротакси AutoFlight в Алатау (19.05.2026): фото © Tengrinews.kz / Болат Айтмолда; телесюжет 24KZ; видео производителя."),
    ]
    y_cur = 1280000
    for idx, (head, body) in enumerate(items):
        bid = sid + 10 + idx * 3
        add_rect(sptree, bid, f"SrcDot_{idx}", 428625, y_cur + 70000, 85000, 85000, fill_hex="3CC1BF" if idx % 2 == 0 else "E9C28E", radius=True)
        sptree.append(make_textbox_sp(
            bid+1, f"SrcRow_{idx}", 590000, y_cur, 11170000, 960000,
            [
                {"text": head, "sz": 1180, "bold": True, "color": "E9C28E" if idx % 2 == 1 else "3CC1BF", "spcAft": 150},
                {"text": body, "sz": 1020, "bold": False, "color": "F8F4ED", "lnSpc": 115000},
            ],
            pad=(0,0,0,0)
        ))
        y_cur += 1020000

def build_sokol_contacts(slide_root):
    """Bilingual RU / EN / KZ Contacts & Investor CTA slide for Sokol."""
    sptree = slide_root.find("p:cSld/p:spTree", NS)
    sid = 10
    sptree.append(make_textbox_sp(
        sid, "Section", 428625, 230000, 9000000, 220000,
        [{"text": "КОНТАКТЫ И ПРИГЛАШЕНИЕ К ПАРТНЁРСТВУ  ·  CONTACTS & PARTNERSHIP", "sz": 1100, "bold": True, "color": "E9C28E"}],
        pad=(0,0,0,0)
    ))
    sptree.append(make_textbox_sp(
        sid+1, "Title", 428625, 520000, 11334750, 500000,
        [{"text": "Приглашаем к переговорам по размещению производств и инвестициям", "sz": 2300, "bold": True, "color": "F8F4ED"}],
        pad=(0,0,0,0)
    ))
    sptree.append(make_textbox_sp(
        sid+2, "SubtitleKZ_EN", 428625, 1020000, 11334750, 280000,
        [{"text": "KZ: Инвесторлар мен серіктестерді ынтымақтастыққа шақырамыз  ·  EN: Start Your Industrial & Logistics Project at SOKOL", "sz": 1150, "bold": False, "color": "3CC1BF"}],
        pad=(0,0,0,0)
    ))

    # Left Card: Contacts
    add_rect(sptree, sid+10, "ContactCard", 428625, 1440000, 5500000, 4650000, fill_hex="132F3D", border_hex="E9C28E", border_w=15875, radius=True)
    sptree.append(make_textbox_sp(
        sid+11, "ContactCardTxt", 640000, 1640000, 5080000, 4300000,
        [
            {"text": "ДИРЕКЦИЯ ПРОЕКТА / PROJECT OFFICE", "sz": 1050, "bold": True, "color": "E9C28E", "spcAft": 450},
            {"text": "SOKOL INDUSTRIAL PARK", "sz": 1850, "bold": True, "color": "F8F4ED", "spcAft": 650},
            {"text": "Локация площадки / Site Location:\nРеспублика Казахстан, Алматинская область, Илийский район,\n76-й км трассы Алматы — Астана (рядом с СТК Sokol)", "sz": 1200, "bold": False, "color": "B8CACC", "spcAft": 550},
            {"text": "Контактное лицо / Key Contact:\n[ФИО, должность руководителя проекта — TBD]", "sz": 1200, "bold": False, "color": "F8F4ED", "spcAft": 550},
            {"text": "Телефон / Tel:  [+7 (727) XXX-XX-XX — TBD]\nE-mail:  [invest@sokol-park.kz — TBD]\nСайт / Web:  [www.sokol-park.kz — TBD]", "sz": 1250, "bold": True, "color": "E9C28E"},
        ],
        pad=(0,0,0,0)
    ))

    # Right Card: 3 tracks + QR
    add_rect(sptree, sid+20, "TracksCard", 6210000, 1440000, 5550000, 4650000, fill_hex="132F3D", border_hex="3CC1BF", border_w=15875, radius=True)
    sptree.append(make_textbox_sp(
        sid+21, "TracksCardTxt", 6430000, 1640000, 5120000, 2750000,
        [
            {"text": "ПАКЕТ ДАННЫХ ДЛЯ СТАРТА ПЕРЕГОВОРОВ (RU / EN)", "sz": 1050, "bold": True, "color": "3CC1BF", "spcAft": 450},
            {"text": "Что направить инициатору / Initial Brief:", "sz": 1600, "bold": True, "color": "F8F4ED", "spcAft": 500},
            {"text": "1. Промышленным резидентам (Manufacturers): продукт, площадь участка (га), нагрузки по воде и э/э, целевой рынок сбыта.\n2. Логистическим операторам (Logistics): грузопоток (TEU/т), требования к ж/д фронту и складам класса А.\n3. Девелоперам города (Urban Investors): формат участия в жилье (32 га), образовании, медицине и сервисе.", "sz": 1100, "bold": False, "color": "B8CACC", "lnSpc": 120000},
        ],
        pad=(0,0,0,0)
    ))
    add_rect(sptree, sid+22, "SokolQR", 6430000, 4520000, 1320000, 1320000, fill_hex="0B202B", border_hex="E9C28E", border_w=12700, radius=True)
    sptree.append(make_textbox_sp(
        sid+23, "SokolQRTxt", 6430000, 4520000, 1320000, 1320000,
        [{"text": "[QR-КОД]\nData Room\n& Генплан\n(TBD)", "sz": 1050, "bold": True, "color": "E9C28E", "align": "ctr"}],
        anchor="ctr"
    ))
    sptree.append(make_textbox_sp(
        sid+24, "SokolSloganBox", 7940000, 4580000, 3600000, 1200000,
        [
            {"text": "ONE INDUSTRIAL CITY.\nCONNECTED TO EURASIA.", "sz": 1300, "bold": True, "color": "F8F4ED", "spcAft": 300},
            {"text": "Бір индустриялық қала — Еуразиямен байланысқан · Единый индустриальный город.", "sz": 1000, "bold": False, "color": "B8CACC"},
        ],
        pad=(0,0,0,0)
    ))

print("slide_builders.py written successfully.")
