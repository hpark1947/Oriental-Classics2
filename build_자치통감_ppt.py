# -*- coding: utf-8 -*-
"""
자치통감(資治通鑑) 발표자료 — 망라적 77장 PPT
원칙: 흰 배경 · 한 슬라이드 한 주제 · 매수 무제약
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

INK    = RGBColor(0x1F, 0x2A, 0x44)
ACCENT = RGBColor(0x8B, 0x1A, 0x1A)
SUB    = RGBColor(0x55, 0x60, 0x70)
RULE   = RGBColor(0xC8, 0xA2, 0x5B)
PALE   = RGBColor(0xF5, 0xEE, 0xDD)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)


def set_white_background(slide):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = WHITE


def add_textbox(slide, left, top, width, height, text, *,
                font_size=18, bold=False, color=INK, align=PP_ALIGN.LEFT,
                font_name='맑은 고딕', anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Emu(0); tf.margin_right = Emu(0)
    tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return tb


def add_paragraphs(slide, left, top, width, height, lines, *,
                   font_size=18, color=INK, font_name='맑은 고딕',
                   line_spacing=1.35, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0); tf.margin_right = Emu(0)
    tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
    for i, item in enumerate(lines):
        if isinstance(item, str):
            text, opts = item, {}
        else:
            text, opts = item
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = opts.get('align', align)
        p.line_spacing = line_spacing
        if opts.get('space_before'):
            p.space_before = Pt(opts['space_before'])
        run = p.add_run()
        run.text = text
        run.font.name = opts.get('font_name', font_name)
        run.font.size = Pt(opts.get('font_size', font_size))
        run.font.bold = opts.get('bold', False)
        run.font.color.rgb = opts.get('color', color)
    return tb


def add_rule(slide, left, top, width, color=RULE, weight=2.0):
    line = slide.shapes.add_connector(1, left, top, left + width, top)
    line.line.color.rgb = color
    line.line.width = Pt(weight)
    return line


def add_filled_rect(slide, left, top, width, height, fill_color, line_color=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_color
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
    shp.shadow.inherit = False
    return shp


def add_page_header(slide, section, page_num, total):
    add_textbox(slide, Inches(0.5), Inches(0.3), Inches(10), Inches(0.3),
                section, font_size=11, color=SUB, bold=True)
    add_textbox(slide, Inches(12.0), Inches(0.3), Inches(1.3), Inches(0.3),
                f'{page_num} / {total}', font_size=10, color=SUB,
                align=PP_ALIGN.RIGHT)
    add_rule(slide, Inches(0.5), Inches(0.65), Inches(12.8))


def add_title(slide, title, subtitle=None):
    add_textbox(slide, Inches(0.5), Inches(0.85), Inches(12.8), Inches(0.7),
                title, font_size=30, bold=True, color=INK)
    if subtitle:
        add_textbox(slide, Inches(0.5), Inches(1.55), Inches(12.8), Inches(0.4),
                    subtitle, font_size=14, color=SUB)


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

SLIDES = []
def S(section):
    def deco(fn):
        SLIDES.append((fn, section))
        return fn
    return deco


# ============== 1. 표지 ==============
@S('표지')
def s_cover(slide, page, total):
    set_white_background(slide)
    add_filled_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.4), RULE)
    add_textbox(slide, Inches(0.5), Inches(1.5), Inches(12.3), Inches(1.8),
                '資 治 通 鑑', font_size=100, bold=True, color=INK,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.6), Inches(12.3), Inches(0.5),
                'Comprehensive Mirror to Aid in Government · 자치통감',
                font_size=20, color=ACCENT, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.4), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(4.65), Inches(12.3), Inches(0.5),
                '사마광(司馬光) 편 — 19년의 사학 · 1,362년의 통사',
                font_size=20, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.2), Inches(12.3), Inches(0.4),
                '16기(紀) 294권 · 약 300만 자 · 1084년 헌상 · 동양 통사의 최고봉',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.8), Inches(12.3), Inches(0.5),
                '"鑑前世之興衰, 考當今之得失"',
                font_size=18, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.3), Inches(12.3), Inches(0.4),
                '— 지난 시대의 흥망을 거울 삼아 지금의 득실을 살핀다 (사마광)',
                font_size=14, color=SUB, align=PP_ALIGN.CENTER)
    add_filled_rect(slide, Inches(0), Inches(7.1), Inches(13.333), Inches(0.4), RULE)


# ============== 2. 목차 ==============
@S('목차')
def s_toc(slide, page, total):
    set_white_background(slide)
    add_textbox(slide, Inches(0.5), Inches(0.4), Inches(12.8), Inches(0.7),
                '목 차', font_size=32, bold=True, color=INK)
    add_rule(slide, Inches(0.5), Inches(1.15), Inches(12.8))
    items = [
        [('Ⅰ', '개요 — 자치통감이란 무엇인가'),
         ('Ⅱ', '사마광 — 19년의 사학(史學)'),
         ('Ⅲ', '책의 구조 — 16기 294권'),
         ('Ⅳ', '시작·끝·사관(史觀)'),
         ('Ⅴ', '16기 시대별 흐름'),
         ('Ⅵ', '핵심 사평 — 신광왈(臣光曰)')],
        [('Ⅶ', '주요 사건·인물 10선'),
         ('Ⅷ', '명구절 8선'),
         ('Ⅸ', '정관정요와의 비교'),
         ('Ⅹ', '후대 영향 — 동아시아 1,000년'),
         ('Ⅺ', '현대적 의의'),
         ('Ⅻ', '마무리')],
    ]
    for col, group in enumerate(items):
        x = 0.7 + col * 6.4
        top = 1.6
        for num, title in group:
            add_textbox(slide, Inches(x), Inches(top), Inches(1.0), Inches(0.5),
                        num, font_size=17, bold=True, color=ACCENT)
            add_textbox(slide, Inches(x + 1.0), Inches(top), Inches(5.3), Inches(0.5),
                        title, font_size=17, color=INK)
            top += 0.7


# ============== Ⅰ. 개요 ==============
@S('Ⅰ. 개요')
def s_what_is(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅰ. 개요', page, total)
    add_title(slide, '자치통감(資治通鑑)이란 무엇인가')
    add_filled_rect(slide, Inches(0.5), Inches(2.2), Inches(12.3), Inches(1.4), PALE)
    add_textbox(slide, Inches(0.9), Inches(2.4), Inches(11.5), Inches(0.5),
                '"다스림에 도움이 되고 역대에 통하여 거울이 되는 책"',
                font_size=22, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.9), Inches(2.95), Inches(11.5), Inches(0.5),
                '주(周) 위열왕 BC 403 — 후주(後周) 세종 959 · 1,362년의 통사',
                font_size=15, color=ACCENT, align=PP_ALIGN.CENTER)
    nums = [('16', '기(紀)'), ('294', '권(卷)'), ('약 300만', '자(字)'), ('19', '년 편찬')]
    for i, (n, lbl) in enumerate(nums):
        x = 0.6 + i * 3.05
        add_textbox(slide, Inches(x), Inches(4.0), Inches(2.9), Inches(1.0),
                    n, font_size=44, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(x), Inches(5.1), Inches(2.9), Inches(0.5),
                    lbl, font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.1), Inches(12.3), Inches(0.4),
                '동양 통사(通史)의 최고봉 · 편년체(編年體)의 정점',
                font_size=15, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.6), Inches(12.3), Inches(0.4),
                '정관정요와 함께 동양 제왕학의 양대 기둥',
                font_size=13, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅰ. 개요')
def s_status(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅰ. 개요', page, total)
    add_title(slide, '책의 위상 — 핵심 정보')
    rows = [
        ('서명',     '자치통감(資治通鑑)',          '"다스림에 도움 되는 통시(通時)의 거울"'),
        ('편자',     '사마광(司馬光, 1019~1086)', '북송(北宋)의 학자·재상'),
        ('편찬',     '1066~1084 · 19년',           '낙양에서 사학(史學) 운영'),
        ('헌상',     '1084년 송 신종에게',          '책 이름은 신종이 명명'),
        ('분량',     '16기 294권 · 약 300만 자',    '동양 통사 최대 분량'),
        ('형식',     '편년체(編年體)',              '연·월·일 순서로 서술'),
        ('수록 기간', '1,362년',                     'BC 403 ~ AD 959'),
        ('사평',     '"신광왈(臣光曰)" 219편',       '사마광의 직접 평론'),
    ]
    top = 1.95
    for i, (tag, val, note) in enumerate(rows):
        bg = RGBColor(0xFA, 0xFA, 0xFA) if i % 2 == 0 else WHITE
        add_filled_rect(slide, Inches(0.5), Inches(top), Inches(2.3), Inches(0.6), PALE)
        add_filled_rect(slide, Inches(2.85), Inches(top), Inches(10.0), Inches(0.6), bg)
        add_textbox(slide, Inches(0.55), Inches(top + 0.15), Inches(2.2), Inches(0.4),
                    tag, font_size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(2.95), Inches(top + 0.05), Inches(4.5), Inches(0.5),
                    val, font_size=15, bold=True, color=ACCENT)
        add_textbox(slide, Inches(7.5), Inches(top + 0.08), Inches(5.3), Inches(0.5),
                    note, font_size=13, color=SUB)
        top += 0.66


@S('Ⅰ. 개요')
def s_simagwang(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅰ. 개요', page, total)
    add_title(slide, '사마광(司馬光, 1019~1086) — 통감의 저자')
    add_textbox(slide, Inches(0.7), Inches(2.3), Inches(3.2), Inches(3.5),
                '司\n馬\n光', font_size=110, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    lines = [
        ('자(字)·호(號)', {'bold': True, 'font_size': 17, 'color': ACCENT}),
        ('  자 군실(君實) · 호 우수(迂叟)',
         {'font_size': 14}),
        ('  사후 시호(諡號) 문정(文正) — 송 최고 영예',
         {'font_size': 14, 'color': SUB}),
        ('가문·출신', {'bold': True, 'font_size': 17, 'color': ACCENT, 'space_before': 8}),
        ('  하주(夏州) 협서성 출신 · 명문 출신',
         {'font_size': 14}),
        ('관직', {'bold': True, 'font_size': 17, 'color': ACCENT, 'space_before': 8}),
        ('  20세 진사 급제 · 한림학사 · 재상에 오름',
         {'font_size': 14}),
        ('학문·성격', {'bold': True, 'font_size': 17, 'color': ACCENT, 'space_before': 8}),
        ('  사학(史學)의 거장 · "성실(誠)"을 평생 강조',
         {'font_size': 14}),
        ('  검소·엄정 — 송 사대부의 모범',
         {'font_size': 14, 'bold': True, 'color': INK}),
    ]
    add_paragraphs(slide, Inches(4.2), Inches(2.3), Inches(8.7), Inches(4.5),
                   lines, line_spacing=1.3)


@S('Ⅰ. 개요')
def s_book_name(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅰ. 개요', page, total)
    add_title(slide, '책 이름의 의미 — "자치통감(資治通鑑)"',
              '"송 신종이 직접 명명 — 황제가 책 이름을 지어준 사실 자체가 이 책의 위상"')
    items = [
        ('資', '자',   '도움이 되다',  '제왕의 통치(治)에 보탬이 되는'),
        ('治', '치',   '다스림',        '국가 경영의 학문 · 정치의 거울'),
        ('通', '통',   '통하다',        '역대(歷代)를 관통 · 1,362년'),
        ('鑑', '감',   '거울',          '과거를 비추어 현재를 살핌'),
    ]
    top = 2.4
    for han, eum, mean, role in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(1.3), Inches(0.85), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.15), Inches(1.3), Inches(0.6),
                    han, font_size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(2.0), Inches(top), Inches(1.5), Inches(0.85), PALE)
        add_textbox(slide, Inches(2.0), Inches(top + 0.1), Inches(1.5), Inches(0.4),
                    eum, font_size=14, color=SUB, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(2.0), Inches(top + 0.45), Inches(1.5), Inches(0.4),
                    mean, font_size=14, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(3.6), Inches(top), Inches(9.2), Inches(0.85),
                        RGBColor(0xFA, 0xFA, 0xFA))
        add_textbox(slide, Inches(3.8), Inches(top + 0.22), Inches(8.9), Inches(0.5),
                    role, font_size=14, color=INK)
        top += 0.97
    add_textbox(slide, Inches(0.5), Inches(6.55), Inches(12.3), Inches(0.5),
                '"다스림에 도움이 되고, 역대에 통하여, 거울이 되는 책"',
                font_size=17, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(7.05), Inches(12.3), Inches(0.4),
                '신종의 명명 — 책의 사명을 한 마디로 압축',
                font_size=12, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅰ. 개요')
def s_two_pillars(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅰ. 개요', page, total)
    add_title(slide, '동양 제왕학의 양대 기둥 — 정관정요와 자치통감')
    add_filled_rect(slide, Inches(0.6), Inches(2.3), Inches(5.9), Inches(4.5), PALE)
    add_textbox(slide, Inches(0.6), Inches(2.5), Inches(5.9), Inches(0.5),
                '정관정요(貞觀政要)', font_size=22, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_paragraphs(slide, Inches(0.9), Inches(3.15), Inches(5.3), Inches(3.5), [
        ('편자', {'bold': True, 'font_size': 15, 'color': ACCENT}),
        ('  오긍(吳兢) · 8세기 초', {'font_size': 13}),
        ('성격', {'bold': True, 'font_size': 15, 'color': ACCENT, 'space_before': 6}),
        ('  주제별 사례집 — 40편', {'font_size': 13}),
        ('  당 태종 23년의 정치 문답', {'font_size': 13}),
        ('초점', {'bold': True, 'font_size': 15, 'color': ACCENT, 'space_before': 6}),
        ('  "What" — 무엇을 할 것인가', {'font_size': 13, 'bold': True}),
        ('  구체적 사례·실용 지침', {'font_size': 13, 'color': SUB}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(6.8), Inches(2.3), Inches(5.9), Inches(4.5),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_textbox(slide, Inches(6.8), Inches(2.5), Inches(5.9), Inches(0.5),
                '자치통감(資治通鑑)', font_size=22, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_paragraphs(slide, Inches(7.1), Inches(3.15), Inches(5.3), Inches(3.5), [
        ('편자', {'bold': True, 'font_size': 15, 'color': ACCENT}),
        ('  사마광(司馬光) · 1084', {'font_size': 13}),
        ('성격', {'bold': True, 'font_size': 15, 'color': ACCENT, 'space_before': 6}),
        ('  편년체 통사 — 16기 294권', {'font_size': 13}),
        ('  1,362년의 흥망 종합', {'font_size': 13}),
        ('초점', {'bold': True, 'font_size': 15, 'color': ACCENT, 'space_before': 6}),
        ('  "Why" — 왜 흥하고 망하는가', {'font_size': 13, 'bold': True}),
        ('  역사의 패턴·구조적 사평', {'font_size': 13, 'color': SUB}),
    ], line_spacing=1.3)
    add_textbox(slide, Inches(0.5), Inches(7.05), Inches(12.3), Inches(0.4),
                '"정관정요는 사례, 자치통감은 통사" — 짝을 이루는 양대 텍스트',
                font_size=13, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)


# ============== Ⅱ. 사마광 ==============
@S('Ⅱ. 사마광')
def s_simagwang_life(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅱ. 사마광', page, total)
    add_title(slide, '사마광의 일생 — 학자·정치가·역사가')
    timeline = [
        ('1019',    '하주(夏州) 출생',                                     False),
        ('1038',   '20세 진사 급제 — 가장 어린 합격자 중 하나',          False),
        ('1066',   '영종에게 『통지(通志)』 8권 헌상',                     True),
        ('1067',   '신종 즉위 — 책 이름을 "자치통감"으로 명명',          True),
        ('1071',   '왕안석 신법에 반대 → 낙양으로 은퇴',                  False),
        ('1071~1084', '낙양에서 19년간 편찬 — 사학(史學) 운영',           True),
        ('1084',   '『자치통감』 완성 · 신종에게 헌상',                    True),
        ('1085',   '신종 사망 · 철종 즉위 · 사마광 재상으로 복귀',         False),
        ('1086',   '사망 (68세) — 시호 문정(文正)',                       False),
    ]
    top = 2.0
    for era, event, is_key in timeline:
        c = ACCENT if is_key else SUB
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.5), Inches(0.55), c)
        add_textbox(slide, Inches(0.6), Inches(top + 0.13), Inches(2.5), Inches(0.4),
                    era, font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        bg = RGBColor(0xFA, 0xE5, 0xE5) if is_key else PALE
        add_filled_rect(slide, Inches(3.2), Inches(top), Inches(9.6), Inches(0.55), bg)
        add_textbox(slide, Inches(3.4), Inches(top + 0.13), Inches(9.3), Inches(0.4),
                    event, font_size=13, color=INK)
        top += 0.62


@S('Ⅱ. 사마광')
def s_two_emperors(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅱ. 사마광', page, total)
    add_title(slide, '두 황제의 후원 — 영종과 신종')
    add_filled_rect(slide, Inches(0.6), Inches(2.3), Inches(5.9), Inches(4.5), PALE)
    add_textbox(slide, Inches(0.6), Inches(2.5), Inches(5.9), Inches(0.5),
                '영종(英宗, 재위 1063~1067)', font_size=18, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_paragraphs(slide, Inches(0.9), Inches(3.15), Inches(5.3), Inches(3.5), [
        ('1066년 — 사마광의 『통지』 8권 헌상', {'font_size': 14, 'bold': True}),
        ('', {'font_size': 6}),
        ('영종이 그 진가를 알아봄', {'font_size': 14, 'space_before': 6}),
        ('"더 큰 통사를 편찬하라"', {'font_size': 14, 'bold': True, 'color': ACCENT}),
        ('국가 자원 지원 — 책 1,000권 빌려줌', {'font_size': 13, 'color': SUB}),
        ('보조 학자 임명', {'font_size': 13, 'color': SUB}),
        ('', {'font_size': 6}),
        ('1067년 영종 붕어 (35세)', {'font_size': 13, 'color': SUB, 'space_before': 6}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(6.8), Inches(2.3), Inches(5.9), Inches(4.5),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_textbox(slide, Inches(6.8), Inches(2.5), Inches(5.9), Inches(0.5),
                '신종(神宗, 재위 1067~1085)', font_size=18, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_paragraphs(slide, Inches(7.1), Inches(3.15), Inches(5.3), Inches(3.5), [
        ('영종의 후원을 그대로 이어받음', {'font_size': 14, 'bold': True}),
        ('', {'font_size': 6}),
        ('1067년 책 이름 명명 — 자치통감', {'font_size': 14, 'bold': True, 'color': ACCENT, 'space_before': 6}),
        ('"鑑前世之興衰" 서문 친저', {'font_size': 13}),
        ('', {'font_size': 6}),
        ('정치적 모순 — 왕안석 신법 추진', {'font_size': 13, 'color': SUB, 'space_before': 6}),
        ('사마광은 신법에 반대 → 낙양으로', {'font_size': 13, 'color': SUB}),
        ('', {'font_size': 6}),
        ('그러나 통감 편찬은 끝까지 지원', {'font_size': 14, 'bold': True, 'color': ACCENT, 'space_before': 4}),
    ], line_spacing=1.3)


@S('Ⅱ. 사마광')
def s_wang_ansik(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅱ. 사마광', page, total)
    add_title(slide, '왕안석(王安石)과의 갈등 — 낙양 19년의 배경')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.0), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(1.7), [
        ('왕안석의 신법(新法, 1069~) — 부국강병의 급진적 개혁',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('청묘법·면역법·시역법·보갑법 등 — 국가 개입 확대',
         {'font_size': 15, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.5),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(4.65), Inches(11.7), Inches(2.2), [
        ('사마광의 반대 — 보수주의자의 신중',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('"백성을 흔드는 급진 개혁은 위험하다"',
         {'font_size': 16, 'bold': True, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('"역사가 가르치는 점진(漸進)의 지혜"',
         {'font_size': 14, 'color': SUB, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('→ 1071년 낙양으로 은퇴 — 그러나 이 은퇴가 자치통감을 낳았다',
         {'font_size': 14, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
    ], line_spacing=1.25)


@S('Ⅱ. 사마광')
def s_19years(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅱ. 사마광', page, total)
    add_title(slide, '낙양 19년 — "두 방을 가득 채운 미완성 원고"')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.0), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(1.7), [
        ('낙양 — 송 부도(副都) · 학문의 중심',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('1071~1084년 — 19년간 정치를 떠나 오직 사학(史學)에 몰두',
         {'font_size': 16, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('"끝나기 전에 죽을까 두렵다"',
         {'font_size': 14, 'color': SUB, 'bold': True, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.5),
                    RGBColor(0xFA, 0xFA, 0xFA))
    add_paragraphs(slide, Inches(0.8), Inches(4.65), Inches(11.7), Inches(2.2), [
        ('편찬 과정의 일화',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('• 모든 사료(史料)를 직접 검토 — "구결(舊缺)" 메모를 남김',
         {'font_size': 14, 'space_before': 6}),
        ('• 작업실 두 방이 미완성 원고로 가득 — 후에 모두 정리',
         {'font_size': 14, 'space_before': 4}),
        ('• 잠을 줄이려 "둥근 베개(警枕)" 사용 — 굴러서 깨도록',
         {'font_size': 14, 'bold': True, 'color': ACCENT, 'space_before': 4}),
        ('• 1084년 완성 — 사마광 65세',
         {'font_size': 14, 'bold': True, 'space_before': 4}),
    ], line_spacing=1.25)


@S('Ⅱ. 사마광')
def s_collaborators(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅱ. 사마광', page, total)
    add_title(slide, '협력자 3인 — 시대를 분담한 학자들')
    items = [
        ('류반(劉攽)', '1023~1089',  '한기(漢紀) 분담',
         '한대 사료의 최고 전문가 · 사기·한서 정통'),
        ('류서(劉恕)', '1032~1078',  '위·진·남북조 분담',
         '"기억의 화신" — 모든 사료를 외운다는 평가'),
        ('범조우(范祖禹)', '1041~1098', '당기(唐紀) 분담',
         '당사 전문가 · 후일 『당감(唐鑑)』 친저'),
    ]
    top = 2.4
    for name, era, role, char in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.5), Inches(1.3), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.32), Inches(2.5), Inches(0.5),
                    name, font_size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(0.6), Inches(top + 0.8), Inches(2.5), Inches(0.4),
                    era, font_size=12, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(3.2), Inches(top), Inches(2.8), Inches(1.3), PALE)
        add_textbox(slide, Inches(3.3), Inches(top + 0.4), Inches(2.6), Inches(0.5),
                    role, font_size=14, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(6.1), Inches(top), Inches(6.7), Inches(1.3),
                        RGBColor(0xFA, 0xFA, 0xFA))
        add_textbox(slide, Inches(6.3), Inches(top + 0.4), Inches(6.4), Inches(0.5),
                    char, font_size=13, color=INK)
        top += 1.45
    add_textbox(slide, Inches(0.5), Inches(6.85), Inches(12.3), Inches(0.4),
                '사마광이 총괄 — 사평(臣光曰)은 모두 그의 친저',
                font_size=13, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)


# ============== Ⅲ. 책의 구조 ==============
@S('Ⅲ. 구조')
def s_form(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 구조', page, total)
    add_title(slide, '편년체(編年體) — 시간 순서의 통사')
    add_filled_rect(slide, Inches(0.6), Inches(2.3), Inches(5.9), Inches(4.5), PALE)
    add_textbox(slide, Inches(0.6), Inches(2.5), Inches(5.9), Inches(0.5),
                '기전체(紀傳體) — 사기·한서', font_size=17, bold=True, color=SUB,
                align=PP_ALIGN.CENTER)
    add_paragraphs(slide, Inches(0.9), Inches(3.15), Inches(5.3), Inches(3.5), [
        ('인물 중심 서술', {'font_size': 16, 'bold': True}),
        ('', {'font_size': 6}),
        ('• 본기 — 제왕의 일생', {'font_size': 13, 'space_before': 6}),
        ('• 열전 — 신하의 일생', {'font_size': 13, 'space_before': 4}),
        ('', {'font_size': 6}),
        ('장점 — 인물의 입체적 조명', {'font_size': 13, 'color': SUB, 'space_before': 6}),
        ('단점 — 시간의 흐름이 끊김', {'font_size': 13, 'color': SUB}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(6.8), Inches(2.3), Inches(5.9), Inches(4.5),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_textbox(slide, Inches(6.8), Inches(2.5), Inches(5.9), Inches(0.5),
                '편년체(編年體) — 자치통감', font_size=17, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_paragraphs(slide, Inches(7.1), Inches(3.15), Inches(5.3), Inches(3.5), [
        ('시간 중심 서술', {'font_size': 16, 'bold': True, 'color': ACCENT}),
        ('', {'font_size': 6}),
        ('• 연(年)·월·일 순', {'font_size': 13, 'space_before': 6}),
        ('• 모든 사건을 시간순으로', {'font_size': 13, 'space_before': 4}),
        ('', {'font_size': 6}),
        ('장점 — 동시대 사건의 연관성 파악',
         {'font_size': 13, 'color': SUB, 'space_before': 6}),
        ('  "거시적 흐름의 거울"',
         {'font_size': 13, 'bold': True, 'color': ACCENT}),
    ], line_spacing=1.3)


@S('Ⅲ. 구조')
def s_16gi(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 구조', page, total)
    add_title(slide, '16기(紀) 개관 — 1,362년의 분할')
    items = [
        ('1', '주기(周紀)',   '5권',  'BC 403~256'),
        ('2', '진기(秦紀)',   '3권',  'BC 255~206'),
        ('3', '한기(漢紀)',   '60권', 'BC 206~AD 220'),
        ('4', '위기(魏紀)',   '10권', '220~265'),
        ('5', '진기(晉紀)',   '40권', '265~419'),
        ('6', '송기(宋紀)',   '16권', '420~478'),
        ('7', '제기(齊紀)',   '10권', '479~501'),
        ('8', '양기(梁紀)',   '22권', '502~556'),
        ('9', '진기(陳紀)',   '10권', '557~588'),
        ('10','수기(隋紀)',   '8권',  '589~617'),
        ('11','당기(唐紀)',   '81권', '618~906'),
        ('12','후량기(後梁)', '6권',  '907~922'),
        ('13','후당기(後唐)', '8권',  '923~935'),
        ('14','후진기(後晉)', '6권',  '936~946'),
        ('15','후한기(後漢)', '4권',  '947~950'),
        ('16','후주기(後周)', '5권',  '951~959'),
    ]
    top = 2.0
    for i, (num, name, vol, era) in enumerate(items):
        col = i % 2
        row = i // 2
        x = 0.5 + col * 6.4
        y = top + row * 0.62
        add_filled_rect(slide, Inches(x), Inches(y), Inches(0.6), Inches(0.55), ACCENT)
        add_textbox(slide, Inches(x), Inches(y + 0.13), Inches(0.6), Inches(0.4),
                    num, font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(x + 0.7), Inches(y), Inches(2.2), Inches(0.55), PALE)
        add_textbox(slide, Inches(x + 0.7), Inches(y + 0.13), Inches(2.2), Inches(0.4),
                    name, font_size=13, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(x + 3.0), Inches(y), Inches(1.2), Inches(0.55),
                        RGBColor(0xFA, 0xFA, 0xFA))
        add_textbox(slide, Inches(x + 3.0), Inches(y + 0.13), Inches(1.2), Inches(0.4),
                    vol, font_size=12, color=SUB, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(x + 4.3), Inches(y), Inches(1.7), Inches(0.55),
                        RGBColor(0xFA, 0xFA, 0xFA))
        add_textbox(slide, Inches(x + 4.3), Inches(y + 0.13), Inches(1.7), Inches(0.4),
                    era, font_size=11, color=INK, align=PP_ALIGN.CENTER)


@S('Ⅲ. 구조')
def s_vol_dist(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 구조', page, total)
    add_title(slide, '권수 분포 — "어디에 가장 무게를 두었나"')
    items = [
        ('당기',        '81권', 28),
        ('한기',        '60권', 20),
        ('진기(晉)',    '40권', 14),
        ('양기',        '22권', 7),
        ('송기',        '16권', 5),
        ('위·제·진·수', '38권', 13),
        ('주·진(秦)',   '8권',  3),
        ('오대',        '29권', 10),
    ]
    top = 2.3
    max_w = 9.5
    max_val = max(v for _, _, v in items)
    for name, vol, val in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.0), Inches(0.55), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.13), Inches(2.0), Inches(0.4),
                    name, font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(2.7), Inches(top), Inches(1.3), Inches(0.55), PALE)
        add_textbox(slide, Inches(2.7), Inches(top + 0.13), Inches(1.3), Inches(0.4),
                    vol, font_size=13, color=ACCENT, bold=True, align=PP_ALIGN.CENTER)
        w = max_w * val / max_val
        add_filled_rect(slide, Inches(4.1), Inches(top + 0.05), Inches(w), Inches(0.45), ACCENT)
        add_textbox(slide, Inches(4.1 + w + 0.1), Inches(top + 0.13), Inches(1.5),
                    Inches(0.4), f'{val}%', font_size=12, color=SUB, bold=True)
        top += 0.62
    add_textbox(slide, Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.4),
                '당기·한기에 절반 — 동양 통사의 두 황금기에 집중',
                font_size=13, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)


@S('Ⅲ. 구조')
def s_aux(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 구조', page, total)
    add_title(slide, '보조 문헌 — 통감을 둘러싼 책들')
    items = [
        ('목록(目錄)',      '사마광 친저 · 30권',     '날짜·사건 일람 — 통감의 색인'),
        ('고이(考異)',      '사마광 친저 · 30권',     '"왜 이 사료를 채택했는가" — 사료 비판'),
        ('통감기사본말',     '원추(袁樞) · 1174',      '사건별 재편 — 기사본말체의 시조'),
        ('통감강목(綱目)',   '주희(朱熹) · 1172',      '도덕적 재해석 — 성리학적 시각'),
        ('통감속편',         '필원(畢沅) · 청대',       '송 이후를 다룬 후속서'),
    ]
    top = 2.4
    for tag, author, char in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.5), Inches(0.7), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.15), Inches(2.5), Inches(0.4),
                    tag, font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(3.2), Inches(top), Inches(3.5), Inches(0.7), PALE)
        add_textbox(slide, Inches(3.35), Inches(top + 0.15), Inches(3.3), Inches(0.4),
                    author, font_size=13, bold=True, color=ACCENT)
        add_filled_rect(slide, Inches(6.8), Inches(top), Inches(6.0), Inches(0.7),
                        RGBColor(0xFA, 0xFA, 0xFA))
        add_textbox(slide, Inches(7.0), Inches(top + 0.15), Inches(5.7), Inches(0.4),
                    char, font_size=12, color=INK)
        top += 0.85


# ============== Ⅳ. 사관 ==============
@S('Ⅳ. 사관')
def s_start(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅳ. 사관', page, total)
    add_title(slide, '시작점 — 삼가분진(三家分晉, BC 403)',
              '"왜 주(周) 위열왕 23년에서 시작하는가?"')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.5), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(2.2), [
        ('BC 403년 — 주 위열왕(威烈王) 23년',
         {'font_size': 17, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('주왕이 진(晉)의 세 대부 한(韓)·위(魏)·조(趙)를 제후로 책봉',
         {'font_size': 15, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('"신하가 제후가 되는 정명(正名)의 붕괴"',
         {'font_size': 15, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.25)
    add_filled_rect(slide, Inches(0.5), Inches(5.0), Inches(12.3), Inches(2.0),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(5.15), Inches(11.7), Inches(1.8), [
        ('사마광의 시작점 선택 의도',
         {'font_size': 17, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('"천하 질서가 무너진 결정적 순간"',
         {'font_size': 16, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('첫 사평(臣光曰) — 명분과 질서의 중요성을 처음 천명',
         {'font_size': 13, 'color': SUB, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.3)


@S('Ⅳ. 사관')
def s_end(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅳ. 사관', page, total)
    add_title(slide, '끝점 — 후주(後周) 세종(世宗) 현덕(顯德) 6년 (959)',
              '"왜 송 건국 직전에서 멈추었나?"')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.5), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(2.2), [
        ('959년 — 후주 세종(柴榮)의 마지막 해',
         {'font_size': 17, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('이듬해(960) 송 태조 조광윤이 진교병변(陳橋兵變)으로 즉위',
         {'font_size': 15, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('= 송(宋) 건국 직전에서 절묘하게 멈춤',
         {'font_size': 15, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.25)
    add_filled_rect(slide, Inches(0.5), Inches(5.0), Inches(12.3), Inches(2.0),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(5.15), Inches(11.7), Inches(1.8), [
        ('"본조(本朝, 송)는 평론하지 않는다"',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('동시대 사관(史官)의 금기 — 객관성 유지',
         {'font_size': 14, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('과거를 거울로 삼아 "지금"을 비추게 하려는 의도',
         {'font_size': 14, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.3)


@S('Ⅳ. 사관')
def s_mirror(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅳ. 사관', page, total)
    add_title(slide, '"감(鑑)"의 의미 — 거울로서의 역사')
    add_textbox(slide, Inches(0.5), Inches(2.2), Inches(12.3), Inches(0.9),
                '鑑 前 世 之 興 衰',
                font_size=42, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.2), Inches(12.3), Inches(0.9),
                '考 當 今 之 得 失',
                font_size=42, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(4.2), Inches(12.3), Inches(0.5),
                '감전세지흥쇠 · 고당금지득실',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(4.7), Inches(12.3), Inches(0.4),
                '— 사마광 「자치통감서(序)」', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(5.3), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(5.6), Inches(12.3), Inches(0.6),
                '"지난 시대의 흥망(興衰)을 거울 삼아',
                font_size=20, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.2), Inches(12.3), Inches(0.6),
                ' 지금의 득실(得失)을 살핀다"',
                font_size=22, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.4),
                '역사는 단순 기록이 아니라 "거울" — 동양 사학의 근본 정신',
                font_size=13, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)


@S('Ⅳ. 사관')
def s_singwangwal(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅳ. 사관', page, total)
    add_title(slide, '"신광왈(臣光曰)" — 사마광의 친저 사평 219편')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.0), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(1.7), [
        ('"신광왈(臣光曰)" — "신하 사마광이 아룁니다"',
         {'font_size': 20, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('통감 전체에 약 219편의 사평(史評)을 친저',
         {'font_size': 16, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('주요 사건마다 사마광의 직접적 해석과 평가',
         {'font_size': 14, 'color': SUB, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.5),
                    RGBColor(0xFA, 0xFA, 0xFA))
    add_paragraphs(slide, Inches(0.8), Inches(4.65), Inches(11.7), Inches(2.2), [
        ('사평의 핵심 주제',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('• 정명(正名) — 군신·부자의 분(分)을 지킴',
         {'font_size': 14, 'space_before': 6}),
        ('• 인재 등용의 원칙 — "재(才)와 덕(德)의 균형"',
         {'font_size': 14, 'space_before': 4}),
        ('• 간언 수용 — 직간자를 보호',
         {'font_size': 14, 'space_before': 4}),
        ('• 검약·신중 — 군주의 자기 절제',
         {'font_size': 14, 'space_before': 4}),
        ('• 변화(漸進) — 급진적 개혁의 위험성',
         {'font_size': 14, 'bold': True, 'color': ACCENT, 'space_before': 4}),
    ], line_spacing=1.3)


# ============== Ⅴ. 16기 시대별 흐름 ==============
@S('Ⅴ. 시대별')
def s_zhou(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 시대별', page, total)
    add_title(slide, '주기(周紀) 5권 — BC 403~256 · 전국시대',
              '"천하 질서의 붕괴와 칠웅(七雄)의 쟁투"')
    items = [
        ('BC 403',  '삼가분진 — 주왕이 한·위·조를 제후로 책봉'),
        ('BC 361',  '진(秦) 효공 즉위 · 상앙(商鞅)의 변법(BC 356)'),
        ('BC 338',  '효공 사후 상앙 거열형 — 그러나 법은 살아남음'),
        ('BC 318',  '5국 합종(合縱) — 진에 대항한 동맹'),
        ('BC 286',  '제(齊)가 송(宋)을 멸함 · 동방의 강국'),
        ('BC 260',  '장평대전 — 진(秦)이 조(趙)의 40만 포로 생매장'),
        ('BC 256',  '진이 동주(東周)를 멸함 — 800년 주(周) 왕실 종말'),
    ]
    top = 2.4
    for era, event in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.0), Inches(0.55), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.13), Inches(2.0), Inches(0.4),
                    era, font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(2.7), Inches(top), Inches(10.1), Inches(0.55), PALE)
        add_textbox(slide, Inches(2.9), Inches(top + 0.13), Inches(9.8), Inches(0.4),
                    event, font_size=13, color=INK)
        top += 0.62
    add_textbox(slide, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
                '첫 사평 — "명분 붕괴의 결정적 순간"이라는 통감 전체의 출발점',
                font_size=13, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)


@S('Ⅴ. 시대별')
def s_qin(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 시대별', page, total)
    add_title(slide, '진기(秦紀) 3권 — BC 255~206 · 통일과 멸망',
              '"15년 통일 후 14년 만에 멸망 — 가장 짧은 제국"')
    items = [
        ('BC 246',  '진왕 정(政, 훗날 시황) 즉위 — 13세'),
        ('BC 221',  '6국 통일 — 중국 최초의 황제 등극'),
        ('BC 213~212', '분서갱유(焚書坑儒) — 사상 통제'),
        ('BC 210',  '시황 순행 중 사망 — 환관 조고의 음모 시작'),
        ('BC 209',  '진승(陳勝)·오광(吳廣)의 봉기 — "왕후장상 영유종호?"'),
        ('BC 207',  '거록의 전투 — 항우가 진의 주력 격파'),
        ('BC 206',  '유방이 함양 입성 → 진(秦) 멸망 · 약법삼장'),
    ]
    top = 2.4
    for era, event in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.2), Inches(0.55), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.13), Inches(2.2), Inches(0.4),
                    era, font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(2.9), Inches(top), Inches(9.9), Inches(0.55), PALE)
        add_textbox(slide, Inches(3.1), Inches(top + 0.13), Inches(9.6), Inches(0.4),
                    event, font_size=13, color=INK)
        top += 0.62
    add_textbox(slide, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
                '통감의 핵심 교훈 — "무력과 법만으로는 천하를 지킬 수 없다"',
                font_size=13, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)


@S('Ⅴ. 시대별')
def s_qianhan(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 시대별', page, total)
    add_title(slide, '한기(漢紀) 60권 ① — 전한(前漢) · BC 206~AD 8',
              '"초한쟁패 · 문경의 치 · 무제의 영광과 그늘"')
    items = [
        ('BC 206~202', '초한쟁패 — 유방 vs 항우 4년 전쟁'),
        ('BC 202',     '유방 즉위 — 한(漢) 건국 · 토사구팽 시작'),
        ('BC 180~141', '문경의 치(文景之治) — 무위(無爲)의 황금기'),
        ('BC 141~87',  '무제(武帝) — 흉노 정벌·실크로드·유가 국교'),
        ('BC 87~74',   '소제·곽광의 섭정 — 외척 정치의 시작'),
        ('BC 73~49',   '선제(宣帝) — 한 중흥의 명군'),
        ('BC 8~AD 8',  '왕망의 찬탈 — 신(新)왕조 (15년)'),
    ]
    top = 2.4
    for era, event in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.5), Inches(0.55), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.13), Inches(2.5), Inches(0.4),
                    era, font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(3.2), Inches(top), Inches(9.6), Inches(0.55), PALE)
        add_textbox(slide, Inches(3.4), Inches(top + 0.13), Inches(9.3), Inches(0.4),
                    event, font_size=13, color=INK)
        top += 0.62
    add_textbox(slide, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
                '문경의 치 — 동양 통사의 모범적 황금기 · 통감 사평의 단골 모범',
                font_size=13, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)


@S('Ⅴ. 시대별')
def s_houhan(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 시대별', page, total)
    add_title(slide, '한기(漢紀) 60권 ② — 후한(後漢) · 25~220',
              '"광무중흥 → 환관·외척 정치 → 황건의 난"')
    items = [
        ('25',   '광무제(光武帝) 즉위 — 한 부흥, "광무중흥"'),
        ('29~57', '광무제 치세 — 명군의 모범'),
        ('57~75', '명제(明帝) — 불교 동전(東傳), 백마사 창건'),
        ('166',  '제1차 당고지화(黨錮之禍) — 환관의 사대부 탄압'),
        ('169',  '제2차 당고지화 — 청류파 대거 학살'),
        ('184',  '황건(黃巾)의 난 — "창천이사 황천당립"'),
        ('189',  '동탁(董卓)의 낙양 입성 — 후한 황제권 붕괴'),
        ('220',  '조비(曹丕)가 헌제를 폐위 → 위(魏) 건국 · 한 멸망'),
    ]
    top = 2.3
    for era, event in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.0), Inches(0.55), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.13), Inches(2.0), Inches(0.4),
                    era, font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(2.7), Inches(top), Inches(10.1), Inches(0.55), PALE)
        add_textbox(slide, Inches(2.9), Inches(top + 0.13), Inches(9.8), Inches(0.4),
                    event, font_size=13, color=INK)
        top += 0.6
    add_textbox(slide, Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.4),
                '통감 교훈 — "외척·환관의 정치 농단이 왕조 멸망의 원인"',
                font_size=13, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)


@S('Ⅴ. 시대별')
def s_3kingdoms(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 시대별', page, total)
    add_title(slide, '위기(魏紀) 10권 + 진기(晉紀) — 삼국·서진 · 220~316')
    items = [
        ('220',  '조비 — 위(魏) 건국 (낙양)'),
        ('221',  '유비 — 촉한(蜀漢) 건국 (성도)'),
        ('229',  '손권 — 오(吳) 건국 (건업) · 삼국정립 완성'),
        ('234',  '제갈량(諸葛亮) 오장원에서 병사'),
        ('249',  '사마의(司馬懿) 고평릉의 변 — 위 권력 장악'),
        ('263',  '위가 촉(蜀)을 멸함'),
        ('265',  '사마염(司馬炎) — 위를 폐하고 진(晉) 건국'),
        ('280',  '진이 오(吳)를 멸함 — 천하 통일'),
        ('291~306', '팔왕(八王)의 난 — 진의 내전'),
        ('316',  '서진(西晉) 멸망 — 영가의 난'),
    ]
    top = 2.0
    for i, (era, event) in enumerate(items):
        col = i % 2
        row = i // 2
        x = 0.5 + col * 6.4
        y = top + row * 0.6
        add_filled_rect(slide, Inches(x), Inches(y), Inches(1.5), Inches(0.5), ACCENT)
        add_textbox(slide, Inches(x), Inches(y + 0.1), Inches(1.5), Inches(0.4),
                    era, font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(x + 1.6), Inches(y), Inches(4.4), Inches(0.5), PALE)
        add_textbox(slide, Inches(x + 1.7), Inches(y + 0.1), Inches(4.3), Inches(0.4),
                    event, font_size=11, color=INK)


@S('Ⅴ. 시대별')
def s_nbcho(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 시대별', page, total)
    add_title(slide, '동진·남북조 — 317~589 · 송·제·양·진(陳)기',
              '"한족이 강남으로 남하 · 북방은 16국 → 북위 → 북주"')
    items = [
        ('317~420', '동진(東晉) — 한족 사대부의 남방 정권 · 사현 등'),
        ('383',     '비수(淝水)대전 — 동진이 전진(前秦) 격파'),
        ('420~478', '송(宋) — 유유(劉裕) 건국 · 원가의 치'),
        ('439',     '북위(北魏) 화북 통일'),
        ('479~501', '제(齊) — 짧은 왕조'),
        ('502~556', '양(梁) — 무제(武帝)의 불교 부흥 · 후경의 난'),
        ('557~588', '진(陳) — 남조 마지막'),
        ('581',     '수(隋) 양견(楊堅) 건국 · 북주 찬탈'),
        ('589',     '수가 진(陳)을 멸함 — 분열 종식, 천하 재통일'),
    ]
    top = 2.3
    for era, event in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.0), Inches(0.5), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.1), Inches(2.0), Inches(0.4),
                    era, font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(2.7), Inches(top), Inches(10.1), Inches(0.5), PALE)
        add_textbox(slide, Inches(2.9), Inches(top + 0.1), Inches(9.8), Inches(0.4),
                    event, font_size=13, color=INK)
        top += 0.55


@S('Ⅴ. 시대별')
def s_sui(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 시대별', page, total)
    add_title(slide, '수기(隋紀) 8권 — 589~617 · 30년의 통일과 멸망',
              '"문제의 영광과 양제의 폭정 — 통감의 핵심 반면교사"')
    items = [
        ('581',     '문제(文帝) 양견 — 북주를 폐하고 수 건국'),
        ('589',     '천하 재통일 — 분열 273년 종식'),
        ('589~604', '문제 치세 — 율령 정비·과거제 시작·균전제'),
        ('604',     '양제(煬帝) 즉위 — 부친 시해 의혹'),
        ('605~610', '대운하 건설 — 수백만 명 동원'),
        ('612~614', '고구려 3차 원정 — 모두 실패 · 백만 군대 손실'),
        ('616',     '양제 강도(江都)로 도피 — 사실상 강남 정권'),
        ('618',     '강도에서 호위군에게 살해 — 수 멸망'),
    ]
    top = 2.3
    for era, event in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.2), Inches(0.5), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.1), Inches(2.2), Inches(0.4),
                    era, font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(2.9), Inches(top), Inches(9.9), Inches(0.5), PALE)
        add_textbox(slide, Inches(3.1), Inches(top + 0.1), Inches(9.6), Inches(0.4),
                    event, font_size=13, color=INK)
        top += 0.55
    add_textbox(slide, Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.4),
                '"수 양제 = 진 시황과 함께 통감의 양대 반면교사"',
                font_size=13, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)


@S('Ⅴ. 시대별')
def s_tang_1(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 시대별', page, total)
    add_title(slide, '당기(唐紀) 81권 ① — 618~755 · 전기',
              '"통감 최대 분량(28%) — 당의 영광"')
    items = [
        ('618',     '당 고조 이연(李淵) 건국'),
        ('626',     '현무문(玄武門)의 변 — 이세민 즉위'),
        ('627~649', '정관(貞觀)의 치 — 통감의 모범 사례'),
        ('630',     '돌궐 평정 · "천가한(天可汗)" 칭호'),
        ('649~683', '고종(高宗) — 측천무후의 부상'),
        ('690~705', '측천무후 — 중국 최초의 여황제 · 주(周)'),
        ('712~756', '현종(玄宗) — 개원(開元)의 치 · 당의 정점'),
        ('755',     '안사(安史)의 난 — 당의 분기점'),
    ]
    top = 2.3
    for era, event in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.0), Inches(0.5), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.1), Inches(2.0), Inches(0.4),
                    era, font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(2.7), Inches(top), Inches(10.1), Inches(0.5), PALE)
        add_textbox(slide, Inches(2.9), Inches(top + 0.1), Inches(9.8), Inches(0.4),
                    event, font_size=13, color=INK)
        top += 0.55
    add_textbox(slide, Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.4),
                '"정관의 치 → 개원의 치 → 안사의 난" — 통감이 가장 깊이 다루는 흐름',
                font_size=13, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)


@S('Ⅴ. 시대별')
def s_tang_2(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 시대별', page, total)
    add_title(slide, '당기(唐紀) 81권 ② — 755~906 · 후기',
              '"안사의 난 이후 150년의 점진적 쇠퇴"')
    items = [
        ('755~763', '안사의 난 — 안녹산·사사명 · 인구 절반 사라짐'),
        ('756',     '양귀비의 죽음 — 마외역(馬嵬驛)'),
        ('780',     '양세법(兩稅法) 시행 — 조용조 → 토지·재산세'),
        ('806~820', '헌종(憲宗) — 원화의 중흥 · 단명'),
        ('835',     '감로(甘露)의 변 — 환관의 사대부 학살'),
        ('874~884', '황소(黃巢)의 난 — 장안 함락'),
        ('907',     '주전충(朱全忠)이 애제 폐위 → 후량(後梁) 건국'),
    ]
    top = 2.4
    for era, event in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.0), Inches(0.55), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.13), Inches(2.0), Inches(0.4),
                    era, font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(2.7), Inches(top), Inches(10.1), Inches(0.55), PALE)
        add_textbox(slide, Inches(2.9), Inches(top + 0.13), Inches(9.8), Inches(0.4),
                    event, font_size=13, color=INK)
        top += 0.62
    add_textbox(slide, Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.4),
                '통감 교훈 — "환관·번진(藩鎭)·환국의 3중 위기가 290년 당을 무너뜨림"',
                font_size=13, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)


@S('Ⅴ. 시대별')
def s_5dynasties(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 시대별', page, total)
    add_title(slide, '오대(五代) 5기 29권 — 907~959 · 53년의 5왕조 교체',
              '"평균 11년의 단명 왕조 — 무인 정치의 위험"')
    items = [
        ('907~923', '후량(後梁) — 주전충 · 16년'),
        ('923~936', '후당(後唐) — 이존욱 · 13년'),
        ('936~947', '후진(後晉) — 석경당 · 거란에 연운16주 할양'),
        ('947~950', '후한(後漢) — 유지원 · 3년 (가장 짧은)'),
        ('951~960', '후주(後周) — 곽위·시영 · 명군 등장'),
        ('959',     '후주 세종 사망 — 통감의 끝점'),
        ('960',     '조광윤 — 송(宋) 건국 (통감 이후)'),
    ]
    top = 2.4
    for era, event in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.2), Inches(0.55), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.13), Inches(2.2), Inches(0.4),
                    era, font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(2.9), Inches(top), Inches(9.9), Inches(0.55), PALE)
        add_textbox(slide, Inches(3.1), Inches(top + 0.13), Inches(9.6), Inches(0.4),
                    event, font_size=13, color=INK)
        top += 0.62
    add_textbox(slide, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
                '통감의 마지막 사평 — "무인의 시대를 끝낼 문치(文治)의 출현을 기대"',
                font_size=13, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)


# ============== Ⅵ. 핵심 사평 "신광왈" ==============
@S('Ⅵ. 사평')
def s_sapyeong_intro(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅵ. 사평', page, total)
    add_title(slide, '"신광왈(臣光曰)" — 사평이란 무엇인가')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.0), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(1.7), [
        ('사평(史評) = 사서(史書) 안의 사가(史家) 해석',
         {'font_size': 19, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('단순 사건 나열이 아닌 "왜 그러했는가"의 해명',
         {'font_size': 15, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('편자의 사관(史觀)·세계관·도덕적 판단의 결정체',
         {'font_size': 14, 'color': SUB, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.5),
                    RGBColor(0xFA, 0xFA, 0xFA))
    add_paragraphs(slide, Inches(0.8), Inches(4.65), Inches(11.7), Inches(2.2), [
        ('사평의 전통',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('• 좌전 — "군자왈(君子曰)"',
         {'font_size': 14, 'space_before': 6}),
        ('• 사기 — "태사공왈(太史公曰)"',
         {'font_size': 14, 'space_before': 4}),
        ('• 한서 — "찬왈(贊曰)"',
         {'font_size': 14, 'space_before': 4}),
        ('• 자치통감 — "신광왈(臣光曰)" · 219편',
         {'font_size': 14, 'bold': True, 'color': ACCENT, 'space_before': 4}),
    ], line_spacing=1.3)


@S('Ⅵ. 사평')
def s_sapyeong_themes(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅵ. 사평', page, total)
    add_title(slide, '신광왈의 5대 주제')
    items = [
        ('정명(正名)',       '명분과 분(分)의 질서',
         '"신하가 제후가 되면 천하가 무너진다"'),
        ('재(才)와 덕(德)',  '인재의 두 기준 — 덕이 재를 이끌어야',
         '"덕이 재의 주인이고, 재는 덕의 보조"'),
        ('간언 수용',          '직간하는 신하를 보호',
         '"군주가 비판을 외면하면 망한다"'),
        ('검약·신중',          '군주의 자기 절제',
         '"수 양제가 사치로 망했다"'),
        ('점진(漸進)',          '급진적 변화의 위험',
         '"천하는 점차 변하는 것이지 급변하지 않는다"'),
    ]
    top = 2.3
    for tag, role, quote in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.5), Inches(0.85), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.22), Inches(2.5), Inches(0.5),
                    tag, font_size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(3.2), Inches(top), Inches(4.5), Inches(0.85), PALE)
        add_textbox(slide, Inches(3.35), Inches(top + 0.22), Inches(4.3), Inches(0.5),
                    role, font_size=13, bold=True, color=ACCENT)
        add_filled_rect(slide, Inches(7.8), Inches(top), Inches(5.0), Inches(0.85),
                        RGBColor(0xFA, 0xFA, 0xFA))
        add_textbox(slide, Inches(7.95), Inches(top + 0.22), Inches(4.8), Inches(0.5),
                    quote, font_size=12, color=INK, align=PP_ALIGN.CENTER)
        top += 0.97


@S('Ⅵ. 사평')
def s_sapyeong_1(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅵ. 사평', page, total)
    add_title(slide, '대표 사평 ① — 첫 사평: "예(禮)·분(分)·명(名)"',
              '"BC 403 삼가분진에 대한 사마광의 첫 평론 — 통감 전체의 사상적 출발점"')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.7), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(2.4), [
        ('"臣光曰: 天子之職 莫大於禮',
         {'font_size': 17, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        (' 禮莫大於分 分莫大於名"',
         {'font_size': 17, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('"신광이 말하기를: 천자의 직무 중 예(禮)보다 큰 것이 없고',
         {'font_size': 14, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        (' 예에서는 분(分, 분별)보다 큰 것이 없으며',
         {'font_size': 14, 'align': PP_ALIGN.CENTER}),
        (' 분에서는 명(名, 명분)보다 큰 것이 없다"',
         {'font_size': 14, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.25)
    add_filled_rect(slide, Inches(0.5), Inches(5.2), Inches(12.3), Inches(1.9),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(5.35), Inches(11.7), Inches(1.7), [
        ('해석',
         {'font_size': 16, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('"군신의 분(分)이 무너지면 — 예가 무너지고 — 천하가 무너진다"',
         {'font_size': 14, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
        ('주왕이 신하(한·위·조)를 제후로 책봉한 것이 1,500년 분열의 시작',
         {'font_size': 13, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
    ], line_spacing=1.3)


@S('Ⅵ. 사평')
def s_sapyeong_2(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅵ. 사평', page, total)
    add_title(slide, '대표 사평 ② — "재(才)와 덕(德)"',
              '"인재 평가의 두 축 — 진(晉)의 지백(智伯)에 대한 사평"')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.7), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(2.4), [
        ('"才者 德之資也   德者 才之帥也"',
         {'font_size': 22, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('재자 덕지자야 · 덕자 재지수야',
         {'font_size': 14, 'color': SUB, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
        ('', {'font_size': 4}),
        ('"재(才)는 덕(德)의 보조요, 덕은 재의 장수(將帥)다"',
         {'font_size': 16, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
    ], line_spacing=1.25)
    add_filled_rect(slide, Inches(0.5), Inches(5.2), Inches(12.3), Inches(1.9),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(5.35), Inches(11.7), Inches(1.7), [
        ('4종류의 인간',
         {'font_size': 16, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('성인(德才兼全) → 군자(德勝才) → 우인(才德皆無) → 소인(才勝德)',
         {'font_size': 14, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
        ('"소인 — 재능이 덕을 능가 — 가장 위험" (지백의 멸문 사례)',
         {'font_size': 13, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
    ], line_spacing=1.3)


@S('Ⅵ. 사평')
def s_sapyeong_3(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅵ. 사평', page, total)
    add_title(slide, '대표 사평 ③ — "당 태종과 위징"',
              '"통감이 칭찬한 군신 관계의 모범"')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.7), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(2.4), [
        ('"臣光曰: 古來人臣以諫見尊者多矣',
         {'font_size': 17, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        (' 未有如魏徵之得君也"',
         {'font_size': 17, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('"신광이 말하기를: 옛부터 간언으로 존경받은 신하가 많지만',
         {'font_size': 14, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        (' 위징처럼 군주의 신뢰를 얻은 자는 없었다"',
         {'font_size': 14, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.25)
    add_filled_rect(slide, Inches(0.5), Inches(5.2), Inches(12.3), Inches(1.9),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(5.35), Inches(11.7), Inches(1.7), [
        ('통감 사평의 패턴 — 위징·당 태종을 모든 시대의 모범으로',
         {'font_size': 16, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('"적이었던 자라도 능력이 있으면 쓴다 + 비판을 끝까지 듣는다"',
         {'font_size': 13, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
        ('이것이 정관(貞觀)의 비결이라 사마광이 거듭 강조',
         {'font_size': 13, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
    ], line_spacing=1.3)


# ============== Ⅶ. 주요 사건·인물 10선 ==============
def make_event_slide(num, title, era, summary, people, lesson):
    @S(f'Ⅶ. 주요 사건')
    def renderer(slide, page, total):
        set_white_background(slide)
        add_page_header(slide, f'Ⅶ. 주요 사건 ({num}/10)', page, total)
        add_title(slide, title, era)
        add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.0), PALE)
        add_textbox(slide, Inches(0.5), Inches(2.45), Inches(12.3), Inches(0.4),
                    '사건 개요', font_size=15, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(0.8), Inches(2.95), Inches(11.7), Inches(1.3),
                    summary, font_size=14, color=INK, align=PP_ALIGN.CENTER,
                    anchor=MSO_ANCHOR.MIDDLE)
        add_filled_rect(slide, Inches(0.5), Inches(4.5), Inches(5.9), Inches(2.4),
                        RGBColor(0xFA, 0xFA, 0xFA))
        add_textbox(slide, Inches(0.5), Inches(4.65), Inches(5.9), Inches(0.4),
                    '핵심 인물', font_size=14, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(0.7), Inches(5.1), Inches(5.5), Inches(1.7),
                    people, font_size=13, color=INK, align=PP_ALIGN.CENTER,
                    anchor=MSO_ANCHOR.MIDDLE)
        add_filled_rect(slide, Inches(6.9), Inches(4.5), Inches(5.9), Inches(2.4),
                        RGBColor(0xFA, 0xE5, 0xE5))
        add_textbox(slide, Inches(6.9), Inches(4.65), Inches(5.9), Inches(0.4),
                    '통감의 교훈', font_size=14, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(7.1), Inches(5.1), Inches(5.5), Inches(1.7),
                    lesson, font_size=13, color=INK, align=PP_ALIGN.CENTER,
                    anchor=MSO_ANCHOR.MIDDLE)
    return renderer


make_event_slide('1', '삼가분진(三家分晉)', 'BC 403 · 주 위열왕 23년',
    '진(晉)의 세 대부 한·위·조가 진을 분할 → 주왕이 그들을 제후로 책봉.\n"신하가 제후가 되는" 명분의 결정적 붕괴.',
    '한건자·위환자·조양자\n주 위열왕(책봉의 주체)',
    '통감의 첫 사건이자 첫 사평\n"명분 붕괴가 1,500년 분열의 시작"')

make_event_slide('2', '진(秦)의 통일과 멸망', 'BC 221~206 · 15년의 제국',
    '진시황 6국 통일(221) → 분서갱유 → 대규모 토목 →\n사후 14년 만에 농민 봉기로 멸망.',
    '진시황 · 이사 · 조고\n진승·오광 · 항우 · 유방',
    '"무력·법치만으로는 천하를 지킬 수 없다"\n— 정관정요·통감의 공통 반면교사')

make_event_slide('3', '초한쟁패(楚漢爭覇)', 'BC 206~202 · 4년 전쟁',
    '유방(劉邦) vs 항우(項羽) — 평민과 귀족,\n인재 활용과 자기 능력의 대결.',
    '유방 · 장량 · 소하 · 한신\n항우 · 범증 · 우희',
    '"리더십은 능력이 아니라 그릇"\n인재를 알아보고 쓰는 자가 천하를 얻는다')

make_event_slide('4', '문경의 치(文景之治)', 'BC 180~141 · 40년',
    '한 문제·경제의 무위지치(無爲之治) —\n세금 경감·법 완화·검약. 동양 황금기의 원형.',
    '한 문제 · 한 경제\n조조(晁錯) — 삭번책',
    '"무위(無爲)의 정치가 가장 큰 다스림"\n검약과 신중이 황금기의 비결')

make_event_slide('5', '한 무제와 흉노 정벌', 'BC 141~87 · 54년',
    '무제 — 흉노 정벌·실크로드 개척·유가 국교화.\n그러나 만년의 무리한 원정과 신선술로 국력 소진.',
    '한 무제 · 위청 · 곽거병\n동중서 · 사마천',
    '"한 군주의 양면 — 영광과 그림자"\n초기의 적극성이 말년의 폐단으로')

make_event_slide('6', '광무중흥(光武中興)', 'AD 25~57 · 후한 부흥',
    '왕망의 신(新)을 멸하고 한 부흥. 광무제 유수는\n절제·검약·인재 활용의 모범으로 평가.',
    '광무제 유수\n등우 · 풍이 · 마원 등 28장(將)',
    '"명군의 본보기 — 자기 절제"\n승리 후에도 겸손함을 잃지 않음')

make_event_slide('7', '수 양제의 폭정', '604~618 · 14년',
    '대운하·만리장성·고구려 3차 원정 동시 추진.\n수백만 명 동원 → 농민 봉기 → 14년 만에 멸망.',
    '수 양제 · 우세기(虞世基)\n이세민·유방 등 군웅의 봉기',
    '"진시황과 함께 통감의 양대 반면교사"\n자기 욕망의 절제 실패의 원형')

make_event_slide('8', '정관(貞觀)의 치', '627~649 · 23년',
    '당 태종 + 위징·방현령·두여회·이정 등 명신.\n중국 역사상 가장 모범적 통치 시대.',
    '당 태종 이세민\n위징 · 방현령 · 두여회',
    '통감 사평의 단골 모범 사례\n"비판 수용 + 인재 활용 + 자기 절제"')

make_event_slide('9', '안사(安史)의 난', '755~763 · 8년',
    '현종 만년의 사치·양귀비 총애 → 안녹산·사사명 봉기.\n인구 절반 사라짐. 당의 결정적 분기점.',
    '당 현종 · 양귀비 · 양국충\n안녹산 · 사사명 · 곽자의',
    '"늙은 명군의 비극"\n초심 상실의 위험성 (정관정요 신종편과 짝)')

make_event_slide('10', '오대(五代) 53년의 혼란', '907~960 · 5왕조',
    '5개 단명 왕조 평균 11년 — 무인 정치의 극치.\n후주 세종이 명군이었으나 일찍 죽음.',
    '주전충·이존욱·석경당\n후주 시영(柴榮) · 조광윤',
    '통감의 끝점 = 송 건국 직전\n"무인의 시대를 끝낼 문치(文治) 기대"')


# ============== Ⅷ. 명구절 8선 ==============
def make_quote_slide(section, hanmun, eum, mean, ref, *, hanmun_size=42):
    def renderer(slide, page, total):
        set_white_background(slide)
        add_page_header(slide, section, page, total)
        add_textbox(slide, Inches(0.5), Inches(1.2), Inches(12.3), Inches(2.4),
                    hanmun, font_size=hanmun_size, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, Inches(0.5), Inches(3.7), Inches(12.3), Inches(0.5),
                    eum, font_size=20, color=SUB, align=PP_ALIGN.CENTER)
        add_rule(slide, Inches(5.5), Inches(4.4), Inches(2.3), color=RULE, weight=1.5)
        add_textbox(slide, Inches(0.5), Inches(4.7), Inches(12.3), Inches(1.5),
                    mean, font_size=20, bold=True, color=INK,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, Inches(0.5), Inches(6.5), Inches(12.3), Inches(0.4),
                    f'— {ref}', font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    return renderer


SLIDES.append((make_quote_slide('Ⅷ. 명구 (1/8)',
    '鑑 前 世 之 興 衰\n考 當 今 之 得 失',
    '감전세지흥쇠 · 고당금지득실',
    '지난 시대의 흥망을 거울 삼아 지금의 득실을 살핀다',
    '사마광 「자치통감서(序)」', hanmun_size=30), 'Ⅷ. 명구'))

SLIDES.append((make_quote_slide('Ⅷ. 명구 (2/8)',
    '天 子 之 職 莫 大 於 禮\n禮 莫 大 於 分   分 莫 大 於 名',
    '천자지직 막대어례 · 예 막대어분 분막대어명',
    '천자의 직무는 예보다 큰 것이 없고, 예는 분(分), 분은 명(名)이 가장 크다',
    '통감 첫 사평 (BC 403 삼가분진)', hanmun_size=22), 'Ⅷ. 명구'))

SLIDES.append((make_quote_slide('Ⅷ. 명구 (3/8)',
    '才 者   德 之 資 也\n德 者   才 之 帥 也',
    '재자 덕지자야 · 덕자 재지수야',
    '재(才)는 덕(德)의 보조요, 덕은 재의 장수(將帥)다 — 인재 평가의 두 축',
    '통감 사평 (진 지백의 멸문에서)', hanmun_size=28), 'Ⅷ. 명구'))

SLIDES.append((make_quote_slide('Ⅷ. 명구 (4/8)',
    '兼 聽 則 明   偏 信 則 暗',
    '겸청즉명 편신즉암',
    '두루 들으면 밝아지고, 한쪽 말만 믿으면 어두워진다',
    '위징 → 당 태종 (정관정요·통감 한기·당기)', hanmun_size=50), 'Ⅷ. 명구'))

SLIDES.append((make_quote_slide('Ⅷ. 명구 (5/8)',
    '民 以 食 爲 天',
    '민 이 식 위 천',
    '백성에게는 먹을 것이 곧 하늘 — 한 무제기·여러 사평에 인용',
    '역이기(酈食其) → 한 고조 (한기)', hanmun_size=80), 'Ⅷ. 명구'))

SLIDES.append((make_quote_slide('Ⅷ. 명구 (6/8)',
    '上 有 所 好   下 必 甚 焉',
    '상유소호 하필심언',
    '윗사람이 좋아하는 바가 있으면 아랫사람이 반드시 그것을 따라 심해진다',
    '통감 사평 — 군주의 호오(好惡)의 위험성', hanmun_size=44), 'Ⅷ. 명구'))

SLIDES.append((make_quote_slide('Ⅷ. 명구 (7/8)',
    '禍 福 無 門   惟 人 自 召',
    '화복무문 유인자소',
    '화와 복에는 정해진 문이 없다 — 오직 사람이 스스로 부른다',
    '통감의 반복 인용 (좌전 기원)', hanmun_size=44), 'Ⅷ. 명구'))

SLIDES.append((make_quote_slide('Ⅷ. 명구 (8/8)',
    '與 善 仁 言   暖 於 布 帛\n傷 人 之 言   深 於 矛 戟',
    '여선인언 난어포백 · 상인지언 심어모극',
    '선한 말은 비단보다 따뜻하고, 사람을 다치게 하는 말은 창보다 깊다',
    '통감 사평 (순자 인용·언어 신중 강조)', hanmun_size=22), 'Ⅷ. 명구'))


# ============== Ⅸ. 정관정요와의 비교 ==============
@S('Ⅸ. 정관정요와 비교')
def s_compare_table(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅸ. 정관정요와 비교', page, total)
    add_title(slide, '정관정요 vs 자치통감 — 양대 기둥의 비교')
    rows = [
        ('편자',      '오긍(吳兢) · 사관',          '사마광(司馬光) · 재상·학자'),
        ('완성',      '8세기 초',                   '1084년'),
        ('분량',      '10권 40편 · 약 460장',        '294권 · 약 300만 자'),
        ('형식',      '주제별 분류',                 '편년체 통사'),
        ('수록 기간', '당 태종 23년',                '1,362년'),
        ('초점',      '"What" — 사례 중심',          '"Why" — 패턴·구조 중심'),
        ('성격',      '경영 핸드북',                 '거시 통사'),
        ('짝짓기',    '실용 사례집',                 '체계적 분석서'),
    ]
    top = 1.95
    add_filled_rect(slide, Inches(0.5), Inches(top), Inches(2.0), Inches(0.55), SUB)
    add_textbox(slide, Inches(0.5), Inches(top + 0.1), Inches(2.0), Inches(0.4),
                '항목', font_size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_filled_rect(slide, Inches(2.55), Inches(top), Inches(5.0), Inches(0.55), SUB)
    add_textbox(slide, Inches(2.55), Inches(top + 0.1), Inches(5.0), Inches(0.4),
                '정관정요', font_size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_filled_rect(slide, Inches(7.6), Inches(top), Inches(5.2), Inches(0.55), ACCENT)
    add_textbox(slide, Inches(7.6), Inches(top + 0.1), Inches(5.2), Inches(0.4),
                '자치통감', font_size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    row_h = 0.6
    for r_idx, row in enumerate(rows):
        y = top + 0.55 + r_idx * row_h
        bg = RGBColor(0xFA, 0xFA, 0xFA) if r_idx % 2 == 0 else WHITE
        add_filled_rect(slide, Inches(0.5), Inches(y), Inches(2.0), Inches(row_h), PALE)
        add_textbox(slide, Inches(0.55), Inches(y + 0.15), Inches(1.9), Inches(0.4),
                    row[0], font_size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(2.55), Inches(y), Inches(5.0), Inches(row_h), bg)
        add_textbox(slide, Inches(2.7), Inches(y + 0.15), Inches(4.7), Inches(0.4),
                    row[1], font_size=13, color=INK, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(7.6), Inches(y), Inches(5.2), Inches(row_h), bg)
        add_textbox(slide, Inches(7.75), Inches(y + 0.15), Inches(4.9), Inches(0.4),
                    row[2], font_size=13, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)


@S('Ⅸ. 정관정요와 비교')
def s_compare_essence(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅸ. 정관정요와 비교', page, total)
    add_title(slide, '두 책의 본질 — "사례"와 "통사"의 짝')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.0), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(1.7), [
        ('"정관정요는 백과사전, 자치통감은 통사"',
         {'font_size': 19, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('정관정요 — 주제별로 펼쳐보는 사례집',
         {'font_size': 15, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('자치통감 — 시간순으로 읽는 큰 강물',
         {'font_size': 15, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.5),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(4.65), Inches(11.7), Inches(2.2), [
        ('이상적 학습 순서',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('① 정관정요 — 한 주제씩 펼쳐 구체 사례를 익힘',
         {'font_size': 14, 'space_before': 6}),
        ('② 자치통감 — 1,362년의 큰 흐름 안에서 그 사례를 재배치',
         {'font_size': 14, 'space_before': 4}),
        ('③ 정관정요로 돌아가 더 깊이 음미',
         {'font_size': 14, 'space_before': 4}),
        ('— 세종이 평생 두 책을 짝지어 학습한 이유 —',
         {'font_size': 13, 'color': SUB, 'bold': True, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
    ], line_spacing=1.3)


# ============== Ⅹ. 후대 영향 ==============
@S('Ⅹ. 후대 영향')
def s_zhuxi(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅹ. 후대 영향', page, total)
    add_title(slide, '주희(朱熹)의 『자치통감강목(綱目)』 — 도덕적 재해석')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(0.8), ACCENT)
    add_textbox(slide, Inches(0.5), Inches(2.45), Inches(12.3), Inches(0.55),
                '주희(朱熹, 1130~1200) · 1172년 완성 · 59권',
                font_size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_filled_rect(slide, Inches(0.5), Inches(3.3), Inches(12.3), Inches(1.8), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(3.45), Inches(11.7), Inches(1.6), [
        ('"강(綱)"과 "목(目)"의 이중 구조',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('• 강(綱) — 큰 사건의 도덕적 판단 (춘추 필법)',
         {'font_size': 14, 'space_before': 4}),
        ('• 목(目) — 세부 사실의 서술',
         {'font_size': 14, 'space_before': 2}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(0.5), Inches(5.3), Inches(12.3), Inches(1.7),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(5.45), Inches(11.7), Inches(1.5), [
        ('성리학적 재해석',
         {'font_size': 17, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('"통감이 사실 중심이라면, 강목은 도덕적 평가"',
         {'font_size': 14, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
        ('조선에서 더 널리 읽힌 책 — 세종의 경연 교재',
         {'font_size': 13, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
    ], line_spacing=1.3)


@S('Ⅹ. 후대 영향')
def s_dunkyong(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅹ. 후대 영향', page, total)
    add_title(slide, '진덕수의 『대학연의』 — 통감 사례의 체계화 (1264)')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.0), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(1.7), [
        ('진덕수(眞德秀, 1178~1235) — 송 후기 학자',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('『대학연의』 43권 — 『대학』의 8조목을 자치통감 사례로 풀어냄',
         {'font_size': 15, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('통감의 통사적 흐름 + 대학의 체계적 이론 = 황제 교과서',
         {'font_size': 14, 'color': SUB, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.5),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(4.65), Inches(11.7), Inches(2.2), [
        ('조선에서의 위상',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('세종이 100번 이상 완독 · 숙종·정조도 평생 학습',
         {'font_size': 15, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('"임금이 반드시 공부해야 하는 책"',
         {'font_size': 14, 'color': SUB, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('통감의 사례가 대학연의를 통해 동아시아 황실 표준이 됨',
         {'font_size': 13, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
    ], line_spacing=1.25)


@S('Ⅹ. 후대 영향')
def s_qing(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅹ. 후대 영향', page, total)
    add_title(slide, '청대 황제들 — 강희·옹정·건륭의 평생 학습서')
    items = [
        ('강희제(康熙)', '1654~1722', '8세 즉위 · 61년 재위 · 청 황금기의 시작',
         '"통감을 매일 한 권씩 읽었다" — 어전 강의 기록'),
        ('옹정제(雍正)', '1678~1735', '강희의 학습 전통 계승 · 율법 정비',
         '"통감의 사평을 자기 정치 판단의 기준으로"'),
        ('건륭제(乾隆)', '1711~1799', '60년 재위 · 청 최성기',
         '『어비통감집람』 — 건륭이 친히 평한 통감 주석본'),
    ]
    top = 2.4
    for name, era, char, quote in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.5), Inches(1.3), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.32), Inches(2.5), Inches(0.5),
                    name, font_size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(0.6), Inches(top + 0.8), Inches(2.5), Inches(0.4),
                    era, font_size=12, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(3.2), Inches(top), Inches(9.6), Inches(1.3), PALE)
        add_textbox(slide, Inches(3.35), Inches(top + 0.18), Inches(9.3), Inches(0.45),
                    char, font_size=13, bold=True, color=INK)
        add_textbox(slide, Inches(3.35), Inches(top + 0.65), Inches(9.3), Inches(0.5),
                    quote, font_size=12, color=ACCENT, bold=True)
        top += 1.45


@S('Ⅹ. 후대 영향')
def s_sejong(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅹ. 후대 영향', page, total)
    add_title(slide, '조선 세종(世宗) — 통감 학습의 모범',
              '"즉위 직후 집현전에 통감 강독을 명함"')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.0),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(1.7), [
        ('세종의 통감 학습',
         {'font_size': 19, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('• 경연(經筵)에서 통감강목·대학연의를 가장 많이 강독',
         {'font_size': 14, 'space_before': 4}),
        ('• 집현전 학자들과 함께 통감 정독',
         {'font_size': 14, 'space_before': 4}),
        ('• "통감의 사례가 곧 정치의 거울"',
         {'font_size': 14, 'bold': True, 'color': ACCENT, 'space_before': 4}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.5), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(4.65), Inches(11.7), Inches(2.2), [
        ('『자치통감훈의(資治通鑑訓義)』 (1438) — 세종의 명저',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('세종이 직접 명령 · 집현전이 편찬',
         {'font_size': 15, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('통감의 한국어(이두) 주석본 — 조선 학자의 필독서',
         {'font_size': 14, 'color': SUB, 'align': PP_ALIGN.CENTER}),
        ('"동양 통감 연구사에서 한국이 남긴 가장 큰 발자국"',
         {'font_size': 13, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
    ], line_spacing=1.25)


@S('Ⅹ. 후대 영향')
def s_japan(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅹ. 후대 영향', page, total)
    add_title(slide, '일본 — 막부 교육의 핵심 텍스트')
    items = [
        ('도쿠가와 이에야스',  '1543~1616',
         '『정관정요』·『자치통감』을 통치의 거울로',
         '"치도(治道)의 두 책" — 막부 300년 평화의 사상적 토대'),
        ('하야시 라잔',         '1583~1657',
         '도쿠가와 시대 막부 어용 학자',
         '에도 막부 관료에게 통감 강독'),
        ('아라이 하쿠세키',     '1657~1725',
         '6대 쇼군 이에노부의 시강(侍講)',
         '통감의 사평을 그대로 막부 정치에 응용'),
        ('요시다 쇼인',         '1830~1859',
         '막부 말기 사상가 · 메이지 유신의 정신적 스승',
         '제자들에게 통감을 반복 강의'),
    ]
    top = 2.4
    for name, era, role, char in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.5), Inches(0.95), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.18), Inches(2.5), Inches(0.4),
                    name, font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(0.6), Inches(top + 0.58), Inches(2.5), Inches(0.3),
                    era, font_size=11, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(3.2), Inches(top), Inches(4.0), Inches(0.95), PALE)
        add_textbox(slide, Inches(3.35), Inches(top + 0.28), Inches(3.8), Inches(0.5),
                    role, font_size=12, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(7.3), Inches(top), Inches(5.5), Inches(0.95),
                        RGBColor(0xFA, 0xFA, 0xFA))
        add_textbox(slide, Inches(7.45), Inches(top + 0.28), Inches(5.3), Inches(0.5),
                    char, font_size=12, color=INK)
        top += 1.07


@S('Ⅹ. 후대 영향')
def s_chronicles(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅹ. 후대 영향', page, total)
    add_title(slide, '동아시아 후속 통사 — 통감 형식의 확산')
    items = [
        ('통감기사본말(通鑑紀事本末)', '원추 · 1174',
         '사건별 재편 → 기사본말체 시조'),
        ('통감강목(通鑑綱目)',          '주희 · 1172',
         '강(綱)·목(目) 구조 · 도덕적 재해석'),
        ('속자치통감장편(續資治通鑑長編)', '이도(李燾) · 1183',
         '북송 시대(960~1100)를 자치통감 형식으로'),
        ('속자치통감(續資治通鑑)',      '필원(畢沅) · 청대',
         '송~원(960~1370) 411년을 통감 형식으로'),
        ('명통감(明通鑑)',              '하섭 · 청대',
         '명대(1368~1644) 277년을 통감 형식으로'),
        ('대일본사(大日本史)',          '미토 · 1657~1906',
         '일본 황통(皇統) 통사 · 통감 영향'),
    ]
    top = 2.3
    for tag, author, char in items:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(3.5), Inches(0.55), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.13), Inches(3.5), Inches(0.4),
                    tag, font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(4.2), Inches(top), Inches(2.8), Inches(0.55), PALE)
        add_textbox(slide, Inches(4.35), Inches(top + 0.13), Inches(2.6), Inches(0.4),
                    author, font_size=12, bold=True, color=ACCENT)
        add_filled_rect(slide, Inches(7.1), Inches(top), Inches(5.7), Inches(0.55),
                        RGBColor(0xFA, 0xFA, 0xFA))
        add_textbox(slide, Inches(7.25), Inches(top + 0.13), Inches(5.5), Inches(0.4),
                    char, font_size=12, color=INK)
        top += 0.62


# ============== Ⅺ. 현대 의의 ==============
@S('Ⅺ. 현대 의의')
def s_modern_macro(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅺ. 현대 의의', page, total)
    add_title(slide, '현대 ① — 통사적(macro) 시각의 회복')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.0), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(1.7), [
        ('"단편적 정보 시대에 통사의 시각이 가장 결핍"',
         {'font_size': 19, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('SNS·뉴스·블로그 — 짧고 자극적인 단편이 지배',
         {'font_size': 15, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('1,362년의 큰 강물을 본 자만이 오늘의 작은 물결을 읽는다',
         {'font_size': 14, 'color': SUB, 'bold': True, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.5),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(4.65), Inches(11.7), Inches(2.2), [
        ('통사적 시각의 실용성',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('• 패턴 인식 — "이번에는 다르다"는 착각의 해독제',
         {'font_size': 14, 'space_before': 6}),
        ('• 인과 추적 — 즉각적 사건의 멀고 깊은 원인',
         {'font_size': 14, 'space_before': 4}),
        ('• 사이클 인식 — 흥(興)과 쇠(衰)의 반복 구조',
         {'font_size': 14, 'space_before': 4}),
        ('• 경영의 long-term thinking · 정치의 100년 안목',
         {'font_size': 14, 'bold': True, 'color': ACCENT, 'space_before': 4}),
    ], line_spacing=1.3)


@S('Ⅺ. 현대 의의')
def s_modern_sapyeong(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅺ. 현대 의의', page, total)
    add_title(slide, '현대 ② — "신광왈" — 사례에 평론을 더하는 사고법')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.0), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(1.7), [
        ('"사실 → 평론" — 정보를 지혜로 바꾸는 사고 회로',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('자치통감은 단순 기록이 아니라 219편의 사평이 동반된 책',
         {'font_size': 15, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('"읽고 생각하지 않으면 통감이 아니다"',
         {'font_size': 14, 'color': SUB, 'bold': True, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.5),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(4.65), Inches(11.7), Inches(2.2), [
        ('현대 적용 — "사평의 습관"',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('• 뉴스 — "왜 이런 일이?"를 자문',
         {'font_size': 14, 'space_before': 6}),
        ('• 사건 — "이전에 같은 패턴이?"를 자문',
         {'font_size': 14, 'space_before': 4}),
        ('• 결정 — "이 결정이 5년 후 어떻게 평가될까?"',
         {'font_size': 14, 'bold': True, 'color': ACCENT, 'space_before': 4}),
    ], line_spacing=1.3)


@S('Ⅺ. 현대 의의')
def s_modern_mirror(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅺ. 현대 의의', page, total)
    add_title(slide, '현대 ③ — "감(鑑)" — 거울로서의 역사')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.0), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(1.7), [
        ('"역사는 단순 과거의 기록이 아니라 현재를 비추는 거울"',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('당 태종의 "三鏡" · 사마광의 "鑑" — 같은 사상',
         {'font_size': 15, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('현재를 객관화하는 가장 강력한 도구 — 시간적 거리',
         {'font_size': 14, 'color': SUB, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.5),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(4.65), Inches(11.7), Inches(2.2), [
        ('현대 응용 — "역사가 보여주는 거울"',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('• 경영 — "코닥·노키아·블랙베리는 왜 무너졌나"',
         {'font_size': 14, 'space_before': 6}),
        ('• 조직 — "엔론·리먼은 왜 망했나"',
         {'font_size': 14, 'space_before': 4}),
        ('• 개인 — "지난 5년의 나는 어떤 군주였나"',
         {'font_size': 14, 'bold': True, 'color': ACCENT, 'space_before': 4}),
    ], line_spacing=1.3)


@S('Ⅺ. 현대 의의')
def s_modern_critique(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅺ. 현대 의의', page, total)
    add_title(slide, '현대 ④ — "고이(考異)" — 사료 비판의 정신')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.0), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(1.7), [
        ('사마광의 『통감고이(通鑑考異)』 30권',
         {'font_size': 19, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('"왜 이 사료를 채택했는가" — 모든 채택의 근거를 명시',
         {'font_size': 15, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('동양 사학의 가장 엄밀한 사료 비판의 원형',
         {'font_size': 14, 'color': SUB, 'bold': True, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.5),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(4.65), Inches(11.7), Inches(2.2), [
        ('현대 응용 — "정보 비판의 습관"',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('• 가짜 뉴스 시대의 사료 비판 — 출처·교차 검증',
         {'font_size': 14, 'space_before': 6}),
        ('• AI 답변의 검증 — 환각(hallucination) 식별',
         {'font_size': 14, 'space_before': 4}),
        ('• "왜 이 사료/정보를 채택했는가"의 사고 습관',
         {'font_size': 14, 'bold': True, 'color': ACCENT, 'space_before': 4}),
    ], line_spacing=1.3)


@S('Ⅺ. 현대 의의')
def s_modern_progressive(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅺ. 현대 의의', page, total)
    add_title(slide, '현대 ⑤ — "점진(漸進)" — 급진적 개혁에 대한 신중')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.0), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(1.7), [
        ('사마광의 평생 사상 — "점진의 정치"',
         {'font_size': 19, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('왕안석 신법과의 대립이 통감의 사관(史觀)에 영향',
         {'font_size': 15, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('"역사는 점차 변하는 것이지 급변하지 않는다"',
         {'font_size': 14, 'color': SUB, 'bold': True, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.5),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(4.65), Inches(11.7), Inches(2.2), [
        ('현대 적용 — 변화 관리의 지혜',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('• 조직 변화 — 급진보다 점진의 효과',
         {'font_size': 14, 'space_before': 6}),
        ('• "왕안석의 실패 — 옳은 이상이 잘못된 실행으로 무산"',
         {'font_size': 14, 'space_before': 4}),
        ('• 진(秦)·수(隋)의 폐망 — 통감의 점진 사상의 실증',
         {'font_size': 14, 'bold': True, 'color': ACCENT, 'space_before': 4}),
    ], line_spacing=1.3)


@S('Ⅺ. 현대 의의')
def s_modern_who(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅺ. 현대 의의', page, total)
    add_title(slide, '현대 ⑥ — 우리 시대의 자치통감은 누가 쓰는가')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(2.0), PALE)
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(1.7), [
        ('"우리 시대도 1,362년의 통감을 갖고 싶다"',
         {'font_size': 19, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('19년간 사학에 몰두한 사마광 같은 학자가 필요한 시대',
         {'font_size': 15, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('단편 정보의 홍수 속에 통합적 시각이 점점 희소',
         {'font_size': 14, 'color': SUB, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.3)
    add_filled_rect(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.5),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(4.65), Inches(11.7), Inches(2.2), [
        ('우리 각자의 "작은 통감"',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 4}),
        ('• 자기 인생의 통감 — 지난 30년을 거울로',
         {'font_size': 14, 'space_before': 6}),
        ('• 조직의 통감 — 회사 10년사를 거울로',
         {'font_size': 14, 'space_before': 4}),
        ('• 가족의 통감 — 3대(代)의 거울로',
         {'font_size': 14, 'space_before': 4}),
        ('"통감의 정신은 누구나 자기 시대의 사관(史官)이 되는 일"',
         {'font_size': 13, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
    ], line_spacing=1.25)


# ============== Ⅻ. 마무리 ==============
@S('Ⅻ. 마무리')
def s_summary(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅻ. 마무리', page, total)
    add_title(slide, '한 문장으로 정리하는 자치통감')
    add_filled_rect(slide, Inches(0.7), Inches(2.0), Inches(11.9), Inches(5.0), PALE)
    add_paragraphs(slide, Inches(1.1), Inches(2.2), Inches(11.1), Inches(4.7), [
        ('사마광이 19년 낙양 사학(史學)에서',
         {'font_size': 17, 'color': INK, 'align': PP_ALIGN.CENTER}),
        ('류반·류서·범조우와 함께 편찬한',
         {'font_size': 17, 'color': INK, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 6}),
        ('주(周) 위열왕 BC 403 — 후주 세종 959',
         {'font_size': 17, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER, 'space_before': 8}),
        ('1,362년의 16기 294권 약 300만 자 통사',
         {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
        ('', {'font_size': 6}),
        ('219편의 "신광왈" 사평으로',
         {'font_size': 17, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 8}),
        ('역사를 거울 삼아 정치의 도(道)를 비추는',
         {'font_size': 19, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
        ('', {'font_size': 4}),
        ('동양 통사(通史)의 최고봉',
         {'font_size': 24, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 8}),
        ('', {'font_size': 4}),
        ('— 정관정요와 함께 동양 제왕학의 양대 기둥 —',
         {'font_size': 15, 'color': SUB, 'bold': True, 'align': PP_ALIGN.CENTER, 'space_before': 12}),
    ], line_spacing=1.2)


@S('Ⅻ. 마무리')
def s_book_meaning(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅻ. 마무리', page, total)
    add_title(slide, '책 이름이 답이다 — "자치통감"의 사명')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(4.6),
                    RGBColor(0xFA, 0xE5, 0xE5))
    add_paragraphs(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(4.3), [
        ('資 治 通 鑑',
         {'font_size': 50, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 8}),
        ('資 — 도움이 되는',
         {'font_size': 18, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 10}),
        ('治 — 다스림에',
         {'font_size': 18, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
        ('通 — 역대(歷代)에 통하는',
         {'font_size': 18, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
        ('鑑 — 거울',
         {'font_size': 22, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
        ('', {'font_size': 6}),
        ('"다스림에 도움이 되고, 역대에 통하여, 거울이 되는 책"',
         {'font_size': 16, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER, 'space_before': 12}),
        ('— 송 신종이 직접 명명한 사명 —',
         {'font_size': 14, 'color': SUB, 'bold': True, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
    ], line_spacing=1.25)


@S('Ⅻ. 마무리')
def s_closing(slide, page, total):
    set_white_background(slide)
    add_filled_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.4), RULE)
    add_textbox(slide, Inches(0.5), Inches(1.5), Inches(12.3), Inches(1.5),
                '臣 光 曰',
                font_size=130, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.2), Inches(12.3), Inches(0.5),
                '신 광 왈', font_size=24, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.0), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(4.3), Inches(12.3), Inches(0.7),
                '"신하 사마광이 아룁니다"',
                font_size=22, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.1), Inches(12.3), Inches(0.5),
                '— 219편의 사평이 1,362년의 흥망을 비추는 한 마디로 시작 —',
                font_size=14, color=ACCENT, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.8), Inches(12.3), Inches(0.5),
                '"역사는 단순한 과거가 아니라 지금을 비추는 거울이다"',
                font_size=15, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.45), Inches(12.3), Inches(0.4),
                '— 사마광의 "鑑前世之興衰, 考當今之得失" —',
                font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.4),
                '감사합니다', font_size=24, bold=True, color=INK,
                align=PP_ALIGN.CENTER)
    add_filled_rect(slide, Inches(0), Inches(7.1), Inches(13.333), Inches(0.4), RULE)


# ============== 빌드 ==============
total_pages = len(SLIDES)
print(f'Total {total_pages} slides — building ...')
for i, (renderer, section) in enumerate(SLIDES, 1):
    slide = prs.slides.add_slide(blank)
    renderer(slide, i, total_pages)

out_path = r'C:\Users\박호군\ClaudeProjects\Oriental-Classics2\발표자료\자치통감_발표자료.pptx'
prs.save(out_path)
print(f'Saved: {out_path}  ({total_pages} slides)')