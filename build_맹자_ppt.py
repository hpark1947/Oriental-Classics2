# -*- coding: utf-8 -*-
"""
맹자 발표자료 재작성 스크립트
원칙:
  1) 배경 무색(흰색)
  2) 한 슬라이드 한 주제
  3) 슬라이드 매수 무제약
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


def set_white_background(slide):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


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
        run.font.italic = opts.get('italic', False)
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
                title, font_size=32, bold=True, color=INK)
    if subtitle:
        add_textbox(slide, Inches(0.5), Inches(1.55), Inches(12.8), Inches(0.4),
                    subtitle, font_size=15, color=SUB)


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


# ---------- 1. 표지 ----------
@S('표지')
def s_cover(slide, page, total):
    set_white_background(slide)
    add_filled_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.4), RULE)
    add_textbox(slide, Inches(0.5), Inches(1.8), Inches(12.3), Inches(1.6),
                '孟 子', font_size=110, bold=True, color=INK,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.7), Inches(12.3), Inches(0.5),
                'Mencius · 맹자', font_size=24, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.5), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(4.8), Inches(12.3), Inches(0.5),
                '맹가(孟軻)의 사상과 언행록 — 아성(亞聖)의 가르침',
                font_size=20, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.3), Inches(12.3), Inches(0.4),
                '전국시대 (BC 372~289경) · 사서(四書) · 7편 14장 · 성선설과 왕도정치',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    add_filled_rect(slide, Inches(0), Inches(7.1), Inches(13.333), Inches(0.4), RULE)


# ---------- 2. 목차 ----------
@S('목차')
def s_toc(slide, page, total):
    set_white_background(slide)
    add_textbox(slide, Inches(0.5), Inches(0.5), Inches(12.8), Inches(0.7),
                '목 차', font_size=36, bold=True, color=INK)
    add_rule(slide, Inches(0.5), Inches(1.3), Inches(12.8))
    items = [
        ('Ⅰ', '개요 — 맹자란 무엇인가'),
        ('Ⅱ', '7편 14장의 구성'),
        ('Ⅲ', '핵심 사상 ① — 성선설(性善說)과 사단(四端)'),
        ('Ⅳ', '핵심 사상 ② — 왕도정치(王道政治)와 민본'),
        ('Ⅴ', '핵심 사상 ③ — 호연지기(浩然之氣)와 수양론'),
        ('Ⅵ', '명구절 10선'),
        ('Ⅶ', '맹자의 구조적 특징'),
        ('Ⅷ', '현대적 의의'),
        ('Ⅸ', '다른 고전과의 비교'),
        ('Ⅹ', '마무리'),
    ]
    top = 1.7
    for num, title in items:
        add_textbox(slide, Inches(1.2), Inches(top), Inches(1.0), Inches(0.4),
                    num, font_size=22, bold=True, color=ACCENT)
        add_textbox(slide, Inches(2.4), Inches(top), Inches(10.0), Inches(0.4),
                    title, font_size=20, color=INK)
        top += 0.5


# ---------- Ⅰ. 개요 ----------
@S('Ⅰ. 개요')
def s_what_is(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅰ. 개요', page, total)
    add_title(slide, '맹자(孟子)란 무엇인가')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(1.4), PALE)
    add_textbox(slide, Inches(0.9), Inches(2.55), Inches(11.5), Inches(0.5),
                '전국시대 유학자 맹자의 사상과 언행을 담은 유교 경전',
                font_size=24, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.9), Inches(3.15), Inches(11.5), Inches(0.5),
                '논어가 간결한 어록이라면, 맹자는 체계적 논변과 비유의 책',
                font_size=17, color=ACCENT, align=PP_ALIGN.CENTER)
    nums = [('7', '편(篇)'), ('14', '장(상·하)'), ('약 260', '장(章)'), ('약 35,000', '자(字)')]
    for i, (n, lbl) in enumerate(nums):
        x = 0.7 + i * 3.1
        add_textbox(slide, Inches(x), Inches(4.4), Inches(3.0), Inches(1.0),
                    n, font_size=56, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(x), Inches(5.6), Inches(3.0), Inches(0.5),
                    lbl, font_size=17, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅰ. 개요')
def s_mencius(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅰ. 개요', page, total)
    add_title(slide, '맹자(孟軻, BC 372~289경)', '아성(亞聖) — 공자 다음의 성인')
    lines = [
        ('출신', {'bold': True, 'font_size': 18, 'color': ACCENT}),
        ('  추(鄒)나라 출신, 공자의 고향 노(魯)나라와 인접', {'font_size': 18}),
        ('수학(受學)', {'bold': True, 'font_size': 18, 'color': ACCENT, 'space_before': 12}),
        ('  자사(子思)의 문인에게 배움 — 공자 학통의 정통 계승자', {'font_size': 18}),
        ('활동', {'bold': True, 'font_size': 18, 'color': ACCENT, 'space_before': 12}),
        ('  제(齊)·위(魏)·송(宋)·등(滕) 등 여러 나라를 유세하며 왕도정치 설파',
         {'font_size': 18}),
        ('존칭', {'bold': True, 'font_size': 18, 'color': ACCENT, 'space_before': 12}),
        ('  아성(亞聖) — 공자에 버금가는 성인',
         {'font_size': 18}),
    ]
    add_paragraphs(slide, Inches(0.9), Inches(2.3), Inches(11.5), Inches(4.5),
                   lines, line_spacing=1.4)


@S('Ⅰ. 개요')
def s_three_moves(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅰ. 개요', page, total)
    add_title(slide, '맹모삼천지교(孟母三遷之敎)', '— 교육 환경의 결정적 중요성')
    moves = [
        ('1차', '묘지 부근',  '아이가 장례 흉내를 내자 이사를 결심'),
        ('2차', '시장 부근',  '상인 흉내를 내자 다시 이사'),
        ('3차', '서당 부근',  '예의·학문을 본받자 비로소 정착'),
    ]
    top = 2.4
    for tag, place, desc in moves:
        add_filled_rect(slide, Inches(0.7), Inches(top), Inches(1.4), Inches(1.1), ACCENT)
        add_textbox(slide, Inches(0.7), Inches(top + 0.3), Inches(1.4), Inches(0.6),
                    tag, font_size=28, bold=True,
                    color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(2.3), Inches(top + 0.05), Inches(10.5), Inches(0.5),
                    place, font_size=22, bold=True, color=INK)
        add_textbox(slide, Inches(2.3), Inches(top + 0.55), Inches(10.5), Inches(0.5),
                    desc, font_size=16, color=SUB)
        top += 1.4
    add_textbox(slide, Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.4),
                '환경이 사람을 만든다 — 후대 동아시아 교육 사상의 원형이 됨',
                font_size=14, color=ACCENT, bold=True, align=PP_ALIGN.CENTER)


@S('Ⅰ. 개요')
def s_dotong(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅰ. 개요', page, total)
    add_title(slide, '도통(道統) 계보', '맹자가 자임한 공자 학통의 정통 후계')
    figures = [
        ('공자', '孔子', '(BC 551~479)', '유교의 창시자'),
        ('증자', '曾子', '(BC 505~436)', '효(孝)의 사상, 대학 전수'),
        ('자사', '子思', '(BC 483~402)', '공자 손자, 중용 전수'),
        ('맹자', '孟子', '(BC 372~289)', '성선설·왕도정치 체계화'),
    ]
    box_w = 2.95
    gap = 0.15
    start_x = (13.333 - (box_w * 4 + gap * 3)) / 2
    for i, (name, hanmun, period, desc) in enumerate(figures):
        x = start_x + i * (box_w + gap)
        add_filled_rect(slide, Inches(x), Inches(2.5), Inches(box_w), Inches(4.2), PALE)
        add_textbox(slide, Inches(x), Inches(2.7), Inches(box_w), Inches(1.0),
                    hanmun, font_size=72, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(x), Inches(4.3), Inches(box_w), Inches(0.5),
                    name, font_size=22, bold=True, color=INK,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(x), Inches(4.9), Inches(box_w), Inches(0.4),
                    period, font_size=13, color=SUB, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(x), Inches(5.5), Inches(box_w), Inches(1.0),
                    desc, font_size=14, color=INK, align=PP_ALIGN.CENTER)
        # 화살표
        if i < 3:
            arrow_x = x + box_w + 0.01
            add_textbox(slide, Inches(arrow_x - 0.03), Inches(4.3), Inches(0.2), Inches(0.5),
                        '▶', font_size=18, color=ACCENT, align=PP_ALIGN.CENTER)


@S('Ⅰ. 개요')
def s_compilation(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅰ. 개요', page, total)
    add_title(slide, '편찬 과정')
    add_filled_rect(slide, Inches(0.6), Inches(2.3), Inches(6.0), Inches(4.5), PALE)
    add_textbox(slide, Inches(0.6), Inches(2.5), Inches(6.0), Inches(0.5),
                '1단계 — 자술설(自述說)', font_size=18, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_paragraphs(slide, Inches(0.9), Inches(3.2), Inches(5.5), Inches(3.5), [
        ('맹자 본인이 직접 편찬', {'font_size': 18, 'bold': True}),
        ('', {'font_size': 6}),
        ('제자 만장(萬章)·공손추(公孫丑)와',
         {'font_size': 16, 'space_before': 6}),
        ('함께 자료를 정리한 것으로 추정',
         {'font_size': 16}),
        ('', {'font_size': 6}),
        ('논어와 달리 사후 편집이 아닌',
         {'font_size': 15, 'color': SUB, 'space_before': 6}),
        ('생전 자술의 성격이 강함',
         {'font_size': 15, 'color': SUB}),
    ], line_spacing=1.4)
    add_filled_rect(slide, Inches(6.8), Inches(2.3), Inches(6.0), Inches(4.5), PALE)
    add_textbox(slide, Inches(6.8), Inches(2.5), Inches(6.0), Inches(0.5),
                '2단계 — 후한의 정본 확정', font_size=18, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_paragraphs(slide, Inches(7.1), Inches(3.2), Inches(5.5), Inches(3.5), [
        ('조기(趙岐)의 주석', {'font_size': 18, 'bold': True}),
        ('', {'font_size': 6}),
        ('후한 시기 조기가 주석을 달면서',
         {'font_size': 16, 'space_before': 6}),
        ('7편을 각각 상·하로 분할 → 14권',
         {'font_size': 16}),
        ('', {'font_size': 6}),
        ('이것이 현재까지 이어지는',
         {'font_size': 15, 'color': SUB, 'space_before': 6}),
        ('맹자의 표준 체제',
         {'font_size': 15, 'color': SUB}),
    ], line_spacing=1.4)


# ---------- Ⅱ. 구성 ----------
@S('Ⅱ. 구성')
def s_structure_overview(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅱ. 구성', page, total)
    add_title(slide, '7편 14장의 구성 개관')
    add_filled_rect(slide, Inches(0.7), Inches(2.3), Inches(11.9), Inches(0.7), ACCENT)
    add_textbox(slide, Inches(0.7), Inches(2.45), Inches(11.9), Inches(0.4),
                '총 7편 · 상·하 분권 14권 · 약 260장',
                font_size=22, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF),
                align=PP_ALIGN.CENTER)
    # 7편 한 줄에 4 + 3
    names = ['양혜왕\n梁惠王', '공손추\n公孫丑', '등문공\n滕文公', '이루\n離婁',
             '만장\n萬章', '고자\n告子', '진심\n盡心']
    # 첫 줄 4편, 둘째 줄 3편
    for i in range(4):
        x = 0.7 + i * 3.0
        add_filled_rect(slide, Inches(x), Inches(3.2), Inches(2.85), Inches(1.7), PALE)
        add_textbox(slide, Inches(x), Inches(3.4), Inches(2.85), Inches(0.4),
                    f'第 {i+1} 편', font_size=13, color=SUB,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(x), Inches(3.8), Inches(2.85), Inches(1.0),
                    names[i], font_size=20, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
    for i in range(3):
        x = 1.45 + i * 3.5
        add_filled_rect(slide, Inches(x), Inches(5.0), Inches(3.35), Inches(1.7), PALE)
        add_textbox(slide, Inches(x), Inches(5.2), Inches(3.35), Inches(0.4),
                    f'第 {i+5} 편', font_size=13, color=SUB,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(x), Inches(5.6), Inches(3.35), Inches(1.0),
                    names[i+4], font_size=20, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.4),
                '각 편은 첫 장의 인물 또는 핵심어로 이름을 삼는 어록체 전통',
                font_size=14, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅱ. 구성')
def s_seven_chapters(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅱ. 구성', page, total)
    add_title(slide, '7편 — 편명과 핵심 주제')
    rows = [
        ('1', '양혜왕(梁惠王)', '상7+하16', '왕도정치, 인정(仁政), 의리(義利) 분별'),
        ('2', '공손추(公孫丑)', '상9+하14', '호연지기(浩然之氣), 사단(四端), 왕패 구분'),
        ('3', '등문공(滕文公)', '상5+하10', '정전제(井田制), 오륜(五倫), 이단 비판'),
        ('4', '이루(離婁)',     '상28+하33', '인의(仁義)의 실천, 수신, 군신 관계'),
        ('5', '만장(萬章)',     '상9+하9',   '성인 행적, 경전 해석'),
        ('6', '고자(告子)',     '상20+하16', '성선설 논변, 의(義)의 내재성'),
        ('7', '진심(盡心)',     '상46+하38', '수양론의 총결산, 사상의 집대성'),
    ]
    top = 2.3
    for num, name, count, desc in rows:
        add_textbox(slide, Inches(0.6), Inches(top), Inches(0.6), Inches(0.4),
                    num, font_size=20, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(1.3), Inches(top), Inches(3.0), Inches(0.4),
                    name, font_size=18, bold=True, color=INK)
        add_textbox(slide, Inches(4.4), Inches(top), Inches(1.8), Inches(0.4),
                    count, font_size=14, color=SUB)
        add_textbox(slide, Inches(6.3), Inches(top), Inches(6.8), Inches(0.4),
                    desc, font_size=15, color=INK)
        top += 0.6


# ---------- Ⅲ. 성선설과 사단 ----------
@S('Ⅲ. 성선설·사단')
def s_seongseon(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 성선설·사단', page, total)
    add_title(slide, '성선설(性善說) — 인간 본성은 선하다')
    add_textbox(slide, Inches(0.9), Inches(2.4), Inches(3.5), Inches(3.5),
                '性\n善', font_size=130, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_paragraphs(slide, Inches(4.8), Inches(2.5), Inches(8.0), Inches(4.5), [
        ('핵심 명제', {'bold': True, 'font_size': 20, 'color': ACCENT}),
        ('인간의 본성(性)은 본디 선(善)하다',
         {'font_size': 18, 'bold': True}),
        ('', {'font_size': 8}),
        ('악(惡)에 대한 설명', {'bold': True, 'font_size': 20, 'color': ACCENT, 'space_before': 10}),
        ('악은 본성이 아니라', {'font_size': 17}),
        ('환경의 영향이나 본성을 기르지 못한 결과',
         {'font_size': 17}),
        ('', {'font_size': 8}),
        ('수양의 목표', {'bold': True, 'font_size': 20, 'color': ACCENT, 'space_before': 10}),
        ('잃어버린 본성을 되찾는 것 — 구방심(求放心)',
         {'font_size': 17}),
    ], line_spacing=1.35)


@S('Ⅲ. 성선설·사단')
def s_water_metaphor(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 성선설·사단', page, total)
    add_title(slide, '물의 비유 — 性善 猶水之就下')
    add_textbox(slide, Inches(0.5), Inches(2.2), Inches(12.3), Inches(1.0),
                '人性之善也   猶水之就下也',
                font_size=44, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.3), Inches(12.3), Inches(0.5),
                '인성지선야 유수지취하야', font_size=20, color=SUB,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.9), Inches(12.3), Inches(0.4),
                '— 고자편 상 2장', font_size=14, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.5), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(4.8), Inches(12.3), Inches(0.6),
                '인간의 본성이 선한 것은, 물이 아래로 흐르는 것과 같다',
                font_size=22, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.7), Inches(12.3), Inches(0.8),
                '물이 위로 튀어 오를 수는 있어도 그것이 본성이 아니듯,\n' +
                '사람이 악을 행할 수는 있어도 그것이 본성은 아니다',
                font_size=16, color=SUB, align=PP_ALIGN.CENTER)


def make_sadan_slide(section, han, eum, virtue, virtue_han, desc, ref):
    def renderer(slide, page, total):
        set_white_background(slide)
        add_page_header(slide, section, page, total)
        add_title(slide, f'사단 — {han}({eum})',
                  f'→ {virtue}({virtue_han})의 싹(端)')
        add_filled_rect(slide, Inches(0.7), Inches(2.3), Inches(5.5), Inches(4.5), PALE)
        add_textbox(slide, Inches(0.7), Inches(2.6), Inches(5.5), Inches(1.6),
                    han, font_size=80, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(0.7), Inches(4.4), Inches(5.5), Inches(0.5),
                    eum, font_size=22, color=SUB, align=PP_ALIGN.CENTER)
        add_rule(slide, Inches(2.0), Inches(5.05), Inches(2.9), color=RULE, weight=1.5)
        add_textbox(slide, Inches(0.7), Inches(5.3), Inches(5.5), Inches(1.2),
                    f'→  {virtue}\n({virtue_han})',
                    font_size=30, bold=True, color=INK, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(6.7), Inches(2.5), Inches(6.2), Inches(0.5),
                    '의미', font_size=18, bold=True, color=ACCENT)
        add_rule(slide, Inches(6.7), Inches(3.0), Inches(6.0), weight=1.5)
        add_textbox(slide, Inches(6.7), Inches(3.2), Inches(6.2), Inches(3.0),
                    desc, font_size=18, color=INK, anchor=MSO_ANCHOR.TOP)
        add_textbox(slide, Inches(6.7), Inches(6.4), Inches(6.2), Inches(0.4),
                    f'— {ref}', font_size=13, color=SUB)
    return renderer


SLIDES.append((make_sadan_slide('Ⅲ. 성선설·사단',
    '惻 隱 之 心', '측은지심', '인(仁)', '仁',
    '남의 고통을 차마 보지 못하는 마음.\n타인의 아픔에 공감하는 동정심.',
    '공손추 상 6장'), 'Ⅲ. 성선설·사단'))

SLIDES.append((make_sadan_slide('Ⅲ. 성선설·사단',
    '羞 惡 之 心', '수오지심', '의(義)', '義',
    '자신의 잘못을 부끄러워하고,\n남의 잘못을 미워하는 마음.\n도덕적 정의감의 뿌리.',
    '공손추 상 6장'), 'Ⅲ. 성선설·사단'))

SLIDES.append((make_sadan_slide('Ⅲ. 성선설·사단',
    '辭 讓 之 心', '사양지심', '예(禮)', '禮',
    '남에게 양보하는 마음.\n자신을 낮추고 타인을 존중하는 태도.',
    '공손추 상 6장'), 'Ⅲ. 성선설·사단'))

SLIDES.append((make_sadan_slide('Ⅲ. 성선설·사단',
    '是 非 之 心', '시비지심', '지(智)', '智',
    '옳고 그름을 가리는 마음.\n도덕적 판단의 능력.',
    '공손추 상 6장'), 'Ⅲ. 성선설·사단'))


@S('Ⅲ. 성선설·사단')
def s_yuja_ipjeong(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 성선설·사단', page, total)
    add_title(slide, '유자입정(孺子入井) — 성선의 결정적 증거')
    add_textbox(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(0.7),
                '孺 子 將 入 於 井',
                font_size=44, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.1), Inches(12.3), Inches(0.5),
                '유자장입어정 — "어린아이가 우물에 빠지려 한다"',
                font_size=18, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.6), Inches(12.3), Inches(0.4),
                '— 공손추 상 6장', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.2), Inches(2.3), color=RULE, weight=1.5)
    add_paragraphs(slide, Inches(1.0), Inches(4.4), Inches(11.3), Inches(2.8), [
        ('어린아이가 우물에 빠지려는 순간', {'font_size': 19, 'align': PP_ALIGN.CENTER}),
        ('누구나 깜짝 놀라 측은한 마음이 든다', {'font_size': 19, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
        ('', {'font_size': 6}),
        ('아이의 부모와 친해서가 아니고', {'font_size': 15, 'color': SUB, 'align': PP_ALIGN.CENTER, 'space_before': 10}),
        ('마을 사람들의 칭찬을 들으려 함도 아니다', {'font_size': 15, 'color': SUB, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 6}),
        ('— 사단(四端)은 누구에게나 본래부터 있다는 증거', {'font_size': 16, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER, 'space_before': 10}),
    ], line_spacing=1.3)


@S('Ⅲ. 성선설·사단')
def s_hwakchung(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 성선설·사단', page, total)
    add_title(slide, '확충(擴充) — 싹을 기르는 일')
    add_filled_rect(slide, Inches(0.6), Inches(2.3), Inches(5.9), Inches(4.5), PALE)
    add_textbox(slide, Inches(0.6), Inches(2.5), Inches(5.9), Inches(0.5),
                '사단은 "싹(端)"이다', font_size=18, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.6), Inches(3.2), Inches(5.9), Inches(1.8),
                '端',
                font_size=180, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.6), Inches(5.5), Inches(5.9), Inches(1.0),
                '완성된 덕(德)이 아니라\n발현 가능한 가능성',
                font_size=18, color=INK, align=PP_ALIGN.CENTER)
    add_filled_rect(slide, Inches(6.8), Inches(2.3), Inches(5.9), Inches(4.5), PALE)
    add_textbox(slide, Inches(6.8), Inches(2.5), Inches(5.9), Inches(0.5),
                '기르고 확장해야 한다', font_size=18, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(6.8), Inches(3.2), Inches(5.9), Inches(1.8),
                '擴 充',
                font_size=110, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(6.8), Inches(5.5), Inches(5.9), Inches(1.0),
                '확충하지 않으면 부모도 섬길 수 없고\n확충하면 천하도 보전할 수 있다',
                font_size=15, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(6.8), Inches(6.5), Inches(5.9), Inches(0.4),
                '— 공손추 상 6장', font_size=12, color=SUB, align=PP_ALIGN.CENTER)


# ---------- Ⅳ. 왕도정치와 민본 ----------
@S('Ⅳ. 왕도·민본')
def s_wangdo_vs_pae(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅳ. 왕도·민본', page, total)
    add_title(slide, '왕도(王道) vs 패도(覇道)')
    add_filled_rect(slide, Inches(0.6), Inches(2.2), Inches(6.0), Inches(0.6), ACCENT)
    add_textbox(slide, Inches(0.6), Inches(2.27), Inches(6.0), Inches(0.5),
                '왕도(王道) — 덕(德)·인(仁)', font_size=20, bold=True,
                color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
    add_filled_rect(slide, Inches(6.7), Inches(2.2), Inches(6.0), Inches(0.6), SUB)
    add_textbox(slide, Inches(6.7), Inches(2.27), Inches(6.0), Inches(0.5),
                '패도(覇道) — 힘·이익', font_size=20, bold=True,
                color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
    pairs = [
        ('덕(德)으로 다스림',       '힘(力)으로 다스림'),
        ('인(仁)으로 행함',         '이익으로 행함'),
        ('백성이 진심으로 복종',     '겉으로만 복종, 속으로 원망'),
        ('지속 가능한 통치',         '단기적 성공, 결국 무너짐'),
        ('성왕(聖王) — 요·순·우·탕·문·무', '춘추오패 — 환공·문공 등'),
    ]
    top = 3.0
    for left, right in pairs:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(6.0), Inches(0.65),
                        PALE)
        add_textbox(slide, Inches(0.7), Inches(top + 0.13), Inches(5.8), Inches(0.5),
                    left, font_size=16, color=INK, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(6.7), Inches(top), Inches(6.0), Inches(0.65),
                        RGBColor(0xF5, 0xF5, 0xF5))
        add_textbox(slide, Inches(6.8), Inches(top + 0.13), Inches(5.8), Inches(0.5),
                    right, font_size=16, color=INK, align=PP_ALIGN.CENTER)
        top += 0.75


@S('Ⅳ. 왕도·민본')
def s_hangsan(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅳ. 왕도·민본', page, total)
    add_title(slide, '항산항심(恒産恒心) — 민생이 곧 정치다')
    add_textbox(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(0.9),
                '無 恒 産 而 有 恒 心 者   惟 士 爲 能',
                font_size=34, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.3), Inches(12.3), Inches(0.5),
                '무항산이유항심자 유사위능', font_size=16, color=SUB,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.8), Inches(12.3), Inches(0.4),
                '— 양혜왕 상 7장', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.4), Inches(2.3), color=RULE, weight=1.5)
    add_paragraphs(slide, Inches(1.0), Inches(4.7), Inches(11.3), Inches(2.5), [
        ('일정한 생업(恒産)이 없으면서도', {'font_size': 18, 'align': PP_ALIGN.CENTER}),
        ('일정한 마음(恒心)을 지닐 수 있는 자는', {'font_size': 18, 'align': PP_ALIGN.CENTER}),
        ('오직 선비뿐이다', {'font_size': 18, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 6}),
        ('→ 일반 백성에게는 먼저 안정된 생업을 보장해야', {'font_size': 16, 'color': SUB, 'align': PP_ALIGN.CENTER, 'space_before': 10}),
        ('도덕적 마음과 사회적 안정이 가능하다',
         {'font_size': 16, 'color': SUB, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.3)


@S('Ⅳ. 왕도·민본')
def s_minwigi(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅳ. 왕도·민본', page, total)
    add_title(slide, '민위귀(民爲貴) — 백성이 가장 귀하다')
    add_textbox(slide, Inches(0.5), Inches(2.2), Inches(12.3), Inches(1.0),
                '民 爲 貴   社 稷 次 之   君 爲 輕',
                font_size=38, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.3), Inches(12.3), Inches(0.5),
                '민위귀   사직차지   군위경', font_size=18, color=SUB,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.85), Inches(12.3), Inches(0.4),
                '— 진심 하 14장', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.4), Inches(2.3), color=RULE, weight=1.5)
    # 위계
    items = [('民', '백성이 가장 귀하다', ACCENT, 64),
             ('社稷', '사직(국가)이 다음이며', INK, 40),
             ('君', '임금이 가장 가볍다', SUB, 64)]
    top = 4.7
    for ch, mean, c, fs in items:
        add_textbox(slide, Inches(2.0), Inches(top), Inches(2.5), Inches(0.8),
                    ch, font_size=fs, bold=True, color=c, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(5.0), Inches(top + 0.15), Inches(7.0), Inches(0.6),
                    mean, font_size=18, color=INK)
        top += 0.8


@S('Ⅳ. 왕도·민본')
def s_revolution(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅳ. 왕도·민본', page, total)
    add_title(slide, '역성혁명(易姓革命) — 폭군은 일부(一夫)다')
    add_textbox(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(0.9),
                '聞 誅 一 夫 紂 矣   未 聞 弑 君 也',
                font_size=32, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.3), Inches(12.3), Inches(0.5),
                '문주일부주의 미문시군야', font_size=16, color=SUB,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.8), Inches(12.3), Inches(0.4),
                '— 양혜왕 하 8장', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.4), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(4.7), Inches(12.3), Inches(0.6),
                '"필부 주(紂)를 죽였다는 말은 들었으나,',
                font_size=18, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.25), Inches(12.3), Inches(0.6),
                '임금을 시해했다는 말은 듣지 못했다"',
                font_size=18, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.1), Inches(12.3), Inches(0.6),
                '덕(德)을 잃은 폭군은 더 이상 군주가 아닌 일부(一夫)에 불과',
                font_size=16, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.6), Inches(12.3), Inches(0.4),
                '→ 동아시아 저항권(right of resistance)의 사상적 기원',
                font_size=14, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅳ. 왕도·민본')
def s_irue(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅳ. 왕도·민본', page, total)
    add_title(slide, '의리론(義利論) — 이익보다 의로움을')
    add_textbox(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(0.9),
                '王 何 必 曰 利   亦 有 仁 義 而 已 矣',
                font_size=32, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.3), Inches(12.3), Inches(0.5),
                '왕하필왈리 역유인의이이의', font_size=16, color=SUB,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.8), Inches(12.3), Inches(0.4),
                '— 양혜왕 상 1장 (맹자가 양혜왕을 만난 첫 만남)',
                font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.4), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(4.7), Inches(12.3), Inches(0.6),
                '"왕께서는 어찌 꼭 이익(利)을 말씀하십니까?',
                font_size=20, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.4), Inches(12.3), Inches(0.6),
                '오직 인의(仁義)가 있을 뿐입니다"',
                font_size=20, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.3), Inches(12.3), Inches(0.6),
                '맹자 사상 전체의 출발 선언 — 이익 추구 시대에 던진 도덕 정치의 깃발',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER)


# ---------- Ⅴ. 호연지기와 수양론 ----------
@S('Ⅴ. 호연지기·수양')
def s_hoyeonjigi(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 호연지기·수양', page, total)
    add_title(slide, '호연지기(浩然之氣) — 도덕적 용기의 원천')
    add_textbox(slide, Inches(0.9), Inches(2.4), Inches(3.5), Inches(3.5),
                '浩\n然\n氣', font_size=100, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_paragraphs(slide, Inches(4.8), Inches(2.4), Inches(8.2), Inches(4.6), [
        ('원전 (공손추 상 2장)', {'bold': True, 'font_size': 16, 'color': ACCENT}),
        ('其爲氣也  至大至剛', {'font_size': 22, 'bold': True, 'color': INK}),
        ('以直養而無害  則塞于天地之間',
         {'font_size': 22, 'bold': True, 'color': INK}),
        ('', {'font_size': 6}),
        ('풀이', {'bold': True, 'font_size': 16, 'color': ACCENT, 'space_before': 10}),
        ('그 기운은 지극히 크고 지극히 강하니',
         {'font_size': 15, 'color': SUB}),
        ('정직(直)함으로 기르고 해치지 않으면',
         {'font_size': 15, 'color': SUB}),
        ('천지 사이에 가득 찬다',
         {'font_size': 15, 'color': SUB}),
        ('', {'font_size': 6}),
        ('성격', {'bold': True, 'font_size': 16, 'color': ACCENT, 'space_before': 10}),
        ('도덕적 실천이 쌓여 생기는 정신적 에너지',
         {'font_size': 15}),
    ], line_spacing=1.3)


@S('Ⅴ. 호연지기·수양')
def s_jojang(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 호연지기·수양', page, total)
    add_title(slide, '조장(助長) — 억지로 자라게 하지 말라')
    add_filled_rect(slide, Inches(0.6), Inches(2.3), Inches(11.9), Inches(2.4), PALE)
    add_paragraphs(slide, Inches(0.9), Inches(2.5), Inches(11.3), Inches(2.0), [
        ('송(宋)나라의 한 농부 이야기', {'font_size': 17, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 6}),
        ('벼가 자라지 않는 것을 걱정한 농부가', {'font_size': 17, 'align': PP_ALIGN.CENTER}),
        ('벼 싹을 일일이 위로 잡아당겨 늘여놓고는', {'font_size': 17, 'align': PP_ALIGN.CENTER}),
        ('"오늘 너무 힘들었다, 벼가 자라도록 도왔다(助苗長)"고 자랑했다', {'font_size': 17, 'align': PP_ALIGN.CENTER}),
        ('이튿날 보니 — 벼는 모두 말라 죽어 있었다', {'font_size': 17, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
    ], line_spacing=1.3)
    add_textbox(slide, Inches(0.5), Inches(5.0), Inches(12.3), Inches(0.4),
                '— 공손추 상 2장', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(5.5), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(5.8), Inches(12.3), Inches(0.6),
                '호연지기는 의(義)의 축적에서 자라는 것 — 억지로 조장하면 안 된다',
                font_size=18, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.5), Inches(12.3), Inches(0.5),
                '도덕적 성장에 지름길은 없다 — 꾸준한 실천만이 답',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅴ. 호연지기·수양')
def s_5_steps(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 호연지기·수양', page, total)
    add_title(slide, '수양의 다섯 단계')
    steps = [
        ('1', '구방심', '求放心',  '잃어버린 마음을 되찾는다'),
        ('2', '존심양성', '存心養性', '마음을 보존하고 본성을 기른다'),
        ('3', '과욕',     '寡欲',     '욕심을 줄여 본성을 드러낸다'),
        ('4', '양기',     '養氣',     '호연지기를 기른다'),
        ('5', '진심지성', '盡心知性', '마음을 다하여 천(天)에 이른다'),
    ]
    top = 2.3
    for num, kor, han, desc in steps:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(0.9), Inches(0.85), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.2), Inches(0.9), Inches(0.5),
                    num, font_size=26, bold=True,
                    color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(1.7), Inches(top), Inches(2.5), Inches(0.85), PALE)
        add_textbox(slide, Inches(1.7), Inches(top + 0.12), Inches(2.5), Inches(0.4),
                    han, font_size=18, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(1.7), Inches(top + 0.48), Inches(2.5), Inches(0.4),
                    kor, font_size=13, color=SUB, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(4.4), Inches(top + 0.25), Inches(8.5), Inches(0.5),
                    desc, font_size=18, color=INK)
        top += 0.97


@S('Ⅴ. 호연지기·수양')
def s_jinsim(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 호연지기·수양', page, total)
    add_title(slide, '진심지성(盡心知性) — 수양론의 정점')
    add_textbox(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(0.9),
                '盡 其 心 者   知 其 性 也',
                font_size=36, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.2), Inches(12.3), Inches(0.9),
                '知 其 性   則 知 天 矣',
                font_size=36, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(4.1), Inches(12.3), Inches(0.5),
                '진기심자 지기성야 · 지기성 즉지천의',
                font_size=16, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(4.6), Inches(12.3), Inches(0.4),
                '— 진심 상 1장', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(5.2), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(5.5), Inches(12.3), Inches(0.6),
                '마음을 다하면 본성을 알고, 본성을 알면 하늘을 안다',
                font_size=22, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.3), Inches(12.3), Inches(0.6),
                '心 → 性 → 天 : 마음·본성·하늘이 한 길로 통한다는 동양 수양론의 정점',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER)


# ---------- Ⅵ. 명구절 ----------
def make_quote_slide(section, hanmun, eum, mean, ref, *, hanmun_size=58):
    def renderer(slide, page, total):
        set_white_background(slide)
        add_page_header(slide, section, page, total)
        add_textbox(slide, Inches(0.5), Inches(1.2), Inches(12.3), Inches(2.4),
                    hanmun, font_size=hanmun_size, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, Inches(0.5), Inches(3.7), Inches(12.3), Inches(0.5),
                    eum, font_size=22, color=SUB, align=PP_ALIGN.CENTER)
        add_rule(slide, Inches(5.5), Inches(4.4), Inches(2.3), color=RULE, weight=1.5)
        add_textbox(slide, Inches(0.5), Inches(4.7), Inches(12.3), Inches(1.3),
                    mean, font_size=22, bold=True, color=INK,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.4),
                    f'— {ref}', font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    return renderer


SLIDES.append((make_quote_slide('Ⅵ. 명구절 (1/10)',
    '王何必曰利  亦有仁義而已矣',
    '왕하필왈리   역유인의이이의',
    '왕께서 하필 이익을 말씀하십니까? 오직 인의(仁義)가 있을 뿐입니다',
    '양혜왕편 상 1장', hanmun_size=42), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (2/10)',
    '惻隱之心  仁之端也',
    '측은지심   인지단야',
    '측은히 여기는 마음은 인(仁)의 싹이다',
    '공손추편 상 6장', hanmun_size=58), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (3/10)',
    '天時不如地利  地利不如人和',
    '천시불여지리   지리불여인화',
    '하늘의 때(天時)는 땅의 이로움(地利)만 못하고,\n땅의 이로움은 사람의 화합(人和)만 못하다',
    '공손추편 하 1장', hanmun_size=42), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (4/10)',
    '民爲貴  社稷次之  君爲輕',
    '민위귀   사직차지   군위경',
    '백성이 가장 귀하고, 사직이 다음이며, 임금이 가장 가볍다',
    '진심편 하 14장', hanmun_size=44), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (5/10)',
    '五十步百步',
    '오 십 보 백 보',
    '오십 보 도망친 자가 백 보 도망친 자를 비웃다 — 정도의 차이일 뿐 본질은 같다',
    '양혜왕편 상 3장', hanmun_size=110), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (6/10)',
    '天將降大任於是人也\n必先苦其心志',
    '천장강대임어시인야   필선고기심지',
    '하늘이 큰 임무를 어떤 사람에게 내리려 할 때,\n반드시 먼저 그 마음과 뜻을 괴롭힌다',
    '고자편 하 15장', hanmun_size=32), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (7/10)',
    '生於憂患  而死於安樂也',
    '생어우환   이사어안락야',
    '근심과 환란 속에서 살아나고, 안락 속에서 죽는다',
    '고자편 하 15장', hanmun_size=48), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (8/10)',
    '捨生而取義者也',
    '사 생 이 취 의 자 야',
    '삶을 버리고 의(義)를 취하는 것이다',
    '고자편 상 10장', hanmun_size=64), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (9/10)',
    '盡其心者  知其性也\n知其性  則知天矣',
    '진기심자 지기성야 · 지기성 즉지천의',
    '마음을 다하면 본성을 알고, 본성을 알면 하늘을 안다',
    '진심편 상 1장', hanmun_size=32), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (10/10)',
    '自暴者  不可與有言也\n自棄者  不可與有爲也',
    '자포자 불가여유언야 · 자기자 불가여유위야',
    '스스로를 해치는 자와는 말할 수 없고,\n스스로를 버리는 자와는 함께 일할 수 없다',
    '이루편 상 10장', hanmun_size=30), 'Ⅵ. 명구절'))


# ---------- Ⅶ. 구조 ----------
@S('Ⅶ. 구조')
def s_argumentation(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅶ. 구조', page, total)
    add_title(slide, '웅변적 논변 — 설득의 기술')
    add_paragraphs(slide, Inches(0.7), Inches(2.3), Inches(11.9), Inches(4.5), [
        ('논어와의 차이', {'font_size': 22, 'bold': True, 'color': ACCENT}),
        ('논어: 단편적 어록 / 맹자: 완결된 논변(論辨)',
         {'font_size': 18, 'space_before': 4}),
        ('', {'font_size': 8}),
        ('맹자 논변의 특징', {'font_size': 22, 'bold': True, 'color': ACCENT, 'space_before': 14}),
        ('• 논리적이면서도 감정에 호소하는 이중적 설득',
         {'font_size': 18, 'space_before': 4}),
        ('• 상대의 말을 받아 반박하는 대화적 구조',
         {'font_size': 18}),
        ('• 단계적으로 상대를 자기 결론에 동의하게 만드는 유도법',
         {'font_size': 18}),
        ('', {'font_size': 8}),
        ('대표 사례', {'font_size': 22, 'bold': True, 'color': ACCENT, 'space_before': 14}),
        ('양혜왕·제선왕과의 대화 — 정치 철학을 논변으로 풀어낸 백미',
         {'font_size': 17, 'color': SUB}),
    ], line_spacing=1.35)


@S('Ⅶ. 구조')
def s_metaphor(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅶ. 구조', page, total)
    add_title(slide, '비유와 우화의 대가')
    metaphors = [
        ('五十步百步', '오십보백보', '본질적 차이 없는 차이를 꼬집다'),
        ('緣木求魚',   '연목구어',   '잘못된 방법으로는 목적 달성 불가'),
        ('助長',       '조장',       '억지로 자라게 하면 망친다'),
        ('牛山之木',   '우산지목',   '본성을 잃어가는 과정의 비유'),
    ]
    top = 2.4
    for han, eum, desc in metaphors:
        add_filled_rect(slide, Inches(0.7), Inches(top), Inches(3.0), Inches(0.95), ACCENT)
        add_textbox(slide, Inches(0.7), Inches(top + 0.15), Inches(3.0), Inches(0.45),
                    han, font_size=22, bold=True,
                    color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(0.7), Inches(top + 0.6), Inches(3.0), Inches(0.35),
                    eum, font_size=12,
                    color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(3.9), Inches(top), Inches(8.8), Inches(0.95), PALE)
        add_textbox(slide, Inches(4.1), Inches(top + 0.27), Inches(8.4), Inches(0.5),
                    desc, font_size=18, color=INK)
        top += 1.13


@S('Ⅶ. 구조')
def s_flow(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅶ. 구조', page, total)
    add_title(slide, '편별 흐름 — 사상의 전개 순서')
    flow = [
        ('양혜왕', '정치의 출발 — 인정(仁政)·왕도 선언'),
        ('공손추', '내면의 힘 — 호연지기·사단'),
        ('등문공', '사회 제도 — 정전제·오륜'),
        ('이루',   '인의(仁義)의 실천 — 일상의 도덕'),
        ('만장',   '역사적 검증 — 성왕의 행적과 경전 해석'),
        ('고자',   '인성론 논변 — 성선설의 이론적 완성'),
        ('진심',   '수양론 완성 — 사상의 집대성'),
    ]
    top = 2.3
    for i, (name, desc) in enumerate(flow):
        add_filled_rect(slide, Inches(0.7), Inches(top), Inches(2.5), Inches(0.6),
                        PALE)
        add_textbox(slide, Inches(0.7), Inches(top + 0.1), Inches(2.5), Inches(0.4),
                    name, font_size=18, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(3.5), Inches(top + 0.1), Inches(9.5), Inches(0.4),
                    desc, font_size=16, color=INK)
        if i < len(flow) - 1:
            add_textbox(slide, Inches(1.7), Inches(top + 0.55), Inches(0.5), Inches(0.3),
                        '▼', font_size=10, color=SUB, align=PP_ALIGN.CENTER)
        top += 0.65


# ---------- Ⅷ. 현대적 의의 ----------
def make_modern_slide(title, kor_subtitle, lines):
    def renderer(slide, page, total):
        set_white_background(slide)
        add_page_header(slide, 'Ⅷ. 현대적 의의', page, total)
        add_title(slide, title, kor_subtitle)
        add_paragraphs(slide, Inches(0.9), Inches(2.4), Inches(11.5), Inches(4.5),
                       lines, line_spacing=1.5, font_size=18)
    return renderer


SLIDES.append((make_modern_slide(
    '현대 ① — 인권과 민주주의',
    '맹자의 민본 사상이 현대 정치 이념으로',
    [
        ('민위귀(民爲貴)', {'font_size': 24, 'bold': True, 'color': ACCENT}),
        ('국민주권·민주주의의 선구적 이념', {'font_size': 18, 'space_before': 4}),
        ('', {'font_size': 8}),
        ('역성혁명론', {'font_size': 24, 'bold': True, 'color': ACCENT, 'space_before': 14}),
        ('저항권(right of resistance)의 동아시아적 기원',
         {'font_size': 18, 'space_before': 4}),
        ('폭정에 대한 정당한 저항의 사상적 근거를 제공',
         {'font_size': 15, 'color': SUB}),
    ]), 'Ⅷ. 현대적 의의'))

SLIDES.append((make_modern_slide(
    '현대 ② — 교육과 인격 형성',
    '성선설의 현대 교육학적 함의',
    [
        ('성선설(性善說)', {'font_size': 24, 'bold': True, 'color': ACCENT}),
        ('긍정 심리학·인본주의 교육과 직결', {'font_size': 18, 'space_before': 4}),
        ('아이는 본래 선한 가능성을 지니고 태어난다',
         {'font_size': 16, 'color': SUB}),
        ('', {'font_size': 8}),
        ('조장(助長) 경계', {'font_size': 24, 'bold': True, 'color': ACCENT, 'space_before': 14}),
        ('과도한 교육 압박의 위험성에 대한 경고', {'font_size': 18, 'space_before': 4}),
        ('자녀를 위한다며 도리어 망친다는 통찰',
         {'font_size': 16, 'color': SUB}),
    ]), 'Ⅷ. 현대적 의의'))

SLIDES.append((make_modern_slide(
    '현대 ③ — 리더십과 경영',
    '왕도·항산 — 도덕 경영의 두 축',
    [
        ('왕도 vs 패도', {'font_size': 24, 'bold': True, 'color': ACCENT}),
        ('덕에 의한 리더십 vs 힘에 의한 리더십', {'font_size': 18, 'space_before': 4}),
        ('지속 가능한 조직은 결국 신뢰에서 나온다',
         {'font_size': 16, 'color': SUB}),
        ('', {'font_size': 8}),
        ('항산항심', {'font_size': 24, 'bold': True, 'color': ACCENT, 'space_before': 14}),
        ('경제적 안정이 사회·조직 안정의 기반', {'font_size': 18, 'space_before': 4}),
        ('생계가 보장되어야 책임감과 윤리가 작동한다',
         {'font_size': 16, 'color': SUB}),
    ]), 'Ⅷ. 현대적 의의'))

SLIDES.append((make_modern_slide(
    '현대 ④ — 윤리와 가치관',
    '의(義)와 자기효능감 — 도덕적 주체성',
    [
        ('의리론(義利之辨)', {'font_size': 24, 'bold': True, 'color': ACCENT}),
        ('이익보다 원칙을 우선하는 기업 윤리', {'font_size': 18, 'space_before': 4}),
        ('단기 수익이 아닌 사회적 책임의 기준',
         {'font_size': 16, 'color': SUB}),
        ('', {'font_size': 8}),
        ('자포자기(自暴自棄) 경계', {'font_size': 24, 'bold': True, 'color': ACCENT, 'space_before': 14}),
        ('자기효능감(Self-efficacy)·성장 마인드셋과 통함', {'font_size': 18, 'space_before': 4}),
        ('스스로를 포기하지 않는 한 가능성은 열려 있다',
         {'font_size': 16, 'color': SUB}),
    ]), 'Ⅷ. 현대적 의의'))


# ---------- Ⅸ. 비교 ----------
@S('Ⅸ. 비교')
def s_compare(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅸ. 비교', page, total)
    add_title(slide, '공자 · 맹자 · 순자 — 유가 삼대 사상가 비교')
    rows = [
        ('핵심 덕목',  '인(仁)',           '인의(仁義)',         '예(禮)'),
        ('인성론',     '직접 언급 적음',   '성선설(性善說) 체계화', '성악설(性惡說)'),
        ('정치사상',   '덕치(德治)',       '왕도·민본·혁명론',   '왕도·예법 병용'),
        ('수양론',     '극기복례',         '존심양성·구방심',    '화성기위'),
        ('문체',       '간결한 어록',      '웅변적 논변',        '체계적 논설문'),
        ('대화 상대',  '주로 제자들',      '제후·학자',           '직하학궁 학자들'),
        ('존칭',       '지성(至聖)',       '아성(亞聖)',          '후성(後聖)'),
    ]
    top = 1.95
    # 헤더
    add_filled_rect(slide, Inches(0.5), Inches(top), Inches(2.2), Inches(0.55), SUB)
    add_textbox(slide, Inches(0.5), Inches(top + 0.1), Inches(2.2), Inches(0.4),
                '항목', font_size=15, bold=True,
                color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
    for i, name in enumerate(['공자(논어)', '맹자', '순자']):
        cx = 0.5 + 2.2 + i * 3.45
        color = ACCENT if name == '맹자' else SUB
        add_filled_rect(slide, Inches(cx), Inches(top), Inches(3.45), Inches(0.55), color)
        add_textbox(slide, Inches(cx), Inches(top + 0.1), Inches(3.45), Inches(0.4),
                    name, font_size=16, bold=True,
                    color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
    row_h = 0.62
    for r_idx, row in enumerate(rows):
        y = top + 0.55 + r_idx * row_h
        bg = RGBColor(0xFA, 0xFA, 0xFA) if r_idx % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF)
        add_filled_rect(slide, Inches(0.5), Inches(y), Inches(2.2), Inches(row_h), PALE)
        add_textbox(slide, Inches(0.55), Inches(y + 0.12), Inches(2.1), Inches(0.4),
                    row[0], font_size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)
        for i in range(3):
            cx = 0.5 + 2.2 + i * 3.45
            add_filled_rect(slide, Inches(cx), Inches(y), Inches(3.45), Inches(row_h), bg)
            text_color = ACCENT if i == 1 else INK
            add_textbox(slide, Inches(cx + 0.05), Inches(y + 0.12),
                        Inches(3.35), Inches(0.4),
                        row[i + 1], font_size=14, color=text_color,
                        bold=(i == 1), align=PP_ALIGN.CENTER)


# ---------- Ⅹ. 마무리 ----------
@S('Ⅹ. 마무리')
def s_summary(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅹ. 마무리', page, total)
    add_title(slide, '한 문장으로 정리하는 맹자')
    add_filled_rect(slide, Inches(0.7), Inches(2.3), Inches(11.9), Inches(4.5), PALE)
    add_paragraphs(slide, Inches(1.1), Inches(2.6), Inches(11.1), Inches(4.1), [
        ('인간은 본래 선한 본성을 지니고 있으며(性善)',
         {'font_size': 22, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 8}),
        ('이 네 가지 싹(四端)을',
         {'font_size': 22, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER, 'space_before': 8}),
        ('기르고 확충(擴充)하는 것이 수양이며',
         {'font_size': 22, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 8}),
        ('덕(德)으로 다스리는 왕도(王道)와',
         {'font_size': 22, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER, 'space_before': 8}),
        ('백성이 가장 귀하다는 민본(民本)을 정치 원칙으로 삼아',
         {'font_size': 22, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 8}),
        ('의로움(義)을 이익(利)보다 앞세우는',
         {'font_size': 22, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 8}),
        ('— 도덕적 용기의 철학 —',
         {'font_size': 22, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.2)


@S('Ⅹ. 마무리')
def s_closing(slide, page, total):
    set_white_background(slide)
    add_filled_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.4), RULE)
    add_textbox(slide, Inches(0.5), Inches(2.0), Inches(12.3), Inches(1.5),
                '民 爲 貴   社 稷 次 之   君 爲 輕',
                font_size=52, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.6), Inches(12.3), Inches(0.5),
                '민위귀 사직차지 군위경', font_size=20, color=SUB,
                align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.4), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(4.7), Inches(12.3), Inches(0.6),
                '"백성이 가장 귀하다"',
                font_size=26, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.4), Inches(12.3), Inches(0.5),
                '— 2,300년 전 맹자가 남긴 정치 철학의 마지막 말',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
                '감사합니다', font_size=28, bold=True, color=INK,
                align=PP_ALIGN.CENTER)
    add_filled_rect(slide, Inches(0), Inches(7.1), Inches(13.333), Inches(0.4), RULE)


# ---------- 빌드 ----------
total_pages = len(SLIDES)
for i, (renderer, section) in enumerate(SLIDES, 1):
    slide = prs.slides.add_slide(blank)
    renderer(slide, i, total_pages)

out_path = r'C:\Users\박호군\ClaudeProjects\Oriental-Classics2\발표자료\맹자_발표자료.pptx'
prs.save(out_path)
print(f'Saved: {out_path}  ({total_pages} slides)')
