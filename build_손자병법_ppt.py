# -*- coding: utf-8 -*-
"""
손자병법 발표자료 재작성 스크립트
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


def set_white_background(slide):
    fill = slide.background.fill
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
    add_textbox(slide, Inches(0.5), Inches(1.6), Inches(12.3), Inches(1.8),
                '孫子兵法', font_size=96, bold=True, color=INK,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.6), Inches(12.3), Inches(0.5),
                'The Art of War · 손자병법', font_size=24, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.4), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(4.7), Inches(12.3), Inches(0.5),
                '손무(孫武)의 13편 — 동양 최고(最古)의 전략 고전',
                font_size=20, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.2), Inches(12.3), Inches(0.4),
                '춘추시대 말기 (BC 6~5세기) · 13편 약 6,000자',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.8), Inches(12.3), Inches(0.4),
                '부전이굴인지병(不戰而屈人之兵) — 싸우지 않고 이기는 것이 최선',
                font_size=14, color=ACCENT, bold=True, align=PP_ALIGN.CENTER)
    add_filled_rect(slide, Inches(0), Inches(7.1), Inches(13.333), Inches(0.4), RULE)


# ---------- 2. 목차 ----------
@S('목차')
def s_toc(slide, page, total):
    set_white_background(slide)
    add_textbox(slide, Inches(0.5), Inches(0.5), Inches(12.8), Inches(0.7),
                '목 차', font_size=36, bold=True, color=INK)
    add_rule(slide, Inches(0.5), Inches(1.3), Inches(12.8))
    items = [
        ('Ⅰ', '개요 — 손자병법이란 무엇인가'),
        ('Ⅱ', '13편의 구성'),
        ('Ⅲ', '핵심 사상 ① — 부전승(不戰勝)과 오사(五事)'),
        ('Ⅳ', '핵심 사상 ② — 허실·기정·세'),
        ('Ⅴ', '핵심 사상 ③ — 지피지기와 변화의 철학'),
        ('Ⅵ', '명구절 10선'),
        ('Ⅶ', '손자병법의 구조적 특징'),
        ('Ⅷ', '현대적 의의'),
        ('Ⅸ', '다른 사상과의 비교'),
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
    add_title(slide, '손자병법(孫子兵法)이란 무엇인가')
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(1.4), PALE)
    add_textbox(slide, Inches(0.9), Inches(2.55), Inches(11.5), Inches(0.5),
                '동양에서 가장 오래되고 가장 영향력 있는 군사 전략서',
                font_size=24, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.9), Inches(3.15), Inches(11.5), Inches(0.5),
                '전쟁 기술서를 넘어서, 전략적 사고의 교과서로 평가받는 고전',
                font_size=16, color=ACCENT, align=PP_ALIGN.CENTER)
    nums = [('13', '편(篇)'), ('약 6,000', '자(字)'), ('2,500+', '년의 영향')]
    for i, (n, lbl) in enumerate(nums):
        x = 1.8 + i * 3.5
        add_textbox(slide, Inches(x), Inches(4.4), Inches(3.0), Inches(1.0),
                    n, font_size=58, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(x), Inches(5.6), Inches(3.0), Inches(0.5),
                    lbl, font_size=17, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.7), Inches(12.3), Inches(0.4),
                '극도로 간결한 문장에 압축된 보편 전략 — 군사·외교·경영·일상까지',
                font_size=14, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅰ. 개요')
def s_sunmu(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅰ. 개요', page, total)
    add_title(slide, '손무(孫武)', '— 손자병법의 저자, 춘추시대의 군사 전략가')
    lines = [
        ('출신', {'bold': True, 'font_size': 18, 'color': ACCENT}),
        ('  제(齊)나라 출신, 후에 오(吳)나라로 망명',
         {'font_size': 18}),
        ('등용', {'bold': True, 'font_size': 18, 'color': ACCENT, 'space_before': 12}),
        ('  오왕(吳王) 합려(闔閭)에게 병법 13편을 올려 장군으로 임명',
         {'font_size': 18}),
        ('전공', {'bold': True, 'font_size': 18, 'color': ACCENT, 'space_before': 12}),
        ('  초(楚)나라 수도 영(郢) 함락 — 약소국 오나라를 패자(覇者) 반열에 올림',
         {'font_size': 18}),
        ('후대 평가', {'bold': True, 'font_size': 18, 'color': ACCENT, 'space_before': 12}),
        ('  "병성(兵聖)" — 병가(兵家)의 성인',
         {'font_size': 18}),
    ]
    add_paragraphs(slide, Inches(0.9), Inches(2.3), Inches(11.5), Inches(4.5),
                   lines, line_spacing=1.4)


@S('Ⅰ. 개요')
def s_episode(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅰ. 개요', page, total)
    add_title(slide, '오궁(吳宮)의 궁녀 훈련 — 손무를 증명한 일화')
    add_filled_rect(slide, Inches(0.6), Inches(2.3), Inches(11.9), Inches(4.0), PALE)
    add_paragraphs(slide, Inches(0.9), Inches(2.5), Inches(11.3), Inches(3.7), [
        ('오왕 합려가 손무를 시험했다.', {'font_size': 17, 'bold': True, 'align': PP_ALIGN.CENTER}),
        ('"이 궁녀들도 훈련시킬 수 있겠소?"', {'font_size': 17, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
        ('', {'font_size': 4}),
        ('손무는 궁녀 180명을 두 부대로 나누고 합려의 총희(寵姬) 두 명을 대장으로 세웠다.',
         {'font_size': 16, 'align': PP_ALIGN.CENTER, 'space_before': 6, 'color': SUB}),
        ('호령을 내리자 궁녀들은 웃기만 했다.', {'font_size': 16, 'align': PP_ALIGN.CENTER, 'color': SUB}),
        ('"법령이 분명하지 않은 것은 장수의 죄"라며 다시 설명하고 명령을 내렸다.',
         {'font_size': 16, 'align': PP_ALIGN.CENTER, 'color': SUB}),
        ('그래도 웃기만 하자 손무는 두 대장을 처형하려 했다.',
         {'font_size': 16, 'align': PP_ALIGN.CENTER, 'color': SUB}),
        ('', {'font_size': 4}),
        ('합려가 만류했지만 손무는 답했다.', {'font_size': 16, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('"장군이 출정하면 임금의 명령도 받지 않는 바가 있습니다(將在軍 君命有所不受)"',
         {'font_size': 17, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('두 총희를 베자 — 궁녀들은 더 이상 웃지 않았다.',
         {'font_size': 16, 'align': PP_ALIGN.CENTER, 'space_before': 6, 'color': SUB}),
    ], line_spacing=1.25)
    add_textbox(slide, Inches(0.5), Inches(6.6), Inches(12.3), Inches(0.4),
                '— 사기(史記) 손자오기열전 · 군기(軍紀)와 장수의 권한을 보여주는 고전 일화',
                font_size=13, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅰ. 개요')
def s_pansa(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅰ. 개요', page, total)
    add_title(slide, '판본과 영향사', '— 2,500년에 걸친 전파의 궤적')
    items = [
        ('AD 3세기',  '조조(曹操)의 주석', '최초의 본격적 주석서를 편찬 — 후대 판본의 기초'),
        ('AD 7세기',  '당송 11가 주석',     '두목·이전 등 11명의 주석가가 가세 — 정본 형성'),
        ('1972년',    '은작산(銀雀山) 죽간 출토', '한묘에서 죽간본 발견 — 현대 손자병법 연구의 전환점'),
        ('18세기',    '유럽 전파',          '예수회 선교사들이 라틴어·프랑스어로 번역 → 서구 진출'),
        ('20세기',    '서구 군사학교 필독서', 'West Point·MBA·실리콘밸리까지 영향력 확장'),
    ]
    top = 2.4
    for era, title, desc in items:
        add_filled_rect(slide, Inches(0.7), Inches(top), Inches(2.3), Inches(0.85), ACCENT)
        add_textbox(slide, Inches(0.7), Inches(top + 0.22), Inches(2.3), Inches(0.5),
                    era, font_size=16, bold=True,
                    color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(3.2), Inches(top), Inches(9.5), Inches(0.85), PALE)
        add_textbox(slide, Inches(3.4), Inches(top + 0.08), Inches(9.1), Inches(0.4),
                    title, font_size=17, bold=True, color=INK)
        add_textbox(slide, Inches(3.4), Inches(top + 0.45), Inches(9.1), Inches(0.4),
                    desc, font_size=13, color=SUB)
        top += 0.95


# ---------- Ⅱ. 구성 ----------
@S('Ⅱ. 구성')
def s_structure_overview(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅱ. 구성', page, total)
    add_title(slide, '13편의 구성 개관 — 세 층위의 논리 전개')
    # 3개 박스
    layers = [
        ('전략론',     '1~6편',  '전쟁의 본질과\n근본 원칙', ACCENT),
        ('전술론',     '7~11편', '실전의 운용과\n지형 활용', INK),
        ('특수전법',   '12~13편', '화공·정보전 등\n특수한 기법', SUB),
    ]
    for i, (name, scope, desc, color) in enumerate(layers):
        x = 0.5 + i * 4.3
        add_filled_rect(slide, Inches(x), Inches(2.5), Inches(4.1), Inches(4.5), PALE)
        add_filled_rect(slide, Inches(x), Inches(2.5), Inches(4.1), Inches(0.8), color)
        add_textbox(slide, Inches(x), Inches(2.65), Inches(4.1), Inches(0.6),
                    name, font_size=24, bold=True,
                    color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(x), Inches(3.6), Inches(4.1), Inches(0.5),
                    scope, font_size=18, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(x), Inches(4.5), Inches(4.1), Inches(2.0),
                    desc, font_size=17, color=INK, align=PP_ALIGN.CENTER,
                    anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, Inches(0.5), Inches(7.05), Inches(12.3), Inches(0.4),
                '총론(전략) → 원칙(전술) → 특수(정보·화공)로 이어지는 체계적 전개',
                font_size=14, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅱ. 구성')
def s_strategy(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅱ. 구성', page, total)
    add_title(slide, '전략론 — 1~6편')
    rows = [
        ('1', '시계(始計)', '전쟁의 기본 계산 — 오사칠계(五事七計)'),
        ('2', '작전(作戰)', '전쟁의 비용 — 속전속결(速戰速決)의 원칙'),
        ('3', '모공(謀攻)', '부전승(不戰勝) — 지피지기(知彼知己)'),
        ('4', '군형(軍形)', '선승이후구전(先勝而後求戰) — 형(形)의 확립'),
        ('5', '병세(兵勢)', '기정(奇正)의 운용 · 세(勢)의 위력'),
        ('6', '허실(虛實)', '주도권 장악 — 적의 빈 곳을 친다, 물의 비유'),
    ]
    top = 2.3
    for num, name, desc in rows:
        add_textbox(slide, Inches(0.7), Inches(top), Inches(0.7), Inches(0.5),
                    num, font_size=22, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(1.6), Inches(top), Inches(3.0), Inches(0.5),
                    name, font_size=20, bold=True, color=INK)
        add_textbox(slide, Inches(4.7), Inches(top + 0.05), Inches(8.2), Inches(0.5),
                    desc, font_size=16, color=SUB)
        top += 0.7


@S('Ⅱ. 구성')
def s_tactics(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅱ. 구성', page, total)
    add_title(slide, '전술론·특수전법 — 7~13편')
    rows = [
        ('7',  '군쟁(軍爭)',  '기동과 주도권 — 풍림화산(風林火山)'),
        ('8',  '구변(九變)',  '아홉 가지 변통 · 장수의 다섯 위험(五危)'),
        ('9',  '행군(行軍)',  '지형별 행군과 적의 동태 관찰'),
        ('10', '지형(地形)',  '여섯 지형과 군대의 여섯 패망'),
        ('11', '구지(九地)',  '아홉 지형 · 사지즉전(死地則戰)'),
        ('12', '화공(火攻)',  '화공의 다섯 기법 · 분노로 군대를 일으키지 말 것'),
        ('13', '용간(用間)',  '다섯 종류의 간첩(五間) · 선지(先知)의 중요성'),
    ]
    top = 2.3
    for num, name, desc in rows:
        add_textbox(slide, Inches(0.7), Inches(top), Inches(0.7), Inches(0.4),
                    num, font_size=20, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(1.6), Inches(top), Inches(3.0), Inches(0.4),
                    name, font_size=18, bold=True, color=INK)
        add_textbox(slide, Inches(4.7), Inches(top + 0.03), Inches(8.2), Inches(0.4),
                    desc, font_size=15, color=SUB)
        top += 0.62


# ---------- Ⅲ. 부전승과 오사 ----------
@S('Ⅲ. 부전승·오사')
def s_bujunseung(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 부전승·오사', page, total)
    add_title(slide, '부전승(不戰勝) — 백전백승은 최선이 아니다')
    add_textbox(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(0.8),
                '百 戰 百 勝   非 善 之 善 者 也',
                font_size=30, bold=True, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.2), Inches(12.3), Inches(0.8),
                '不 戰 而 屈 人 之 兵   善 之 善 者 也',
                font_size=32, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(4.1), Inches(12.3), Inches(0.5),
                '백전백승 비선지선자야 · 부전이굴인지병 선지선자야',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(4.6), Inches(12.3), Inches(0.4),
                '— 모공편 제3', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(5.2), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(5.5), Inches(12.3), Inches(0.6),
                '백번 싸워 백번 이기는 것은 최선이 아니다',
                font_size=18, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.1), Inches(12.3), Inches(0.6),
                '싸우지 않고 적의 군대를 굴복시키는 것이 최선의 최선이다',
                font_size=20, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.8), Inches(12.3), Inches(0.4),
                '— 전쟁을 미화하지 않은, 가장 신중한 전략의 고전',
                font_size=14, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅲ. 부전승·오사')
def s_4_grades(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 부전승·오사', page, total)
    add_title(slide, '승리의 4등급 — 어떤 승리를 추구할 것인가')
    grades = [
        ('최상',  '伐謀', '벌모', '적의 전략(謀)을 무력화 — 미연에 막는다', ACCENT),
        ('차선',  '伐交', '벌교', '적의 동맹(交)을 와해 — 고립시킨다', RGBColor(0xB5, 0x4B, 0x4B)),
        ('하책',  '伐兵', '벌병', '적의 군대(兵)를 공격 — 정면 충돌', RGBColor(0x70, 0x70, 0x70)),
        ('최하책','攻城', '공성', '적의 성(城)을 공격 — 최대 손실', SUB),
    ]
    top = 2.3
    for rank, han, eum, desc, color in grades:
        add_filled_rect(slide, Inches(0.7), Inches(top), Inches(1.6), Inches(1.0), color)
        add_textbox(slide, Inches(0.7), Inches(top + 0.28), Inches(1.6), Inches(0.5),
                    rank, font_size=18, bold=True,
                    color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(2.5), Inches(top), Inches(2.0), Inches(1.0), PALE)
        add_textbox(slide, Inches(2.5), Inches(top + 0.1), Inches(2.0), Inches(0.5),
                    han, font_size=24, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(2.5), Inches(top + 0.6), Inches(2.0), Inches(0.4),
                    eum, font_size=13, color=SUB, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(4.7), Inches(top + 0.3), Inches(8.2), Inches(0.5),
                    desc, font_size=16, color=INK)
        top += 1.15


@S('Ⅲ. 부전승·오사')
def s_5_factors(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 부전승·오사', page, total)
    add_title(slide, '오사(五事) — 전쟁을 결정짓는 다섯 근본 요소', '— 시계편 제1')
    factors = [
        ('道', '도', '군주와 백성의 일체감 — 대의명분',  '명분과 비전'),
        ('天', '천', '기후·계절·타이밍',                  '외부 환경'),
        ('地', '지', '지형·거리·경쟁 환경',                '시장·지형'),
        ('將', '장', '지(智)·신(信)·인(仁)·용(勇)·엄(嚴)','리더십'),
        ('法', '법', '편제·규율·보급 — 조직 시스템',      '체계와 시스템'),
    ]
    top = 2.2
    for han, eum, desc, modern in factors:
        add_filled_rect(slide, Inches(0.7), Inches(top), Inches(1.0), Inches(0.85), ACCENT)
        add_textbox(slide, Inches(0.7), Inches(top + 0.1), Inches(1.0), Inches(0.6),
                    han, font_size=32, bold=True,
                    color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(1.9), Inches(top), Inches(1.0), Inches(0.85), PALE)
        add_textbox(slide, Inches(1.9), Inches(top + 0.22), Inches(1.0), Inches(0.5),
                    eum, font_size=18, bold=True, color=INK, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(3.1), Inches(top + 0.18), Inches(6.5), Inches(0.5),
                    desc, font_size=15, color=INK)
        add_filled_rect(slide, Inches(9.8), Inches(top), Inches(2.9), Inches(0.85), RGBColor(0xFA, 0xFA, 0xFA))
        add_textbox(slide, Inches(9.8), Inches(top + 0.22), Inches(2.9), Inches(0.5),
                    f'현대: {modern}', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
        top += 0.97


# ---------- Ⅳ. 허실·기정·세 ----------
@S('Ⅳ. 허실·기정·세')
def s_heoshil(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅳ. 허실·기정·세', page, total)
    add_title(slide, '허실(虛實) — 빈 곳을 치고, 내 빈 곳을 감추라')
    add_textbox(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(0.9),
                '避  實  而  擊  虛',
                font_size=44, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.3), Inches(12.3), Inches(0.5),
                '피실이격허 — 실(實)한 곳을 피하고 허(虛)한 곳을 친다',
                font_size=18, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.9), Inches(12.3), Inches(0.4),
                '— 허실편 제6', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.5), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(4.8), Inches(12.3), Inches(0.9),
                '致 人 而 不 致 於 人',
                font_size=38, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.85), Inches(12.3), Inches(0.5),
                '치인이부치어인 — 적을 끌어오지, 끌려가지 않는다',
                font_size=18, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.6), Inches(12.3), Inches(0.4),
                '핵심은 주도권(initiative)의 장악',
                font_size=15, bold=True, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅳ. 허실·기정·세')
def s_gijeong(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅳ. 허실·기정·세', page, total)
    add_title(slide, '기정(奇正) — 정공(正)과 기습(奇)의 조화')
    add_textbox(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(0.9),
                '以 正 合   以 奇 勝',
                font_size=44, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.3), Inches(12.3), Inches(0.5),
                '이정합 이기승 — 정(正)으로 맞서고, 기(奇)로 이긴다',
                font_size=17, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.85), Inches(12.3), Inches(0.4),
                '— 병세편 제5', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    # 좌우 정 vs 기
    add_filled_rect(slide, Inches(0.7), Inches(4.5), Inches(5.9), Inches(2.4), PALE)
    add_textbox(slide, Inches(0.7), Inches(4.65), Inches(5.9), Inches(0.5),
                '正 — 정공(正攻)', font_size=22, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.9), Inches(5.3), Inches(5.5), Inches(1.5),
                '정규적·예측 가능한 본대\n적을 묶어두는 주력',
                font_size=15, color=INK, align=PP_ALIGN.CENTER)
    add_filled_rect(slide, Inches(6.8), Inches(4.5), Inches(5.9), Inches(2.4), PALE)
    add_textbox(slide, Inches(6.8), Inches(4.65), Inches(5.9), Inches(0.5),
                '奇 — 기습(奇襲)', font_size=22, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(7.0), Inches(5.3), Inches(5.5), Inches(1.5),
                '비정규적·기상천외한 별동대\n결정적 승리를 가르는 카드',
                font_size=15, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.4),
                '奇正相生 — 정·기는 끝없이 순환하며 무한한 변화를 낳는다',
                font_size=15, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)


@S('Ⅳ. 허실·기정·세')
def s_se(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅳ. 허실·기정·세', page, total)
    add_title(slide, '세(勢) — 형세(形勢)를 만들어 압도하라')
    add_textbox(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(0.9),
                '激 水 之 疾   至 於 漂 石 者   勢 也',
                font_size=30, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.2), Inches(12.3), Inches(0.5),
                '격수지질 지어표석자 세야',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.7), Inches(12.3), Inches(0.4),
                '— 병세편 제5', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.3), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(0.6),
                '"급류가 빠르게 흘러 돌을 떠내려 보내는 것 — 그것이 세(勢)다"',
                font_size=18, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_filled_rect(slide, Inches(0.7), Inches(5.4), Inches(11.9), Inches(1.5), PALE)
    add_paragraphs(slide, Inches(0.9), Inches(5.55), Inches(11.5), Inches(1.3), [
        ('求之於勢   不責於人', {'font_size': 22, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('세(勢)에서 승리를 구하지, 개인에게 책임을 돌리지 않는다',
         {'font_size': 16, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
        ('좋은 리더는 시스템을 설계하여 평범한 사람도 탁월한 성과를 내게 한다',
         {'font_size': 14, 'color': SUB, 'align': PP_ALIGN.CENTER, 'space_before': 4}),
    ], line_spacing=1.3)


# ---------- Ⅴ. 지피지기와 변화 ----------
@S('Ⅴ. 지·변화')
def s_jipijigi(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 지·변화', page, total)
    add_title(slide, '지피지기(知彼知己) — 지(知)의 세 수준')
    levels = [
        ('知 彼 知 己',  '지피지기',    '百 戰 不 殆', '백전불태', '백 번 싸워도 위태롭지 않다', ACCENT),
        ('不知彼而知己', '부지피이지기', '一 勝 一 負', '일승일부', '한 번 이기고 한 번 진다',     RGBColor(0x90, 0x60, 0x60)),
        ('不知彼不知己', '부지피부지기', '每 戰 必 殆', '매전필태', '매번 싸울 때마다 위태롭다', SUB),
    ]
    top = 2.4
    for han, eum, han2, eum2, mean, color in levels:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(4.2), Inches(1.2), PALE)
        add_textbox(slide, Inches(0.6), Inches(top + 0.15), Inches(4.2), Inches(0.5),
                    han, font_size=20, bold=True, color=color, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(0.6), Inches(top + 0.72), Inches(4.2), Inches(0.4),
                    eum, font_size=13, color=SUB, align=PP_ALIGN.CENTER)
        # 화살표
        add_textbox(slide, Inches(4.8), Inches(top + 0.4), Inches(0.5), Inches(0.5),
                    '▶', font_size=18, color=color, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(5.4), Inches(top), Inches(3.2), Inches(1.2), color)
        add_textbox(slide, Inches(5.4), Inches(top + 0.15), Inches(3.2), Inches(0.5),
                    han2, font_size=20, bold=True,
                    color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(5.4), Inches(top + 0.72), Inches(3.2), Inches(0.4),
                    eum2, font_size=13,
                    color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(8.8), Inches(top + 0.4), Inches(4.3), Inches(0.5),
                    mean, font_size=15, color=INK)
        top += 1.4
    add_textbox(slide, Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.4),
                '※ "백전백승"이 아니라 "백전불태" — 화려한 승리보다 안전을 추구하는 사상',
                font_size=14, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)


@S('Ⅴ. 지·변화')
def s_water(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 지·변화', page, total)
    add_title(slide, '병형상수(兵形象水) — 군대는 물과 같다')
    add_textbox(slide, Inches(0.5), Inches(2.2), Inches(12.3), Inches(0.9),
                '兵   形   象   水',
                font_size=64, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.3), Inches(12.3), Inches(0.5),
                '병형상수 — "군대의 형태는 물과 같다"',
                font_size=17, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.85), Inches(12.3), Inches(0.4),
                '— 허실편 제6', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.5), Inches(2.3), color=RULE, weight=1.5)
    items = [
        ('높은 곳을 피하고 낮은 곳으로',  '강한 곳 피하고 약한 곳 공격'),
        ('지형에 따라 형태가 변한다',    '상황에 따라 전략을 바꾼다'),
        ('일정한 형태가 없다(無常形)',    '병법에 일정한 형세는 없다(兵無常勢)'),
    ]
    top = 4.75
    for left, right in items:
        add_textbox(slide, Inches(0.7), Inches(top), Inches(5.8), Inches(0.4),
                    f'• {left}', font_size=15, color=ACCENT, bold=True, align=PP_ALIGN.RIGHT)
        add_textbox(slide, Inches(6.8), Inches(top), Inches(5.8), Inches(0.4),
                    f'→  {right}', font_size=15, color=INK)
        top += 0.55


@S('Ⅴ. 지·변화')
def s_seonseung(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 지·변화', page, total)
    add_title(slide, '선승이후구전(先勝而後求戰) — 준비의 철학')
    add_textbox(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(0.9),
                '勝 兵 先 勝 而 後 求 戰',
                font_size=36, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.2), Inches(12.3), Inches(0.9),
                '敗 兵 先 戰 而 後 求 勝',
                font_size=30, bold=True, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(4.15), Inches(12.3), Inches(0.5),
                '승병선승이후구전 · 패병선전이후구승',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(4.65), Inches(12.3), Inches(0.4),
                '— 군형편 제4', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(5.25), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(5.55), Inches(12.3), Inches(0.6),
                '이기는 군대는 먼저 이겨놓고 싸우고,',
                font_size=20, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.2), Inches(12.3), Inches(0.6),
                '지는 군대는 먼저 싸운 뒤에 이기려 한다',
                font_size=20, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.9), Inches(12.3), Inches(0.4),
                '승리는 전장에서 결정되는 것이 아니라, 전장에 들어서기 전에 이미 결정된다',
                font_size=13, color=SUB, align=PP_ALIGN.CENTER)


# ---------- Ⅵ. 명구절 ----------
def make_quote_slide(section, hanmun, eum, mean, ref, *, hanmun_size=46):
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
                    mean, font_size=22, bold=True, color=INK,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, Inches(0.5), Inches(6.5), Inches(12.3), Inches(0.4),
                    f'— {ref}', font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    return renderer


SLIDES.append((make_quote_slide('Ⅵ. 명구절 (1/10)',
    '兵者 國之大事\n死生之地 存亡之道',
    '병자 국지대사 · 사생지지 존망지도',
    '전쟁은 나라의 중대사이니, 백성의 생사와 나라의 존망이 달린 길이라\n살피지 않을 수 없다',
    '시계편 제1', hanmun_size=34), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (2/10)',
    '不戰而屈人之兵\n善之善者也',
    '부전이굴인지병 · 선지선자야',
    '싸우지 않고 적의 군대를 굴복시키는 것이 최선의 최선이다',
    '모공편 제3', hanmun_size=36), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (3/10)',
    '知 彼 知 己   百 戰 不 殆',
    '지피지기 백전불태',
    '적을 알고 나를 알면 백 번 싸워도 위태롭지 않다',
    '모공편 제3', hanmun_size=42), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (4/10)',
    '勝 兵 先 勝 而 後 求 戰',
    '승병선승이후구전',
    '이기는 군대는 먼저 이겨놓고 싸운다',
    '군형편 제4', hanmun_size=40), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (5/10)',
    '凡 戰 者   以 正 合   以 奇 勝',
    '범전자 이정합 이기승',
    '무릇 전투란 정(正)으로 맞서고 기(奇)로 이기는 것이다',
    '병세편 제5', hanmun_size=34), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (6/10)',
    '故 善 戰 者\n致 人 而 不 致 於 人',
    '고선전자 · 치인이부치어인',
    '잘 싸우는 자는 적을 끌어오지, 끌려가지 않는다 — 주도권의 원리',
    '허실편 제6', hanmun_size=34), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (7/10)',
    '其疾如風  其徐如林\n侵掠如火  不動如山',
    '기질여풍 기서여림 · 침략여화 부동여산',
    '바람처럼 빠르고, 숲처럼 고요하며,\n불처럼 공격하고, 산처럼 움직이지 않는다 — 풍림화산(風林火山)',
    '군쟁편 제7', hanmun_size=28), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (8/10)',
    '投之亡地然後存\n陷之死地然後生',
    '투지망지연후존 · 함지사지연후생',
    '망할 곳에 던져야 살고, 죽을 곳에 빠뜨려야 산다 — 사지즉전(死地則戰)',
    '구지편 제11', hanmun_size=34), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (9/10)',
    '主不可以怒而興師\n將不可以慍而致戰',
    '주불가이노이흥사 · 장불가이온이치전',
    '임금은 분노로 군대를 일으켜서는 안 되고,\n장수는 격분하여 싸움을 일으켜서는 안 된다',
    '화공편 제12', hanmun_size=30), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (10/10)',
    '明君賢將  所以動而勝人者\n先 知 也',
    '명군현장 소이동이승인자 · 선지야',
    '명군(明君)과 현장(賢將)이 움직이는 곳마다 승리하는 까닭은\n미리 아는 것(先知) 때문이다 — 정보전의 중요성',
    '용간편 제13', hanmun_size=28), 'Ⅵ. 명구절'))


# ---------- Ⅶ. 구조 ----------
@S('Ⅶ. 구조')
def s_logic_structure(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅶ. 구조', page, total)
    add_title(slide, '13편의 논리 구조 — 체계적 전개')
    blocks = [
        ('총론',     '시계·작전',          '1~2편',  '전쟁 자체의 본질과 비용'),
        ('전략 원칙', '모공·군형·병세·허실', '3~6편',  '부전승·선승·기정·허실의 4대 원칙'),
        ('실전 운용', '군쟁·구변·행군·지형·구지', '7~11편', '기동·변통·지형 활용'),
        ('특수 기법', '화공·용간',          '12~13편', '화공과 정보전'),
    ]
    top = 2.3
    for tag, name, scope, desc in blocks:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.4), Inches(1.0), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.3), Inches(2.4), Inches(0.5),
                    tag, font_size=20, bold=True,
                    color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(3.2), Inches(top), Inches(9.5), Inches(1.0), PALE)
        add_textbox(slide, Inches(3.4), Inches(top + 0.12), Inches(5.5), Inches(0.4),
                    name, font_size=16, bold=True, color=INK)
        add_textbox(slide, Inches(9.5), Inches(top + 0.12), Inches(3.0), Inches(0.4),
                    scope, font_size=14, color=SUB, align=PP_ALIGN.RIGHT)
        add_textbox(slide, Inches(3.4), Inches(top + 0.55), Inches(9.1), Inches(0.4),
                    desc, font_size=14, color=SUB)
        top += 1.15


@S('Ⅶ. 구조')
def s_brevity(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅶ. 구조', page, total)
    add_title(slide, '간결함의 미학 — "한 문장 = 한 원칙"')
    add_textbox(slide, Inches(0.7), Inches(2.4), Inches(5.5), Inches(2.0),
                '약 6,000', font_size=140, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.7), Inches(4.5), Inches(5.5), Inches(0.5),
                '자(字)', font_size=22, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.7), Inches(5.0), Inches(5.5), Inches(0.5),
                '— 손자병법 전체 분량',
                font_size=14, color=SUB, align=PP_ALIGN.CENTER, bold=True)
    add_paragraphs(slide, Inches(7.0), Inches(2.4), Inches(6.0), Inches(4.5), [
        ('극도로 압축된 문장', {'bold': True, 'font_size': 20, 'color': ACCENT}),
        ('논어보다도 짧은 분량 안에', {'font_size': 16}),
        ('완성된 전략 체계를 담아냈다', {'font_size': 16}),
        ('', {'font_size': 6}),
        ('독립적 인용 가능성', {'bold': True, 'font_size': 20, 'color': ACCENT, 'space_before': 10}),
        ('한 문장이 하나의 원칙을 담고 있어',
         {'font_size': 16}),
        ('어디서든 단독으로 인용·활용 가능',
         {'font_size': 16}),
        ('', {'font_size': 6}),
        ('보편성의 비밀', {'bold': True, 'font_size': 20, 'color': ACCENT, 'space_before': 10}),
        ('구체 사례 대신 원리만 남겨',
         {'font_size': 16}),
        ('어느 시대·어느 분야에도 적용 가능',
         {'font_size': 16}),
    ], line_spacing=1.3)


@S('Ⅶ. 구조')
def s_metaphors(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅶ. 구조', page, total)
    add_title(slide, '비유의 힘 — 추상을 이미지로')
    metaphors = [
        ('水',   '물',      '유연한 변화와 적응의 상징'),
        ('激水', '급류',    '돌을 떠내려 보내는 세(勢)의 위력'),
        ('率然', '상산의 뱀', '머리·꼬리·중간이 유기적으로 호응하는 조직'),
        ('風林火山', '풍림화산', '상황별 군대 운용의 네 가지 모습'),
    ]
    top = 2.3
    for han, kor, desc in metaphors:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(2.5), Inches(1.0), ACCENT)
        add_textbox(slide, Inches(0.6), Inches(top + 0.2), Inches(2.5), Inches(0.6),
                    han, font_size=28, bold=True,
                    color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(3.3), Inches(top), Inches(2.5), Inches(1.0), PALE)
        add_textbox(slide, Inches(3.3), Inches(top + 0.32), Inches(2.5), Inches(0.5),
                    kor, font_size=20, bold=True, color=INK, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(6.0), Inches(top), Inches(6.7), Inches(1.0),
                        RGBColor(0xFA, 0xFA, 0xFA))
        add_textbox(slide, Inches(6.2), Inches(top + 0.3), Inches(6.4), Inches(0.5),
                    desc, font_size=16, color=INK)
        top += 1.15


@S('Ⅶ. 구조')
def s_dual(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅶ. 구조', page, total)
    add_title(slide, '이중 구조 — 전쟁론이자 반전론')
    add_filled_rect(slide, Inches(0.6), Inches(2.3), Inches(6.0), Inches(4.5), PALE)
    add_textbox(slide, Inches(0.6), Inches(2.5), Inches(6.0), Inches(0.5),
                '전쟁론의 얼굴', font_size=20, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_paragraphs(slide, Inches(0.9), Inches(3.3), Inches(5.5), Inches(3.3), [
        ('• 13편 전체가 전쟁의 기술서', {'font_size': 17}),
        ('• 정보·기동·기습·심리·화공까지', {'font_size': 17, 'space_before': 8}),
        ('  전쟁의 모든 국면을 다룬다', {'font_size': 16, 'color': SUB}),
        ('• 동서양 군사학의 정전(正典)', {'font_size': 17, 'space_before': 8}),
        ('  — 가장 뛰어난 병서', {'font_size': 14, 'color': ACCENT, 'bold': True}),
    ], line_spacing=1.35)
    add_filled_rect(slide, Inches(6.8), Inches(2.3), Inches(6.0), Inches(4.5), PALE)
    add_textbox(slide, Inches(6.8), Inches(2.5), Inches(6.0), Inches(0.5),
                '반전론의 얼굴', font_size=20, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_paragraphs(slide, Inches(7.1), Inches(3.3), Inches(5.5), Inches(3.3), [
        ('• "싸우지 않고 이기라" — 부전승', {'font_size': 17}),
        ('• "분노로 싸우지 마라" — 화공편 경고',
         {'font_size': 17, 'space_before': 8}),
        ('• "전쟁은 국지대사 — 살피지 않을 수 없다"',
         {'font_size': 16, 'color': SUB, 'space_before': 8}),
        ('• 전쟁의 파괴력에 대한 가장 신중한 경고',
         {'font_size': 14, 'color': ACCENT, 'bold': True, 'space_before': 10}),
    ], line_spacing=1.35)


# ---------- Ⅷ. 현대적 의의 ----------
def make_modern_slide(title, subtitle, lines):
    def renderer(slide, page, total):
        set_white_background(slide)
        add_page_header(slide, 'Ⅷ. 현대적 의의', page, total)
        add_title(slide, title, subtitle)
        add_paragraphs(slide, Inches(0.9), Inches(2.4), Inches(11.5), Inches(4.6),
                       lines, line_spacing=1.5, font_size=18)
    return renderer


SLIDES.append((make_modern_slide(
    '현대 ① — 경영 전략',
    '손자병법이 비즈니스 전략의 고전이 된 까닭',
    [
        ('부전승(不戰勝)', {'font_size': 22, 'bold': True, 'color': ACCENT}),
        ('  → 블루오션 전략 · 경쟁 회피 · 시장 재정의', {'font_size': 16, 'space_before': 4}),
        ('속전속결(速戰速決)', {'font_size': 22, 'bold': True, 'color': ACCENT, 'space_before': 10}),
        ('  → MVP(Minimum Viable Product) · 린 스타트업', {'font_size': 16, 'space_before': 4}),
        ('지피지기(知彼知己)', {'font_size': 22, 'bold': True, 'color': ACCENT, 'space_before': 10}),
        ('  → 시장 조사 · 경쟁사 분석 · SWOT', {'font_size': 16, 'space_before': 4}),
        ('피실격허(避實擊虛)', {'font_size': 22, 'bold': True, 'color': ACCENT, 'space_before': 10}),
        ('  → 니치 마켓 공략 · 선택과 집중', {'font_size': 16, 'space_before': 4}),
    ]), 'Ⅷ. 현대적 의의'))

SLIDES.append((make_modern_slide(
    '현대 ② — 리더십',
    '시스템 설계자로서의 리더',
    [
        ('세(勢)에서 승리를 구하라 — 求之於勢', {'font_size': 22, 'bold': True, 'color': ACCENT}),
        ('  개인의 영웅성이 아니라 시스템 설계가 승부를 가른다',
         {'font_size': 16, 'space_before': 4}),
        ('  → 좋은 리더는 평범한 사람도 탁월한 성과를 내게 한다',
         {'font_size': 15, 'color': SUB}),
        ('', {'font_size': 8}),
        ('장수의 다섯 덕목 (智·信·仁·勇·嚴)', {'font_size': 22, 'bold': True, 'color': ACCENT, 'space_before': 12}),
        ('  지혜·신뢰·인간미·용기·엄정함의 균형',
         {'font_size': 16, 'space_before': 4}),
        ('  → 어느 한쪽으로 치우치면 무너진다',
         {'font_size': 15, 'color': SUB}),
    ]), 'Ⅷ. 현대적 의의'))

SLIDES.append((make_modern_slide(
    '현대 ③ — 일상의 지혜',
    '개인의 삶에 적용되는 손자',
    [
        ('선승이후구전(先勝而後求戰)', {'font_size': 22, 'bold': True, 'color': ACCENT}),
        ('  → 준비 없이 도전하지 않는다 · 사전 준비의 가치',
         {'font_size': 16, 'space_before': 4}),
        ('노이흥사(怒而興師) 금지', {'font_size': 22, 'bold': True, 'color': ACCENT, 'space_before': 10}),
        ('  → 감정에 휩쓸린 중요 결정을 피하라',
         {'font_size': 16, 'space_before': 4}),
        ('병형상수(兵形象水)', {'font_size': 22, 'bold': True, 'color': ACCENT, 'space_before': 10}),
        ('  → 고정된 패턴에 얽매이지 마라 · 상황에 맞게 변하라',
         {'font_size': 16, 'space_before': 4}),
    ]), 'Ⅷ. 현대적 의의'))


# ---------- Ⅸ. 비교 ----------
@S('Ⅸ. 비교')
def s_compare(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅸ. 비교', page, total)
    add_title(slide, '손자병법과 다른 사상의 비교')
    rows = [
        ('클라우제비츠 (전쟁론)', '전쟁을 정치의 연장으로 봄', '손자: 부전승 / 클라우제비츠: 결전 중시'),
        ('마키아벨리 (군주론)',   '현실주의적 권력관',         '손자: 도(道) 우선 / 마키아벨리: 권모술수'),
        ('노자 (도덕경)',         '유연함과 물의 비유',         '노자: 무위(無爲) / 손자: 적극적 전략'),
        ('공자 (논어)',           '인(仁)의 가치 중시',         '공자: 도덕 우선 / 손자: 전략적 효용'),
        ('순자',                  '현실주의적 인간관',           '순자: 제도 설계 / 손자: 전장 설계'),
    ]
    top = 2.0
    # 헤더
    add_filled_rect(slide, Inches(0.5), Inches(top), Inches(3.5), Inches(0.55), ACCENT)
    add_textbox(slide, Inches(0.5), Inches(top + 0.1), Inches(3.5), Inches(0.4),
                '비교 대상', font_size=15, bold=True,
                color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
    add_filled_rect(slide, Inches(4.1), Inches(top), Inches(4.0), Inches(0.55), SUB)
    add_textbox(slide, Inches(4.1), Inches(top + 0.1), Inches(4.0), Inches(0.4),
                '공통점', font_size=15, bold=True,
                color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
    add_filled_rect(slide, Inches(8.2), Inches(top), Inches(4.6), Inches(0.55), SUB)
    add_textbox(slide, Inches(8.2), Inches(top + 0.1), Inches(4.6), Inches(0.4),
                '차이점', font_size=15, bold=True,
                color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
    row_h = 0.85
    for r_idx, row in enumerate(rows):
        y = top + 0.55 + r_idx * row_h
        bg = RGBColor(0xFA, 0xFA, 0xFA) if r_idx % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF)
        add_filled_rect(slide, Inches(0.5), Inches(y), Inches(3.5), Inches(row_h), PALE)
        add_textbox(slide, Inches(0.55), Inches(y + 0.2), Inches(3.4), Inches(0.5),
                    row[0], font_size=14, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(4.1), Inches(y), Inches(4.0), Inches(row_h), bg)
        add_textbox(slide, Inches(4.2), Inches(y + 0.22), Inches(3.8), Inches(0.5),
                    row[1], font_size=13, color=INK, align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(8.2), Inches(y), Inches(4.6), Inches(row_h), bg)
        add_textbox(slide, Inches(8.3), Inches(y + 0.22), Inches(4.4), Inches(0.5),
                    row[2], font_size=13, color=INK, align=PP_ALIGN.CENTER)


# ---------- Ⅹ. 마무리 ----------
@S('Ⅹ. 마무리')
def s_summary(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅹ. 마무리', page, total)
    add_title(slide, '한 문장으로 정리하는 손자병법')
    add_filled_rect(slide, Inches(0.7), Inches(2.3), Inches(11.9), Inches(4.5), PALE)
    add_paragraphs(slide, Inches(1.1), Inches(2.7), Inches(11.1), Inches(4.1), [
        ('전쟁의 가부(可否)를 신중히 판단하고 (오사칠계)',
         {'font_size': 22, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 8}),
        ('가능하면 싸우지 않고 이기며 (부전승)',
         {'font_size': 22, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER, 'space_before': 8}),
        ('', {'font_size': 8}),
        ('반드시 싸워야 한다면 먼저 이길 조건을 갖춘다 (선승이후구전)',
         {'font_size': 22, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER, 'space_before': 8}),
        ('', {'font_size': 8}),
        ('이 모든 것의 기반은',
         {'font_size': 22, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 8}),
        ('정보(지피지기)와 유연한 변화(병무상세)다',
         {'font_size': 22, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 8}),
        ('— 전쟁을 다루지만, 결국은 신중함의 철학 —',
         {'font_size': 18, 'bold': True, 'color': SUB, 'align': PP_ALIGN.CENTER, 'space_before': 10}),
    ], line_spacing=1.2)


@S('Ⅹ. 마무리')
def s_closing(slide, page, total):
    set_white_background(slide)
    add_filled_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.4), RULE)
    add_textbox(slide, Inches(0.5), Inches(1.8), Inches(12.3), Inches(1.5),
                '不 戰 而 屈 人 之 兵',
                font_size=54, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.2), Inches(12.3), Inches(1.5),
                '善 之 善 者 也',
                font_size=44, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(4.4), Inches(12.3), Inches(0.5),
                '부전이굴인지병 선지선자야', font_size=18, color=SUB,
                align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(5.1), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(5.4), Inches(12.3), Inches(0.6),
                '"싸우지 않고 적을 굴복시키는 것이 최선의 최선이다"',
                font_size=22, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.1), Inches(12.3), Inches(0.4),
                '— 모공편 제3',
                font_size=14, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.8), Inches(12.3), Inches(0.5),
                '감사합니다', font_size=26, bold=True, color=INK,
                align=PP_ALIGN.CENTER)
    add_filled_rect(slide, Inches(0), Inches(7.1), Inches(13.333), Inches(0.4), RULE)


# ---------- 빌드 ----------
total_pages = len(SLIDES)
for i, (renderer, section) in enumerate(SLIDES, 1):
    slide = prs.slides.add_slide(blank)
    renderer(slide, i, total_pages)

out_path = r'C:\Users\박호군\ClaudeProjects\Oriental-Classics2\발표자료\손자병법_발표자료.pptx'
prs.save(out_path)
print(f'Saved: {out_path}  ({total_pages} slides)')
