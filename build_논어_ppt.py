# -*- coding: utf-8 -*-
"""
논어 발표자료 재작성 스크립트
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
from pptx.oxml.ns import qn
from lxml import etree

# 색상 팔레트 (흰 배경 + 가독성 있는 어두운 글자)
INK        = RGBColor(0x1F, 0x2A, 0x44)   # 본문 (짙은 남색-검정)
ACCENT     = RGBColor(0x8B, 0x1A, 0x1A)   # 강조 (와인 레드 — 한자/구절)
SUB        = RGBColor(0x55, 0x60, 0x70)   # 보조 (회색)
RULE       = RGBColor(0xC8, 0xA2, 0x5B)   # 장식선 (골드)
PALE       = RGBColor(0xF5, 0xEE, 0xDD)   # 강조 박스 배경 (옅은 베이지)


def set_white_background(slide):
    """슬라이드 배경을 명시적으로 흰색으로 채운다."""
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
    """lines: list of (text, {opts}) 또는 str"""
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
    """장식용 가로선."""
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
    """모든 본문 슬라이드 상단 공통 헤더 (섹션명 + 페이지)"""
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


# ---------------------- 슬라이드 생성 ----------------------

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

# 모든 슬라이드 만들기 위해 먼저 총 개수를 추정하고 그 다음 페이지 번호를 채울 것.
# 단순화: 본문 슬라이드(3번부터)에만 페이지 번호를 표시.

SLIDES = []  # (renderer, section)

def S(section):
    def deco(fn):
        SLIDES.append((fn, section))
        return fn
    return deco


# 1. 표지
@S('표지')
def s_cover(slide, page, total):
    set_white_background(slide)
    # 상단 장식
    add_filled_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.4), RULE)
    # 큰 한자
    add_textbox(slide, Inches(0.5), Inches(1.8), Inches(12.3), Inches(1.6),
                '論 語', font_size=110, bold=True, color=INK,
                align=PP_ALIGN.CENTER, font_name='맑은 고딕')
    add_textbox(slide, Inches(0.5), Inches(3.7), Inches(12.3), Inches(0.5),
                'The Analects · 논어', font_size=24, color=ACCENT,
                align=PP_ALIGN.CENTER)
    # 가운데 장식선
    add_rule(slide, Inches(5.5), Inches(4.5), Inches(2.3), color=RULE, weight=1.5)
    # 부제
    add_textbox(slide, Inches(0.5), Inches(4.8), Inches(12.3), Inches(0.5),
                '공자(孔子)와 제자들의 언행록', font_size=20, color=INK,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.3), Inches(12.3), Inches(0.4),
                '춘추시대 (BC 551~479) · 사서(四書)의 으뜸 · 20편 482장',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    # 하단 장식
    add_filled_rect(slide, Inches(0), Inches(7.1), Inches(13.333), Inches(0.4), RULE)


# 2. 목차
@S('목차')
def s_toc(slide, page, total):
    set_white_background(slide)
    add_textbox(slide, Inches(0.5), Inches(0.5), Inches(12.8), Inches(0.7),
                '목 차', font_size=36, bold=True, color=INK)
    add_rule(slide, Inches(0.5), Inches(1.3), Inches(12.8))
    items = [
        ('Ⅰ',  '개요 — 논어란 무엇인가'),
        ('Ⅱ',  '20편의 구성'),
        ('Ⅲ',  '핵심 사상 ① — 인(仁)'),
        ('Ⅳ',  '핵심 사상 ② — 예(禮)와 덕치(德治)'),
        ('Ⅴ',  '핵심 사상 ③ — 군자(君子)'),
        ('Ⅵ',  '명구절 10선'),
        ('Ⅶ',  '논어의 구조적 특징'),
        ('Ⅷ',  '현대적 의의'),
        ('Ⅸ',  '다른 고전과의 비교'),
        ('Ⅹ',  '마무리'),
    ]
    top = 1.7
    for num, title in items:
        add_textbox(slide, Inches(1.2), Inches(top), Inches(1.0), Inches(0.4),
                    num, font_size=22, bold=True, color=ACCENT)
        add_textbox(slide, Inches(2.4), Inches(top), Inches(10.0), Inches(0.4),
                    title, font_size=20, color=INK)
        top += 0.5


# ---------------------- Ⅰ. 개요 ----------------------

@S('Ⅰ. 개요')
def s_what_is(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅰ. 개요', page, total)
    add_title(slide, '논어(論語)란 무엇인가')
    # 큰 정의 박스
    add_filled_rect(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(1.4), PALE)
    add_textbox(slide, Inches(0.9), Inches(2.55), Inches(11.5), Inches(0.5),
                '공자와 제자들의 언행을 기록한 유교 핵심 경전',
                font_size=26, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.9), Inches(3.15), Inches(11.5), Inches(0.5),
                '"論(논의하다) + 語(말씀)" — 토론하여 정리한 말씀',
                font_size=17, color=ACCENT, align=PP_ALIGN.CENTER)
    # 수치
    nums = [('20', '편(篇)'), ('482', '장(章)'), ('약 600', '문장')]
    for i, (n, lbl) in enumerate(nums):
        x = 1.8 + i * 3.5
        add_textbox(slide, Inches(x), Inches(4.4), Inches(3.0), Inches(1.0),
                    n, font_size=64, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(x), Inches(5.6), Inches(3.0), Inches(0.5),
                    lbl, font_size=18, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅰ. 개요')
def s_confucius(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅰ. 개요', page, total)
    add_title(slide, '공자(孔子, BC 551~479)', '유교의 창시자, 동아시아 사상의 원천')
    lines = [
        ('이름·자(字)', {'bold': True, 'font_size': 18, 'color': ACCENT}),
        ('  공구(孔丘) — 이름  |  중니(仲尼) — 자', {'font_size': 18}),
        ('출신·신분', {'bold': True, 'font_size': 18, 'color': ACCENT, 'space_before': 12}),
        ('  노(魯)나라 출신  |  정치인·사상가·교육자', {'font_size': 18}),
        ('제자', {'bold': True, 'font_size': 18, 'color': ACCENT, 'space_before': 12}),
        ('  약 3,000명, 그중 뛰어난 자 72명(七十二賢)', {'font_size': 18}),
        ('생애의 분기점', {'bold': True, 'font_size': 18, 'color': ACCENT, 'space_before': 12}),
        ('  14년간 천하를 주유(周遊) 후 고향에 돌아와 후학 양성에 전념',
         {'font_size': 18}),
    ]
    add_paragraphs(slide, Inches(0.9), Inches(2.3), Inches(11.5), Inches(4.5),
                   lines, line_spacing=1.4)


@S('Ⅰ. 개요')
def s_compilation(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅰ. 개요', page, total)
    add_title(slide, '편찬 과정', '한 사람의 저작이 아닌, 여러 세대에 걸친 집단 편집')
    # 3단계 박스
    stages = [
        ('1차', '직계 제자 주도', '중궁·자유·자하 등이 공자 사후 어록을 모아 1차 편집'),
        ('2차', '증자 사후 보충', '증자(曾子) 학파가 자료를 보충하여 확장'),
        ('3차', '전국시대 추가·확정', '전한 말 장우(張禹) 편집본을 거쳐 후한대에 현재 형태로 확정'),
    ]
    top = 2.3
    for tag, title, desc in stages:
        add_filled_rect(slide, Inches(0.7), Inches(top), Inches(1.4), Inches(1.1), ACCENT)
        add_textbox(slide, Inches(0.7), Inches(top + 0.3), Inches(1.4), Inches(0.6),
                    tag, font_size=28, bold=True,
                    color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(2.3), Inches(top + 0.05), Inches(10.5), Inches(0.5),
                    title, font_size=22, bold=True, color=INK)
        add_textbox(slide, Inches(2.3), Inches(top + 0.55), Inches(10.5), Inches(0.5),
                    desc, font_size=16, color=SUB)
        top += 1.45


# ---------------------- Ⅱ. 구성 ----------------------

@S('Ⅱ. 구성')
def s_structure_overview(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅱ. 구성', page, total)
    add_title(slide, '20편의 구성 개관', '상론(上論) 10편 + 하론(下論) 10편 = 총 20편 482장')
    # 좌우 분할
    add_filled_rect(slide, Inches(0.7), Inches(2.4), Inches(5.9), Inches(4.4), PALE)
    add_filled_rect(slide, Inches(6.8), Inches(2.4), Inches(5.9), Inches(4.4), PALE)
    # 좌: 상론
    add_textbox(slide, Inches(0.7), Inches(2.6), Inches(5.9), Inches(0.5),
                '상론(上論) 1~10편', font_size=24, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(2.5), Inches(3.15), Inches(2.3), color=RULE, weight=1.2)
    add_paragraphs(slide, Inches(1.0), Inches(3.4), Inches(5.3), Inches(3.2), [
        ('학이 · 위정 · 팔일 · 이인', {'font_size': 17, 'bold': True, 'align': PP_ALIGN.CENTER}),
        ('공야장 · 옹야 · 술이', {'font_size': 17, 'bold': True, 'align': PP_ALIGN.CENTER}),
        ('태백 · 자한 · 향당', {'font_size': 17, 'bold': True, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('', {'font_size': 8}),
        ('수신·인·예·공자의 일상까지', {'font_size': 15, 'color': SUB, 'align': PP_ALIGN.CENTER, 'space_before': 12}),
        ('— 논어의 총론적 부분', {'font_size': 15, 'color': SUB, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.4)
    # 우: 하론
    add_textbox(slide, Inches(6.8), Inches(2.6), Inches(5.9), Inches(0.5),
                '하론(下論) 11~20편', font_size=24, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(8.6), Inches(3.15), Inches(2.3), color=RULE, weight=1.2)
    add_paragraphs(slide, Inches(7.1), Inches(3.4), Inches(5.3), Inches(3.2), [
        ('선진 · 안연 · 자로', {'font_size': 17, 'bold': True, 'align': PP_ALIGN.CENTER}),
        ('헌문 · 위령공 · 계씨', {'font_size': 17, 'bold': True, 'align': PP_ALIGN.CENTER}),
        ('양화 · 미자 · 자장 · 요왈', {'font_size': 17, 'bold': True, 'align': PP_ALIGN.CENTER, 'space_before': 6}),
        ('', {'font_size': 8}),
        ('정치·군자론·역대 성왕에 이르기까지', {'font_size': 15, 'color': SUB, 'align': PP_ALIGN.CENTER, 'space_before': 12}),
        ('— 논어의 각론·심화부', {'font_size': 15, 'color': SUB, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.4)


@S('Ⅱ. 구성')
def s_upper(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅱ. 구성', page, total)
    add_title(slide, '상론(上論) — 1~10편의 흐름')
    rows = [
        ('1', '학이(學而)', '배움·효제·수신 — 논어의 총론'),
        ('2', '위정(爲政)', '덕치·효도·학문의 단계'),
        ('3', '팔일(八佾)', '예악의 본질, 인과 예의 관계'),
        ('4', '이인(里仁)', '인(仁)의 내면, 충서(忠恕)'),
        ('5–6', '공야장 · 옹야', '인물 평가, 문질빈빈(文質彬彬)'),
        ('7–10', '술이 · 태백 · 자한 · 향당', '공자의 학문관·일상·절조'),
    ]
    top = 2.3
    for num, name, desc in rows:
        add_textbox(slide, Inches(0.9), Inches(top), Inches(1.3), Inches(0.4),
                    num, font_size=20, bold=True, color=ACCENT)
        add_textbox(slide, Inches(2.3), Inches(top), Inches(3.5), Inches(0.4),
                    name, font_size=20, bold=True, color=INK)
        add_textbox(slide, Inches(5.8), Inches(top), Inches(7.0), Inches(0.4),
                    desc, font_size=17, color=SUB)
        top += 0.65


@S('Ⅱ. 구성')
def s_lower(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅱ. 구성', page, total)
    add_title(slide, '하론(下論) — 11~20편의 흐름')
    rows = [
        ('11–12', '선진 · 안연', '과유불급(過猶不及), 극기복례(克己復禮)'),
        ('13–15', '자로 · 헌문 · 위령공', '정명(正名)·정치·군자론의 총결산'),
        ('16–18', '계씨 · 양화 · 미자', '수양·예악·은일(隱逸)과 입세(入世)'),
        ('19',    '자장(子張)',         '제자들의 발언 모음 — 공자 어록이 없는 유일한 편'),
        ('20',    '요왈(堯曰)',         '요·순·우 등 성왕(聖王)의 이상, 논어의 결어'),
    ]
    top = 2.4
    for num, name, desc in rows:
        add_textbox(slide, Inches(0.9), Inches(top), Inches(1.5), Inches(0.4),
                    num, font_size=20, bold=True, color=ACCENT)
        add_textbox(slide, Inches(2.5), Inches(top), Inches(3.5), Inches(0.4),
                    name, font_size=20, bold=True, color=INK)
        add_textbox(slide, Inches(6.0), Inches(top), Inches(7.0), Inches(0.4),
                    desc, font_size=17, color=SUB)
        top += 0.75


# ---------------------- Ⅲ. 인(仁) ----------------------

@S('Ⅲ. 인(仁)')
def s_ren_concept(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 인(仁)', page, total)
    add_title(slide, '인(仁) — 사람다움의 최고 덕목')
    # 큰 한자
    add_textbox(slide, Inches(0.9), Inches(2.4), Inches(3.5), Inches(3.5),
                '仁', font_size=240, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 해설
    add_filled_rect(slide, Inches(4.8), Inches(2.5), Inches(8.0), Inches(4.2), PALE)
    add_paragraphs(slide, Inches(5.1), Inches(2.8), Inches(7.5), Inches(3.8), [
        ('어원', {'bold': True, 'font_size': 20, 'color': ACCENT}),
        ('사람 人(인) + 둘 二(이)', {'font_size': 18}),
        ('= 사람 둘이 함께하는 모양', {'font_size': 18}),
        ('', {'font_size': 8}),
        ('의미', {'bold': True, 'font_size': 20, 'color': ACCENT, 'space_before': 8}),
        ('사람과 사람의 바람직한 관계', {'font_size': 18}),
        ('곧 "사람다움" 그 자체', {'font_size': 18}),
        ('', {'font_size': 8}),
        ('논어에서의 위치', {'bold': True, 'font_size': 20, 'color': ACCENT, 'space_before': 8}),
        ('전체 등장 횟수 109회 — 단연 최다 핵심어',
         {'font_size': 18, 'color': INK}),
    ], line_spacing=1.35)


@S('Ⅲ. 인(仁)')
def s_keji(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 인(仁)', page, total)
    add_title(slide, '克己復禮爲仁 — 극기복례위인')
    # 한자 구절
    add_textbox(slide, Inches(0.5), Inches(2.2), Inches(12.3), Inches(1.0),
                '克 己 復 禮 爲 仁', font_size=64, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    # 음
    add_textbox(slide, Inches(0.5), Inches(3.3), Inches(12.3), Inches(0.5),
                '극 기 복 례 위 인', font_size=24, color=SUB,
                align=PP_ALIGN.CENTER)
    # 출처
    add_textbox(slide, Inches(0.5), Inches(3.9), Inches(12.3), Inches(0.4),
                '— 안연편 1장', font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.5), Inches(2.3), color=RULE, weight=1.5)
    # 풀이
    add_textbox(slide, Inches(0.5), Inches(4.8), Inches(12.3), Inches(0.6),
                '자기를 이기고 예(禮)로 돌아감이 곧 인(仁)이다',
                font_size=24, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.6), Inches(12.3), Inches(0.6),
                '자신의 욕망을 절제하고 사회적 규범과 조화를 이루는 것이 인의 출발',
                font_size=17, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅲ. 인(仁)')
def s_aeren(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 인(仁)', page, total)
    add_title(slide, '仁者愛人 — 인자애인')
    add_textbox(slide, Inches(0.5), Inches(2.2), Inches(12.3), Inches(1.0),
                '仁 者 愛 人', font_size=72, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.4), Inches(12.3), Inches(0.5),
                '인 자 애 인', font_size=24, color=SUB,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(4.0), Inches(12.3), Inches(0.4),
                '— 안연편 22장', font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.6), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(4.9), Inches(12.3), Inches(0.6),
                '인(仁)이란, 사람을 사랑하는 것이다',
                font_size=26, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.7), Inches(12.3), Inches(0.6),
                '제자 번지(樊遲)의 "인이 무엇입니까?" 라는 물음에 대한 공자의 가장 간명한 답',
                font_size=17, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅲ. 인(仁)')
def s_chukigeup(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 인(仁)', page, total)
    add_title(slide, '추기급인(推己及人) — 자기에서 출발하는 사랑')
    # 두 구절 좌우
    # 좌
    add_filled_rect(slide, Inches(0.6), Inches(2.3), Inches(6.0), Inches(4.5), PALE)
    add_textbox(slide, Inches(0.6), Inches(2.5), Inches(6.0), Inches(0.5),
                '적극적 표현', font_size=18, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.7), Inches(3.1), Inches(5.8), Inches(1.5),
                '己欲立而立人\n己欲達而達人',
                font_size=32, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.7), Inches(4.7), Inches(5.8), Inches(1.0),
                '자기가 서고자 하면 남을 세워주고\n자기가 도달하고자 하면 남을 도달하게 한다',
                font_size=15, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.7), Inches(6.3), Inches(5.8), Inches(0.4),
                '— 옹야편 30장', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    # 우
    add_filled_rect(slide, Inches(6.8), Inches(2.3), Inches(6.0), Inches(4.5), PALE)
    add_textbox(slide, Inches(6.8), Inches(2.5), Inches(6.0), Inches(0.5),
                '소극적 표현 (황금률)', font_size=18, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(6.9), Inches(3.1), Inches(5.8), Inches(1.5),
                '己所不欲\n勿施於人',
                font_size=36, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(6.9), Inches(4.7), Inches(5.8), Inches(1.0),
                '자기가 원하지 않는 것을\n남에게 하지 말라',
                font_size=15, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(6.9), Inches(6.3), Inches(5.8), Inches(0.4),
                '— 위령공편 24장', font_size=13, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅲ. 인(仁)')
def s_chungseo(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 인(仁)', page, total)
    add_title(slide, '충서(忠恕) — 인의 실천 방법',
              '"夫子之道 忠恕而已矣" — 선생님의 도는 충서일 뿐이다 (이인편 15장)')
    # 두 글자 큰 박스
    # 충
    add_filled_rect(slide, Inches(0.8), Inches(2.5), Inches(5.8), Inches(4.4), PALE)
    add_textbox(slide, Inches(0.8), Inches(2.7), Inches(5.8), Inches(1.6),
                '忠', font_size=130, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.8), Inches(4.5), Inches(5.8), Inches(0.5),
                '충 (盡己)', font_size=22, bold=True, color=INK,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.8), Inches(5.1), Inches(5.8), Inches(1.5),
                '자기 마음을 다하는 것\n— 안으로 향하는 진정성',
                font_size=17, color=SUB, align=PP_ALIGN.CENTER)
    # 서
    add_filled_rect(slide, Inches(6.8), Inches(2.5), Inches(5.8), Inches(4.4), PALE)
    add_textbox(slide, Inches(6.8), Inches(2.7), Inches(5.8), Inches(1.6),
                '恕', font_size=130, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(6.8), Inches(4.5), Inches(5.8), Inches(0.5),
                '서 (推己及人)', font_size=22, bold=True, color=INK,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(6.8), Inches(5.1), Inches(5.8), Inches(1.5),
                '자기를 미루어 남에게 미치는 것\n— 밖으로 향하는 공감',
                font_size=17, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅲ. 인(仁)')
def s_hyo(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅲ. 인(仁)', page, total)
    add_title(slide, '효(孝) — 인의 출발점')
    add_textbox(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(1.0),
                '孝 弟 也 者   其 爲 仁 之 本 與',
                font_size=44, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.4), Inches(12.3), Inches(0.5),
                '효제야자 기위인지본여', font_size=20, color=SUB,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(4.0), Inches(12.3), Inches(0.4),
                '— 학이편 2장 (유자 발언)', font_size=15, color=SUB,
                align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.6), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(4.9), Inches(12.3), Inches(0.6),
                '효(孝)와 제(弟)는 인을 이루는 근본이다',
                font_size=26, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.7), Inches(12.3), Inches(0.8),
                '가장 가까운 가족 관계에서 시작된 사랑이\n점차 이웃·사회·천하로 확장되어 인(仁)이 된다',
                font_size=17, color=SUB, align=PP_ALIGN.CENTER)


# ---------------------- Ⅳ. 예와 덕치 ----------------------

@S('Ⅳ. 예·덕치')
def s_li(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅳ. 예·덕치', page, total)
    add_title(slide, '예(禮) — 인(仁)의 외적 표현')
    add_textbox(slide, Inches(0.9), Inches(2.4), Inches(3.0), Inches(2.5),
                '禮', font_size=200, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_paragraphs(slide, Inches(4.5), Inches(2.4), Inches(8.5), Inches(4.5), [
        ('의미의 확장', {'bold': True, 'font_size': 20, 'color': ACCENT}),
        ('종교 제례(祭禮) → 일상 행위 규범으로 확장',
         {'font_size': 17}),
        ('', {'font_size': 8}),
        ('인과 예의 관계', {'bold': True, 'font_size': 20, 'color': ACCENT, 'space_before': 10}),
        ('人而不仁 如禮何', {'font_size': 24, 'bold': True, 'color': INK}),
        ('어질지 못하면 예를 어찌하겠는가 (팔일편 3장)',
         {'font_size': 15, 'color': SUB}),
        ('', {'font_size': 8}),
        ('두 축은 한 몸', {'bold': True, 'font_size': 20, 'color': ACCENT, 'space_before': 10}),
        ('인(仁) 없는 예 = 빈 껍데기',
         {'font_size': 17}),
        ('예(禮) 없는 인 = 실현 불가능한 추상',
         {'font_size': 17}),
    ], line_spacing=1.35)


@S('Ⅳ. 예·덕치')
def s_deokchi(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅳ. 예·덕치', page, total)
    add_title(slide, '덕치(德治) — 덕(德)으로 다스림')
    # 구절 1
    add_filled_rect(slide, Inches(0.8), Inches(2.3), Inches(11.7), Inches(1.9), PALE)
    add_textbox(slide, Inches(0.8), Inches(2.45), Inches(11.7), Inches(0.6),
                '爲政以德   譬如北辰',
                font_size=32, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.8), Inches(3.15), Inches(11.7), Inches(0.5),
                '덕으로 정치를 하면, 마치 북극성과 같다',
                font_size=18, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.8), Inches(3.7), Inches(11.7), Inches(0.4),
                '— 위정편 1장', font_size=14, color=SUB, align=PP_ALIGN.CENTER)
    # 구절 2
    add_filled_rect(slide, Inches(0.8), Inches(4.5), Inches(11.7), Inches(1.9), PALE)
    add_textbox(slide, Inches(0.8), Inches(4.65), Inches(11.7), Inches(0.6),
                '其身正   不令而行',
                font_size=32, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.8), Inches(5.35), Inches(11.7), Inches(0.5),
                '자신이 바르면, 명령하지 않아도 행해진다',
                font_size=18, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.8), Inches(5.9), Inches(11.7), Inches(0.4),
                '— 자로편 6장', font_size=14, color=SUB, align=PP_ALIGN.CENTER)
    # 요약
    add_textbox(slide, Inches(0.8), Inches(6.6), Inches(11.7), Inches(0.4),
                '솔선수범(率先垂範) — 통치자의 인격이 곧 정치의 정당성',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER, bold=True)


@S('Ⅳ. 예·덕치')
def s_jeongmyeong(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅳ. 예·덕치', page, total)
    add_title(slide, '정명(正名) — 이름을 바르게 함')
    add_textbox(slide, Inches(0.5), Inches(2.3), Inches(12.3), Inches(0.9),
                '名 不 正 則 言 不 順',
                font_size=44, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.3), Inches(12.3), Inches(0.5),
                '이름(명분)이 바르지 않으면 말이 순조롭지 못하고,',
                font_size=18, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.8), Inches(12.3), Inches(0.5),
                '말이 순조롭지 않으면 일이 이루어지지 못한다',
                font_size=18, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(4.3), Inches(12.3), Inches(0.4),
                '— 자로편 3장', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.95), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(5.2), Inches(12.3), Inches(0.7),
                '君  君   臣  臣   父  父   子  子',
                font_size=32, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.5),
                '임금은 임금답게, 신하는 신하답게, 아버지는 아버지답게, 자식은 자식답게',
                font_size=16, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.5), Inches(12.3), Inches(0.4),
                '— 안연편 11장 · 각자 자기 역할에 충실할 것',
                font_size=13, color=SUB, align=PP_ALIGN.CENTER)


@S('Ⅳ. 예·덕치')
def s_education(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅳ. 예·덕치', page, total)
    add_title(slide, '공자의 교육 철학')
    # 두 구절
    add_filled_rect(slide, Inches(0.6), Inches(2.3), Inches(6.1), Inches(4.5), PALE)
    add_textbox(slide, Inches(0.6), Inches(2.5), Inches(6.1), Inches(0.5),
                '배움과 사유의 균형', font_size=18, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.7), Inches(3.2), Inches(5.9), Inches(1.5),
                '學而不思則罔\n思而不學則殆',
                font_size=28, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.7), Inches(5.0), Inches(5.9), Inches(1.3),
                '배우기만 하고 생각하지 않으면 어둡고\n생각하기만 하고 배우지 않으면 위태롭다',
                font_size=15, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.7), Inches(6.4), Inches(5.9), Inches(0.4),
                '— 위정편 15장', font_size=13, color=SUB, align=PP_ALIGN.CENTER)
    # 우
    add_filled_rect(slide, Inches(6.9), Inches(2.3), Inches(6.0), Inches(4.5), PALE)
    add_textbox(slide, Inches(6.9), Inches(2.5), Inches(6.0), Inches(0.5),
                '교육 기회의 평등', font_size=18, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(6.9), Inches(3.3), Inches(6.0), Inches(1.4),
                '有 敎 無 類',
                font_size=64, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(6.9), Inches(5.0), Inches(6.0), Inches(1.3),
                '가르침에는 차별이 없다\n신분·빈부와 무관하게 누구나',
                font_size=15, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(6.9), Inches(6.4), Inches(6.0), Inches(0.4),
                '— 위령공편 39장', font_size=13, color=SUB,
                align=PP_ALIGN.CENTER)


# ---------------------- Ⅴ. 군자 ----------------------

@S('Ⅴ. 군자(君子)')
def s_gunja(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 군자(君子)', page, total)
    add_title(slide, '군자(君子) — 이상적 인간상')
    add_textbox(slide, Inches(0.9), Inches(2.4), Inches(3.5), Inches(3.5),
                '君子', font_size=140, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_paragraphs(slide, Inches(4.8), Inches(2.5), Inches(8.0), Inches(4.5), [
        ('어원적 의미', {'bold': True, 'font_size': 20, 'color': ACCENT}),
        ('본래 "임금의 아들" — 귀족·통치 계층을 지칭하는 신분 용어',
         {'font_size': 16, 'color': SUB}),
        ('', {'font_size': 8}),
        ('공자의 재정의', {'bold': True, 'font_size': 20, 'color': ACCENT, 'space_before': 10}),
        ('혈통이 아니라 인격으로 정의되는',
         {'font_size': 18}),
        ('"도덕적으로 수양된 이상적 인간"',
         {'font_size': 18, 'bold': True, 'color': INK}),
        ('', {'font_size': 8}),
        ('사상사적 의의', {'bold': True, 'font_size': 20, 'color': ACCENT, 'space_before': 10}),
        ('계급 개념을 도덕 개념으로 전환 — 누구나 군자가 될 수 있다',
         {'font_size': 16, 'color': SUB}),
    ], line_spacing=1.35)


@S('Ⅴ. 군자(君子)')
def s_gunja_vs(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 군자(君子)', page, total)
    add_title(slide, '군자 vs 소인 — 다섯 가지 대비')
    # 헤더
    add_filled_rect(slide, Inches(0.6), Inches(2.2), Inches(6.0), Inches(0.5), ACCENT)
    add_filled_rect(slide, Inches(6.7), Inches(2.2), Inches(6.0), Inches(0.5), SUB)
    add_textbox(slide, Inches(0.6), Inches(2.25), Inches(6.0), Inches(0.4),
                '군자(君子)', font_size=18, bold=True,
                color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(6.7), Inches(2.25), Inches(6.0), Inches(0.4),
                '소인(小人)', font_size=18, bold=True,
                color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
    pairs = [
        ('和而不同 — 조화롭되 같지 않다',  '同而不和 — 같되 조화롭지 않다', '자로 23'),
        ('周而不比 — 두루 사귀되 편당하지 않는다', '比而不周 — 편당하되 두루 사귀지 못한다', '위정 14'),
        ('坦蕩蕩 — 마음이 넓고 평탄하다',  '長戚戚 — 항상 근심하고 초조하다',  '술이 37'),
        ('喩於義 — 의(義)에 밝다',           '喩於利 — 이(利)에 밝다',             '이인 16'),
        ('求諸己 — 자신에게서 구한다',       '求諸人 — 남에게서 구한다',           '위령공 21'),
    ]
    top = 2.85
    for left, right, ref in pairs:
        add_filled_rect(slide, Inches(0.6), Inches(top), Inches(6.0), Inches(0.7),
                        RGBColor(0xFA, 0xFA, 0xFA))
        add_filled_rect(slide, Inches(6.7), Inches(top), Inches(6.0), Inches(0.7),
                        RGBColor(0xFA, 0xFA, 0xFA))
        add_textbox(slide, Inches(0.7), Inches(top + 0.08), Inches(5.8), Inches(0.5),
                    left, font_size=15, color=INK, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(6.8), Inches(top + 0.08), Inches(5.8), Inches(0.5),
                    right, font_size=15, color=INK, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(0.6), Inches(top + 0.38), Inches(12.1), Inches(0.3),
                    f'({ref})', font_size=10, color=SUB, align=PP_ALIGN.CENTER)
        top += 0.78


@S('Ⅴ. 군자(君子)')
def s_samdaldeok(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅴ. 군자(君子)', page, total)
    add_title(slide, '삼달덕(三達德) — 군자의 세 덕목', '— 자한편 29장')
    items = [
        ('知', '知者不惑', '지자불혹', '지혜로운 자는 미혹되지 않는다'),
        ('仁', '仁者不憂', '인자불우', '어진 자는 근심하지 않는다'),
        ('勇', '勇者不懼', '용자불구', '용감한 자는 두려워하지 않는다'),
    ]
    for i, (ch, hanmun, eum, mean) in enumerate(items):
        x = 0.5 + i * 4.3
        add_filled_rect(slide, Inches(x), Inches(2.3), Inches(4.1), Inches(4.7), PALE)
        add_textbox(slide, Inches(x), Inches(2.5), Inches(4.1), Inches(1.6),
                    ch, font_size=130, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(x), Inches(4.5), Inches(4.1), Inches(0.6),
                    hanmun, font_size=26, bold=True, color=INK,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(x), Inches(5.2), Inches(4.1), Inches(0.4),
                    eum, font_size=15, color=SUB, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(x), Inches(5.8), Inches(4.1), Inches(1.0),
                    mean, font_size=15, color=INK, align=PP_ALIGN.CENTER)


# ---------------------- Ⅵ. 명구절 ----------------------

def make_quote_slide(section, page, total, hanmun, eum, mean, ref, *, hanmun_size=58):
    def renderer(slide, _page, _total):
        set_white_background(slide)
        add_page_header(slide, section, _page, _total)
        # 한자 구절 가운데 크게
        add_textbox(slide, Inches(0.5), Inches(1.2), Inches(12.3), Inches(2.4),
                    hanmun, font_size=hanmun_size, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # 음
        add_textbox(slide, Inches(0.5), Inches(3.7), Inches(12.3), Inches(0.5),
                    eum, font_size=22, color=SUB, align=PP_ALIGN.CENTER)
        # 장식선
        add_rule(slide, Inches(5.5), Inches(4.4), Inches(2.3), color=RULE, weight=1.5)
        # 뜻
        add_textbox(slide, Inches(0.5), Inches(4.7), Inches(12.3), Inches(1.3),
                    mean, font_size=24, bold=True, color=INK,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # 출처
        add_textbox(slide, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.4),
                    f'— {ref}', font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    return renderer


SLIDES.append((make_quote_slide('Ⅵ. 명구절 (1/10)', 0, 0,
    '學而時習之  不亦說乎',
    '학이시습지   불역열호',
    '배우고 때때로 익히면 또한 기쁘지 아니한가',
    '학이편 1장'), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (2/10)', 0, 0,
    '溫故而知新  可以爲師矣',
    '온고이지신   가이위사의',
    '옛것을 익히고 새것을 알면 스승이 될 만하다',
    '위정편 11장'), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (3/10)', 0, 0,
    '朝聞道  夕死可矣',
    '조문도   석사가의',
    '아침에 도(道)를 들으면 저녁에 죽어도 좋다',
    '이인편 8장', hanmun_size=64), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (4/10)', 0, 0,
    '三人行  必有我師焉',
    '삼인행   필유아사언',
    '세 사람이 길을 가면 반드시 그 중에 나의 스승이 있다',
    '술이편 22장', hanmun_size=58), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (5/10)', 0, 0,
    '歲寒然後  知松柏之後彫也',
    '세한연후   지송백지후조야',
    '날이 추워진 뒤에야 소나무·잣나무가 늦게 시드는 줄을 안다',
    '자한편 28장', hanmun_size=48), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (6/10)', 0, 0,
    '克己復禮爲仁',
    '극기복례위인',
    '자기를 이기고 예(禮)로 돌아감이 인(仁)이다',
    '안연편 1장', hanmun_size=70), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (7/10)', 0, 0,
    '己所不欲  勿施於人',
    '기소불욕   물시어인',
    '자기가 원하지 않는 것을 남에게 행하지 말라',
    '위령공편 24장', hanmun_size=60), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (8/10)', 0, 0,
    '君子和而不同  小人同而不和',
    '군자화이부동   소인동이불화',
    '군자는 조화롭되 같지 않고, 소인은 같되 조화롭지 않다',
    '자로편 23장', hanmun_size=44), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (9/10)', 0, 0,
    '過猶不及',
    '과유불급',
    '지나친 것은 미치지 못함과 같다',
    '선진편 16장', hanmun_size=110), 'Ⅵ. 명구절'))

SLIDES.append((make_quote_slide('Ⅵ. 명구절 (10/10)', 0, 0,
    '志士仁人  有殺身以成仁',
    '지사인인   유살신이성인',
    '뜻있는 선비와 어진 사람은 목숨을 바쳐서라도 인(仁)을 이룬다',
    '위령공편 9장', hanmun_size=48), 'Ⅵ. 명구절'))


# ---------------------- Ⅶ. 구조적 특징 ----------------------

@S('Ⅶ. 구조')
def s_style(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅶ. 구조', page, total)
    add_title(slide, '논어의 문체 — 네 가지 형식')
    styles = [
        ('子曰 (자왈)', '공자의 직접 발언 — 가장 기본적인 형식'),
        ('문답체',     '제자의 질문 + 공자의 답 (안연·자로편 등에 집중)'),
        ('서사체',     '사건과 일화를 서술 (미자·향당편)'),
        ('제자 발언',  '제자의 어록 — 자장(子張)편은 공자 발언이 없는 유일한 편'),
    ]
    top = 2.4
    for tag, desc in styles:
        add_filled_rect(slide, Inches(0.8), Inches(top), Inches(3.0), Inches(0.9), ACCENT)
        add_textbox(slide, Inches(0.8), Inches(top + 0.2), Inches(3.0), Inches(0.5),
                    tag, font_size=20, bold=True,
                    color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
        add_filled_rect(slide, Inches(4.0), Inches(top), Inches(8.7), Inches(0.9), PALE)
        add_textbox(slide, Inches(4.2), Inches(top + 0.2), Inches(8.3), Inches(0.5),
                    desc, font_size=16, color=INK)
        top += 1.05


@S('Ⅶ. 구조')
def s_ren_pattern(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅶ. 구조', page, total)
    add_title(slide, '인(仁)의 등장 패턴',
              '논어의 핵심 개념이 어떻게 분포되어 있는가')
    # 큰 숫자 (109회)
    add_textbox(slide, Inches(0.5), Inches(2.3), Inches(6.0), Inches(2.0),
                '109', font_size=150, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(4.5), Inches(6.0), Inches(0.5),
                '논어 전체에서 仁이 등장하는 횟수',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.1), Inches(6.0), Inches(0.5),
                '— 단연 최다 핵심어',
                font_size=14, color=SUB, align=PP_ALIGN.CENTER, bold=True)
    # 우측 설명
    add_paragraphs(slide, Inches(7.0), Inches(2.5), Inches(6.0), Inches(4.5), [
        ('집중되는 편', {'bold': True, 'font_size': 20, 'color': ACCENT}),
        ('안연편 · 이인편 — 인(仁) 논의의 본진',
         {'font_size': 16}),
        ('', {'font_size': 8}),
        ('등장하지 않는 편', {'bold': True, 'font_size': 20, 'color': ACCENT, 'space_before': 10}),
        ('계씨편 — 仁이 단 한 번도 나오지 않는다',
         {'font_size': 16}),
        ('', {'font_size': 8}),
        ('의의', {'bold': True, 'font_size': 20, 'color': ACCENT, 'space_before': 10}),
        ('제자마다 다른 인의 정의 — 인재시교(因材施敎)',
         {'font_size': 16}),
        ('상대에 따라 다르게 설명한다는 교육 원리의 증거',
         {'font_size': 14, 'color': SUB}),
    ], line_spacing=1.35)


@S('Ⅶ. 구조')
def s_sumisanggwan(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅶ. 구조', page, total)
    add_title(slide, '수미상관(首尾相關) — 닫힌 구조')
    # 시작
    add_filled_rect(slide, Inches(0.7), Inches(2.4), Inches(5.9), Inches(4.5), PALE)
    add_textbox(slide, Inches(0.7), Inches(2.6), Inches(5.9), Inches(0.5),
                '첫 머리 — 학이편 1장', font_size=18, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.7), Inches(3.3), Inches(5.9), Inches(1.5),
                '學而時習之\n不亦君子乎',
                font_size=28, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.7), Inches(5.0), Inches(5.9), Inches(1.5),
                '배우고 때때로 익히면\n이 또한 군자가 아니겠는가',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    # 끝
    add_filled_rect(slide, Inches(6.8), Inches(2.4), Inches(5.9), Inches(4.5), PALE)
    add_textbox(slide, Inches(6.8), Inches(2.6), Inches(5.9), Inches(0.5),
                '맨 끝 — 요왈편 3장', font_size=18, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(6.8), Inches(3.3), Inches(5.9), Inches(1.5),
                '不知命\n無以爲君子也',
                font_size=28, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(6.8), Inches(5.0), Inches(5.9), Inches(1.5),
                '천명(天命)을 알지 못하면\n군자가 될 수 없다',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    # 하단 결론
    add_textbox(slide, Inches(0.5), Inches(7.05), Inches(12.3), Inches(0.4),
                '배움(學)에서 출발 → 군자(君子)로 귀결 — 처음과 끝이 호응하는 닫힌 구조',
                font_size=15, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)


# ---------------------- Ⅷ. 현대적 의의 ----------------------

def make_modern_slide(title, hanmun, eum, modern_title, lines):
    def renderer(slide, page, total):
        set_white_background(slide)
        add_page_header(slide, 'Ⅷ. 현대적 의의', page, total)
        add_title(slide, title)
        # 좌측: 원전
        add_filled_rect(slide, Inches(0.7), Inches(2.3), Inches(5.4), Inches(4.6), PALE)
        add_textbox(slide, Inches(0.7), Inches(2.5), Inches(5.4), Inches(0.5),
                    '원전(原典)', font_size=14, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(0.7), Inches(3.3), Inches(5.4), Inches(1.5),
                    hanmun, font_size=32, bold=True, color=ACCENT,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(0.7), Inches(5.0), Inches(5.4), Inches(0.5),
                    eum, font_size=15, color=SUB, align=PP_ALIGN.CENTER)
        # 우측: 현대 적용
        add_textbox(slide, Inches(6.5), Inches(2.5), Inches(6.4), Inches(0.5),
                    modern_title, font_size=20, bold=True, color=INK)
        add_rule(slide, Inches(6.5), Inches(3.0), Inches(6.0), weight=1.5)
        add_paragraphs(slide, Inches(6.5), Inches(3.2), Inches(6.4), Inches(3.5),
                       lines, line_spacing=1.4, font_size=16)
    return renderer


SLIDES.append((make_modern_slide(
    '현대 ① — 화이부동(和而不同)',
    '和 而 不 同', '화이부동 (자로 23장)',
    '다원주의 사회의 원리',
    [
        ('서로 다른 종교·문화·가치관이', {'font_size': 17}),
        ('공존하는 다양성 속의 조화', {'font_size': 17, 'bold': True}),
        ('', {'font_size': 8}),
        ('획일적 동일성이 아니라', {'font_size': 16, 'color': SUB, 'space_before': 10}),
        ('차이를 인정하면서 어울리는 사회의 원리', {'font_size': 16, 'color': SUB}),
    ]), 'Ⅷ. 현대적 의의'))

SLIDES.append((make_modern_slide(
    '현대 ② — 기소불욕 물시어인',
    '己所不欲\n勿施於人', '위령공 24장',
    '보편 윤리 · 황금률',
    [
        ('동서양을 관통하는 황금률(Golden Rule)', {'font_size': 17, 'bold': True}),
        ('', {'font_size': 8}),
        ('적용 영역', {'font_size': 16, 'color': ACCENT, 'bold': True, 'space_before': 10}),
        ('• 세계인권선언 정신의 동양 원형', {'font_size': 16}),
        ('• 기업 윤리 강령의 핵심', {'font_size': 16}),
        ('• AI 윤리 가이드라인의 기초 원리', {'font_size': 16}),
    ]), 'Ⅷ. 현대적 의의'))

SLIDES.append((make_modern_slide(
    '현대 ③ — 학이불사즉망',
    '學而不思則罔\n思而不學則殆', '위정 15장',
    '비판적 사고 (Critical Thinking)',
    [
        ('정보 과잉의 시대에 요구되는', {'font_size': 17}),
        ('암기가 아닌 비판적 사고와 성찰', {'font_size': 17, 'bold': True}),
        ('', {'font_size': 8}),
        ('생성형 AI 시대에 더욱 중요해진', {'font_size': 16, 'color': SUB, 'space_before': 10}),
        ('"배움 + 사유"의 균형이라는 통찰', {'font_size': 16, 'color': SUB}),
    ]), 'Ⅷ. 현대적 의의'))

SLIDES.append((make_modern_slide(
    '현대 ④ — 덕치와 솔선수범',
    '其身正\n不令而行', '자로 6장',
    '리더십 · 수기치인(修己治人)',
    [
        ('서번트 리더십 · 윤리적 리더십과 상통', {'font_size': 17, 'bold': True}),
        ('', {'font_size': 8}),
        ('핵심 원리', {'font_size': 16, 'color': ACCENT, 'bold': True, 'space_before': 10}),
        ('• 권위는 명령이 아니라 인격에서 나온다', {'font_size': 16}),
        ('• 자기를 다스린 자만이 남을 다스릴 수 있다', {'font_size': 16}),
        ('— 修己治人 (수기치인)', {'font_size': 14, 'color': SUB}),
    ]), 'Ⅷ. 현대적 의의'))

SLIDES.append((make_modern_slide(
    '현대 ⑤ — 仁, AI 시대의 인문학',
    '仁  者\n愛  人', '안연 22장',
    '인공지능 시대, "사람다움"의 좌표',
    [
        ('기술만으로 해결할 수 없는', {'font_size': 17}),
        ('"사람다움"의 방향을 제시', {'font_size': 17, 'bold': True}),
        ('', {'font_size': 8}),
        ('AI가 무엇을 할 수 있는가가 아니라', {'font_size': 16, 'color': SUB, 'space_before': 10}),
        ('사람이 사람에게 무엇이어야 하는가', {'font_size': 16, 'color': SUB}),
        ('— 2,500년 전의 물음이 다시 현재형', {'font_size': 14, 'color': ACCENT, 'bold': True, 'space_before': 6}),
    ]), 'Ⅷ. 현대적 의의'))


# ---------------------- Ⅸ. 비교 ----------------------

@S('Ⅸ. 비교')
def s_compare(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅸ. 비교', page, total)
    add_title(slide, '논어 · 맹자 · 순자 — 유가 삼대 경전 비교')
    # 표 헤더
    cols = [('항목', 2.2), ('논어', 3.3), ('맹자', 3.3), ('순자', 3.3)]
    rows = [
        ('시대',       '춘추 말기',         '전국 중기',          '전국 말기'),
        ('형식',       '어록·문답',         '대화·논변',          '논설문(論說文)'),
        ('인성론',     '성상근(性相近)',    '성선설(性善說)',     '성악설(性惡說)'),
        ('핵심 덕목',  '인(仁)',            '인의(仁義)',         '예(禮)'),
        ('수양 방법',  '학·극기복례',       '존심양성·확충',      '화성기위·적(積)'),
        ('천(天)관',   '도덕적 천',         '의지적 천',          '자연적 천'),
        ('정치론',     '덕치·정명',         '왕도·역성혁명',      '왕도·예법 병용'),
    ]
    # 헤더
    x = 0.5
    top = 1.95
    add_filled_rect(slide, Inches(x), Inches(top), Inches(2.2), Inches(0.55), ACCENT)
    add_textbox(slide, Inches(x), Inches(top + 0.1), Inches(2.2), Inches(0.4),
                '항목', font_size=15, bold=True,
                color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
    for i, name in enumerate(['논어', '맹자', '순자']):
        cx = 0.5 + 2.2 + i * 3.45
        color = ACCENT if name == '논어' else SUB
        add_filled_rect(slide, Inches(cx), Inches(top), Inches(3.45), Inches(0.55), color)
        add_textbox(slide, Inches(cx), Inches(top + 0.1), Inches(3.45), Inches(0.4),
                    name, font_size=16, bold=True,
                    color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
    # 데이터 행
    row_h = 0.62
    for r_idx, row in enumerate(rows):
        y = top + 0.55 + r_idx * row_h
        bg = RGBColor(0xFA, 0xFA, 0xFA) if r_idx % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF)
        # 항목 열
        add_filled_rect(slide, Inches(0.5), Inches(y), Inches(2.2), Inches(row_h), PALE)
        add_textbox(slide, Inches(0.55), Inches(y + 0.12), Inches(2.1), Inches(0.4),
                    row[0], font_size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)
        for i in range(3):
            cx = 0.5 + 2.2 + i * 3.45
            add_filled_rect(slide, Inches(cx), Inches(y), Inches(3.45), Inches(row_h), bg)
            text_color = ACCENT if i == 0 else INK
            add_textbox(slide, Inches(cx + 0.05), Inches(y + 0.12),
                        Inches(3.35), Inches(0.4),
                        row[i + 1], font_size=14, color=text_color,
                        bold=(i == 0), align=PP_ALIGN.CENTER)


# ---------------------- Ⅹ. 마무리 ----------------------

@S('Ⅹ. 마무리')
def s_summary(slide, page, total):
    set_white_background(slide)
    add_page_header(slide, 'Ⅹ. 마무리', page, total)
    add_title(slide, '한 문장으로 정리하는 논어')
    add_filled_rect(slide, Inches(0.7), Inches(2.3), Inches(11.9), Inches(4.5), PALE)
    add_paragraphs(slide, Inches(1.1), Inches(2.7), Inches(11.1), Inches(4.1), [
        ('배움(學)에서 출발하여', {'font_size': 22, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 8}),
        ('인(仁)을 핵심으로', {'font_size': 22, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER, 'space_before': 8}),
        ('예(禮)를 형식으로', {'font_size': 22, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('충서(忠恕)를 방법으로', {'font_size': 22, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('군자(君子)를 이상으로 삼아', {'font_size': 22, 'bold': True, 'color': ACCENT, 'align': PP_ALIGN.CENTER}),
        ('', {'font_size': 8}),
        ('수기(修己)에서 치인(治人)으로 나아가는', {'font_size': 22, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 8}),
        ('', {'font_size': 8}),
        ('— 사람다움의 총체적 설계도 —', {'font_size': 22, 'bold': True, 'color': INK, 'align': PP_ALIGN.CENTER, 'space_before': 8}),
    ], line_spacing=1.2)


@S('Ⅹ. 마무리')
def s_closing(slide, page, total):
    set_white_background(slide)
    add_filled_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.4), RULE)
    # 명언
    add_textbox(slide, Inches(0.5), Inches(2.0), Inches(12.3), Inches(1.5),
                '半 部 論 語 治 天 下',
                font_size=66, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(3.6), Inches(12.3), Inches(0.5),
                '반 부 논 어 치 천 하', font_size=22, color=SUB,
                align=PP_ALIGN.CENTER)
    add_rule(slide, Inches(5.5), Inches(4.4), Inches(2.3), color=RULE, weight=1.5)
    add_textbox(slide, Inches(0.5), Inches(4.7), Inches(12.3), Inches(0.6),
                '"논어 반 권으로 천하를 다스린다"',
                font_size=24, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(5.4), Inches(12.3), Inches(0.5),
                '— 송(宋) 재상 조보(趙普)의 말',
                font_size=15, color=SUB, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
                '감사합니다', font_size=28, bold=True, color=INK,
                align=PP_ALIGN.CENTER)
    add_filled_rect(slide, Inches(0), Inches(7.1), Inches(13.333), Inches(0.4), RULE)


# ---------------------- 빌드 ----------------------

total_pages = len(SLIDES)
for i, (renderer, section) in enumerate(SLIDES, 1):
    slide = prs.slides.add_slide(blank)
    renderer(slide, i, total_pages)

out_path = r'C:\Users\박호군\ClaudeProjects\Oriental-Classics2\발표자료\논어_발표자료.pptx'
prs.save(out_path)
print(f'Saved: {out_path}  ({total_pages} slides)')
