# -*- coding: utf-8 -*-
"""Generate Portfolio_SangilNa.pptx using python-pptx"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import os

OUT = os.path.join(os.path.dirname(__file__), "Portfolio_SangilNa.pptx")

NAVY  = RGBColor(0x1a, 0x27, 0x44)
RED   = RGBColor(0xc0, 0x39, 0x2b)
WHITE = RGBColor(0xff, 0xff, 0xff)
LIGHT = RGBColor(0xf5, 0xf6, 0xfa)
GRAY  = RGBColor(0x7f, 0x8c, 0x8d)
MID   = RGBColor(0x2c, 0x3e, 0x50)
BLUE2 = RGBColor(0x24, 0x3b, 0x6e)

W = Inches(13.33)   # widescreen 16:9
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

def blank_layout(prs):
    return prs.slide_layouts[6]   # completely blank

def add_rect(slide, x, y, w, h, fill=None, line=None, line_w=None):
    shp = slide.shapes.add_shape(1, x, y, w, h)   # MSO_SHAPE_TYPE.RECTANGLE = 1
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = line
        if line_w: shp.line.width = line_w
    else:
        shp.line.fill.background()
    return shp

def add_text(slide, text, x, y, w, h,
             font_name="Malgun Gothic", font_size=18,
             bold=False, color=RGBColor(0,0,0),
             align=PP_ALIGN.LEFT, wrap=True,
             italic=False):
    txb = slide.shapes.add_textbox(x, y, w, h)
    txb.word_wrap = wrap
    tf  = txb.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name      = font_name
    run.font.size      = Pt(font_size)
    run.font.bold      = bold
    run.font.italic    = italic
    run.font.color.rgb = color
    return txb

def add_para(tf, text, font_name="Malgun Gothic", font_size=10,
             bold=False, color=RGBColor(0x33,0x33,0x33),
             align=PP_ALIGN.LEFT, space_before=0, italic=False):
    p   = tf.add_paragraph()
    p.alignment   = align
    p.space_before = Pt(space_before)
    run = p.add_run()
    run.text = text
    run.font.name      = font_name
    run.font.size      = Pt(font_size)
    run.font.bold      = bold
    run.font.italic    = italic
    run.font.color.rgb = color
    return p

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SLIDE 1 — COVER
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
sl = prs.slides.add_slide(blank_layout(prs))

# background
add_rect(sl, 0, 0, W, H, fill=NAVY)
# red accent bar (left)
add_rect(sl, 0, 0, Inches(.18), H, fill=RED)
# bottom accent band
add_rect(sl, 0, Inches(6.2), W, Inches(1.3), fill=BLUE2)

# Name
add_text(sl, "나상일  |  Sangil Na",
         Inches(.6), Inches(1.6), Inches(12), Inches(1.2),
         font_size=52, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
# subtitle
add_text(sl, "Senior Research Engineer  ·  AI & Computer Vision for Smart Infrastructure",
         Inches(.6), Inches(2.9), Inches(12), Inches(.7),
         font_size=17, color=RGBColor(0x8f, 0xa8, 0xd4), align=PP_ALIGN.CENTER)

# divider
add_rect(sl, Inches(3.5), Inches(3.8), Inches(6.3), Inches(.04), fill=RED)

# contact
add_text(sl, "(+82) 10-3158-8461   |   billnin@naver.com   |   Suwon, South Korea",
         Inches(.6), Inches(4.0), Inches(12), Inches(.5),
         font_size=12, color=RGBColor(0xcd, 0xd2, 0xe0), align=PP_ALIGN.CENTER)

# tags
tags = ["Computer Vision", "Structural Health Monitoring", "Eulerian Video Magnification",
        "Deep Learning", "TMD Design", "Python · PyTorch · OpenCV", "AutoCAD 2D/3D"]
x_pos = Inches(.5)
for tag in tags:
    w_tag = Inches(1.6)
    box = add_rect(sl, x_pos, Inches(4.9), w_tag, Inches(.38),
                   fill=RGBColor(0x2c, 0x48, 0x7a))
    add_text(sl, tag, x_pos + Inches(.06), Inches(4.93), w_tag - Inches(.1), Inches(.32),
             font_size=9, color=RGBColor(0xdc, 0xe6, 0xf5), align=PP_ALIGN.CENTER)
    x_pos += w_tag + Inches(.08)

# Sungkyunkwan University footer
add_text(sl, "Smart Construction Information Technology Lab  |  Sungkyunkwan University  |  SMARTINSIDE AI",
         Inches(.6), Inches(6.35), Inches(12), Inches(.5),
         font_size=11, color=RGBColor(0x8a, 0x9b, 0xbf), align=PP_ALIGN.CENTER)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# helper: section title slide
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def section_divider(title, subtitle=""):
    sl = prs.slides.add_slide(blank_layout(prs))
    add_rect(sl, 0, 0, W, H, fill=NAVY)
    add_rect(sl, 0, 0, W, Inches(.12), fill=RED)
    add_rect(sl, 0, H - Inches(.12), W, Inches(.12), fill=RED)
    add_text(sl, title,
             Inches(1), Inches(2.5), Inches(11.3), Inches(1.5),
             font_size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    if subtitle:
        add_text(sl, subtitle,
                 Inches(1), Inches(4.1), Inches(11.3), Inches(.8),
                 font_size=16, color=RGBColor(0x8f,0xa8,0xd4), align=PP_ALIGN.CENTER)

# helper: content slide with left navy bar
def content_slide(title):
    sl = prs.slides.add_slide(blank_layout(prs))
    add_rect(sl, 0, 0, W, H, fill=LIGHT)
    add_rect(sl, 0, 0, W, Inches(.7), fill=NAVY)
    add_rect(sl, 0, Inches(.7), Inches(.06), H - Inches(.7), fill=RED)
    add_text(sl, title,
             Inches(.2), Inches(.08), Inches(13), Inches(.55),
             font_size=20, bold=True, color=WHITE)
    return sl

# helper: add a card box
def card(sl, x, y, w, h, title, period, body_lines):
    add_rect(sl, x, y, w, h, fill=WHITE)
    add_rect(sl, x, y, w, Inches(.05), fill=RED)
    add_text(sl, title,
             x + Inches(.15), y + Inches(.12), w - Inches(.3), Inches(.45),
             font_size=11, bold=True, color=NAVY)
    add_text(sl, period,
             x + Inches(.15), y + Inches(.58), w - Inches(.3), Inches(.3),
             font_size=9, color=RED, italic=True)
    ty = y + Inches(.88)
    for line in body_lines:
        add_text(sl, "• " + line,
                 x + Inches(.15), ty, w - Inches(.3), Inches(.28),
                 font_size=8.5, color=RGBColor(0x44,0x44,0x44))
        ty += Inches(.27)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SLIDE 2 — EXPERIENCE (1) SMARTINSIDE AI
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
section_divider("Experience", "경력 사항")

sl = content_slide("Experience  —  SMARTINSIDE AI")
# company card
add_rect(sl, Inches(.3), Inches(.85), Inches(5.6), Inches(6.4), fill=WHITE)
add_rect(sl, Inches(.3), Inches(.85), Inches(5.6), Inches(.06), fill=RED)
add_text(sl, "SMARTINSIDE AI",
         Inches(.5), Inches(.95), Inches(5.2), Inches(.55),
         font_size=20, bold=True, color=NAVY)
add_text(sl, "Mar. 2026 – Present",
         Inches(.5), Inches(1.52), Inches(5.2), Inches(.35),
         font_size=11, color=RED, italic=True)
add_text(sl, "Senior Research Engineer  ·  AI Research Team  ·  Suwon, South Korea",
         Inches(.5), Inches(1.9), Inches(5.2), Inches(.4),
         font_size=10, color=GRAY)

bullets_si = [
    ("Advanced Structural Dynamics", "CV 기반 보도육교 동적 거동 분석 및 TMD 효율 평가 연구 주도"),
    ("Vision-based Cable Force", "사장교 stay-cable 비접촉 장력 추정 방법론 개발"),
    ("AI-Integrated Maintenance", "VLM 기반 교통 인프라 유지관리 자동화 연구 수행 중"),
    ("SCI 논문 투고", "Mechanical Systems and Signal Processing 등 제1저자 논문 투고 중"),
]
ty = Inches(2.4)
for label, detail in bullets_si:
    add_text(sl, "▸  " + label,
             Inches(.55), ty, Inches(4.8), Inches(.3),
             font_size=10, bold=True, color=NAVY)
    add_text(sl, "    " + detail,
             Inches(.55), ty + Inches(.28), Inches(4.8), Inches(.28),
             font_size=9, color=RGBColor(0x44,0x44,0x44))
    ty += Inches(.64)

# right panel — key research areas
add_rect(sl, Inches(6.3), Inches(.85), Inches(6.7), Inches(6.4), fill=BLUE2)
add_text(sl, "Key Research Areas",
         Inches(6.5), Inches(1.0), Inches(6.3), Inches(.5),
         font_size=14, bold=True, color=WHITE)

areas = [
    ("Computer Vision", "Eulerian Video Magnification (EVM)\nOptical Flow · Co-tracker · Segmentation"),
    ("Structural Health Monitoring", "TMD Performance Evaluation\nNon-contact Dynamic Sensing"),
    ("AI / Deep Learning", "CNN · LSTM · Transformer\nVision-Language Model (VLM)"),
    ("Cable-stayed Bridge", "Stay-cable Tension Estimation\nInfrastructure Maintenance Decision Support"),
]
ay = Inches(1.65)
for title_a, detail_a in areas:
    add_rect(sl, Inches(6.5), ay, Inches(6.3), Inches(.04), fill=RED)
    add_text(sl, title_a, Inches(6.5), ay + Inches(.1), Inches(6.3), Inches(.35),
             font_size=11, bold=True, color=WHITE)
    add_text(sl, detail_a, Inches(6.5), ay + Inches(.44), Inches(6.3), Inches(.45),
             font_size=9, color=RGBColor(0xcd,0xd2,0xe0))
    ay += Inches(1.2)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SLIDE 3 — EXPERIENCE (2) BRITEC
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
sl = content_slide("Experience  —  BRITEC")

add_rect(sl, Inches(.3), Inches(.85), Inches(4.1), Inches(6.4), fill=WHITE)
add_rect(sl, Inches(.3), Inches(.85), Inches(4.1), Inches(.06), fill=RED)
add_text(sl, "BRITEC",
         Inches(.5), Inches(.95), Inches(3.8), Inches(.5),
         font_size=20, bold=True, color=NAVY)
add_text(sl, "Mar. 2021 – Dec. 2023",
         Inches(.5), Inches(1.48), Inches(3.8), Inches(.35),
         font_size=11, color=RED, italic=True)
add_text(sl, "Assistant Manager (대리)  ·  R&D  ·  Suwon",
         Inches(.5), Inches(1.85), Inches(3.8), Inches(.35),
         font_size=9.5, color=GRAY)

britec_items = [
    ("구조 & 기계 설계", "AutoCAD 2D/3D 보도교 난간·점검로·TMD 설계\n재설치 리스크 최소화"),
    ("신제품 개발", "짚라인 구조 장비·철도 침하 방지 장치 등\n5종+ 신제품 기획→납품 전 과정 주도"),
    ("진동 사용성 평가", "현장 계측 데이터 기반 TMD 최적 설계\n보행자 사용성 향상 보고서 작성"),
    ("기술 시각화·문서화", "AutoCAD 3D 렌더링·Photoshop 단독 제작\n기업 카탈로그·나라장터 이미지 등록"),
    ("특허 취득", "방호 울타리 아두이노 초음파 센서 시스템\n특허 등록 No.10-2021-0090084 (주발명자)"),
]
ty = Inches(2.3)
for lbl, det in britec_items:
    add_text(sl, "▸  " + lbl,
             Inches(.55), ty, Inches(3.5), Inches(.28),
             font_size=9.5, bold=True, color=NAVY)
    add_text(sl, det,
             Inches(.75), ty + Inches(.28), Inches(3.3), Inches(.38),
             font_size=8.5, color=RGBColor(0x44,0x44,0x44))
    ty += Inches(.74)

# right: skill icons / highlights
for xi, (icon, lbl, val) in enumerate([
    ("🔧", "AutoCAD", "2D/3D/Rendering\n고급 (5년)"),
    ("🎨", "Photoshop", "Technical Viz\n중급 (3년)"),
    ("🔌", "Arduino", "Sensor Integration\nCircuit Design"),
    ("📐", "TMD Design", "Mass/Spring Calc\nSite Measurement"),
    ("📄", "Documentation", "Tech Reports\n나라장터 Catalog"),
    ("🏆", "Patent", "No.10-2021-0090084\n주발명자"),
]):
    col = xi % 3
    row = xi // 3
    bx = Inches(4.8) + col * Inches(2.8)
    by = Inches(1.0) + row * Inches(3.0)
    add_rect(sl, bx, by, Inches(2.5), Inches(2.5), fill=BLUE2)
    add_text(sl, icon,         bx + Inches(.9), by + Inches(.2),  Inches(.8), Inches(.5), font_size=22, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(sl, lbl,          bx + Inches(.1), by + Inches(.8),  Inches(2.3), Inches(.4), font_size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(sl, val,          bx + Inches(.1), by + Inches(1.25), Inches(2.3), Inches(.8), font_size=9,  color=RGBColor(0xcd,0xd2,0xe0), align=PP_ALIGN.CENTER)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SLIDE 4 — EDUCATION
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
section_divider("Education", "학력")

sl = content_slide("Education")
edu = [
    ("성균관대학교\nSungkyunkwan Univ.",
     "Mar. 2024 – Feb. 2026",
     "M.S. · Dept. of Global Smart City",
     ["GPA: 4.25 / 4.5",
      "석사학위논문: 비접촉 카메라 기반 영상 증폭을\n통한 진동 저감 장치 성능 평가 기술",
      "Advisor: Prof. Seunghee Park\nSmart Construction Information Tech. Lab"]),
    ("경기대학교\nKyonggi University",
     "Mar. 2018 – Aug. 2020",
     "B.S. · Dept. of Mechanical System Engineering",
     ["GPA: 3.45 / 4.5"]),
    ("세명대학교\nSemyung University",
     "Mar. 2013 – Feb. 2018",
     "Undergraduate · Bioenvironmental Engineering",
     ["GPA: 3.33 / 4.5"]),
]
col_w = Inches(4.1)
for i, (univ, period, dept, items) in enumerate(edu):
    bx = Inches(.3) + i * (col_w + Inches(.2))
    by = Inches(.9)
    bh = Inches(6.3)
    add_rect(sl, bx, by, col_w, bh, fill=WHITE)
    add_rect(sl, bx, by, col_w, Inches(.06), fill=RED)
    # rank badge
    add_rect(sl, bx + Inches(.12), by + Inches(.18), Inches(.5), Inches(.5),
             fill=NAVY)
    add_text(sl, str(i+1), bx + Inches(.12), by + Inches(.18), Inches(.5), Inches(.5),
             font_size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(sl, univ,
             bx + Inches(.72), by + Inches(.18), col_w - Inches(.9), Inches(.65),
             font_size=11, bold=True, color=NAVY)
    add_text(sl, period,
             bx + Inches(.12), by + Inches(.9), col_w - Inches(.3), Inches(.3),
             font_size=9.5, color=RED, italic=True)
    add_text(sl, dept,
             bx + Inches(.12), by + Inches(1.2), col_w - Inches(.3), Inches(.35),
             font_size=9, color=GRAY)
    add_rect(sl, bx + Inches(.1), by + Inches(1.58), col_w - Inches(.2), Inches(.02), fill=RGBColor(0xe0,0xe4,0xec))
    ty2 = by + Inches(1.7)
    for item in items:
        add_text(sl, "• " + item,
                 bx + Inches(.15), ty2, col_w - Inches(.3), Inches(.55),
                 font_size=9, color=RGBColor(0x33,0x33,0x33))
        ty2 += Inches(.55)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SLIDE 5 — RESEARCH PROJECTS (1)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
section_divider("Research Projects", "연구 과제 참여")

sl = content_slide("Research Projects")
# 4 cards 2x2
proj_data = [
    ("원전 지진응답 예측\nAI 알고리즘 자문",
     "Jul. 2025 – Dec. 2025",
     "KEPCO / 한국전력기술",
     ["SASSI·STRATA 통합 자동 파이프라인 → 5,000+ 데이터 세트 생성",
      "처리 기간 30일+ → 1주일 미만 단축 (프로세스 자동화)",
      "CNN·Attention 등 모델 비교 → 최적 모델 확인"]),
    ("인공지능 플러스\nK건설인프라 ERC",
     "Feb. 2025 – Jul. 2025",
     "NRF 한국연구재단",
     ["ERC 제안서 기술 프레임워크 설계 및 본문 작성",
      "복잡한 연구 개념 인포그래픽 시각화 (Photoshop·AutoCAD)",
      "서면 통과 → 대면 평가 PT → 성균관대학교 ERC 최종 선정"]),
    ("초장대 K-지하고속도로\n인프라 안전·효율 향상",
     "Jul. 2024 – Dec. 2024",
     "국토교통부 / 한국건설기술연구원",
     ["수직구 시공 중 작업자 안전 위험 요소 파악 및 기술 동향 분석",
      "PC 세그먼트 수직도 모니터링 틸트 센서 선정·연동 (Arduino)",
      "Lab-scale 실험 환경 구축 → 차년도 연구 기반 마련"]),
    ("작업자 행동 기반\n안전 모니터링·예측 기술",
     "Mar. 2024 – Dec. 2024",
     "행정안전부",
     ["비계 현장 실험 보조: 센서 부착·촬영 세팅·데이터 라벨링",
      "최종 평가 보고서 및 발표자료 제작 참여",
      "병행 개인연구(AI 기반 비계 모니터링)로 학술대회 발표"]),
]
for i, (title, period, org, bullets) in enumerate(proj_data):
    col = i % 2
    row = i // 2
    bx = Inches(.3) + col * Inches(6.4)
    by = Inches(.95) + row * Inches(3.1)
    bw = Inches(6.2)
    bh = Inches(2.9)
    add_rect(sl, bx, by, bw, bh, fill=WHITE)
    add_rect(sl, bx, by, bw, Inches(.06), fill=RED)
    add_text(sl, title,
             bx + Inches(.15), by + Inches(.1), bw - Inches(.3), Inches(.65),
             font_size=12, bold=True, color=NAVY)
    add_text(sl, period + "  |  " + org,
             bx + Inches(.15), by + Inches(.77), bw - Inches(.3), Inches(.28),
             font_size=9, color=RED, italic=True)
    ty = by + Inches(1.1)
    for b in bullets:
        add_text(sl, "• " + b,
                 bx + Inches(.2), ty, bw - Inches(.4), Inches(.35),
                 font_size=8.8, color=RGBColor(0x33,0x33,0x33))
        ty += Inches(.52)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SLIDE 6 — PUBLICATIONS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
section_divider("Publications", "논문 실적")

sl = content_slide("Publications  —  Journal Papers")

pub_groups = [
    ("SCIE Submitted", BLUE2,
     ["[S1]  Sangil Na, et al.  High-Precision Non-Contact Dynamic Performance Evaluation of TMDs via Smartphone-Based EVM.\n"
      "Mechanical Systems and Signal Processing  (Submitted 2026.03.12)  [1st Author]",
      "[S2]  Sangil Na, et al.  Objective Frequency-Band Selection for EVM-Based Non-Contact Dynamic Assessment of TMD-Equipped Footbridges.\n"
      "Journal of Computing in Civil Engineering  (Under Review)  [1st Author]"]),
    ("SCIE Published", RGBColor(0x1a,0x52,0x76),
     ["[P1]  Dongyoung Ko, Minsoo Park, Sangil Na, Juyoung Jang, Yong Kwon Cho, Seunghee Park†.\n"
      "Computer vision-based steel strand tension evaluation.\n"
      "Engineering Applications of Artificial Intelligence, Vol.166, Part B, Jan 2026."]),
    ("KCI Published", RGBColor(0x1e,0x84,0x49),
     ["[K1]  Sangil Na, et al.  영상 증폭 기반 비접촉 계측을 통한 보도교 진동 사용성 평가.\n"
      "한국구조물진단유지관리공학회 논문집, vol.29, no.5, pp.124–131, 2025.10.  [1st Author]",
      "[K2]  Seungwoo Kim, Woonggyu Choi, Sangil Na, Seunghee Park†.\n"
      "Computer Vision 기반 도심지 침수 영역 면적 추정에 관한 연구.\n"
      "한국구조물진단유지관리공학회, vol.29, no.6, pp.49–57, 2025.10."]),
    ("M.S. Thesis", RGBColor(0x6c,0x34,0x83),
     ["나상일.  비접촉 카메라 기반 영상 증폭을 통한 진동 저감 장치 성능 평가 기술.\n"
      "성균관대학교 석사학위논문, 2026.02.25  [1st Author]"]),
]

ty = Inches(0.88)
for label, bg, items in pub_groups:
    add_rect(sl, Inches(.3), ty, Inches(2.2), Inches(.38), fill=bg)
    add_text(sl, label,
             Inches(.38), ty + Inches(.04), Inches(2.0), Inches(.3),
             font_size=9.5, bold=True, color=WHITE)
    ty += Inches(.42)
    for item in items:
        add_rect(sl, Inches(.3), ty, Inches(12.7), Inches(.04), fill=RGBColor(0xe0,0xe4,0xec))
        add_text(sl, item,
                 Inches(.5), ty + Inches(.06), Inches(12.4), Inches(.55),
                 font_size=8.5, color=RGBColor(0x33,0x33,0x33))
        ty += Inches(.68)
    ty += Inches(.1)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SLIDE 7 — CONFERENCE
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
sl = content_slide("Publications  —  Conference Papers (국내 학술대회)")

conf_items = [
    ("2026.04", "한국전산구조공학회 2026", "스마트폰 카메라 기반 EVM을 활용한 보도육교 및 TMD 성능 평가", "1st Author"),
    ("2025.02", "한국방재학회 2025", "작업자 안전 사고 예방을 위한 지능형 자동화 Digital Twin 모니터링 시스템", "Co-Author"),
    ("2024.10", "KSCE 2024 CONVENTION", "IMU 센서와 딥러닝을 통한 건설현장 안전고리 체결 유무 판단", "Co-Author"),
    ("2024.10", "KSCE 2024 CONVENTION", "비계 붕괴 사고 예방을 위한 AI 기반 동적 거동 모니터링 시스템 개발", "1st Author"),
    ("2024.10", "한국구조물진단유지관리공학회 2024", "단일 IMU센서를 활용한 딥러닝 기반 건설현장 작업자 아차사고 분석", "Co-Author"),
]
ty = Inches(0.88)
for date, conf, title, role in conf_items:
    add_rect(sl, Inches(.3), ty, Inches(12.7), Inches(.78), fill=WHITE)
    add_rect(sl, Inches(.3), ty, Inches(.06), Inches(.78), fill=RED)
    # date badge
    add_rect(sl, Inches(.4), ty + Inches(.12), Inches(1.1), Inches(.55), fill=NAVY)
    add_text(sl, date,
             Inches(.4), ty + Inches(.18), Inches(1.1), Inches(.4),
             font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(sl, conf,
             Inches(1.65), ty + Inches(.06), Inches(5.0), Inches(.3),
             font_size=9, color=GRAY)
    add_text(sl, title,
             Inches(1.65), ty + Inches(.35), Inches(8.8), Inches(.35),
             font_size=10, bold=True, color=NAVY)
    bg_r = RED if role == "1st Author" else GRAY
    add_rect(sl, Inches(11.0), ty + Inches(.17), Inches(1.8), Inches(.38), fill=bg_r)
    add_text(sl, role,
             Inches(11.0), ty + Inches(.22), Inches(1.8), Inches(.3),
             font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    ty += Inches(.88)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SLIDE 8 — PATENTS & IP
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
section_divider("Patents & Intellectual Property", "특허 및 지식재산권")

sl = content_slide("Patents & Intellectual Property")

ip_items = [
    ("특허 출원", "10-2025-0050707", "현장 감시 카메라를 이용한 건설 현장 안전 감시 장치 및 방법", "2025.04.18", "공동발명자 (4/6)", False),
    ("특허 등록", "10-2021-0090084", "보안 기능이 강화된 방호 울타리", "2022.01.19", "주발명자 (1/2)", True),
    ("프로그램 등록", "C-2025-055245", "고정형 단안 카메라를 이용한 실시간 수직구 수직도 계측 알고리즘", "2025.12.05", "공동발명가 (7/12)", False),
    ("프로그램 등록", "C-2025-010254", "건설현장 침수 안전 프로그램 (영상 분석 프로그램)", "2025.03.19", "공동발명자 (5/11)", False),
    ("프로그램 등록", "C-2024-048437", "ConvNext 모델 활용 건설현장 작업자 아차사고 행동 분류 판단 알고리즘", "2024.11.27", "공동발명자 (4/10)", False),
]

headers = ["구분", "번호", "명칭", "날짜", "역할"]
col_ws  = [Inches(1.5), Inches(2.0), Inches(6.0), Inches(1.5), Inches(2.0)]

tx = Inches(.3)
ty = Inches(.9)
row_h = Inches(.38)

# header row
hx = tx
for j, (hdr, cw) in enumerate(zip(headers, col_ws)):
    add_rect(sl, hx, ty, cw, row_h, fill=NAVY)
    add_text(sl, hdr, hx + Inches(.08), ty + Inches(.06), cw - Inches(.1), row_h - Inches(.08),
             font_size=10, bold=True, color=WHITE)
    hx += cw
ty += row_h

for k, (kind, num, name, date, role, highlight) in enumerate(ip_items):
    bg = RGBColor(0xfd,0xec,0xea) if highlight else (WHITE if k % 2 == 0 else LIGHT)
    hx = tx
    for j, (val, cw) in enumerate(zip([kind, num, name, date, role], col_ws)):
        add_rect(sl, hx, ty, cw, row_h, fill=bg)
        fc = RED if highlight and j == 0 else RGBColor(0x33,0x33,0x33)
        add_text(sl, val, hx + Inches(.08), ty + Inches(.04), cw - Inches(.1), row_h - Inches(.06),
                 font_size=9, color=fc, bold=(highlight and j == 4))
        hx += cw
    ty += row_h

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SLIDE 9 — SKILLS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
section_divider("Skills", "기술 역량")

sl = content_slide("Skills")

skill_cards = [
    ("Programming & AI", "Python (Pandas, NumPy, SciPy)\nPyTorch · OpenCV · Scikit-learn\nVS Code Environment"),
    ("Computer Vision / DL", "Eulerian Video Magnification (EVM)\nOptical Flow · Co-tracker\nCNN · LSTM · Transformer\nVLM (Studying)"),
    ("Engineering / Design", "AutoCAD 2D/3D/Rendering\nTMD Design & SHM\nSeismic Response Prediction\nArduino 센서 시스템"),
    ("Hardware & Sensing", "Acceleration Sensors (진동 계측)\nEM Sensors · Load Cells\nUltrasonic Sensors\nArduino Integration"),
    ("Software & Viz", "Adobe Photoshop (중급, 3년)\nMS Office (특급, 10년+)\nOverleaf (LaTeX)\nCATIA 3D Modeling"),
    ("Languages", "한국어 — Native\nEnglish — TOEIC 750"),
]
card_w = Inches(4.1)
card_h = Inches(2.8)
for i, (title_s, body_s) in enumerate(skill_cards):
    col = i % 3
    row = i // 3
    sx = Inches(.3) + col * (card_w + Inches(.1))
    sy = Inches(.88) + row * (card_h + Inches(.14))
    add_rect(sl, sx, sy, card_w, card_h, fill=WHITE)
    add_rect(sl, sx, sy, card_w, Inches(.06), fill=RED)
    add_rect(sl, sx, sy + Inches(.06), card_w, Inches(.5), fill=NAVY)
    add_text(sl, title_s,
             sx + Inches(.12), sy + Inches(.1), card_w - Inches(.24), Inches(.38),
             font_size=11, bold=True, color=WHITE)
    add_text(sl, body_s,
             sx + Inches(.12), sy + Inches(.65), card_w - Inches(.24), Inches(2.0),
             font_size=9.5, color=RGBColor(0x33,0x33,0x33))

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SLIDE 10 — CERTIFICATIONS & CLOSING
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
sl = content_slide("Certifications & Summary")

# left cert list
add_rect(sl, Inches(.3), Inches(.88), Inches(6.2), Inches(6.3), fill=WHITE)
add_rect(sl, Inches(.3), Inches(.88), Inches(6.2), Inches(.06), fill=RED)
add_text(sl, "Certifications", Inches(.5), Inches(.96), Inches(5.8), Inches(.45),
         font_size=14, bold=True, color=NAVY)

certs = [
    ("TOEIC 750", "2026.03", ""),
    ("MOS Word Expert", "2016.06", "Microsoft Office Specialist"),
    ("MOS PowerPoint", "2013.12", "Microsoft Office Specialist"),
    ("MOS Excel", "2013.06", "Microsoft Office Specialist"),
    ("자동차운전면허 1종 보통", "2020.07", "서울지방경찰청"),
    ("부품제조설계실무자양성과정", "2019.06–08", "CATIA 3D · 299시간 · 한국건설기계산업협회"),
]
ty = Inches(1.5)
for nm, dt, det in certs:
    add_rect(sl, Inches(.5), ty, Inches(5.9), Inches(.68), fill=LIGHT)
    add_text(sl, nm,   Inches(.65), ty + Inches(.06), Inches(3.5), Inches(.3), font_size=10, bold=True, color=NAVY)
    add_text(sl, dt,   Inches(.65), ty + Inches(.36), Inches(1.2), Inches(.25), font_size=8.5, color=RED)
    add_text(sl, det,  Inches(2.0), ty + Inches(.36), Inches(4.2), Inches(.25), font_size=8.5, color=GRAY)
    ty += Inches(.76)

# right: summary stats
add_rect(sl, Inches(6.9), Inches(.88), Inches(6.1), Inches(6.3), fill=BLUE2)
add_text(sl, "Research Summary", Inches(7.1), Inches(1.0), Inches(5.7), Inches(.45),
         font_size=14, bold=True, color=WHITE)

stats = [
    ("2", "SCIE 논문 투고 (제1저자)"),
    ("1", "SCIE 논문 게재"),
    ("2", "KCI 논문 게재"),
    ("5", "국내 학술대회 발표"),
    ("4", "정부 연구 과제 참여"),
    ("5", "특허·프로그램 등록"),
]
ty = Inches(1.6)
for num, lbl in stats:
    add_rect(sl, Inches(7.1), ty, Inches(1.0), Inches(.7), fill=RED)
    add_text(sl, num, Inches(7.1), ty + Inches(.1), Inches(1.0), Inches(.5),
             font_size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(sl, lbl, Inches(8.25), ty + Inches(.18), Inches(4.5), Inches(.35),
             font_size=11, color=WHITE)
    ty += Inches(.82)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CLOSING SLIDE
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
sl = prs.slides.add_slide(blank_layout(prs))
add_rect(sl, 0, 0, W, H, fill=NAVY)
add_rect(sl, 0, 0, W, Inches(.12), fill=RED)
add_rect(sl, 0, H - Inches(.12), W, Inches(.12), fill=RED)
add_text(sl, "Thank you",
         Inches(1), Inches(1.8), Inches(11.3), Inches(1.6),
         font_size=52, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text(sl, "나상일  |  Sangil Na",
         Inches(1), Inches(3.5), Inches(11.3), Inches(.7),
         font_size=22, color=RGBColor(0x8f,0xa8,0xd4), align=PP_ALIGN.CENTER)
add_rect(sl, Inches(4.0), Inches(4.3), Inches(5.3), Inches(.04), fill=RED)
add_text(sl, "(+82) 10-3158-8461   |   billnin@naver.com",
         Inches(1), Inches(4.5), Inches(11.3), Inches(.45),
         font_size=13, color=RGBColor(0xcd,0xd2,0xe0), align=PP_ALIGN.CENTER)
add_text(sl, "Smart Construction Information Technology Lab  |  Sungkyunkwan University",
         Inches(1), Inches(5.1), Inches(11.3), Inches(.4),
         font_size=11, color=RGBColor(0x8a,0x9b,0xbf), align=PP_ALIGN.CENTER)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
prs.save(OUT)
print(f"PPTX saved → {OUT}")
