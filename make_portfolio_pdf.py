# -*- coding: utf-8 -*-
"""Generate Portfolio_SangilNa.pdf using ReportLab"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os, glob

# ── fonts ──────────────────────────────────────────────────────────
WIN_FONTS = r"C:\Windows\Fonts"

def reg(name, regular, bold=None, italic=None):
    path = os.path.join(WIN_FONTS, regular)
    if os.path.exists(path):
        pdfmetrics.registerFont(TTFont(name, path))
    if bold:
        p = os.path.join(WIN_FONTS, bold)
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(name + "-Bold", p))
    if italic:
        p = os.path.join(WIN_FONTS, italic)
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(name + "-Italic", p))

reg("Malgun", "malgun.ttf",   "malgunbd.ttf")
reg("NanumG", "NanumGothic.ttf", "NanumGothicBold.ttf")

# pick best available CJK font
for fn in ("Malgun", "NanumG"):
    try:
        pdfmetrics.getFont(fn)
        BODY_FONT  = fn
        BOLD_FONT  = fn + "-Bold"
        break
    except Exception:
        pass
else:
    BODY_FONT = BOLD_FONT = "Helvetica"

# ── colours ────────────────────────────────────────────────────────
NAVY  = colors.HexColor("#1a2744")
RED   = colors.HexColor("#c0392b")
GRAY  = colors.HexColor("#7f8c8d")
LIGHT = colors.HexColor("#f5f6fa")
WHITE = colors.white
MID   = colors.HexColor("#2c3e50")

# ── styles ─────────────────────────────────────────────────────────
def S(name, **kw):
    base = kw.pop("base", "Normal")
    ss   = getSampleStyleSheet()
    b    = ss[base]
    return ParagraphStyle(
        name,
        parent     = b,
        fontName   = kw.pop("fontName",   BODY_FONT),
        fontSize   = kw.pop("fontSize",   9),
        leading    = kw.pop("leading",    13),
        textColor  = kw.pop("textColor",  colors.black),
        alignment  = kw.pop("alignment",  TA_LEFT),
        spaceAfter = kw.pop("spaceAfter", 2),
        **kw
    )

sTitle    = S("sTitle",    fontName=BOLD_FONT, fontSize=26, leading=32,
               textColor=NAVY, alignment=TA_CENTER, spaceAfter=4)
sSubtitle = S("sSubtitle", fontName=BODY_FONT, fontSize=11, leading=16,
               textColor=MID,  alignment=TA_CENTER, spaceAfter=4)
sContact  = S("sContact",  fontName=BODY_FONT, fontSize=8,  leading=12,
               textColor=GRAY, alignment=TA_CENTER, spaceAfter=8)
sSecHead  = S("sSecHead",  fontName=BOLD_FONT, fontSize=13, leading=18,
               textColor=NAVY, spaceAfter=6)
sSubHead  = S("sSubHead",  fontName=BOLD_FONT, fontSize=9.5, leading=13,
               textColor=NAVY, spaceAfter=2)
sPeriod   = S("sPeriod",   fontName=BODY_FONT, fontSize=8.5, leading=12,
               textColor=RED,  alignment=TA_RIGHT, spaceAfter=0)
sRole     = S("sRole",     fontName=BODY_FONT, fontSize=8.5, leading=12,
               textColor=GRAY, spaceAfter=4)
sBody     = S("sBody",     fontName=BODY_FONT, fontSize=8.5, leading=13,
               textColor=colors.HexColor("#333333"), spaceAfter=3)
sBullet   = S("sBullet",   fontName=BODY_FONT, fontSize=8.3, leading=12.5,
               textColor=colors.HexColor("#444444"), leftIndent=12,
               bulletIndent=2, spaceAfter=2)
sPubLabel = S("sPubLabel", fontName=BOLD_FONT, fontSize=9, leading=13,
               textColor=WHITE, spaceAfter=4)
sPubText  = S("sPubText",  fontName=BODY_FONT, fontSize=8.3, leading=13,
               textColor=colors.HexColor("#333333"), spaceAfter=3)
sSmall    = S("sSmall",    fontName=BODY_FONT, fontSize=8, leading=12,
               textColor=GRAY, spaceAfter=2)

def bold(text):
    return f'<font name="{BOLD_FONT}">{text}</font>'
def red(text):
    return f'<font color="#c0392b">{text}</font>'
def gray(text):
    return f'<font color="#7f8c8d">{text}</font>'

# ── helpers ────────────────────────────────────────────────────────
def section_header(title):
    return [
        Spacer(1, 10),
        Paragraph(title, sSecHead),
        HRFlowable(width="100%", thickness=2, color=RED, spaceAfter=8),
    ]

def tl_entry(org, period, role, bullets):
    row = Table(
        [[Paragraph(bold(org), sSubHead), Paragraph(red(period), sPeriod)]],
        colWidths=[None, 60*mm]
    )
    row.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP"),
                              ("LEFTPADDING", (0,0), (-1,-1), 0),
                              ("RIGHTPADDING", (0,0), (-1,-1), 0),
                              ("TOPPADDING", (0,0), (-1,-1), 0),
                              ("BOTTOMPADDING", (0,0), (-1,-1), 0)]))
    elems = [row, Paragraph(role, sRole)]
    for b in bullets:
        elems.append(Paragraph(f"• {b}", sBullet))
    elems.append(Spacer(1, 6))
    return KeepTogether(elems)

def pub_block(label_text, label_bg, items):
    label = Table(
        [[Paragraph(label_text, sPubLabel)]],
        colWidths=[None]
    )
    label.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), label_bg),
        ("LEFTPADDING",  (0,0), (-1,-1), 8),
        ("RIGHTPADDING", (0,0), (-1,-1), 8),
        ("TOPPADDING",   (0,0), (-1,-1), 4),
        ("BOTTOMPADDING",(0,0), (-1,-1), 4),
        ("ROUNDEDCORNERS", (0,0), (-1,-1), 4),
    ]))
    elems = [label, Spacer(1, 4)]
    for it in items:
        elems.append(Paragraph(it, sPubText))
        elems.append(HRFlowable(width="100%", thickness=0.5,
                                color=colors.HexColor("#e0e4ec"), spaceAfter=4))
    return elems

# ── BUILD DOC ──────────────────────────────────────────────────────
out_path = os.path.join(os.path.dirname(__file__), "Portfolio_SangilNa.pdf")
doc = SimpleDocTemplate(
    out_path,
    pagesize        = A4,
    topMargin       = 18*mm,
    bottomMargin    = 18*mm,
    leftMargin      = 20*mm,
    rightMargin     = 20*mm,
    title           = "Portfolio – Sangil Na",
    author          = "Sangil Na",
)

story = []

# ── HEADER BLOCK ───────────────────────────────────────────────────
header_data = [[
    Paragraph("나상일  |  Sangil Na", sTitle),
]]
header_tbl = Table(header_data, colWidths=[doc.width])
header_tbl.setStyle(TableStyle([
    ("BACKGROUND",    (0,0), (-1,-1), NAVY),
    ("TOPPADDING",    (0,0), (-1,-1), 14),
    ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ("LEFTPADDING",   (0,0), (-1,-1), 10),
    ("RIGHTPADDING",  (0,0), (-1,-1), 10),
]))
story.append(header_tbl)

subtitle_tbl = Table([[
    Paragraph("Senior Research Engineer · AI &amp; Computer Vision for Smart Infrastructure", sSubtitle)
]], colWidths=[doc.width])
subtitle_tbl.setStyle(TableStyle([
    ("BACKGROUND",    (0,0), (-1,-1), MID),
    ("TOPPADDING",    (0,0), (-1,-1), 6),
    ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ("LEFTPADDING",   (0,0), (-1,-1), 10),
    ("RIGHTPADDING",  (0,0), (-1,-1), 10),
]))
story.append(subtitle_tbl)

contact_tbl = Table([[
    Paragraph(
        "(+82) 10-3158-8461  |  billnin@naver.com  |  Sungkyunkwan University M.S. (2026)  |  Suwon, South Korea",
        sContact)
]], colWidths=[doc.width])
contact_tbl.setStyle(TableStyle([
    ("BACKGROUND",    (0,0), (-1,-1), LIGHT),
    ("TOPPADDING",    (0,0), (-1,-1), 6),
    ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ("LEFTPADDING",   (0,0), (-1,-1), 10),
    ("RIGHTPADDING",  (0,0), (-1,-1), 10),
]))
story.append(contact_tbl)
story.append(Spacer(1, 6))

# ── EXPERIENCE ─────────────────────────────────────────────────────
story += section_header("Experience")

story.append(tl_entry(
    "SMARTINSIDE AI", "Mar. 2026 – Present",
    "Senior Research Engineer · AI Research Team · Suwon, South Korea",
    [
        f"{bold('Advanced Structural Dynamics Analysis:')} CV 기반 보도육교 동적 거동 분석 및 TMD 효율 평가 연구 주도",
        f"{bold('Vision-based Cable Force Estimation:')} 사장교 stay-cable 비접촉 장력 추정 방법론 개발",
        f"{bold('AI-Integrated Maintenance:')} VLM 기반 교통 인프라 유지관리 자동화 연구 주도",
        f"{bold('SCI 논문:')} Mechanical Systems and Signal Processing 등 제1저자 논문 투고 중",
    ]
))

story.append(tl_entry(
    "BRITEC", "Mar. 2021 – Dec. 2023",
    "Assistant Manager (대리) · R&D Division · Suwon, South Korea",
    [
        f"{bold('Structural & Mechanical Design:')} AutoCAD 2D/3D 활용 보도교 난간·점검로·TMD 설계, 재설치 리스크 최소화",
        f"{bold('New Product Development:')} 짚라인 구조 장비·철도 궤도 침하 방지 장치 등 5종 이상 신제품 개발 (기획→납품)",
        f"{bold('Vibration Performance Optimization:')} 현장 계측 데이터 기반 TMD 최적 설계 → 보행자 사용성 향상",
        f"{bold('Technical Visualization:')} AutoCAD 3D·Photoshop 3D 렌더링 제품 이미지·기술 문서 단독 제작",
        f"{bold('Patent:')} 방호 울타리 아두이노 초음파 센서 시스템 설계 → 특허 등록 (No.10-2021-0090084, 주발명자)",
    ]
))

# ── EDUCATION ──────────────────────────────────────────────────────
story += section_header("Education")

story.append(tl_entry(
    "성균관대학교 (Sungkyunkwan University)", "Mar. 2024 – Feb. 2026",
    "M.S. · Department of Global Smart City · Suwon, South Korea",
    [
        "GPA: 4.25 / 4.5",
        "석사학위논문: 비접촉 카메라 기반 영상 증폭을 통한 진동 저감 장치 성능 평가 기술",
        "Advisor: Prof. Seunghee Park / Smart Construction Information Technology Lab",
    ]
))
story.append(tl_entry(
    "경기대학교 (Kyonggi University)", "Mar. 2018 – Aug. 2020",
    "B.S. · Department of Mechanical System Engineering · Suwon, South Korea",
    ["GPA: 3.45 / 4.5"]
))
story.append(tl_entry(
    "세명대학교 (Semyung University)", "Mar. 2013 – Feb. 2018",
    "Undergraduate Coursework · Department of Bioenvironmental Engineering · Jecheon, South Korea",
    ["GPA: 3.33 / 4.5"]
))

# ── RESEARCH PROJECTS ──────────────────────────────────────────────
story += section_header("Research Projects")

projects = [
    ("원전 지진응답 예측을 위한 AI 알고리즘 자문",
     "Jul. 2025 – Dec. 2025",
     "Korea Electric Power Corporation (KEPCO) / 성균관대학교 · 참여연구원",
     [
        "SASSI·STRATA 통합 자동 데이터 파이프라인 구축 → 5,000+ 지진 응답 데이터 세트 생성",
        "데이터 처리 기간 30일 → 1주일 미만으로 단축 (VS Code 기반 프로세스 자동화)",
        "CNN·Logistic Regression·Attention 모델 비교: 가속도 응답→CNN, FRS→Attention 모델이 최적",
     ]),
    ("인공지능 플러스 K건설인프라 레질리언스 연구센터 (ERC)",
     "Feb. 2025 – Jul. 2025",
     "National Research Foundation of Korea (NRF) / 성균관대학교 · 참여연구원",
     [
        "AI 융합 인프라 레질리언스 ERC 제안서 기술 프레임워크 설계 및 제안서 본문 작성",
        "Photoshop·AutoCAD 3D 인포그래픽 제작으로 복잡한 연구 개념 시각화",
        "서면 심사 통과 후 대면 평가 PT 참여 → 성균관대학교 ERC 센터 최종 선정",
     ]),
    ("초장대 K-지하고속도로 인프라 안전 및 효율 향상 기술 개발",
     "Jul. 2024 – Dec. 2024",
     "Ministry of Land, Infrastructure and Transport / 성균관대학교 · 참여연구원",
     [
        "지하고속도로 수직구 시공 중 안전 위험 요소 파악 및 국내외 기술 동향 분석",
        "PC 세그먼트 수직도 모니터링용 틸트 센서 선정·조달·연동 (아두이노 기반 계측 시스템)",
        "Lab-scale 실험 환경 및 데이터 수집 인프라 구축 → 차년도 연구 기반 마련",
     ]),
    ("작업자 행동 기반 안전 모니터링 및 예측 기술 개발",
     "Mar. 2024 – Dec. 2024",
     "Ministry of the Interior and Safety (행정안전부) / 성균관대학교 · 참여연구원",
     [
        "비시스템 비계 현장 작업자 행동 모니터링 실험 보조 — 센서 부착·촬영 장비 세팅·데이터 라벨링",
        "최종 평가 보고서 및 발표자료 제작 참여 → 정부 과제 성공적 종료",
        "병행 개인연구(비계 AI 동적 거동 모니터링)로 국내 학술대회 발표",
     ]),
]
for org, period, role, bullets in projects:
    story.append(tl_entry(org, period, role, bullets))

# ── PUBLICATIONS ───────────────────────────────────────────────────
story += section_header("Publications")

# SCIE Submitted
story += pub_block(
    "SCIE / SSCI — Submitted",
    colors.HexColor("#243b6e"),
    [
        f"[1] {bold('Sangil Na')}, Minsoo Park, Dongyoung Ko, Seunghee Park†. "
        "High-Precision Non-Contact Dynamic Performance Evaluation of Tuned Mass Dampers for Footbridges via "
        "Smartphone-Based Eulerian Video Magnification. "
        f"{gray('<i>Mechanical Systems and Signal Processing</i>')}, Submitted 2026.03.12. "
        f"{red('[1st Author]')}",

        f"[2] {bold('Sangil Na')}, Minsoo Park, Dongyoung Ko, Seunghee Park†. "
        "Objective Frequency-Band Selection for Eulerian Video Magnification-Based Non-Contact Dynamic "
        "Assessment of TMD-Equipped Footbridges. "
        f"{gray('<i>Journal of Computing in Civil Engineering</i>')}, Under Review. "
        f"{red('[1st Author]')}",
    ]
)
story.append(Spacer(1, 6))

# SCIE Published
story += pub_block(
    "SCIE / SSCI — Published",
    colors.HexColor("#1a5276"),
    [
        "Dongyoung Ko, Minsoo Park, " + bold("Sangil Na") + ", Juyoung Jang, Yong Kwon Cho, Seunghee Park†. "
        "Computer vision-based steel strand tension evaluation. "
        + gray("<i>Engineering Applications of Artificial Intelligence</i>")
        + ", Vol.166, Part B, No.113718, pp.1–17, Jan 2026.",
    ]
)
story.append(Spacer(1, 6))

# KCI Published
story += pub_block(
    "KCI — Published (국내 학술지)",
    colors.HexColor("#1e8449"),
    [
        bold("Sangil Na") + ", Woonggyu Choi, Seungwoo Kim, Seunghee Park†. "
        "영상 증폭 기반 비접촉 계측을 통한 보도교 진동 사용성 평가. "
        + gray("<i>한국구조물진단유지관리공학회 논문집</i>")
        + ", vol.29, no.5, pp.124–131, 2025.10. " + red("[1st Author]"),

        "Seungwoo Kim, Woonggyu Choi, " + bold("Sangil Na") + ", Seunghee Park†. "
        "Computer Vision 기반 도심지 침수 영역 면적 추정에 관한 연구. "
        + gray("<i>한국구조물진단유지관리공학회 논문집</i>")
        + ", vol.29, no.6, pp.49–57, 2025.10.",
    ]
)
story.append(Spacer(1, 6))

# Thesis
story += pub_block(
    "Thesis (석사학위논문)",
    colors.HexColor("#6c3483"),
    [
        bold("나상일") + ". 비접촉 카메라 기반 영상 증폭을 통한 진동 저감 장치 성능 평가 기술. "
        + gray("<i>성균관대학교 석사학위논문</i>") + ", 2026.02.25. " + red("[1st Author]"),
    ]
)
story.append(Spacer(1, 6))

# Conference
story += pub_block(
    "Conference — Domestic (국내 학술대회)",
    colors.HexColor("#784212"),
    [
        bold("Sangil Na") + ", Minsoo Park, Dongyoung Ko, Woonggyu Choi, Kyeunghoon Cheon, Seunghee Park†. "
        "스마트폰 카메라 기반 EVM을 활용한 보도육교 및 TMD 성능 평가. "
        + gray("<i>한국전산구조공학회 2026년도 정기학술대회</i>") + ", 2026.04.17. " + red("[1st Author]"),

        "Kyongmin Kim, " + bold("Sangil Na") + ", Seungsoo Lee, Minsoo Park, Seunghee Park†. "
        "작업자 안전 사고 예방을 위한 지능형 자동화 Digital Twin 모니터링 시스템. "
        + gray("<i>2025 한국방재학회 학술발표대회</i>") + ", 2025.02.20.",

        "Kyongmin Kim, " + bold("Sangil Na") + ", Seongwoo Son, Seungsoo Lee, Minsoo Park, Seunghee Park†. "
        "IMU 센서와 딥러닝을 통한 건설현장 안전고리 체결 유무 판단. "
        + gray("<i>KSCE 2024 CONVENTION (대한토목학회)</i>") + ", 2024.10.18.",

        bold("Sangil Na") + ", Minsoo Park, Seungsoo Lee, Seongwoo Son, Kyoungmin Kim, Seunghee Park†. "
        "비계 붕괴 사고 예방을 위한 AI 기반 동적 거동 모니터링 시스템 개발. "
        + gray("<i>KSCE 2024 CONVENTION (대한토목학회)</i>") + ", 2024.10.17. " + red("[1st Author]"),

        "Seongwoo Son, Sujin Jin, " + bold("Sangil Na") + ", Jaeyeon Won, Seunghee Park†. "
        "단일 IMU센서를 활용한 딥러닝 기반 건설현장 작업자 아차사고 분석. "
        + gray("<i>한국구조물진단유지관리공학회 2024 가을 학술발표회</i>") + ", 2024.10.01.",
    ]
)
story.append(Spacer(1, 6))

# ── PATENTS & IP ───────────────────────────────────────────────────
story += section_header("Patents & Intellectual Property")

def ip_table(headers, rows):
    col_data = [headers] + rows
    col_w = [16*mm, None, 36*mm, 20*mm, 28*mm]
    t = Table(col_data, colWidths=col_w, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0),  NAVY),
        ("TEXTCOLOR",     (0,0), (-1,0),  WHITE),
        ("FONTNAME",      (0,0), (-1,0),  BOLD_FONT),
        ("FONTSIZE",      (0,0), (-1,-1), 7.5),
        ("FONTNAME",      (0,1), (-1,-1), BODY_FONT),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [WHITE, LIGHT]),
        ("GRID",          (0,0), (-1,-1), 0.5, colors.HexColor("#e0e4ec")),
        ("VALIGN",        (0,0), (-1,-1), "TOP"),
        ("TOPPADDING",    (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("LEFTPADDING",   (0,0), (-1,-1), 6),
        ("RIGHTPADDING",  (0,0), (-1,-1), 6),
    ]))
    return t

story.append(Paragraph(bold("특허 (Patents)"), sSubHead))
story.append(Spacer(1, 4))
story.append(ip_table(
    ["상태", "명칭", "출원번호", "날짜", "역할"],
    [
        ["출원", "현장 감시 카메라를 이용한 건설 현장 안전 감시 장치 및 방법",
         "10-2025-0050707", "2025.04.18", "공동발명자 (4/6)"],
        ["등록", "보안 기능이 강화된 방호 울타리",
         "10-2021-0090084", "2022.01.19", "주발명자 (1/2)"],
    ]
))
story.append(Spacer(1, 10))
story.append(Paragraph(bold("프로그램 등록 (Software Registration)"), sSubHead))
story.append(Spacer(1, 4))
story.append(ip_table(
    ["상태", "명칭", "등록번호", "등록일", "역할"],
    [
        ["등록", "고정형 단안 카메라를 이용한 실시간 수직구 수직도 계측 알고리즘",
         "C-2025-055245", "2025.12.05", "공동발명가 (7/12)"],
        ["등록", "건설현장 침수 안전 프로그램 (영상 분석 프로그램)",
         "C-2025-010254", "2025.03.19", "공동발명자 (5/11)"],
        ["등록", "ConvNext 모델 활용 건설현장 작업자 아차사고 행동 분류 판단 알고리즘",
         "C-2024-048437", "2024.11.27", "공동발명자 (4/10)"],
    ]
))
story.append(Spacer(1, 6))

# ── SKILLS ─────────────────────────────────────────────────────────
story += section_header("Skills")

skill_rows = [
    [
        Paragraph(bold("Programming & AI"), sSubHead),
        Paragraph(bold("Computer Vision / DL"), sSubHead),
        Paragraph(bold("Engineering / Design"), sSubHead),
    ],
    [
        Paragraph("Python (Pandas, NumPy, SciPy)\nPyTorch · OpenCV · Scikit-learn", sBody),
        Paragraph("Eulerian Video Magnification (EVM)\nOptical Flow · Co-tracker\nCNN · LSTM · Transformer\nVLM (Studying)", sBody),
        Paragraph("AutoCAD 2D/3D/Rendering (고급, 5년)\nTMD Design & SHM\nSeismic Response Prediction\nArduino 센서 시스템 설계", sBody),
    ],
    [
        Paragraph(bold("Hardware & Instrumentation"), sSubHead),
        Paragraph(bold("Software & Visualization"), sSubHead),
        Paragraph(bold("Languages"), sSubHead),
    ],
    [
        Paragraph("Acceleration Sensors (진동 계측)\nEM Sensors · Load Cells\nUltrasonic Sensors\nArduino Integration", sBody),
        Paragraph("Adobe Photoshop (중급, 3년)\nMS Office (특급, 10년+)\nOverleaf (LaTeX)\nCATIA 3D Modeling", sBody),
        Paragraph("한국어 — Native\nEnglish — TOEIC 750", sBody),
    ],
]
skill_tbl = Table(skill_rows, colWidths=[doc.width/3]*3)
skill_tbl.setStyle(TableStyle([
    ("GRID",          (0,0), (-1,-1), 0.5, colors.HexColor("#e0e4ec")),
    ("BACKGROUND",    (0,0), (-1,-1), LIGHT),
    ("BACKGROUND",    (0,0), (2,0),   colors.HexColor("#dce6f5")),
    ("BACKGROUND",    (0,2), (2,2),   colors.HexColor("#dce6f5")),
    ("VALIGN",        (0,0), (-1,-1), "TOP"),
    ("TOPPADDING",    (0,0), (-1,-1), 6),
    ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ("LEFTPADDING",   (0,0), (-1,-1), 8),
    ("RIGHTPADDING",  (0,0), (-1,-1), 8),
]))
story.append(skill_tbl)
story.append(Spacer(1, 6))

# ── CERTIFICATIONS ─────────────────────────────────────────────────
story += section_header("Certifications & Others")

cert_rows = [
    ["TOEIC 750", "2026.03", "—"],
    ["MOS Word Expert", "2016.06", "Microsoft Office Specialist"],
    ["MOS PowerPoint", "2013.12", "Microsoft Office Specialist"],
    ["MOS Excel", "2013.06", "Microsoft Office Specialist"],
    ["자동차운전면허 1종 보통", "2020.07", "서울지방경찰청"],
    ["부품제조설계실무자양성과정", "2019.06–08", "한국건설기계산업협회 · CATIA 3D (299h)"],
]
cert_tbl = Table(
    [["자격/어학", "취득일", "발급기관"]] + cert_rows,
    colWidths=[60*mm, 30*mm, None],
    repeatRows=1
)
cert_tbl.setStyle(TableStyle([
    ("BACKGROUND",    (0,0), (-1,0),  NAVY),
    ("TEXTCOLOR",     (0,0), (-1,0),  WHITE),
    ("FONTNAME",      (0,0), (-1,0),  BOLD_FONT),
    ("FONTSIZE",      (0,0), (-1,-1), 8),
    ("FONTNAME",      (0,1), (-1,-1), BODY_FONT),
    ("ROWBACKGROUNDS",(0,1), (-1,-1), [WHITE, LIGHT]),
    ("GRID",          (0,0), (-1,-1), 0.5, colors.HexColor("#e0e4ec")),
    ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
    ("TOPPADDING",    (0,0), (-1,-1), 5),
    ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ("LEFTPADDING",   (0,0), (-1,-1), 8),
    ("RIGHTPADDING",  (0,0), (-1,-1), 8),
]))
story.append(cert_tbl)

# ── BUILD ──────────────────────────────────────────────────────────
doc.build(story)
print(f"PDF saved → {out_path}")
