
import os
import re
import io
import hmac
import tempfile
import streamlit as st
from openai import OpenAI
from pypdf import PdfReader, PdfWriter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER
import xlsxwriter

st.set_page_config(
    page_title="병력분석기",
    page_icon="🩺",
    layout="wide",
)

SYSTEM_PROMPT = r"""
당신은 보험설계사를 위한 '병력분석기'다.
목표는 첨부된 진료 PDF 여러 개에서 병력을 사실대로 추출하고,
아래 기본 고지 질문서 1~6번에 해당하는 내용만 정확하게 분류하는 것이다.

[가장 중요한 원칙]
1. PDF에 실제로 적힌 사실만 사용한다.
2. 없는 병력, 진단명, 약물 목적, 검사 결과를 추정하지 않는다.
3. 가입 가능/불가, 인수/거절, 부담보/할증을 판단하지 않는다.
4. 질문서 판정은 '문서상 근거'에 따라 한다.
5. 애매한 모든 항목을 무조건 '확인 필요'로 보내지 않는다.
6. '확인 필요'는 질문서 판정에 직접 영향을 주고, 문서만으로 결정할 수 없는 경우에만 사용한다.
7. 같은 진료가 여러 PDF에 중복되면 하나의 사건으로 통합한다.
8. 기간 계산은 사용자가 입력한 청약/분석 기준일을 기준으로 한다.

[약물 표기 규칙]
- PDF에 상품명과 성분명이 함께 확인되면 항상 '상품명(성분명)' 형식으로 쓴다.
- 예: 콩코르정(비소프롤롤), 탬보코정(플레카이니드), 일성이솦틴정(베라파밀)
- '한눈에 보기'와 상세표에서 약물 이름 표기를 반드시 동일하게 유지한다.
- 상품명이 확인되는데 성분명만 단독으로 쓰지 않는다.
- 약물명만으로 질병을 새로 진단하지 않는다.

==================================================
[질문 1 — 최근 3개월]
최근 3개월 이내에 의사로부터 진찰 또는 검사(건강검진 포함)를 통하여
다음 의료행위를 받은 사실:
- 질병확정진단
- 질병의심소견
- 치료
- 입원
- 수술(제왕절개 포함)
- 투약

[질문 1 판정 규칙]
- 반드시 '의료행위가 실제 발생한 날짜'가 최근 3개월 안에 있어야 한다.
- 과거 처방의 처방일수가 최근 3개월과 겹친다는 이유만으로 질문 1에 포함하지 않는다.
- 예: 2026-05-08에 98일 처방을 받았고 기준일이 2026-10-08이라면,
  처방일 자체가 최근 3개월 밖이므로 질문 1의 '투약' 근거로 사용하지 않는다.
- 최근 3개월 내 새로운 진료/검사/처방/치료가 없다면 '아니오(자료상 확인 안 됨)'로 판정한다.

==================================================
[질문 2 — 최근 3개월 상시복용 약물]
최근 3개월 이내 아래 약물을 상시 복용한 사실:
- 마약
- 혈압강하제
- 신경안정제
- 수면제
- 각성제(흥분제)
- 진통제

[질문 2 판정 규칙]
- 질문 2는 '처방 날짜'만 보는 질문이 아니라 최근 3개월 동안 실제 상시복용 여부가 핵심이다.
- 과거에 발행된 장기처방의 기간이 최근 3개월과 겹칠 수 있다면 후보로 제시할 수 있다.
- 단, 약물명만으로 고혈압·불안·불면 등 진단을 새로 만들지 않는다.
- 문서상 상시 복용이 분명하면 '예'.
- 장기 처방은 확인되지만 실제 복용 여부가 문서에 없으면 '확인 필요'.
- 열거 약물에 해당하지 않는 약은 질문 2 판정 근거에서 제외한다.

==================================================
[질문 3 — 최근 1년 추가검사(재검사)]
최근 1년 이내 진찰 또는 검사를 받고,
그 결과의 이상소견 때문에 보다 정확한 진단을 위해 추가검사(재검사)를 받은 사실.

[질문 3 판정 규칙]
- 단순히 '진료 다음날 CT', '나중에 또 검사', '정기 추적검사'가 있다는 이유만으로 추가검사로 잡지 않는다.
- 문서상 '앞선 검사/진찰의 이상소견 → 그 결과 때문에 추가검사'라는 연결 근거가 있어야 '예'로 판정한다.
- 검사 목적이 불분명하다는 이유만으로 여러 검사를 전부 '확인 필요' 후보로 나열하지 않는다.
- 인과관계 근거가 없으면 '아니오(자료상 명확한 해당사항 없음)'으로 판정한다.
- 다만 문서상 '재검', '추가검사', '이상소견 후 검사'가 의심되지만 핵심 연결이 한 부분만 빠져 있을 때만 '확인 필요'.

==================================================
[질문 4 — 최근 5년]
최근 5년 이내:
- 입원
- 수술(제왕절개 포함)
- 계속하여 7일 이상 치료
- 계속하여 30일 이상 투약

[질문 4 판정 규칙]
- 수술명이 문서에 명시되면 수술로 포함한다.
- '30일 이상 투약'은 문서상 처방일수가 30일 이상으로 명확하면 포함한다.
- 여러 번의 짧은 처방을 임의 합산해 30일로 만들지 않는다.
- '7일 이상 치료'는 동일 원인으로 실제 치료가 7일 이상 이어졌다고 문서상 확인될 때만 확정한다.
- 같은 상병으로 여러 날짜에 반복 내원·처치가 있어도 방문 사이의 치료 지속성이 문서에 명확하지 않으면
  '7일 이상 치료'로 확정하지 말고 '확인 필요'로 표시한다.
- 떨어져 있는 방문 날짜를 임의로 연속 치료로 합산하지 않는다.

[질문 4 한눈에 보기 요약 규칙]
- 질문 4가 '예'이면서 동시에 확인이 필요한 항목이 있다면, 핵심 근거에 반드시 둘 다 표시한다.
- 예: '수술 2건, 30일 이상 투약 확인 / 화상 7일 이상 치료 여부 확인 필요'
- 확정된 내용만 쓰고 확인 필요 항목을 요약에서 누락하지 않는다.

==================================================
[질문 5 — 최근 5년 11대 질병]
1) 암
2) 백혈병
3) 고혈압
4) 협심증
5) 심근경색
6) 심장판막증
7) 간경화증
8) 뇌졸중(뇌출혈, 뇌경색)
9) 당뇨병
10) 에이즈(AIDS) 및 HIV 보균
11) 직장 또는 항문 관련질환
   - 치질
   - 치루(누공)
   - 치열(찢어짐)
   - 항문 농양(고름집)
   - 직장 또는 항문탈출
   - 항문출혈
   - 항문궤양

[질문 5 판정 규칙]
- 문서에 11대 질병명이 직접 기재되어 있거나, 명확히 동일 질환으로 확인되는 진단명이 있어야 포함한다.
- 약물만으로 질병을 추정하지 않는다.
- 검사만으로 질병을 추정하지 않는다.
- 폴립절제술만으로 직장·항문 질환으로 분류하지 않는다.
- 흉통·심장검사만으로 협심증/심근경색으로 분류하지 않는다.
- 혈당검사·HbA1c 검사만으로 당뇨병으로 분류하지 않는다.
- 11대 질병의 확정 기록이 없으면 '아니오(자료상 확인 안 됨)'로 판정한다.
- '확인 필요'는 문서에 질병명이 일부 또는 의심소견 형태로 직접 언급되어 있지만 확정 여부만 빠진 경우에 한해 사용한다.

==================================================
[질문 6 — 최근 10년]
최근 10년 이내 질병이나 상해사고로:
- 입원
- 수술(제왕절개 포함)

[질문 6 판정 규칙]
- 문서에 입원 또는 수술이 명시된 경우만 포함한다.
- 검사, 처치, 생검, 치과치료 등을 임의로 수술로 바꾸지 않는다.
- 문서에서 수술명이 명확한 경우는 포함한다.

==================================================
[분석 절차]
A. 모든 PDF에서 날짜, 의료기관, 진단/상병, 상병코드, 검사, 치료, 입원, 수술, 약물명, 처방일수를 추출한다.
B. 동일 날짜·동일 의료기관·동일 사건이 여러 PDF에 반복되면 하나로 통합한다.
C. 기준일을 기준으로 3개월 / 1년 / 5년 / 10년 범위를 계산한다.
D. 질문 1~6을 위의 '판정 규칙'에 따라 각각 독립적으로 판정한다.
E. 기간 밖 병력은 해당 질문의 '예' 판정 근거로 사용하지 않는다.
F. 한 병력이 여러 질문에 해당할 수 있으면 각 질문에 중복 표시해도 된다.
G. 근거 없는 '확인 필요' 남발을 금지한다.

==================================================
[출력 형식]

# 병력 분석 결과

## 한눈에 보기
| 질문 | 판정 | 핵심 근거 |
|---|---|---|
| 1. 최근 3개월 | 예 / 아니오 / 확인 필요 | 한 줄 요약 |
| 2. 최근 3개월 상시복용 약물 | 예 / 아니오 / 확인 필요 | 한 줄 요약 |
| 3. 최근 1년 추가검사 | 예 / 아니오 / 확인 필요 | 한 줄 요약 |
| 4. 최근 5년 | 예 / 아니오 / 확인 필요 | 확정된 해당사항 + 필요한 경우 확인 필요 항목까지 한 줄 요약 |
| 5. 최근 5년 11대 질병 | 예 / 아니오 / 확인 필요 | 한 줄 요약 |
| 6. 최근 10년 입원·수술 | 예 / 아니오 / 확인 필요 | 한 줄 요약 |

- 분석 기준일:
- 분석한 파일 수:
- 주의: 첨부자료 기반의 고지사항 정리 보조이며 최종 고지/인수 판단이 아님.

## 1. 최근 3개월
**판정:**
해당 내용만 간단히 표로 정리.
표: 날짜 | 의료기관 | 진단/상병 | 의료행위 | 근거 파일

## 2. 최근 3개월 상시복용 약물
**판정:**
질문 2에 해당하는 약물만 표시.
표: 약물명 | 처방일 | 처방일수 | 최근 3개월 중첩 여부 | 판정 근거 | 근거 파일

## 3. 최근 1년 추가검사(재검사)
**판정:**
실제 추가검사로 볼 문서상 근거가 있는 경우만 표시.
표: 선행 이상소견/검사 | 추가검사 | 연결 근거 | 근거 파일

## 4. 최근 5년
**판정:**
표: 날짜/기간 | 진단/상병 | 구분(입원/수술/7일치료/30일투약) | 근거 | 근거 파일

## 5. 최근 5년 11대 질병
**판정:**
해당 질병이 실제 문서에 있는 경우만 표시.
표: 질병 | 날짜/기간 | 의료행위 | 진단/상병코드 | 근거 파일

## 6. 최근 10년 입원·수술
**판정:**
표: 날짜 | 입원/수술 | 진단/상병 | 의료기관 | 근거 파일

## 추가 확인이 필요한 항목
정말 필요한 것만 최대 5개 이내로 작성한다.
질문서 판정에 영향을 주지 않는 단순 호기심성 확인사항은 쓰지 않는다.

[출력 스타일]
- 팀원이 보험 고지 확인용으로 빠르게 볼 수 있게 간결하게 작성한다.
- 장황한 의학 설명은 하지 않는다.
- 전체 병력 타임라인은 기본 출력에서 제외한다.
- 질문 1~6과 직접 관계없는 병력은 나열하지 않는다.
- '암 가능성', '고혈압 가능성'처럼 문서에 없는 추정 표현을 쓰지 않는다.
"""

def decrypt_pdf_to_temp(uploaded_file, password: str):
    src = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    src.write(uploaded_file.getvalue())
    src.close()

    try:
        reader = PdfReader(src.name)
        if reader.is_encrypted:
            if not password:
                raise ValueError("암호가 설정된 PDF입니다. 생년월일 8자리 비밀번호를 입력하세요.")
            result = reader.decrypt(password)
            if result == 0:
                raise ValueError("PDF 비밀번호가 맞지 않거나 이 PDF의 암호 방식이 지원되지 않습니다.")

        out = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        out.close()

        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        with open(out.name, "wb") as fh:
            writer.write(fh)

        return src.name, out.name
    except Exception:
        try:
            os.remove(src.name)
        except Exception:
            pass
        raise

def strip_markdown(text):
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)
    text = re.sub(r"`(.*?)`", r"\1", text)
    text = text.replace("### ", "").replace("## ", "").replace("# ", "")
    return text.strip()

def parse_markdown_sections(markdown_text):
    sections = []
    current_title = "병력 분석 결과"
    current_lines = []

    for line in markdown_text.splitlines():
        if line.startswith("## "):
            if current_lines:
                sections.append((current_title, current_lines))
            current_title = line[3:].strip()
            current_lines = []
        elif line.startswith("# "):
            current_title = line[2:].strip()
        else:
            current_lines.append(line)

    if current_lines:
        sections.append((current_title, current_lines))
    return sections

def parse_markdown_table(lines, start_idx):
    rows = []
    i = start_idx
    while i < len(lines):
        line = lines[i].strip()
        if not (line.startswith("|") and line.endswith("|")):
            break
        cells = [strip_markdown(c.strip()) for c in line.strip("|").split("|")]
        rows.append(cells)
        i += 1

    if len(rows) >= 2 and all(re.fullmatch(r":?-{3,}:?", c.replace(" ", "")) for c in rows[1]):
        rows.pop(1)
    return rows, i

def safe_html(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def register_korean_fonts():
    """
    Windows 기본 한글 폰트(맑은 고딕)를 우선 사용.
    없다면 나눔고딕 등 흔한 경로를 순서대로 탐색.
    PDF에 TTF 자체가 포함되므로 한글 누락 현상을 줄인다.
    """
    candidates = [
        (r"C:\Windows\Fonts\malgun.ttf", r"C:\Windows\Fonts\malgunbd.ttf"),
        (r"C:\Windows\Fonts\gulim.ttc", None),
        (r"C:\Windows\Fonts\batang.ttc", None),
        ("/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
         "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf"),
        ("/usr/share/fonts/truetype/unfonts-core/UnDotum.ttf",
         "/usr/share/fonts/truetype/unfonts-core/UnDotumBold.ttf"),
    ]

    for regular_path, bold_path in candidates:
        if os.path.exists(regular_path) and regular_path.lower().endswith(".ttf"):
            try:
                pdfmetrics.registerFont(TTFont("KoreanRegular", regular_path))
                if bold_path and os.path.exists(bold_path) and bold_path.lower().endswith(".ttf"):
                    pdfmetrics.registerFont(TTFont("KoreanBold", bold_path))
                else:
                    pdfmetrics.registerFont(TTFont("KoreanBold", regular_path))
                pdfmetrics.registerFontFamily(
                    "Korean",
                    normal="KoreanRegular",
                    bold="KoreanBold",
                    italic="KoreanRegular",
                    boldItalic="KoreanBold",
                )
                return "KoreanRegular", "KoreanBold"
            except Exception:
                continue

    raise RuntimeError(
        "한글 PDF 폰트를 찾지 못했습니다. Windows 기본 '맑은 고딕' 폰트가 정상 설치되어 있는지 확인하세요."
    )

def make_table(rows, regular_font, bold_font, cell_style, widths=None):
    if not rows:
        return None
    page_width = A4[0] - 56
    col_count = max(len(r) for r in rows)
    col_widths = [page_width / col_count] * col_count
    if widths and len(widths) == col_count:
        col_widths = [page_width * x for x in widths]

    formatted = []
    for r_idx, row in enumerate(rows):
        style = cell_style
        formatted.append([
            Paragraph(safe_html(cell), style)
            for cell in row
        ])

    tbl = Table(formatted, colWidths=col_widths, repeatRows=1, hAlign="LEFT")
    tbl.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), regular_font),
        ("FONTNAME", (0,0), (-1,0), bold_font),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#E9EEF5")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.HexColor("#111827")),
        ("GRID", (0,0), (-1,-1), 0.35, colors.HexColor("#B8C2CC")),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ]))
    return tbl

def make_pdf_bytes(result_text, analysis_date):
    buffer = io.BytesIO()
    regular_font, bold_font = register_korean_fonts()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=28,
        rightMargin=28,
        topMargin=28,
        bottomMargin=30,
        title=f"병력분석_{analysis_date}",
    )

    title_style = ParagraphStyle(
        "KTitle",
        fontName=bold_font,
        fontSize=18,
        leading=22,
        alignment=TA_CENTER,
        spaceAfter=8,
    )
    subtitle_style = ParagraphStyle(
        "KSubtitle",
        fontName=regular_font,
        fontSize=9,
        textColor=colors.HexColor("#667085"),
        leading=12,
        alignment=TA_CENTER,
        spaceAfter=12,
    )
    h_style = ParagraphStyle(
        "KH",
        fontName=bold_font,
        fontSize=12,
        leading=16,
        spaceBefore=6,
        spaceAfter=7,
    )
    body_style = ParagraphStyle(
        "KBody",
        fontName=regular_font,
        fontSize=8.7,
        leading=12.2,
        spaceAfter=5,
    )
    body_bold = ParagraphStyle(
        "KBodyBold",
        fontName=bold_font,
        fontSize=8.7,
        leading=12.2,
        spaceAfter=5,
    )
    body_small = ParagraphStyle(
        "KBodySmall",
        fontName=regular_font,
        fontSize=8.0,
        leading=11,
        spaceAfter=4,
    )
    cell_style = ParagraphStyle(
        "KCell",
        fontName=regular_font,
        fontSize=7.5,
        leading=9.6,
    )
    cell_small = ParagraphStyle(
        "KCellSmall",
        fontName=regular_font,
        fontSize=6.9,
        leading=8.8,
    )

    story = [
        Paragraph("병력분석기", title_style),
        Paragraph("기본 고지 질문서 기준 분석 결과", subtitle_style),
    ]

    sections = parse_markdown_sections(result_text)
    summary = None
    others = []

    for title, lines in sections:
        if "한눈에 보기" in title:
            summary = (title, lines)
        else:
            others.append((title, lines))

    # 1페이지: 한눈에 보기 + 기본 정보
    if summary:
        _, lines = summary
        story.append(Paragraph("한눈에 보기", h_style))
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if not line:
                i += 1
                continue
            if line.startswith("|") and line.endswith("|"):
                rows, next_i = parse_markdown_table(lines, i)
                if rows:
                    tbl = make_table(
                        rows, regular_font, bold_font, cell_style,
                        widths=[0.28, 0.15, 0.57] if len(rows[0]) == 3 else None
                    )
                    story.append(tbl)
                    story.append(Spacer(1, 8))
                i = next_i
                continue
            txt = strip_markdown(line)
            if txt:
                story.append(Paragraph(safe_html(txt), body_small))
            i += 1

    meta = []
    for _, lines in others:
        for line in lines:
            t = strip_markdown(line)
            if t.startswith("- 분석 기준일:") or t.startswith("- 분석한 파일 수:") or t.startswith("- 주의:"):
                meta.append(t.lstrip("- ").strip())
    for t in meta:
        story.append(Paragraph(safe_html(t), body_small))

    story.append(PageBreak())

    # 상세페이지
    for title, lines in others:
        filtered = []
        for line in lines:
            t = strip_markdown(line)
            if t.startswith("- 분석 기준일:") or t.startswith("- 분석한 파일 수:") or t.startswith("- 주의:"):
                continue
            filtered.append(line)

        if not any(x.strip() for x in filtered):
            continue

        story.append(Paragraph(strip_markdown(title), h_style))
        i = 0
        while i < len(filtered):
            line = filtered[i].strip()

            if not line:
                story.append(Spacer(1, 3))
                i += 1
                continue

            if line.startswith("|") and line.endswith("|"):
                rows, next_i = parse_markdown_table(filtered, i)
                if rows:
                    widths = None
                    style = cell_style

                    if "상시복용 약물" in title and len(rows[0]) == 6:
                        widths = [0.22, 0.12, 0.10, 0.16, 0.25, 0.15]
                        style = cell_small
                    elif title.strip().startswith("4.") and len(rows[0]) == 5:
                        widths = [0.16, 0.19, 0.18, 0.29, 0.18]
                        style = cell_small
                    elif "10년" in title and len(rows[0]) == 5:
                        widths = [0.15, 0.20, 0.24, 0.23, 0.18]
                        style = cell_small
                    elif "추가검사" in title and len(rows[0]) == 4:
                        widths = [0.27, 0.22, 0.33, 0.18]
                    elif "11대 질병" in title and len(rows[0]) == 5:
                        widths = [0.18, 0.17, 0.18, 0.27, 0.20]
                    elif "최근 3개월" in title and "상시복용" not in title and len(rows[0]) == 5:
                        widths = [0.15, 0.21, 0.24, 0.22, 0.18]

                    tbl = make_table(rows, regular_font, bold_font, style, widths)
                    story.append(tbl)
                    story.append(Spacer(1, 7))
                i = next_i
                continue

            txt = strip_markdown(line)
            if txt:
                if txt.startswith("판정:"):
                    story.append(Paragraph(safe_html(txt), body_bold))
                else:
                    story.append(Paragraph(safe_html(txt), body_style))
            i += 1

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

def make_excel_bytes(result_text, analysis_date):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})

    fmt_title = workbook.add_format({
        "bold": True, "font_size": 16, "align": "center", "valign": "vcenter"
    })
    fmt_section = workbook.add_format({
        "bold": True, "font_size": 12, "bg_color": "#E9EEF5",
        "border": 1, "valign": "vcenter"
    })
    fmt_header = workbook.add_format({
        "bold": True, "bg_color": "#DDE7F2", "border": 1,
        "text_wrap": True, "valign": "top", "align": "center"
    })
    fmt_cell = workbook.add_format({
        "border": 1, "text_wrap": True, "valign": "top"
    })
    fmt_text = workbook.add_format({
        "text_wrap": True, "valign": "top"
    })
    fmt_note = workbook.add_format({
        "italic": True, "font_color": "#666666", "text_wrap": True
    })

    sections = parse_markdown_sections(result_text)

    ws = workbook.add_worksheet("분석결과")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 3)
    ws.set_column("B:B", 24)
    ws.set_column("C:H", 22)
    ws.merge_range("B1:H2", "병력분석기 - 분석 결과", fmt_title)
    ws.write("B3", f"분석 기준일: {analysis_date}", fmt_note)

    row = 4
    for title, lines in sections:
        ws.merge_range(row, 1, row, 7, strip_markdown(title), fmt_section)
        row += 1
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if not line:
                row += 1
                i += 1
                continue

            if line.startswith("|") and line.endswith("|"):
                table_rows, next_i = parse_markdown_table(lines, i)
                if table_rows:
                    for c, value in enumerate(table_rows[0]):
                        ws.write(row, 1+c, value, fmt_header)
                    row += 1
                    for data_row in table_rows[1:]:
                        for c, value in enumerate(data_row):
                            ws.write(row, 1+c, value, fmt_cell)
                        row += 1
                    row += 1
                i = next_i
                continue

            txt = strip_markdown(line)
            if txt:
                ws.merge_range(row, 1, row, 7, txt, fmt_text)
                row += 1
            i += 1

    ws.freeze_panes(4, 0)

    summary_rows = None
    for title, lines in sections:
        if "한눈에 보기" in title:
            for idx, line in enumerate(lines):
                if line.strip().startswith("|") and line.strip().endswith("|"):
                    summary_rows, _ = parse_markdown_table(lines, idx)
                    break

    if summary_rows:
        sm = workbook.add_worksheet("한눈에보기")
        sm.hide_gridlines(2)
        sm.set_column("A:A", 28)
        sm.set_column("B:B", 14)
        sm.set_column("C:C", 70)
        sm.merge_range("A1:C2", "병력분석기 - 한눈에 보기", fmt_title)
        for c, value in enumerate(summary_rows[0]):
            sm.write(3, c, value, fmt_header)
        for r_idx, data_row in enumerate(summary_rows[1:], start=4):
            for c, value in enumerate(data_row):
                sm.write(r_idx, c, value, fmt_cell)

    workbook.close()
    output.seek(0)
    return output.getvalue()

# 배포용 설정: API 키와 팀 비밀번호는 서버 환경변수에만 저장합니다.
api_key = os.getenv("OPENAI_API_KEY", "").strip()
model = os.getenv("OPENAI_MODEL", "gpt-6-sol").strip() or "gpt-6-sol"
app_password = os.getenv("APP_PASSWORD", "").strip()

def require_login():
    if not app_password:
        st.error("서버에 팀 접속 비밀번호가 설정되지 않았습니다. 관리자에게 문의하세요.")
        st.stop()

    if st.session_state.get("authenticated"):
        return

    st.title("🩺 병력분석기")
    st.caption("팀 전용 병력 분석 도구")
    st.markdown("#### 팀 접속 비밀번호")
    entered = st.text_input("비밀번호", type="password", label_visibility="collapsed", placeholder="팀 비밀번호를 입력하세요")
    if st.button("접속하기", type="primary", use_container_width=True):
        if hmac.compare_digest(entered, app_password):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")
    st.stop()

require_login()

if not api_key:
    st.error("서버에 OpenAI API 키가 설정되지 않았습니다. 관리자에게 문의하세요.")
    st.stop()

with st.sidebar:
    st.caption("팀 전용")
    if st.button("로그아웃", use_container_width=True):
        st.session_state.clear()
        st.rerun()
    st.markdown("---")
    st.warning(
        "실제 고객자료 사용 전 회사의 개인정보·민감정보 처리 기준과 "
        "외부 AI 사용 허용 여부를 확인하세요."
    )

st.title("🩺 병력분석기")
st.caption("심평원 진료정보 PDF를 분석해 기본 고지 질문서 기준으로 정리합니다.")

analysis_date = st.date_input("청약/분석 기준일")

pdf_password = st.text_input(
    "PDF 비밀번호 (생년월일 8자리)",
    type="password",
    max_chars=8,
    placeholder="예: 19920921",
    help="심평원 PDF를 열 때 사용하는 비밀번호입니다. 분석 중에만 사용하며 저장하지 않습니다.",
)

uploaded_files = st.file_uploader(
    "고객 진료정보 PDF를 모두 첨부하세요",
    type=["pdf"],
    accept_multiple_files=True,
)

st.info("비밀번호가 설정된 PDF도 그대로 첨부할 수 있습니다. 원본 파일은 변경되지 않습니다.")

if uploaded_files:
    st.write(f"첨부된 PDF: **{len(uploaded_files)}개**")
    with st.expander("파일 목록"):
        for f in uploaded_files:
            st.write(f"- {f.name} ({len(f.getvalue())/1024/1024:.2f} MB)")

analyze = st.button("병력 분석하기", type="primary", use_container_width=True)

if analyze:
    if not api_key:
        st.error("서버에 OpenAI API 키가 설정되지 않았습니다.")
        st.stop()
    if not uploaded_files:
        st.error("분석할 PDF를 1개 이상 첨부하세요.")
        st.stop()

    client = OpenAI(api_key=api_key)
    created_file_ids = []
    temp_paths = []

    try:
        with st.spinner("PDF를 확인하고 병력을 분석하고 있습니다..."):
            content = [{
                "type": "input_text",
                "text": (
                    f"청약/분석 기준일은 {analysis_date.isoformat()} 입니다.\n"
                    "첨부된 모든 PDF를 서로 대조해서 중복을 제거하고, "
                    "기본 고지 질문서 1~6번 기준으로만 간결하게 분석해 주세요."
                ),
            }]

            decrypt_errors = []

            for uf in uploaded_files:
                try:
                    src_path, decrypted_path = decrypt_pdf_to_temp(uf, pdf_password)
                    temp_paths.extend([src_path, decrypted_path])
                except Exception as e:
                    decrypt_errors.append(f"{uf.name}: {e}")
                    continue

                with open(decrypted_path, "rb") as fh:
                    uploaded = client.files.create(
                        file=(uf.name, fh, "application/pdf"),
                        purpose="user_data",
                        expires_after={"anchor": "created_at", "seconds": 3600},
                    )
                created_file_ids.append(uploaded.id)
                content.append({"type": "input_file", "file_id": uploaded.id})

            if decrypt_errors:
                st.error("일부 PDF를 열지 못했습니다.")
                for msg in decrypt_errors:
                    st.write(f"- {msg}")

            if not created_file_ids:
                st.stop()

            response = client.responses.create(
                model=model,
                input=[
                    {
                        "role": "developer",
                        "content": [{"type": "input_text", "text": SYSTEM_PROMPT}],
                    },
                    {
                        "role": "user",
                        "content": content,
                    },
                ],
            )

            result = response.output_text

        st.success("병력 분석이 완료되었습니다.")
        st.markdown(result)

        txt_bytes = result.encode("utf-8-sig")
        pdf_bytes = make_pdf_bytes(result, analysis_date.isoformat())
        xlsx_bytes = make_excel_bytes(result, analysis_date.isoformat())

        st.markdown("### 결과 저장")
        col1, col2, col3 = st.columns(3)

        with col1:
            st.download_button(
                "PDF 저장",
                data=pdf_bytes,
                file_name=f"병력분석_{analysis_date.isoformat()}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        with col2:
            st.download_button(
                "Excel 저장",
                data=xlsx_bytes,
                file_name=f"병력분석_{analysis_date.isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with col3:
            st.download_button(
                "TXT 저장",
                data=txt_bytes,
                file_name=f"병력분석_{analysis_date.isoformat()}.txt",
                mime="text/plain",
                use_container_width=True,
            )

    except Exception as e:
        st.error("분석 중 오류가 발생했습니다.")
        st.exception(e)

    finally:
        for fid in created_file_ids:
            try:
                client.files.delete(fid)
            except Exception:
                pass

        for p in temp_paths:
            try:
                os.remove(p)
            except Exception:
                pass
