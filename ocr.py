import pdfplumber
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 구글 API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# 구글 API 키 적용
filename = "translation-418422-02dd548565fc.json"
creds = ServiceAccountCredentials.from_json_keyfile_name(filename, scope)
gc = gspread.authorize(creds)

# 구글 스프레드 시트 경로
sheet = gc.open_by_url("https://docs.google.com/spreadsheets/d/18ZPGcV0qElVQq8JaMMp3nib0MT6_znnIETQvGsumCwM/edit?gid=0#gid=0").sheet1


PDF_PATH = "RW_Rules_FINAL_Med_Res.pdf"
OUTPUT_PATH = "output.txt"

# ---------------------------
# 설정값 (문서마다 조정 가능)
# ---------------------------
HEADER_RATIO = 0.05   # 상단 8% 제거
FOOTER_RATIO = 0.93   # 하단 8% 제거
COLUMN_SPLIT_RATIO = 0.5  # 좌우 단 분리 기준


# ---------------------------
# 제목 레벨 함수
# ---------------------------
def build_title_patterns(title_level):
    patterns = []

    # 숫자 제목 패턴 생성
    for depth in range(1, title_level + 2):
        patterns.append(
            rf"^\d+(\.\d+){{{depth}}}\s+.+"
        )

    # 대문자 제목 (부록 등)
    patterns.append(r"^[A-Z][A-Z\s\-]{5,}$")

    return patterns



# ---------------------------
# 제목 판단
# ---------------------------
def is_title(line):
    if len(line) < 3:
        return False

    for p in TITLE_PATTERNS:
        if re.match(p, line):
            return True

    return False


# ---------------------------
# 문단 합치기
# ---------------------------
def merge_paragraphs(lines):
    paragraphs = []
    buffer = ""
    pending_break = False

    for line in lines:
        line = line.strip()
        if not line:
            if buffer:
                pending_break = True
            continue

        # 소문자로 시작하면 무조건 이어붙이기 (줄바꿈 무시)
        if line and line[0].islower():
            if buffer:
                buffer += " " + line
            else:
                buffer = line
            pending_break = False
            continue

        if is_title(line):
            if buffer:
                paragraphs.append(buffer.strip() + "\n")
            paragraphs.append(line + "\n")
            buffer = ""
            pending_break = False
        else:
            # 문단 분리 조건 판별
            should_split = False
            if pending_break:
                should_split = True
            elif buffer:
                # 이전 줄이 마침표로 끝나고, 현재 줄이 대문자/숫자 등으로 시작하면 분리
                if buffer.strip()[-1] in ['.', '?', '!', '"', '”'] and (line[0].isupper() or line[0].isdigit() or line[0] in ['•', '-']):
                    should_split = True

            if should_split:
                paragraphs.append(buffer.strip() + "\n")
                buffer = line
                pending_break = False
            else:
                if buffer:
                    buffer += " " + line
                else:
                    buffer = line

    if buffer:
        paragraphs.append(buffer.strip() + "\n")

    return paragraphs


# ---------------------------
# 페이지에서 본문 추출
# ---------------------------
def extract_page_text(page):
    width = page.width
    height = page.height

    # 좌표 계산 (pdfplumber는 (x0, top, x1, bottom) 형식을 사용)
    header_bottom = height * HEADER_RATIO
    footer_top = height * FOOTER_RATIO
    column_split = width * COLUMN_SPLIT_RATIO

    # 1. 헤더/푸터가 제거된 메인 영역 설정
    # crop box: (x0, top, x1, bottom)
    # 주의: crop은 해당 영역만큼 페이지를 잘라낸 객체를 반환합니다.
    
    # 왼쪽 단 영역
    left_bbox = (0, header_bottom, column_split, footer_top)
    # 오른쪽 단 영역
    right_bbox = (column_split, header_bottom, width, footer_top)

    try:
        left_crop = page.crop(left_bbox)
        left_text = left_crop.extract_text() or ""
    except ValueError:
        # 영역이 유효하지 않은 경우 (예: 빈 페이지)
        left_text = ""

    try:
        right_crop = page.crop(right_bbox)
        right_text = right_crop.extract_text() or ""
    except ValueError:
        right_text = ""

    # 왼쪽 단 -> 오른쪽 단 순서로 결합
    return left_text + "\n" + right_text

def extract_page_text_legacy1(page):
    width = page.width
    height = page.height

    header_cut = height * HEADER_RATIO
    footer_cut = height * FOOTER_RATIO
    column_split = width * COLUMN_SPLIT_RATIO

    words = page.extract_words()

    left_words = []
    right_words = []

    for w in words:
        y = w["top"]

        # 머릿글/꼬리글 제거
        if y < header_cut or y > footer_cut:
            continue

        if w["x0"] < column_split:
            left_words.append(w)
        else:
            right_words.append(w)

    # y 좌표 기준 정렬
    left_words.sort(key=lambda x: (x["top"], x["x0"]))
    right_words.sort(key=lambda x: (x["top"], x["x0"]))

    left_text = " ".join(w["text"] for w in left_words)
    right_text = " ".join(w["text"] for w in right_words)

    # 왼쪽 단 → 오른쪽 단 순서로 결합
    return left_text + "\n" + right_text


# ---------------------------
# 메인 처리
# ---------------------------
all_lines = []

# 0 = 1.0 /  1 = 1.0, 1.1 / 2 = 1.0, 1.1, 1.1.1
TITLE_LEVEL = 3  # 여기만 바꾸면 됨
TITLE_PATTERNS = build_title_patterns(TITLE_LEVEL)

with pdfplumber.open(PDF_PATH) as pdf:
    for page in pdf.pages:
        text = extract_page_text(page)
        lines = text.split("\n")
        all_lines.extend(lines)

paragraphs = merge_paragraphs(all_lines)

with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
    for p in paragraphs:
        f.write(p)

# 구글 스프레드시트 B2 셀에 결과 붙여넣기
data = [[p.strip()] for p in paragraphs]
if data:
    sheet.update(range_name=f"B2:B{len(data) + 1}", values=data)

print("완료:", OUTPUT_PATH)
