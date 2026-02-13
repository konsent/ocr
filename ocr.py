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


PDF_PATH = "RS_ScenarioBook_Final.pdf"
OUTPUT_PATH = "output.txt"

# ---------------------------
# 설정값 (문서마다 조정 가능)
# ---------------------------
HEADER_RATIO = 0.05   # 상단 8% 제거
FOOTER_RATIO = 0.93   # 하단 8% 제거
COLUMN_SPLIT_RATIO = 0.5  # 좌우 단 분리 기준
TITLE_LEVEL = 3  # 0 = 1.0 /  1 = 1.0, 1.1 / 2 = 1.0, 1.1, 1.1.1

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

    # 문장이 이어지는지 판단하기 위한 접속사/전치사 등 (소문자)
    continuation_words = {'and', 'or', 'but', 'the', 'a', 'an', 'of', 'in', 'on', 'at', 'to', 'by', 'for', 'with', 'from', 'is', 'are', 'was', 'were', 'that', 'which', 'who'}

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
            
            # 1. 불릿 포인트로 시작하면 무조건 분리
            if line[0] in ['•', '●', '-']:
                should_split = True
            
            # 2. 빈 줄이 있었으면 분리
            elif pending_break:
                should_split = True
            
            elif buffer:
                buffer_stripped = buffer.strip()
                # 3. 이전 줄이 마침표 등으로 끝남 -> 분리
                if buffer_stripped[-1] in ['.', '?', '!', '"', '”']:
                    if line[0].isupper() or line[0].isdigit():
                        should_split = True
                # 4. 이전 줄이 마침표로 끝나지 않았지만, 짧고(헤더 가능성), 다음 줄이 대문자로 시작 -> 분리
                elif line[0].isupper() or line[0].isdigit():
                    last_word = buffer_stripped.split()[-1].lower() if buffer_stripped else ""
                    # 접속사/전치사로 끝나지 않고, 길이가 짧으면(60자 미만) 헤더로 간주하여 분리
                    if last_word not in continuation_words and len(buffer_stripped) < 60:
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
    # 1. 단어 단위 추출 (위치 정보 포함)
    words = page.extract_words(x_tolerance=3, y_tolerance=3, keep_blank_chars=False)
    
    # 왼쪽 단 영역
    left_bbox = (0, header_bottom, column_split, footer_top)
    # 오른쪽 단 영역
    right_bbox = (column_split, header_bottom, width, footer_top)
    # 2. 헤더/푸터 영역 제외
    body_words = [w for w in words if header_bottom <= w['top'] <= footer_top]
    
    # 3. 줄(Line) 단위로 그룹화 (Y좌표 기준 정렬 후 클러스터링)
    body_words.sort(key=lambda w: w['top'])
    lines = []
    if body_words:
        current_line = [body_words[0]]
        for w in body_words[1:]:
            # 같은 줄로 볼 수 있는 Y 좌표 차이 (약 5픽셀 이내)
            if abs(w['top'] - current_line[-1]['top']) < 5:
                current_line.append(w)
            else:
                lines.append(current_line)
                current_line = [w]
        lines.append(current_line)

    try:
        left_crop = page.crop(left_bbox)
        left_text = left_crop.extract_text() or ""
    except ValueError:
        # 영역이 유효하지 않은 경우 (예: 빈 페이지)
        left_text = ""
    final_text_parts = []
    left_buffer = []
    right_buffer = []
    
    # 단 간격 기준 (이보다 작으면 같은 문장으로 간주, 픽셀 단위)
    GUTTER_THRESHOLD = 20 

    try:
        right_crop = page.crop(right_bbox)
        right_text = right_crop.extract_text() or ""
    except ValueError:
        right_text = ""
    for line in lines:
        # 줄 내부 단어들을 X좌표 순으로 정렬
        line.sort(key=lambda w: w['x0'])
        
        # A. 중앙을 가로지르는 단어가 있는지 확인 (예: 큰 제목)
        crossing_word = any(w['x0'] < column_split and w['x1'] > column_split for w in line)
        
        # B. 좌우에 단어가 다 있는데, 사이 간격이 매우 좁은지 확인 (예: 잘린 단어)
        small_gap = False
        left_side = [w for w in line if w['x1'] <= column_split]
        right_side = [w for w in line if w['x0'] >= column_split]
        
        if left_side and right_side:
            gap = right_side[0]['x0'] - left_side[-1]['x1']
            if gap < GUTTER_THRESHOLD:
                small_gap = True
        
        # 가로지르거나 간격이 좁으면 -> 통짜 라인(Spanning Line)으로 처리
        if crossing_word or small_gap:
            # 기존에 쌓인 좌/우 버퍼 비우기 (순서: 왼쪽 다 출력 -> 오른쪽 다 출력)
            if left_buffer:
                final_text_parts.append("\n".join(left_buffer))
                left_buffer = []
            if right_buffer:
                final_text_parts.append("\n".join(right_buffer))
                right_buffer = []
            
            # 현재 줄을 통째로 추가
            line_text = " ".join(w['text'] for w in line)
            final_text_parts.append(line_text)
            
        else:
            # 일반적인 2단 분리 라인
            l_text = " ".join(w['text'] for w in left_side)
            r_text = " ".join(w['text'] for w in right_side)
            
            if l_text: left_buffer.append(l_text)
            if r_text: right_buffer.append(r_text)

    # 왼쪽 단 -> 오른쪽 단 순서로 결합
    return left_text + "\n" + right_text
    # 남은 버퍼 처리
    if left_buffer:
        final_text_parts.append("\n".join(left_buffer))
    if right_buffer:
        final_text_parts.append("\n".join(right_buffer))

    return "\n".join(final_text_parts)

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
