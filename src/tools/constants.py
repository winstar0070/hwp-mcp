"""
HWP MCP 프로젝트의 상수 정의 파일
모든 하드코딩된 값들을 중앙 집중화하여 관리합니다.
"""

# ============== HWP 단위 변환 상수 ==============
# HWP는 HWPUNIT이라는 고유 단위를 사용합니다
HWPUNIT_PER_MM = 2834.64   # 1mm = 2834.64 HWPUNIT
HWPUNIT_PER_CM = 28346.4   # 1cm = 28346.4 HWPUNIT
HWPUNIT_PER_INCH = 72009.6 # 1inch = 72009.6 HWPUNIT
HWPUNIT_PER_PT = 100       # 1pt = 100 HWPUNIT

# ============== 표 관련 상수 ==============
TABLE_MAX_ROWS = 100       # 표 최대 행 수
TABLE_MAX_COLS = 100       # 표 최대 열 수
TABLE_DEFAULT_WIDTH = 8000 # 표 기본 너비 (HWPUNIT)
TABLE_DEFAULT_HEIGHT = 1000 # 셀 기본 높이 (HWPUNIT)

# 표 스타일
TABLE_STYLES = {
    "default": {
        "border_type": 1,
        "border_width": 0.7
    },
    "simple": {
        "border_type": 1,
        "border_width": 0.5,
        "header_color": "#F0F0F0"
    },
    "professional": {
        "border_type": 1,
        "border_width": 1.0,
        "header_color": "#4472C4",
        "header_text_color": "#FFFFFF"
    },
    "colorful": {
        "border_type": 1,
        "border_width": 0.5,
        "alternating_colors": ["#F2F2F2", "#FFFFFF"]
    },
    "dark": {
        "border_type": 2,
        "border_width": 1.0,
        "header_color": "#2B2B2B",
        "header_text_color": "#FFFFFF",
        "body_color": "#3A3A3A",
        "body_text_color": "#E0E0E0"
    }
}

# ============== 이미지 관련 상수 ==============
ALLOWED_IMAGE_FORMATS = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff']
IMAGE_EMBED_MODE = 1       # 0: 링크, 1: 파일 포함
IMAGE_DEFAULT_WIDTH = 0    # 0: 원본 크기 유지
IMAGE_DEFAULT_HEIGHT = 0   # 0: 원본 크기 유지

# ============== 페이지 설정 상수 ==============
# 용지 크기 (HWPUNIT)
PAPER_SIZES = {
    "A3": {"width": 119903, "height": 169613},
    "A4": {"width": 59528, "height": 84188},
    "A5": {"width": 41976, "height": 59528},
    "B4": {"width": 72855, "height": 103185},
    "B5": {"width": 51590, "height": 72855},
    "Letter": {"width": 61199, "height": 79370},
    "Legal": {"width": 61199, "height": 100711}
}

# 페이지 방향
PAGE_ORIENTATION = {
    "portrait": 0,   # 세로
    "landscape": 1   # 가로
}

# 기본 여백 (mm 단위)
DEFAULT_MARGINS = {
    "top": 20,
    "bottom": 20,
    "left": 30,
    "right": 30,
    "header": 15,
    "footer": 15
}

# ============== 색상 관련 상수 ==============
# 하이라이트 색상 매핑
HIGHLIGHT_COLORS = {
    "yellow": 4,     # 노란색
    "green": 3,      # 초록색
    "cyan": 6,       # 청록색
    "magenta": 5,    # 자홍색
    "red": 2,        # 빨간색
    "blue": 1,       # 파란색
    "gray": 7        # 회색
}

# ============== 텍스트 서식 상수 ==============
# 정렬 옵션
TEXT_ALIGNMENT = {
    "left": 0,
    "center": 1,
    "right": 2,
    "justify": 3
}

# 밑줄 스타일
UNDERLINE_STYLES = {
    "none": 0,
    "single": 1,
    "double": 2,
    "thick": 3,
    "dotted": 4,
    "dashed": 5,
    "wave": 6
}

# ============== 필드 타입 상수 ==============
FIELD_TYPES = {
    "date": "InsertFieldDate",
    "time": "InsertFieldTime",
    "page": "InsertPageNumber",
    "total_pages": "InsertPageCount",
    "filename": "InsertFieldFileName",
    "path": "InsertFieldPath",
    "author": "InsertFieldAuthor"
}

# ============== PDF 변환 상수 ==============
PDF_QUALITY = {
    "low": 0,
    "medium": 1,
    "high": 2,
    "print": 3
}

# ============== 찾기/바꾸기 옵션 상수 ==============
FIND_OPTIONS = {
    "direction_forward": 0,
    "direction_backward": 1,
    "match_case": 1,
    "ignore_case": 0,
    "whole_word": 1,
    "partial_word": 0
}

# ============== 도형 타입 상수 ==============
SHAPE_TYPES = {
    "rectangle": 0,
    "ellipse": 1,
    "line": 2,
    "polygon": 3,
    "arrow": 4,
    "textbox": 5
}

# ============== 보안 모듈 경로 ==============
SECURITY_MODULE_NAME = "FilePathCheckerModuleExample"
SECURITY_MODULE_DEFAULT_PATH = "D:/hwp-mcp/security_module/FilePathCheckerModuleExample.dll"

# ============== 시간 제한 상수 ==============
DEFAULT_TIMEOUT = 30       # 기본 타임아웃 (초)
BATCH_OPERATION_TIMEOUT = 300  # 배치 작업 타임아웃 (초)

# ============== 배치 작업 상수 ==============
BATCH_CHUNK_SIZE = 100     # 배치 작업 청크 크기
MAX_RETRY_COUNT = 3        # 최대 재시도 횟수
RETRY_DELAY = 1            # 재시도 지연 시간 (초)

# ============== 템플릿 관련 상수 ==============
TEMPLATE_DIR = "HWP_Templates"  # 템플릿 디렉토리 이름
TEMP_DOCUMENT_NAME = "temp_document.hwp"  # 임시 문서 파일명

# ============== 특수 문자열 상수 ==============
NUMBER_SEQUENCE_KOREAN = "1부터 10까지"  # 숫자 시퀀스 한국어 표현
VERTICAL_KOREAN = "세로"  # 세로 방향 한국어