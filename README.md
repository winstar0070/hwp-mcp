# HWP-MCP (한글 Model Context Protocol)

[![GitHub](https://img.shields.io/github/license/jkf87/hwp-mcp)](https://github.com/jkf87/hwp-mcp)

HWP-MCP는 한글 워드 프로세서(HWP)를 Claude와 같은 AI 모델이 제어할 수 있도록 해주는 Model Context Protocol(MCP) 서버입니다. 이 프로젝트는 한글 문서를 자동으로 생성, 편집, 조작하는 기능을 AI에게 제공합니다.

## 주요 기능

- 문서 생성 및 관리: 새 문서 생성, 열기, 저장 기능
- 텍스트 편집: 텍스트 삽입, 글꼴 설정, 단락 추가
- 테이블 작업: 테이블 생성, 데이터 채우기, 셀 내용 설정
- 완성된 문서 생성: 템플릿 기반 보고서 및 편지 자동 생성
- 일괄 작업: 여러 작업을 한 번에 실행하는 배치 기능

## 시스템 요구사항

- Windows 운영체제
- 한글(HWP) 프로그램 설치
- Python 3.7 이상
- 필수 Python 패키지 (requirements.txt 참조)

## 설치 방법

1. 저장소 클론:
```bash
git clone https://github.com/winstar0070/hwp-mcp.git
cd hwp-mcp
```

2. 의존성 설치:
```bash
pip install -r requirements.txt
```

3. (선택사항) MCP 패키지 설치:
```bash
pip install mcp
```

## 사용 방법

### Claude와 함께 사용하기

Claude 데스크톱 설정 파일에 다음과 같이 HWP-MCP 서버를 등록하세요:

```json
{
  "mcpServers": {
    "hwp": {
      "command": "python",
      "args": ["경로/hwp-mcp/hwp_mcp_stdio_server.py"]
    }
  }
}
```

### 주요 기능 예시

#### 새 문서 생성
```python
hwp_create()
```

#### 텍스트 삽입
```python
# 일반 텍스트 삽입
hwp_insert_text("원하는 텍스트를 입력하세요.")

# 서식이 적용된 텍스트 삽입
hwp_insert_text_with_font(
    text="굵은 14pt 텍스트입니다.",
    font_name="맑은 고딕",
    font_size=14,
    bold=True
)
```

#### 테이블 생성 및 데이터 입력
```python
# 테이블 생성
hwp_insert_table(rows=5, cols=2)

# 테이블에 데이터 채우기
hwp_fill_table_with_data([
    ["월", "판매량"], 
    ["1월", "120"], 
    ["2월", "150"], 
    ["3월", "180"], 
    ["4월", "200"]
], has_header=True)

# 표에 연속된 숫자 채우기
hwp_fill_column_numbers(start=1, end=10, column=1, from_first_cell=True)
```

#### 문서 저장
```python
hwp_save("경로/문서명.hwp")
```

#### 글자 서식 설정
```python
# 선택된 텍스트에 서식 적용
hwp_apply_font_to_selection(
    font_name="바탕",
    font_size=12,
    italic=True,
    underline=True
)

# 다음에 입력할 텍스트의 서식 설정
hwp_set_font(name="굴림", size=16, bold=True)
```

#### 이미지 삽입
```python
# 이미지 삽입
hwp_insert_image(
    image_path="이미지.jpg",
    width=100,  # mm 단위
    height=75,
    align="center"
)
```

#### 찾기/바꾸기
```python
# 텍스트 찾기/바꾸기
hwp_find_replace(
    find_text="찾을텍스트",
    replace_text="바꿀텍스트",
    match_case=False,
    replace_all=True
)
```

#### PDF 변환
```python
# 현재 문서를 PDF로 변환
hwp_export_pdf(
    output_path="문서.pdf",
    quality="high"
)
```

#### 페이지 설정
```python
# 페이지 크기, 방향, 여백 설정
hwp_set_page(
    paper_size="A4",
    orientation="portrait",
    top_margin=20,
    bottom_margin=20,
    left_margin=30,
    right_margin=30
)
```

#### 머리말/꼬리말 설정
```python
# 머리말과 꼬리말 설정
hwp_set_header_footer(
    header_text="문서 제목",
    footer_text="페이지 ",
    show_page_number=True,
    page_number_position="footer-center"
)
```

#### 문단 서식
```python
# 문단 서식 설정
hwp_set_paragraph(
    alignment="justify",  # left, center, right, justify
    line_spacing=1.5,
    indent_first=10,
    space_after=5
)
```

#### 일괄 작업 예시
```python
hwp_batch_operations([
    {"operation": "hwp_create"},
    {"operation": "hwp_set_page", "params": {"paper_size": "A4"}},
    {"operation": "hwp_insert_text_with_font", 
     "params": {"text": "제목", "font_size": 20, "bold": True}},
    {"operation": "hwp_insert_paragraph"},
    {"operation": "hwp_insert_text", "params": {"text": "본문 내용"}},
    {"operation": "hwp_export_pdf", "params": {"output_path": "문서.pdf"}},
    {"operation": "hwp_save", "params": {"path": "문서.hwp"}}
])
```

## 프로젝트 구조

```
hwp-mcp/
├── hwp_mcp_stdio_server.py  # 메인 서버 스크립트
├── requirements.txt         # 의존성 패키지 목록
├── hwp-mcp-구조설명.md       # 프로젝트 구조 설명 문서
├── CLAUDE.md               # Claude Code용 프로젝트 가이드
├── src/
│   ├── tools/
│   │   ├── hwp_controller.py      # 한글 제어 핵심 컨트롤러
│   │   ├── hwp_table_tools.py     # 테이블 관련 기능 전문 모듈
│   │   ├── hwp_advanced_features.py # 고급 기능 모듈 (이미지, PDF, 서식 등)
│   │   ├── hwp_document_features.py # 문서 편집 고급 기능 모듈
│   │   ├── hwp_exceptions.py      # 예외 클래스 정의
│   │   ├── constants.py           # 상수 정의
│   │   └── config.py              # 설정 관리
│   ├── utils/                     # 유틸리티 함수
│   └── __tests__/                 # 테스트 모듈
└── security_module/
    └── FilePathCheckerModuleExample.dll  # 보안 모듈
```

## 트러블슈팅

### 보안 모듈 관련 문제
기본적으로 한글 프로그램은 외부에서 파일 접근 시 보안 경고를 표시합니다. 이를 우회하기 위해 `FilePathCheckerModuleExample.dll` 모듈을 사용합니다. 만약 보안 모듈 등록에 실패해도 기능은 작동하지만, 파일 열기/저장 시 보안 대화 상자가 표시될 수 있습니다.

### 한글 연결 실패
한글 프로그램이 실행 중이지 않을 경우 연결에 실패할 수 있습니다. 한글 프로그램이 설치되어 있고 정상 작동하는지 확인하세요.

### 테이블 데이터 입력 문제
테이블에 데이터를 입력할 때 커서 위치가 예상과 다르게 동작하는 경우가 있었으나, 현재 버전에서는 이 문제가 해결되었습니다. 테이블의 모든 셀에 정확하게 데이터가 입력됩니다.

## 변경 로그

### 2025-06-20 (최신)
- **문서 편집 고급 기능 대량 추가**
  - `hwp_document_features.py` 모듈 신규 생성
  - 각주/미주 삽입: `hwp_insert_footnote`, `hwp_insert_endnote`
  - 하이퍼링크: `hwp_insert_hyperlink`
  - 북마크: `hwp_insert_bookmark`, `hwp_goto_bookmark`
  - 주석: `hwp_insert_comment`
  - 검색 및 하이라이트: `hwp_search_and_highlight`
  - 워터마크: `hwp_insert_watermark`
  - 필드 코드: `hwp_insert_field` (날짜, 시간, 페이지 번호 등)
  - 문서 보안: `hwp_set_document_password`
  
- **표 고급 기능 추가**
  - `hwp_table_tools.py` 모듈 확장
  - 표 스타일 적용: `hwp_apply_table_style` (simple, professional, colorful, dark)
  - 표 정렬: `hwp_sort_table`
  - 셀 병합: `hwp_merge_cells`
  - 셀 분할: `hwp_split_cell`
  
- **코드 품질 개선**
  - `hwp_exceptions.py`: 구체적인 예외 클래스 정의
    - HwpConnectionError, HwpDocumentError, HwpTableError 등
    - 각 기능별 특화된 예외 처리
  - `constants.py`: 모든 하드코딩된 값 중앙화
    - HWPUNIT 변환 상수
    - 표, 이미지, 색상 관련 상수
    - 페이지 설정, PDF 품질 등
  - `config.py`: 런타임 설정 관리
    - 싱글톤 패턴의 ConfigManager
    - JSON 기반 설정 저장/로드
    - 환경별 구성 지원
  
- **기존 기능 개선**
  - HwpController에 예외 처리 적용
  - `is_connected()` 메서드 추가: HWP 연결 상태 확인
  - `get_document_features()` 메서드 추가: 문서 고급 기능 접근
  
- 글자 서식 설정 기능 개선
  - `insert_text_with_font` 메서드 추가: 서식이 적용된 텍스트를 직접 삽입
  - `apply_font_to_selection` 메서드 추가: 선택된 텍스트에 서식 적용
  - 기존 `set_font` 메서드의 스타일 설정 로직 개선 (명시적 1/0 값 사용)
  - 글자 크기, 굵게, 기울임꼴, 밑줄 등 모든 서식이 정상 작동하도록 수정
  - MCP 서버에 새로운 도구 추가: `hwp_insert_text_with_font`, `hwp_apply_font_to_selection`
- 고급 기능 대량 추가
  - **1단계 - 즉시 유용한 기능들**
    - `hwp_insert_image`: 이미지 삽입 (크기 조절, 정렬 옵션 포함)
    - `hwp_find_replace`: 찾기/바꾸기 (대소문자 구분, 전체 단어 일치 옵션)
    - `hwp_export_pdf`: PDF 변환 (품질 설정 가능)
  - **2단계 - 문서 품질 향상 기능들**
    - `hwp_set_page`: 페이지 설정 (용지 크기, 방향, 여백)
    - `hwp_set_header_footer`: 머리말/꼬리말 설정 (페이지 번호 포함)
    - `hwp_set_paragraph`: 문단 서식 (정렬, 줄 간격, 들여쓰기)
  - **3단계 - 고급 기능들**
    - `hwp_create_toc`: 목차 자동 생성
    - `hwp_insert_shape`: 도형 삽입 (사각형, 원, 화살표 등)
    - `hwp_save_as_template`: 템플릿 저장
    - `hwp_apply_template`: 템플릿 적용
  - 새로운 모듈 추가: `hwp_advanced_features.py`
- 프로젝트 관리 개선
  - CLAUDE.md 파일 추가 (Claude Code용 프로젝트 가이드)
  - .gitignore에 CLAUDE.md 추가 (사용자별 설정 파일)
  - 테스트 코드 추가: `test_font_features.py`, `test_advanced_features.py`

### 2025-03-27
- 표 생성 및 데이터 채우기 기능 개선
  - 표 안에 표가 중첩되는 문제 해결
  - 표 생성과 데이터 채우기 기능 분리
  - 표 생성 전 현재 커서 위치 확인 로직 추가
  - 기존 표에 데이터만 채우는 기능 개선
- 프로젝트 관리 개선
  - .gitignore 파일 추가 (임시 파일, 캐시 파일 등 제외)

### 2025-03-25
- 테이블 데이터 입력 기능 개선
  - 첫 번째 셀부터 정확하게 데이터 입력 가능
  - 셀 선택 및 커서 위치 설정 로직 개선
  - 텍스트 입력 시 커서 위치 유지 기능 추가
- 테이블 전용 도구 모듈(`hwp_table_tools.py`) 추가
- `hwp_fill_column_numbers` 함수에 `from_first_cell` 옵션 추가

## 설정 관리 (config.py)

HWP-MCP는 유연한 설정 관리 시스템을 제공합니다.

### 설정 파일 위치

설정 파일은 다음 순서로 검색됩니다:
1. 환경 변수 `HWP_MCP_CONFIG`에 지정된 경로
2. 현재 디렉토리의 `hwp_config.json`
3. 사용자 홈 디렉토리의 `~/.hwp_mcp/config.json`
4. 프로젝트 루트의 `hwp_config.json`

### 설정 파일 예시

```json
{
  "auto_connect": true,
  "visible": true,
  "register_security_module": true,
  "security_module_path": "D:/hwp-mcp/security_module/FilePathCheckerModuleExample.dll",
  "batch_size": 100,
  "timeout": 300,
  "retry_count": 3,
  "default_font": "맑은 고딕",
  "default_font_size": 10,
  "default_paper_size": "A4",
  "log_level": "INFO",
  "debug_mode": false
}
```

### 주요 설정 항목

#### 일반 설정
- `auto_connect`: HWP 자동 연결 여부 (기본값: true)
- `visible`: HWP 창 표시 여부 (기본값: true)
- `register_security_module`: 보안 모듈 자동 등록 (기본값: true)

#### 성능 설정
- `batch_size`: 배치 작업 청크 크기 (기본값: 100)
- `timeout`: 작업 타임아웃 (초) (기본값: 300)
- `retry_count`: 재시도 횟수 (기본값: 3)

#### 문서 기본값
- `default_font`: 기본 글꼴 (기본값: "맑은 고딕")
- `default_font_size`: 기본 글꼴 크기 (기본값: 10)
- `default_paper_size`: 기본 용지 크기 (기본값: "A4")
- `default_orientation`: 기본 용지 방향 (기본값: "portrait")

#### 로깅 설정
- `log_level`: 로그 레벨 (기본값: "INFO")
- `log_file`: 로그 파일 이름 (기본값: "hwp_mcp.log")
- `debug_mode`: 디버그 모드 (기본값: false)

### 프로그래밍 방식으로 설정 관리

```python
from src.tools.config import get_config, update_config, save_config

# 현재 설정 가져오기
config = get_config()
print(f"현재 글꼴: {config.default_font}")

# 설정 업데이트
update_config(
    default_font="바탕",
    default_font_size=12,
    debug_mode=True
)

# 설정 저장
save_config("my_config.json")
```

### 환경별 설정

개발, 테스트, 프로덕션 환경별로 다른 설정을 사용하려면:

```bash
# 개발 환경
export HWP_MCP_CONFIG=config/development.json

# 프로덕션 환경
export HWP_MCP_CONFIG=config/production.json
```

## 라이선스

이 프로젝트는 MIT 라이선스에 따라 배포됩니다. 자세한 내용은 [LICENSE](LICENSE) 파일을 참조하세요.

## 기여 방법

1. 이슈 제보 또는 기능 제안: GitHub 이슈를 사용하세요.
2. 코드 기여: 변경사항을 포함한 Pull Request를 제출하세요.

## 관련 프로젝트

- [HWP SDK](https://www.hancom.com/product/sdk): 한글과컴퓨터의 공식 SDK
- [Cursor MCP](https://docs.cursor.com/context/model-context-protocol#configuration-locations)
- [Smithery](https://smithery.ai/server/@jkf87/hwp-mcp)

## 연락처

프로젝트 관련 문의는 GitHub 이슈, [코난쌤](https://www.youtube.com/@conanssam)를 통해 해주세요. 
