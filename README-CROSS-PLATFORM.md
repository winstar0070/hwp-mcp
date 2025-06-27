# Cross-Platform Document MCP Server

이 서버는 macOS, Linux, Windows에서 모두 동작하는 문서 처리 MCP 서버입니다.

## 특징

- **크로스 플랫폼**: python-docx를 사용하여 모든 OS에서 동작
- **다양한 문서 형식 지원**: DOCX, PDF 변환 (pypandoc 사용)
- **한글 완벽 지원**: Unicode 기반으로 한글 텍스트 문제없이 처리
- **MCP 프로토콜**: Claude와 직접 통합

## 설치 방법

### 1. 의존성 설치

```bash
# macOS/Linux
pip install -r requirements-cross-platform.txt

# Windows
pip install -r requirements-cross-platform.txt
```

### 2. (선택사항) Pandoc 설치

PDF 변환 기능을 사용하려면 pandoc이 필요합니다:

```bash
# macOS
brew install pandoc

# Ubuntu/Debian
sudo apt-get install pandoc

# Windows
# https://pandoc.org/installing.html 에서 다운로드

# 또는 Python 패키지로 설치
pip install pypandoc-binary
```

## Claude Desktop 설정

Claude Desktop의 설정 파일에 다음을 추가하세요:

### macOS
`~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "cross-platform-doc": {
      "command": "python3",
      "args": ["/절대/경로/cross_platform_mcp_server.py"]
    }
  }
}
```

### Windows
`%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "cross-platform-doc": {
      "command": "python",
      "args": ["C:\\절대\\경로\\cross_platform_mcp_server.py"]
    }
  }
}
```

## 사용 가능한 기능

### 문서 관리
- `create_new_document`: 새 문서 생성
- `open_document`: 기존 문서 열기
- `save_document`: 문서 저장
- `get_document_info`: 문서 정보 조회

### 텍스트 작업
- `insert_text`: 텍스트 삽입
- `insert_heading`: 제목 삽입 (레벨 1-9)
- `find_and_replace`: 찾기 및 바꾸기
- `set_font`: 폰트 설정

### 표 작업
- `insert_table`: 표 삽입
- 데이터와 함께 표 생성 가능

### 목록
- `insert_bullet_list`: 글머리 기호 목록
- `insert_numbered_list`: 번호 매기기 목록

### 기타
- `insert_image`: 이미지 삽입
- `add_page_break`: 페이지 나누기
- `export_to_pdf`: PDF로 내보내기
- `list_available_styles`: 사용 가능한 스타일 목록

## 사용 예시

Claude Desktop에서:

```
"새 문서를 만들고 제목을 '회의록'으로 설정해줘"
"오늘 날짜로 부제목을 추가하고 참석자 명단을 표로 만들어줘"
"문서를 PDF로 저장해줘"
```

## HWP 파일 지원

HWP 파일을 직접 편집할 수는 없지만, 다음과 같은 워크플로우가 가능합니다:

1. HWP → DOCX 변환 (별도 도구 사용)
2. 이 서버로 DOCX 편집
3. DOCX → HWP 변환 (필요시)

## 개발 계획

- [ ] Excel 지원 추가 (openpyxl)
- [ ] PowerPoint 지원 추가 (python-pptx)
- [ ] 더 많은 포맷팅 옵션
- [ ] 템플릿 시스템
- [ ] 일괄 처리 기능

## 문제 해결

### "No module named 'docx'" 오류
```bash
pip install python-docx
```

### PDF 변환 실패
```bash
# pandoc 설치 확인
pandoc --version

# 또는 pypandoc-binary 사용
pip install pypandoc-binary
```

### 한글 폰트 문제 (PDF)
시스템에 한글 폰트가 설치되어 있는지 확인하세요.

## 라이선스

MIT License