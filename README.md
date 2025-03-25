# HWP-MCP (한글 Model Context Protocol)

[![GitHub](https://img.shields.io/github/license/jkf87/hwp-mcp)](https://github.com/jkf87/hwp-mcp)

HWP-MCP는 한글 워드 프로세서(HWP)를 Claude와 같은 AI 모델이 제어할 수 있도록 해주는 Model Context Protocol(MCP) 서버입니다. 이 프로젝트는 한글 문서를 자동으로 생성, 편집, 조작하는 기능을 AI에게 제공합니다.

## 주요 기능

- 문서 생성 및 관리: 새 문서 생성, 열기, 저장 기능
- 텍스트 편집: 텍스트 삽입, 글꼴 설정, 단락 추가
- 테이블 작업: 테이블 생성, 셀 병합, 셀 내용 설정
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
git clone https://github.com/jkf87/hwp-mcp.git
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
hwp_insert_text("원하는 텍스트를 입력하세요.")
```

#### 테이블 생성(구현중)
```python
hwp_insert_table(rows=3, cols=4)
```

#### 문서 저장
```python
hwp_save("경로/문서명.hwp")
```

#### 일괄 작업 예시
```python
hwp_batch_operations([
    {"operation": "hwp_create"},
    {"operation": "hwp_insert_text", "params": {"text": "제목"}},
    {"operation": "hwp_set_font", "params": {"size": 20, "bold": True}},
    {"operation": "hwp_save", "params": {"path": "경로/문서명.hwp"}}
])
```

## 프로젝트 구조

```
hwp-mcp/
├── hwp_mcp_stdio_server.py  # 메인 서버 스크립트
├── requirements.txt         # 의존성 패키지 목록
├── hwp-mcp-구조설명.md       # 프로젝트 구조 설명 문서
├── src/
│   ├── tools/
│   │   └── hwp_controller.py  # 한글 제어 핵심 컨트롤러
│   ├── utils/                 # 유틸리티 함수
│   └── __tests__/             # 테스트 모듈
└── security_module/
    └── FilePathCheckerModuleExample.dll  # 보안 모듈
```

## 트러블슈팅

### 보안 모듈 관련 문제
기본적으로 한글 프로그램은 외부에서 파일 접근 시 보안 경고를 표시합니다. 이를 우회하기 위해 `FilePathCheckerModuleExample.dll` 모듈을 사용합니다. 만약 보안 모듈 등록에 실패해도 기능은 작동하지만, 파일 열기/저장 시 보안 대화 상자가 표시될 수 있습니다.

### 한글 연결 실패
한글 프로그램이 실행 중이지 않을 경우 연결에 실패할 수 있습니다. 한글 프로그램이 설치되어 있고 정상 작동하는지 확인하세요.

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