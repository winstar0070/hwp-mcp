# hwp-mcp

한글 워드 프로세서(HWP)를 Claude와 같은 AI 모델이 제어할 수 있도록 해주는 Model Context Protocol(MCP) 서버입니다.

## 개요

hwp-mcp는 한글 워드 프로세서(HWP)를 Claude와 같은 AI 모델이 제어할 수 있도록 해주는 Model Context Protocol(MCP) 서버입니다. 이 프로젝트는 한글 문서를 자동으로 생성, 편집, 조작하는 기능을 Claude에게 제공합니다.

## 설치 및 설정

1. 필수 라이브러리 설치:
   ```
   pip install -r requirements.txt
   ```

2. Claude 데스크톱 설정:
   - Claude 데스크톱 설정 파일을 다음과 같이 구성해야 합니다:
   ```json
   {
     "mcpServers": {
       "hwp": {
         "command": "python",
         "args": ["D:\\hwp-mcp\\hwp_mcp_stdio_server.py"]
       }
     }
   }
   ```
   - 경로는 실제 설치 위치에 맞게 수정하세요.

## 작동 방식

1. **Claude 호출 과정**:
   - Claude 데스크톱 애플리케이션에서 hwp-mcp를 호출하면 `hwp_mcp_stdio_server.py` 스크립트가 실행됩니다.
   - 이 스크립트는 표준 입출력(stdio)을 통해 Claude와 통신하는 FastMCP 서버를 생성합니다.
   - 서버는 HWP를 조작하기 위한 여러 도구(tool)들을 등록합니다.

2. **기능 실행 과정**:
   - Claude가 HWP 관련 기능을 요청하면, 등록된 도구가 호출됩니다.
   - 도구는 `HwpController`를 사용하여 Windows COM 자동화를 통해 HWP 프로그램과 상호작용합니다.
   - `win32com` 라이브러리를 통해 HWP에 명령을 전송합니다.

## 주요 기능

- 문서 생성 및 열기
- 텍스트 삽입 및 편집
- 표 삽입 및 조작
- 서식 지정
- 문서 저장 및 닫기

## 프로젝트 구조

- **hwp_mcp_stdio_server.py**: 메인 서버 실행 파일
- **hwp-mcp/**: 주요 모듈 디렉토리
  - **src/tools/hwp_controller.py**: HWP와 상호작용하는 핵심 클래스
  - **src/utils/**: 유틸리티 모듈
  - **security_module/**: 보안 모듈

## 요구사항

- Python 3.8 이상
- 한글 워드 프로세서(HWP) 설치됨
- pywin32
- 기타 requirements.txt에 명시된 패키지 