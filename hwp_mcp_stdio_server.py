#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import json
import traceback
import logging
import ssl
from threading import Thread

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    filename="hwp_mcp_stdio_server.log",
    filemode="a",
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)

# 추가 스트림 핸들러 설정 (별도로 추가)
stderr_handler = logging.StreamHandler(sys.stderr)
stderr_handler.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s"))
logger = logging.getLogger("hwp-mcp-stdio-server")
logger.addHandler(stderr_handler)

# Optional: Disable SSL certificate validation for development
ssl._create_default_https_context = ssl._create_unverified_context

# Set up paths
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)

try:
    # Import FastMCP library
    from mcp.server.fastmcp import FastMCP
    logger.info("FastMCP successfully imported")
except ImportError as e:
    logger.error(f"Failed to import FastMCP: {str(e)}")
    print(f"Error: Failed to import FastMCP. Please install with 'pip install mcp'", file=sys.stderr)
    sys.exit(1)

# Try to import HwpController
try:
    from src.tools.hwp_controller import HwpController
    logger.info("HwpController imported successfully")
except ImportError as e:
    logger.error(f"Failed to import HwpController: {str(e)}")
    # Try alternate paths
    try:
        sys.path.append(os.path.join(current_dir, "src"))
        sys.path.append(os.path.join(current_dir, "src", "tools"))
        from hwp_controller import HwpController
        logger.info("HwpController imported from alternate path")
    except ImportError as e2:
        logger.error(f"Could not find HwpController in any path: {str(e2)}")
        print(f"Error: Could not find HwpController module", file=sys.stderr)
        sys.exit(1)

# Initialize FastMCP server
mcp = FastMCP(
    "hwp-mcp",
    version="0.1.0",
    description="HWP MCP Server for controlling Hangul Word Processor",
    dependencies=["pywin32>=305"],
    env_vars={}
)

# Global HWP controller instance
hwp_controller = None

def get_hwp_controller():
    """Get or create HwpController instance."""
    global hwp_controller
    if hwp_controller is None:
        logger.info("Creating HwpController instance...")
        try:
            hwp_controller = HwpController()
            if not hwp_controller.connect(visible=True):
                logger.error("Failed to connect to HWP program")
                return None
            logger.info("Successfully connected to HWP program")
        except Exception as e:
            logger.error(f"Error creating HwpController: {str(e)}", exc_info=True)
            return None
    return hwp_controller

@mcp.tool()
def hwp_create() -> str:
    """Create a new HWP document."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        if hwp.create_new_document():
            logger.info("Successfully created new document")
            return "New document created successfully"
        else:
            return "Error: Failed to create new document"
    except Exception as e:
        logger.error(f"Error creating document: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_open(path: str) -> str:
    """Open an existing HWP document."""
    try:
        if not path:
            return "Error: File path is required"
        
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        if hwp.open_document(path):
            logger.info(f"Successfully opened document: {path}")
            return f"Document opened: {path}"
        else:
            return "Error: Failed to open document"
    except Exception as e:
        logger.error(f"Error opening document: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_save(path: str = None) -> str:
    """Save the current HWP document."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        if path:
            if hwp.save_document(path):
                logger.info(f"Successfully saved document to: {path}")
                return f"Document saved to: {path}"
            else:
                return "Error: Failed to save document"
        else:
            temp_path = os.path.join(os.getcwd(), "temp_document.hwp")
            if hwp.save_document(temp_path):
                logger.info(f"Successfully saved document to temporary location: {temp_path}")
                return f"Document saved to: {temp_path}"
            else:
                return "Error: Failed to save document"
    except Exception as e:
        logger.error(f"Error saving document: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_insert_text(text: str) -> str:
    """Insert text at the current cursor position."""
    try:
        if not text:
            return "Error: Text is required"
        
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        if hwp.insert_text(text):
            logger.info("Successfully inserted text")
            return "Text inserted successfully"
        else:
            return "Error: Failed to insert text"
    except Exception as e:
        logger.error(f"Error inserting text: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_set_font(
    name: str = None, 
    size: int = None, 
    bold: bool = False, 
    italic: bool = False, 
    underline: bool = False
) -> str:
    """Set font properties for selected text."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        if hwp.set_font(name, size, bold, italic):
            logger.info("Successfully set font")
            return "Font set successfully"
        else:
            return "Error: Failed to set font"
    except Exception as e:
        logger.error(f"Error setting font: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_insert_table(rows: int, cols: int) -> str:
    """Insert a table at the current cursor position."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        if hwp.insert_table(rows, cols):
            logger.info(f"Successfully inserted {rows}x{cols} table")
            return f"Table inserted with {rows} rows and {cols} columns"
        else:
            return "Error: Failed to insert table"
    except Exception as e:
        logger.error(f"Error inserting table: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_insert_paragraph() -> str:
    """Insert a new paragraph."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        if hwp.insert_paragraph():
            logger.info("Successfully inserted paragraph")
            return "Paragraph inserted successfully"
        else:
            return "Error: Failed to insert paragraph"
    except Exception as e:
        logger.error(f"Error inserting paragraph: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_get_text() -> str:
    """Get the text content of the current document."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        text = hwp.get_text()
        if text is not None:
            logger.info("Successfully retrieved document text")
            return text
        else:
            return "Error: Failed to get document text"
    except Exception as e:
        logger.error(f"Error getting text: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_close(save: bool = True) -> str:
    """Close the HWP document and connection."""
    try:
        global hwp_controller
        if hwp_controller and hwp_controller.is_hwp_running:
            if hwp_controller.disconnect():
                logger.info("Successfully closed HWP connection")
                hwp_controller = None
                return "HWP connection closed successfully"
            else:
                return "Error: Failed to close HWP connection"
        else:
            return "HWP is already closed"
    except Exception as e:
        logger.error(f"Error closing HWP: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_ping_pong(message: str = "핑") -> str:
    """
    핑퐁 테스트용 함수입니다. 핑을 보내면 퐁을 응답하고, 퐁을 보내면 핑을 응답합니다.
    
    Args:
        message: 테스트 메시지 (기본값: "핑")
        
    Returns:
        str: 응답 메시지
    """
    try:
        logger.info(f"핑퐁 테스트 함수 호출됨: 메시지 - {message}")
        
        # 메시지에 따라 응답 생성
        if message == "핑":
            response = "퐁"
        elif message == "퐁":
            response = "핑"
        else:
            response = f"모르는 메시지입니다: {message} (핑 또는 퐁을 보내주세요)"
        
        # 응답 생성 시간 기록
        import datetime
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # 응답 데이터 구성
        result = {
            "response": response,
            "original_message": message,
            "timestamp": current_time
        }
        
        return json.dumps(result, ensure_ascii=False)
    except Exception as e:
        logger.error(f"핑퐁 테스트 함수 오류: {str(e)}", exc_info=True)
        return f"테스트 오류 발생: {str(e)}"

@mcp.tool()
def hwp_create_table_pyhwpx(rows: int, cols: int, data: str = None, has_header: bool = False) -> str:
    """
    pyhwpx를 사용하여 현재 커서 위치에 표를 생성합니다.
    
    Args:
        rows: 표의 행 수
        cols: 표의 열 수
        data: 표에 채울 데이터 (JSON 형식의 2차원 배열 문자열, 예: '[["항목1", "항목2"], ["값1", "값2"]]')
        has_header: 첫 번째 행을 헤더로 처리할지 여부
        
    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        # pyhwpx 방식으로 표 생성
        js_code = f"""
        function createTableWithPyhwpx() {{
            var hwp = this;
            
            // 표 생성
            hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet);
            hwp.HParameterSet.HTableCreation.Rows = {rows};
            hwp.HParameterSet.HTableCreation.Cols = {cols};
            hwp.HParameterSet.HTableCreation.WidthType = 0;  // 단에 맞춤
            hwp.HParameterSet.HTableCreation.HeightType = 1;  // 절대값
            hwp.HParameterSet.HTableCreation.WidthValue = 0;
            hwp.HParameterSet.HTableCreation.HeightValue = 1000;
            
            // 열 너비 설정
            var colWidth = 8000 / {cols};  // 균등하게 분배
            hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", {cols});
            for (var i = 0; i < {cols}; i++) {{
                hwp.HParameterSet.HTableCreation.ColWidth.SetItem(i, colWidth);
            }}
            
            hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet);
            
            return "표 생성 완료";
        }}
        
        createTableWithPyhwpx();
        """
        
        # 표 생성 실행
        result = hwp.execute_script(js_code)
        logger.info(f"표 생성 결과: {result}")
        
        # 데이터가 제공된 경우 표 채우기
        if data:
            try:
                # JSON 문자열을 파이썬 객체로 변환
                data_array = json.loads(data)
                
                # 표의 크기 확인
                if len(data_array) > rows:
                    data_array = data_array[:rows]  # 행 수 초과 시 자르기
                
                # 각 셀에 데이터 입력
                for row_idx, row_data in enumerate(data_array):
                    if len(row_data) > cols:
                        row_data = row_data[:cols]  # 열 수 초과 시 자르기
                    
                    for col_idx, cell_data in enumerate(row_data):
                        # 셀 위치 계산 (List 인덱스는 0부터 시작, 표는 생성 직후 위치 찾기)
                        cell_list_idx = 3 + row_idx * cols + col_idx
                        
                        # 셀 데이터 입력을 위한 JavaScript 코드
                        cell_js = f"""
                        function fillCell() {{
                            var hwp = this;
                            
                            // 셀로 이동
                            hwp.SetPos({cell_list_idx}, 0, 0);
                            
                            // 셀 내용 선택 및 삭제
                            hwp.SelectAll();
                            hwp.Delete();
                            
                            // 텍스트 입력
                            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
                            hwp.HParameterSet.HInsertText.Text = "{str(cell_data)}";
                            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
                            
                            return "셀 데이터 입력 완료";
                        }}
                        
                        fillCell();
                        """
                        
                        # 셀 데이터 입력 실행
                        hwp.execute_script(cell_js)
            
                # 헤더 스타일 설정 (첫 번째 행이 헤더인 경우)
                if has_header and data_array and len(data_array) > 0:
                    for col_idx in range(min(cols, len(data_array[0]))):
                        header_js = f"""
                        function styleHeader() {{
                            var hwp = this;
                            
                            // 헤더 셀로 이동
                            hwp.SetPos({3 + col_idx}, 0, 0);
                            
                            // 셀 내용 선택
                            hwp.SelectAll();
                            
                            // 글자 굵게 설정
                            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet);
                            hwp.HParameterSet.HCharShape.Bold = 1;
                            hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet);
                            
                            // 가운데 정렬
                            hwp.HAction.GetDefault("ParaShape", hwp.HParameterSet.HParaShape.HSet);
                            hwp.HParameterSet.HParaShape.Align = 1;  // 가운데 정렬
                            hwp.HAction.Execute("ParaShape", hwp.HParameterSet.HParaShape.HSet);
                            
                            return "헤더 스타일 적용 완료";
                        }}
                        
                        styleHeader();
                        """
                        
                        # 헤더 스타일 적용 실행
                        hwp.execute_script(header_js)
                        
                return f"표 생성 및 데이터 입력 완료 ({rows}x{cols} 크기)"
            except Exception as data_error:
                logger.error(f"표 데이터 입력 중 오류: {str(data_error)}", exc_info=True)
                return f"표는 생성되었으나 데이터 입력 중 오류 발생: {str(data_error)}"
        
        return f"표 생성 완료 ({rows}x{cols} 크기)"
    except Exception as e:
        logger.error(f"pyhwpx 표 생성 중 오류: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_table_set_cell_text(row: int, col: int, text: str) -> str:
    """
    pyhwpx를 사용하여 표의 특정 셀에 텍스트를 입력합니다.
    
    Args:
        row: 셀의 행 번호 (1부터 시작)
        col: 셀의 열 번호 (1부터 시작)
        text: 입력할 텍스트
        
    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        # 셀 이동 및 텍스트 입력을 위한 JavaScript 코드
        js_code = f"""
        function setCellText() {{
            var hwp = this;
            
            // 문서의 현재 표를 찾아서 이동
            // 주의: 현재 커서 위치의 표를 기준으로 작동
            
            // 현재 표에서 특정 셀로 이동
            var moveSuccess = false;
            
            try {{
                // 표의 첫 번째 셀로 이동
                hwp.TableCellRefresh();
                hwp.HAction.Run("TableCellBlock");
                hwp.HAction.Run("Cancel");
                
                // 지정된 행으로 이동 (1행부터 시작)
                for (var r = 1; r < {row}; r++) {{
                    hwp.HAction.Run("TableLowerCell");
                }}
                
                // 지정된 열로 이동 (1열부터 시작)
                for (var c = 1; c < {col}; c++) {{
                    hwp.HAction.Run("TableRightCell");
                }}
                
                moveSuccess = true;
            }} catch (e) {{
                return "셀 이동 실패: " + e.message;
            }}
            
            if (!moveSuccess) {{
                return "지정된 셀({row}, {col})로 이동할 수 없습니다.";
            }}
            
            // 셀 내용 선택 및 삭제
            hwp.SelectAll();
            hwp.Delete();
            
            // 텍스트 입력
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
            hwp.HParameterSet.HInsertText.Text = "{text}";
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            
            return "셀 텍스트 설정 완료";
        }}
        
        setCellText();
        """
        
        # JavaScript 실행
        result = hwp.execute_script(js_code)
        logger.info(f"셀 텍스트 설정 결과: {result}")
        
        return f"셀({row}, {col})에 텍스트 입력 완료"
    except Exception as e:
        logger.error(f"셀 텍스트 설정 중 오류: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_table_merge_cells(start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    """
    pyhwpx를 사용하여 표의 특정 범위의 셀을 병합합니다.
    
    Args:
        start_row: 시작 행 번호 (1부터 시작)
        start_col: 시작 열 번호 (1부터 시작)
        end_row: 종료 행 번호 (1부터 시작)
        end_col: 종료 열 번호 (1부터 시작)
        
    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        # 셀 병합을 위한 JavaScript 코드
        js_code = f"""
        function mergeCells() {{
            var hwp = this;
            
            try {{
                // 시작 셀로 이동
                hwp.TableCellRefresh();
                hwp.HAction.Run("TableCellBlock");
                hwp.HAction.Run("Cancel");
                
                // 지정된 시작 행으로 이동
                for (var r = 1; r < {start_row}; r++) {{
                    hwp.HAction.Run("TableLowerCell");
                }}
                
                // 지정된 시작 열로 이동
                for (var c = 1; c < {start_col}; c++) {{
                    hwp.HAction.Run("TableRightCell");
                }}
                
                // 셀 선택 모드로 전환
                hwp.HAction.Run("TableCellBlock");
                
                // 종료 셀까지 선택 영역 확장
                var rowDiff = {end_row} - {start_row};
                var colDiff = {end_col} - {start_col};
                
                // 행 방향 확장
                for (var r = 0; r < rowDiff; r++) {{
                    hwp.HAction.Run("TableLowerCell");
                    hwp.HAction.Run("TableCellBlockExtend");
                }}
                
                // 열 방향 확장
                for (var c = 0; c < colDiff; c++) {{
                    hwp.HAction.Run("TableRightCell");
                    hwp.HAction.Run("TableCellBlockExtend");
                }}
                
                // 셀 병합 실행
                hwp.HAction.Run("TableMergeCell");
                
                return "셀 병합 완료";
            }} catch (e) {{
                return "셀 병합 실패: " + e.message;
            }}
        }}
        
        mergeCells();
        """
        
        # JavaScript 실행
        result = hwp.execute_script(js_code)
        logger.info(f"셀 병합 결과: {result}")
        
        return f"셀 병합 완료 ({start_row},{start_col}) - ({end_row},{end_col})"
    except Exception as e:
        logger.error(f"셀 병합 중 오류: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

if __name__ == "__main__":
    logger.info("Starting HWP MCP stdio server")
    try:
        # Run the FastMCP server with stdio transport
        mcp.run(transport="stdio")
    except Exception as e:
        logger.error(f"Error running server: {str(e)}", exc_info=True)
        print(f"Error: {str(e)}", file=sys.stderr)
        sys.exit(1) 