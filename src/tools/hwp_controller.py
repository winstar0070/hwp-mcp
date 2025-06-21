"""
한글(HWP) 문서를 제어하기 위한 컨트롤러 모듈
win32com을 이용하여 한글 프로그램을 자동화합니다.
"""

import os
import win32com.client
import win32gui
import win32con
import time
import logging
from typing import Optional, List, Dict, Any, Tuple

logger = logging.getLogger(__name__)
try:
    from .hwp_exceptions import (
        HwpError, HwpConnectionError, HwpNotRunningError, HwpDocumentError,
        HwpDocumentNotFoundError, HwpDocumentAccessError, HwpDocumentSaveError,
        HwpTableError, HwpTableCellError, HwpTableRangeError, HwpTableNotFoundError,
        HwpImageError, HwpImageNotFoundError, HwpImageFormatError,
        HwpInvalidParameterError, HwpOperationError,
        handle_hwp_error
    )
    from .constants import (
        HWPUNIT_PER_PT, TABLE_MAX_ROWS, TABLE_MAX_COLS,
        TABLE_DEFAULT_WIDTH, TABLE_DEFAULT_HEIGHT,
        ALLOWED_IMAGE_FORMATS, IMAGE_EMBED_MODE,
        SECURITY_MODULE_NAME, SECURITY_MODULE_DEFAULT_PATH
    )
    from .config import get_config
    from .hwp_utils import (
        require_hwp_connection, safe_hwp_operation,
        set_font_properties, move_to_table_cell,
        validate_table_coordinates, log_operation_result
    )
except ImportError:
    # 예외 클래스가 없는 경우 기본 Exception 사용
    HwpError = Exception
    HwpConnectionError = Exception
    HwpNotRunningError = Exception
    HwpDocumentError = Exception
    HwpDocumentNotFoundError = Exception
    HwpDocumentAccessError = Exception
    HwpDocumentSaveError = Exception
    HwpTableError = Exception
    HwpTableCellError = Exception
    HwpTableRangeError = Exception
    HwpTableNotFoundError = Exception
    HwpImageError = Exception
    HwpImageNotFoundError = Exception
    HwpImageFormatError = Exception
    HwpInvalidParameterError = Exception
    HwpOperationError = Exception
    handle_hwp_error = lambda x: x
    
    # 상수 기본값
    HWPUNIT_PER_PT = 100
    TABLE_MAX_ROWS = 100
    
    # 유틸 함수 더미
    def require_hwp_connection(func):
        return func
    def safe_hwp_operation(name, default=None):
        def decorator(func):
            return func
        return decorator
    def set_font_properties(*args, **kwargs):
        return False
    def move_to_table_cell(*args, **kwargs):
        return False
    def validate_table_coordinates(*args, **kwargs):
        pass
    def log_operation_result(*args, **kwargs):
        pass
    TABLE_MAX_COLS = 100
    TABLE_DEFAULT_WIDTH = 8000
    TABLE_DEFAULT_HEIGHT = 1000
    ALLOWED_IMAGE_FORMATS = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff']
    IMAGE_EMBED_MODE = 1
    SECURITY_MODULE_NAME = "FilePathCheckerModuleExample"
    SECURITY_MODULE_DEFAULT_PATH = "D:/hwp-mcp/security_module/FilePathCheckerModuleExample.dll"
    get_config = lambda: None


class HwpController:
    """한글 문서를 제어하는 클래스"""

    def __init__(self):
        """한글 애플리케이션 인스턴스를 초기화합니다."""
        self.hwp = None
        self.visible = True
        self.is_hwp_running = False
        self.current_document_path = None

    def connect(self, visible: bool = True, register_security_module: bool = True) -> bool:
        """
        한글 프로그램에 연결합니다.
        
        Args:
            visible (bool): 한글 창을 화면에 표시할지 여부
            register_security_module (bool): 보안 모듈을 등록할지 여부
            
        Returns:
            bool: 연결 성공 여부
        """
        # HWP 객체 생성
        try:
            self.hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        except OSError as e:
            raise HwpConnectionError(f"한글 프로그램 COM 객체 생성 실패: {e} (HWP가 설치되지 않았거나 등록되지 않았습니다)")
        
        # 보안 모듈 등록
        if register_security_module:
            self._register_security_module()
        
        # 창 표시 설정
        try:
            self.visible = visible
            self.hwp.XHwpWindows.Item(0).Visible = visible
            self.is_hwp_running = True
            return True
        except AttributeError as e:
            raise HwpConnectionError(f"한글 프로그램 창 설정 실패: {e} (HWP API 호환성 문제)")
        except Exception as e:
            raise HwpConnectionError(f"한글 프로그램 초기화 실패: {e}")
    
    def _register_security_module(self):
        """보안 모듈을 등록합니다. (파일 경로 체크 보안 경고창 방지)"""
        try:
            config = get_config()
            module_path = os.path.abspath(config.security_module_path if config else SECURITY_MODULE_DEFAULT_PATH)
            self.hwp.RegisterModule(SECURITY_MODULE_NAME, module_path)
            logger.info(f"보안 모듈이 등록되었습니다: {module_path}")
        except FileNotFoundError:
            logger.warning(f"보안 모듈 파일을 찾을 수 없습니다: {module_path} (보안 경고가 나타날 수 있습니다)")
        except AttributeError as e:
            logger.warning(f"보안 모듈 등록 API 호출 실패: {e} (HWP 버전이 오래되었을 수 있습니다)")
        except Exception as e:
            logger.warning(f"보안 모듈 등록 중 예상치 못한 오류: {e} (계속 진행합니다)")

    def disconnect(self) -> bool:
        """
        한글 프로그램 연결을 종료합니다.
        
        Returns:
            bool: 종료 성공 여부
        """
        try:
            if self.is_hwp_running:
                # HwpObject를 해제합니다
                self.hwp = None
                self.is_hwp_running = False
                
            return True
        except AttributeError as e:
            logger.error(f"HWP 객체 해제 중 속성 오류: {e}")
            return False
        except Exception as e:
            logger.error(f"한글 프로그램 연결 해제 중 예상치 못한 오류: {e}")
            return False

    def create_new_document(self) -> bool:
        """
        새 문서를 생성합니다.
        
        Returns:
            bool: 생성 성공 여부
        """
        try:
            if not self.is_hwp_running:
                self.connect()
            
            self.hwp.Run("FileNew")
            self.current_document_path = None
            return True
        except HwpConnectionError:
            raise
        except AttributeError as e:
            logger.error(f"HWP API 메서드 호출 실패: {e}")
            return False
        except OSError as e:
            logger.error(f"새 문서 생성 중 시스템 오류: {e}")
            return False
        except Exception as e:
            logger.error(f"새 문서 생성 중 예상치 못한 오류: {e}")
            return False

    def open_document(self, file_path: str) -> bool:
        """
        문서를 엽니다.
        
        Args:
            file_path (str): 열 문서의 경로
            
        Returns:
            bool: 열기 성공 여부
        """
        try:
            if not self.is_hwp_running:
                raise HwpNotRunningError()
            
            abs_path = os.path.abspath(file_path)
            
            # 파일 존재 확인
            if not os.path.exists(abs_path):
                raise HwpDocumentNotFoundError(abs_path)
            
            result = self.hwp.Open(abs_path)
            if result:
                self.current_document_path = abs_path
                return True
            else:
                raise HwpDocumentAccessError(abs_path)
        except (HwpConnectionError, HwpNotRunningError, HwpDocumentError):
            raise
        except Exception as e:
            raise HwpDocumentError(f"문서 열기 오류: {e}")

    def save_document(self, file_path: Optional[str] = None) -> bool:
        """
        문서를 저장합니다.
        
        Args:
            file_path (str, optional): 저장할 경로. None이면 현재 경로에 저장.
            
        Returns:
            bool: 저장 성공 여부
        """
        try:
            if not self.is_hwp_running:
                raise HwpNotRunningError()
            
            if file_path:
                abs_path = os.path.abspath(file_path)
                
                # 디렉토리 존재 확인 및 생성
                dir_path = os.path.dirname(abs_path)
                if not os.path.exists(dir_path):
                    os.makedirs(dir_path)
                
                # 파일 형식과 경로 모두 지정하여 저장
                result = self.hwp.SaveAs(abs_path, "HWP", "")
                if result:
                    self.current_document_path = abs_path
                    return True
                else:
                    raise HwpDocumentSaveError(abs_path)
            else:
                if self.current_document_path:
                    result = self.hwp.Save()
                    if not result:
                        raise HwpDocumentSaveError(self.current_document_path)
                else:
                    # 저장 대화 상자 표시 (파라미터 없이 호출)
                    self.hwp.SaveAs()
                    # 대화 상자에서 사용자가 선택한 경로를 알 수 없으므로 None 유지
            
            return True
        except (HwpConnectionError, HwpNotRunningError, HwpDocumentError):
            raise
        except OSError as e:
            raise HwpDocumentSaveError(file_path or self.current_document_path or "현재 문서", f"파일 시스템 오류: {e}")
        except AttributeError as e:
            logger.error(f"HWP API 호출 오류: {e}")
            raise HwpDocumentSaveError(file_path or self.current_document_path or "현재 문서", f"API 호출 실패: {e}")
        except Exception as e:
            raise HwpDocumentSaveError(file_path or self.current_document_path or "현재 문서", str(e))

    def insert_text(self, text: str, preserve_linebreaks: bool = True) -> bool:
        """
        현재 커서 위치에 텍스트를 삽입합니다.
        
        Args:
            text (str): 삽입할 텍스트
            preserve_linebreaks (bool): 줄바꿈 유지 여부
            
        Returns:
            bool: 삽입 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            if preserve_linebreaks and '\n' in text:
                # 줄바꿈이 포함된 경우 줄 단위로 처리
                lines = text.split('\n')
                for i, line in enumerate(lines):
                    if i > 0:  # 첫 줄이 아니면 줄바꿈 추가
                        self.insert_paragraph()
                    if line.strip():  # 빈 줄이 아니면 텍스트 삽입
                        self._insert_text_direct(line)
                return True
            else:
                # 줄바꿈이 없거나 유지하지 않는 경우 한 번에 처리
                return self._insert_text_direct(text)
        except AttributeError as e:
            logger.error(f"텍스트 삽입 API 호출 실패: {e} (HWP 연결 상태를 확인하세요)")
            return False
        except Exception as e:
            logger.error(f"텍스트 삽입 중 예상치 못한 오류: {e}")
            return False

    def _set_table_cursor(self) -> bool:
        """
        표 안에서 커서 위치를 제어하는 내부 메서드입니다.
        현재 셀을 선택하고 취소하여 커서를 셀 안에 위치시킵니다.
        
        Returns:
            bool: 성공 여부
        """
        try:
            # 현재 셀 선택
            self.hwp.Run("TableSelCell")
            # 선택 취소 (커서는 셀 안에 위치)
            self.hwp.Run("Cancel")
            # 셀 내부로 커서 이동을 확실히
            self.hwp.Run("CharRight")
            self.hwp.Run("CharLeft")
            return True
        except Exception as e:
            logger.error(f"표 셀 선택 실패: {str(e)}")
            return False

    def _insert_text_direct(self, text: str) -> bool:
        """
        텍스트를 직접 삽입하는 내부 메서드입니다.
        
        Args:
            text (str): 삽입할 텍스트
            
        Returns:
            bool: 삽입 성공 여부
        """
        try:
            # 텍스트 삽입을 위한 액션 초기화
            self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            self.hwp.HParameterSet.HInsertText.Text = text
            self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            return True
        except AttributeError as e:
            logger.error(f"InsertText 액션 실행 실패: {e} (HWP API 버전을 확인하세요)")
            return False
        except Exception as e:
            logger.error(f"텍스트 직접 삽입 중 예상치 못한 오류: {e}")
            return False

    def set_font(self, font_name: str, font_size: int, bold: bool = False, italic: bool = False, 
                select_previous_text: bool = False) -> bool:
        """
        글꼴 속성을 설정합니다. 현재 위치에서 다음에 입력할 텍스트에 적용됩니다.
        
        Args:
            font_name (str): 글꼴 이름
            font_size (int): 글꼴 크기
            bold (bool): 굵게 여부
            italic (bool): 기울임꼴 여부
            select_previous_text (bool): 이전에 입력한 텍스트를 선택할지 여부
            
        Returns:
            bool: 설정 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 새로운 구현: set_font_style 메서드 사용
            return self.set_font_style(
                font_name=font_name,
                font_size=font_size,
                bold=bold,
                italic=italic,
                underline=False,
                select_previous_text=select_previous_text
            )
        except AttributeError as e:
            logger.error(f"글꼴 설정 API 호출 실패: {e}")
            return False
        except Exception as e:
            logger.error(f"글꼴 설정 중 예상치 못한 오류: {e}")
            return False

    def set_font_style(self, font_name: str = None, font_size: int = None, 
                     bold: bool = False, italic: bool = False, underline: bool = False,
                     select_previous_text: bool = False) -> bool:
        """
        현재 선택된 텍스트의 글꼴 스타일을 설정합니다.
        선택된 텍스트가 없으면, 다음 입력될 텍스트에 적용됩니다.
        
        Args:
            font_name (str, optional): 글꼴 이름. None이면 현재 글꼴 유지.
            font_size (int, optional): 글꼴 크기. None이면 현재 크기 유지.
            bold (bool): 굵게 여부
            italic (bool): 기울임꼴 여부
            underline (bool): 밑줄 여부
            select_previous_text (bool): 이전에 입력한 텍스트를 선택할지 여부
            
        Returns:
            bool: 설정 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 이전 텍스트 선택 옵션이 활성화된 경우 현재 단락의 이전 텍스트 선택
            if select_previous_text:
                self.select_last_text()
            
            # hwp_utils의 공통 함수 사용
            return set_font_properties(
                self.hwp,
                font_name=font_name,
                font_size=font_size,
                bold=bold,
                italic=italic,
                underline=underline
            )
            
        except AttributeError as e:
            logger.error(f"글꼴 스타일 API 호출 실패: {e}")
            return False
        except Exception as e:
            logger.error(f"글꼴 스타일 설정 중 예상치 못한 오류: {e}")
            return False

    def _get_current_position(self):
        """
        현재 커서 위치 정보를 가져옵니다.
        
        Returns:
            tuple or None: (위치 유형, List ID, Para ID, CharPos) 튜플 또는 실패 시 None
        """
        try:
            # GetPos()는 현재 위치 정보를 (위치 유형, List ID, Para ID, CharPos)의 튜플로 반환
            return self.hwp.GetPos()
        except Exception as e:
            logger.error(f"현재 위치 정보 가져오기 실패: {str(e)}")
            # 실패 시 None 반환
            return None

    def _set_position(self, pos):
        """
        커서 위치를 지정된 위치로 변경합니다.
        
        Args:
            pos: GetPos()로 얻은 위치 정보 튜플
            
        Returns:
            bool: 성공 여부
        """
        try:
            if pos:
                self.hwp.SetPos(*pos)
            return True
        except TypeError as e:
            logger.error(f"위치 설정 실패 - 잘못된 매개변수 타입: {str(e)}")
            return False
        except AttributeError as e:
            logger.error(f"위치 설정 실패 - HWP API 호출 오류: {str(e)}")
            return False
        except Exception as e:
            logger.error(f"위치 설정 실패: {str(e)}")
            return False

    def insert_table(self, rows: int, cols: int) -> bool:
        """
        현재 커서 위치에 표를 삽입합니다.
        
        Args:
            rows (int): 행 수
            cols (int): 열 수
            
        Returns:
            bool: 삽입 성공 여부
        """
        try:
            if not self.is_hwp_running:
                raise HwpNotRunningError()
            
            # 매개변수 유효성 검사
            if rows <= 0 or cols <= 0:
                raise HwpInvalidParameterError("rows/cols", f"rows={rows}, cols={cols}", "1 이상의 정수")
            
            if rows > TABLE_MAX_ROWS or cols > TABLE_MAX_COLS:
                raise HwpInvalidParameterError("rows/cols", f"rows={rows}, cols={cols}", f"{TABLE_MAX_ROWS} 이하의 정수")
            
            self.hwp.HAction.GetDefault("TableCreate", self.hwp.HParameterSet.HTableCreation.HSet)
            self.hwp.HParameterSet.HTableCreation.Rows = rows
            self.hwp.HParameterSet.HTableCreation.Cols = cols
            self.hwp.HParameterSet.HTableCreation.WidthType = 0  # 0: 단에 맞춤, 1: 절대값
            self.hwp.HParameterSet.HTableCreation.HeightType = 1  # 0: 자동, 1: 절대값
            self.hwp.HParameterSet.HTableCreation.WidthValue = 0  # 단에 맞춤이므로 무시됨
            self.hwp.HParameterSet.HTableCreation.HeightValue = TABLE_DEFAULT_HEIGHT  # 셀 높이(hwpunit)
            
            # 각 열의 너비를 설정 (모두 동일하게)
            # PageWidth 대신 고정 값 사용
            col_width = TABLE_DEFAULT_WIDTH // cols  # 전체 너비를 열 수로 나눔
            self.hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", cols)
            for i in range(cols):
                self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(i, col_width)
                
            self.hwp.HAction.Execute("TableCreate", self.hwp.HParameterSet.HTableCreation.HSet)
            return True
        except (HwpConnectionError, HwpNotRunningError, HwpInvalidParameterError):
            raise
        except AttributeError as e:
            raise HwpTableError(f"HWP API 호출 오류: {e} (표 생성 기능을 지원하지 않는 HWP 버전일 수 있습니다)")
        except ValueError as e:
            raise HwpInvalidParameterError("rows/cols", f"{rows}x{cols}", "양의 정수")
        except Exception as e:
            raise HwpTableError(f"표 삽입 실패: {e}")

    def insert_image(self, image_path: str, width: int = 0, height: int = 0) -> bool:
        """
        현재 커서 위치에 이미지를 삽입합니다.
        
        Args:
            image_path (str): 이미지 파일 경로
            width (int): 이미지 너비(0이면 원본 크기)
            height (int): 이미지 높이(0이면 원본 크기)
            
        Returns:
            bool: 삽입 성공 여부
        """
        try:
            if not self.is_hwp_running:
                raise HwpNotRunningError()
            
            abs_path = os.path.abspath(image_path)
            if not os.path.exists(abs_path):
                raise HwpImageNotFoundError(abs_path)
            
            # 허용된 이미지 형식 확인
            ext = os.path.splitext(abs_path)[1].lower()
            if ext not in ALLOWED_IMAGE_FORMATS:
                raise HwpImageFormatError(ext)
                
            self.hwp.HAction.GetDefault("InsertPicture", self.hwp.HParameterSet.HInsertPicture.HSet)
            self.hwp.HParameterSet.HInsertPicture.FileName = abs_path
            self.hwp.HParameterSet.HInsertPicture.Width = width
            self.hwp.HParameterSet.HInsertPicture.Height = height
            self.hwp.HParameterSet.HInsertPicture.Embed = IMAGE_EMBED_MODE  # 0: 링크, 1: 파일 포함
            self.hwp.HAction.Execute("InsertPicture", self.hwp.HParameterSet.HInsertPicture.HSet)
            return True
        except (HwpConnectionError, HwpNotRunningError, HwpImageError):
            raise
        except Exception as e:
            raise HwpImageError(f"이미지 삽입 실패: {e}")

    def find_text(self, text: str) -> bool:
        """
        문서에서 텍스트를 찾습니다.
        
        Args:
            text (str): 찾을 텍스트
            
        Returns:
            bool: 찾기 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 간단한 매크로 명령 사용
            self.hwp.Run("MoveDocBegin")  # 문서 처음으로 이동
            
            # 찾기 명령 실행 (매크로 사용)
            result = self.hwp.Run(f'FindText "{text}" 1')  # 1=정방향검색
            return result  # True 또는 False 반환
        except AttributeError as e:
            logger.error(f"텍스트 찾기 API 호출 실패: {e}")
            return False
        except Exception as e:
            logger.error(f"텍스트 찾기 중 예상치 못한 오류: {e}")
            return False

    def replace_text(self, find_text: str, replace_text: str, replace_all: bool = False) -> bool:
        """
        문서에서 텍스트를 찾아 바꿉니다.
        
        Args:
            find_text (str): 찾을 텍스트
            replace_text (str): 바꿀 텍스트
            replace_all (bool): 모두 바꾸기 여부
            
        Returns:
            bool: 바꾸기 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 매크로 명령 사용
            self.hwp.Run("MoveDocBegin")  # 문서 처음으로 이동
            
            if replace_all:
                # 모두 바꾸기 명령 실행
                result = self.hwp.Run(f'ReplaceAll "{find_text}" "{replace_text}" 0 0 0 0 0 0')
                return bool(result)
            else:
                # 하나만 바꾸기 (찾고 바꾸기)
                found = self.hwp.Run(f'FindText "{find_text}" 1')
                if found:
                    result = self.hwp.Run(f'Replace "{replace_text}"')
                    return bool(result)
                return False
        except AttributeError as e:
            logger.error(f"텍스트 바꾸기 API 호출 실패: {e}")
            return False
        except Exception as e:
            logger.error(f"텍스트 바꾸기 중 예상치 못한 오류: {e}")
            return False

    def get_text(self) -> str:
        """
        현재 문서의 전체 텍스트를 가져옵니다.
        
        Returns:
            str: 문서 텍스트
        """
        try:
            if not self.is_hwp_running:
                return ""
            
            return self.hwp.GetTextFile("TEXT", "")
        except AttributeError as e:
            logger.error(f"텍스트 가져오기 API 호출 실패: {e}")
            return ""
        except Exception as e:
            logger.error(f"텍스트 가져오기 중 예상치 못한 오류: {e}")
            return ""

    def set_page_setup(self, orientation: str = "portrait", margin_left: int = 1000, 
                     margin_right: int = 1000, margin_top: int = 1000, margin_bottom: int = 1000) -> bool:
        """
        페이지 설정을 변경합니다.
        
        Args:
            orientation (str): 용지 방향 ('portrait' 또는 'landscape')
            margin_left (int): 왼쪽 여백(hwpunit)
            margin_right (int): 오른쪽 여백(hwpunit)
            margin_top (int): 위쪽 여백(hwpunit)
            margin_bottom (int): 아래쪽 여백(hwpunit)
            
        Returns:
            bool: 설정 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 매크로 명령 사용
            orient_val = 0 if orientation.lower() == "portrait" else 1
            
            # 페이지 설정 매크로
            result = self.hwp.Run(f"PageSetup3 {orient_val} {margin_left} {margin_right} {margin_top} {margin_bottom}")
            return bool(result)
        except AttributeError as e:
            logger.error(f"페이지 설정 API 호출 실패: {e}")
            return False
        except ValueError as e:
            logger.error(f"페이지 설정 값 오류: {e} (orientation은 'portrait' 또는 'landscape'여야 합니다)")
            return False
        except Exception as e:
            logger.error(f"페이지 설정 중 예상치 못한 오류: {e}")
            return False

    def insert_paragraph(self) -> bool:
        """
        새 단락을 삽입합니다.
        
        Returns:
            bool: 삽입 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            self.hwp.HAction.Run("BreakPara")
            return True
        except AttributeError as e:
            logger.error(f"단락 삽입 API 호출 실패: {e}")
            return False
        except Exception as e:
            logger.error(f"단락 삽입 중 예상치 못한 오류: {e}")
            return False

    def select_all(self) -> bool:
        """
        문서 전체를 선택합니다.
        
        Returns:
            bool: 선택 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            self.hwp.Run("SelectAll")
            return True
        except AttributeError as e:
            logger.error(f"전체 선택 API 호출 실패: {e}")
            return False
        except Exception as e:
            logger.error(f"전체 선택 중 예상치 못한 오류: {e}")
            return False

    def fill_cell_field(self, field_name: str, value: str, n: int = 1) -> bool:
        """
        동일한 이름의 셀필드 중 n번째에만 값을 채웁니다.
        위키독스 예제: https://wikidocs.net/261646
        
        Args:
            field_name (str): 필드 이름
            value (str): 채울 값
            n (int): 몇 번째 필드에 값을 채울지 (1부터 시작)
            
        Returns:
            bool: 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
                
            # 1. 필드 목록 가져오기
            # HGO_GetFieldList은 현재 문서에 있는 모든 필드 목록을 가져옵니다.
            self.hwp.HAction.GetDefault("HGo_GetFieldList", self.hwp.HParameterSet.HGo.HSet)
            self.hwp.HAction.Execute("HGo_GetFieldList", self.hwp.HParameterSet.HGo.HSet)
            
            # 2. 필드 이름이 동일한 모든 셀필드 찾기
            field_list = []
            field_count = self.hwp.HParameterSet.HGo.FieldList.Count
            
            for i in range(field_count):
                field_info = self.hwp.HParameterSet.HGo.FieldList.Item(i)
                if field_info.FieldName == field_name:
                    field_list.append((field_info.FieldName, i))
            
            # 3. n번째 필드가 존재하는지 확인 (인덱스는 0부터 시작하므로 n-1)
            if len(field_list) < n:
                print(f"해당 이름의 필드가 충분히 없습니다. 필요: {n}, 존재: {len(field_list)}")
                return False
                
            # 4. n번째 필드의 위치로 이동
            target_field_idx = field_list[n-1][1]
            
            # HGo_SetFieldText를 사용하여 해당 필드 위치로 이동한 후 텍스트 설정
            self.hwp.HAction.GetDefault("HGo_SetFieldText", self.hwp.HParameterSet.HGo.HSet)
            self.hwp.HParameterSet.HGo.HSet.SetItem("FieldIdx", target_field_idx)
            self.hwp.HParameterSet.HGo.HSet.SetItem("Text", value)
            self.hwp.HAction.Execute("HGo_SetFieldText", self.hwp.HParameterSet.HGo.HSet)
            
            return True
        except AttributeError as e:
            logger.error(f"셀필드 API 호출 실패: {e}")
            return False
        except IndexError as e:
            logger.error(f"셀필드 인덱스 오류: {e} (n={n}이 유효한 범위를 벗어났습니다)")
            return False
        except Exception as e:
            logger.error(f"셀필드 값 채우기 중 예상치 못한 오류: {e}")
            return False
        
    def select_last_text(self) -> bool:
        """
        현재 단락의 마지막으로 입력된 텍스트를 선택합니다.
        
        Returns:
            bool: 선택 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 현재 위치 저장
            current_pos = self.hwp.GetPos()
            if not current_pos:
                return False
                
            # 현재 단락의 시작으로 이동
            self.hwp.Run("MoveLineStart")
            start_pos = self.hwp.GetPos()
            
            # 이전 위치로 돌아가서 선택 영역 생성
            self.hwp.SetPos(*start_pos)
            self.hwp.SelectText(start_pos, current_pos)
            
            return True
        except AttributeError as e:
            logger.error(f"텍스트 선택 API 호출 실패: {e}")
            return False
        except Exception as e:
            logger.error(f"텍스트 선택 중 예상치 못한 오류: {e}")
            return False

    def fill_table_with_data(self, data: List[List[str]], start_row: int = 1, start_col: int = 1, has_header: bool = False) -> bool:
        """
        현재 커서 위치의 표에 데이터를 채웁니다.
        
        Args:
            data (List[List[str]]): 채울 데이터 2차원 리스트 (행 x 열)
            start_row (int): 시작 행 번호 (1부터 시작)
            start_col (int): 시작 열 번호 (1부터 시작)
            has_header (bool): 첫 번째 행을 헤더로 처리할지 여부
            
        Returns:
            bool: 작업 성공 여부
        """
        if not self.is_hwp_running:
            return False
        
        try:
            # 현재 위치 저장 (나중에 복원을 위해)
            original_pos = self.hwp.GetPos()
        except Exception as e:
            logger.warning(f"현재 위치 저장 실패: {e}")
            original_pos = None
        
        try:
            # 표의 시작 위치로 이동
            if not self._move_to_table_start(start_row, start_col):
                return False
            
            # 데이터 채우기
            for row_idx, row_data in enumerate(data):
                if not self._fill_table_row(row_data, row_idx, has_header):
                    logger.error(f"{row_idx + 1}번째 행 처리 실패")
                    return False
                    
                # 다음 행으로 이동 (마지막 행이 아닌 경우)
                if row_idx < len(data) - 1:
                    if not self._move_to_next_row(len(row_data)):
                        return False
            
            # 표 밖으로 커서 이동
            return self._exit_table()
            
        except Exception as e:
            logger.error(f"표 데이터 채우기 중 예상치 못한 오류: {e}")
            return False
    
    def _move_to_table_start(self, start_row: int, start_col: int) -> bool:
        """표의 시작 위치로 이동합니다."""
        try:
            self.hwp.Run("TableSelCell")  # 현재 셀 선택
            self.hwp.Run("TableSelTable") # 표 전체 선택
            self.hwp.Run("Cancel")        # 선택 취소 (커서는 표의 시작 부분에 위치)
            self.hwp.Run("TableSelCell")  # 첫 번째 셀 선택
            self.hwp.Run("Cancel")        # 선택 취소
            
            # 시작 위치로 이동
            for _ in range(start_row - 1):
                self.hwp.Run("TableLowerCell")
                
            for _ in range(start_col - 1):
                self.hwp.Run("TableRightCell")
                
            return True
        except Exception as e:
            logger.error(f"표 시작 위치로 이동 실패: {e}")
            return False
    
    def _fill_table_row(self, row_data: List[str], row_idx: int, has_header: bool) -> bool:
        """표의 한 행을 채웁니다."""
        try:
            for col_idx, cell_value in enumerate(row_data):
                # 셀 선택 및 내용 삭제
                self.hwp.Run("TableSelCell")
                self.hwp.Run("Delete")
                
                # 셀에 값 입력
                if has_header and row_idx == 0:
                    self.set_font_style(bold=True)
                    self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
                    self.hwp.HParameterSet.HInsertText.Text = cell_value
                    self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
                    self.set_font_style(bold=False)
                else:
                    self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
                    self.hwp.HParameterSet.HInsertText.Text = cell_value
                    self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
                
                # 다음 셀로 이동 (마지막 셀이 아닌 경우)
                if col_idx < len(row_data) - 1:
                    self.hwp.Run("TableRightCell")
                    
            return True
        except Exception as e:
            logger.error(f"표 행 채우기 실패: {e}")
            return False
    
    def _move_to_next_row(self, col_count: int) -> bool:
        """다음 행의 첫 번째 셀로 이동합니다."""
        try:
            for _ in range(col_count - 1):
                self.hwp.Run("TableLeftCell")
            self.hwp.Run("TableLowerCell")
            return True
        except Exception as e:
            logger.error(f"다음 행으로 이동 실패: {e}")
            return False
    
    def _exit_table(self) -> bool:
        """표 밖으로 커서를 이동합니다."""
        try:
            self.hwp.Run("TableSelCell")  # 현재 셀 선택
            self.hwp.Run("Cancel")        # 선택 취소
            self.hwp.Run("MoveDown")      # 아래로 이동
            return True
        except Exception as e:
            logger.error(f"표 밖으로 이동 실패: {e}")
            return False
    
    def insert_text_with_font(self, text: str, font_name: str = None, font_size: int = None, 
                             bold: bool = False, italic: bool = False, underline: bool = False) -> bool:
        """
        서식이 적용된 텍스트를 삽입합니다.
        먼저 서식을 설정한 후 텍스트를 입력하는 방식으로 작동합니다.
        
        Args:
            text (str): 삽입할 텍스트
            font_name (str, optional): 글꼴 이름
            font_size (int, optional): 글꼴 크기 (포인트 단위)
            bold (bool): 굵게 여부
            italic (bool): 기울임꼴 여부
            underline (bool): 밑줄 여부
            
        Returns:
            bool: 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 1단계: 먼저 글자 서식을 설정
            if font_name or font_size or bold or italic or underline:
                # CharShape를 사용하여 다음에 입력될 텍스트의 서식 설정
                self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
                
                # 글꼴 이름 설정
                if font_name:
                    self.hwp.HParameterSet.HCharShape.FaceNameHangul = font_name
                    self.hwp.HParameterSet.HCharShape.FaceNameLatin = font_name
                    self.hwp.HParameterSet.HCharShape.FaceNameHanja = font_name
                    self.hwp.HParameterSet.HCharShape.FaceNameJapanese = font_name
                    self.hwp.HParameterSet.HCharShape.FaceNameOther = font_name
                    self.hwp.HParameterSet.HCharShape.FaceNameSymbol = font_name
                    self.hwp.HParameterSet.HCharShape.FaceNameUser = font_name
                
                # 글꼴 크기 설정 (hwpunit, 1pt = HWPUNIT_PER_PT)
                if font_size:
                    self.hwp.HParameterSet.HCharShape.Height = font_size * HWPUNIT_PER_PT
                
                # 스타일 설정
                self.hwp.HParameterSet.HCharShape.Bold = 1 if bold else 0
                self.hwp.HParameterSet.HCharShape.Italic = 1 if italic else 0
                self.hwp.HParameterSet.HCharShape.UnderlineType = 1 if underline else 0
                
                # 서식 적용
                self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
            
            # 2단계: 텍스트 삽입
            return self._insert_text_direct(text)
            
        except AttributeError as e:
            logger.error(f"서식 텍스트 삽입 API 호출 실패: {e}")
            return False
        except Exception as e:
            logger.error(f"서식이 적용된 텍스트 삽입 중 예상치 못한 오류: {e}")
            return False
    
    def apply_font_to_selection(self, font_name: str = None, font_size: int = None, 
                               bold: bool = False, italic: bool = False, underline: bool = False) -> bool:
        """
        현재 선택된 텍스트에 서식을 적용합니다.
        
        Args:
            font_name (str, optional): 글꼴 이름
            font_size (int, optional): 글꼴 크기 (포인트 단위)
            bold (bool): 굵게 여부
            italic (bool): 기울임꼴 여부
            underline (bool): 밑줄 여부
            
        Returns:
            bool: 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 선택된 텍스트가 있는지 확인
            selected_text = self.hwp.GetSelectedText()
            if not selected_text:
                print("선택된 텍스트가 없습니다.")
                return False
            
            # set_font_style 메서드를 사용하여 서식 적용
            return self.set_font_style(
                font_name=font_name,
                font_size=font_size,
                bold=bold,
                italic=italic,
                underline=underline,
                select_previous_text=False  # 이미 선택된 상태이므로 False
            )
            
        except AttributeError as e:
            logger.error(f"선택 텍스트 서식 적용 API 호출 실패: {e}")
            return False
        except Exception as e:
            logger.error(f"선택된 텍스트에 서식 적용 중 예상치 못한 오류: {e}")
            return False
    
    def get_advanced_features(self):
        """
        고급 기능 인스턴스를 반환합니다.
        
        Returns:
            HwpAdvancedFeatures: 고급 기능 인스턴스
        """
        from .hwp_advanced_features import HwpAdvancedFeatures
        return HwpAdvancedFeatures(self)
    
    def get_document_features(self):
        """
        문서 편집 고급 기능 인스턴스를 반환합니다.
        
        Returns:
            HwpDocumentFeatures: 문서 편집 고급 기능 인스턴스
        """
        from .hwp_document_features import HwpDocumentFeatures
        return HwpDocumentFeatures(self)
    
    def is_connected(self) -> bool:
        """
        HWP와의 연결 상태를 확인합니다.
        
        Returns:
            bool: 연결 상태
        """
        try:
            if not self.is_hwp_running or not self.hwp:
                return False
            
            # 간단한 명령 실행으로 연결 상태 확인
            self.hwp.Version
            return True
        except Exception as e:
            logger.error(f"HWP 연결 상태 확인 실패: {str(e)}")
            self.is_hwp_running = False
            return False