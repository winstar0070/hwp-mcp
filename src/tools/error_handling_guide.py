"""
HWP-MCP 프로젝트 에러 처리 가이드

이 모듈은 프로젝트 전체에서 일관된 에러 처리를 위한 가이드와 헬퍼 함수를 제공합니다.
"""

from typing import Type, Callable, Any, Dict, Optional
from functools import wraps
import logging
from .hwp_exceptions import *

logger = logging.getLogger(__name__)

# 예외 타입 매핑
EXCEPTION_MAPPING: Dict[Type[Exception], Type[HwpError]] = {
    # 파일 시스템 관련
    FileNotFoundError: HwpDocumentNotFoundError,
    PermissionError: HwpDocumentAccessError,
    OSError: HwpDocumentError,
    
    # 타입 관련
    TypeError: HwpInvalidParameterError,
    ValueError: HwpInvalidParameterError,
    
    # 인덱스 관련
    IndexError: HwpTableRangeError,
    
    # 속성 관련
    AttributeError: HwpOperationError,
    
    # 키 관련
    KeyError: HwpParameterError,
}


def enhanced_error_handler(operation_name: str, 
                         default_return: Any = None,
                         raise_on_error: bool = True):
    """
    향상된 에러 처리 데코레이터
    
    Args:
        operation_name: 작업 이름
        default_return: 에러 시 반환할 기본값
        raise_on_error: True면 예외 발생, False면 기본값 반환
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            
            # HWP 관련 예외는 그대로 전파
            except HwpError:
                raise
            
            # 파일 관련 예외
            except FileNotFoundError as e:
                logger.error(f"{operation_name} - 파일을 찾을 수 없음: {e}")
                if raise_on_error:
                    if hasattr(e, 'filename'):
                        raise HwpDocumentNotFoundError(e.filename)
                    raise HwpDocumentNotFoundError(str(e))
                return default_return
            
            except PermissionError as e:
                logger.error(f"{operation_name} - 파일 접근 권한 없음: {e}")
                if raise_on_error:
                    if hasattr(e, 'filename'):
                        raise HwpDocumentAccessError(e.filename)
                    raise HwpDocumentAccessError(str(e))
                return default_return
            
            except OSError as e:
                logger.error(f"{operation_name} - 파일 시스템 오류: {e}")
                if raise_on_error:
                    raise HwpDocumentError(f"파일 시스템 오류: {e}")
                return default_return
            
            # 타입 관련 예외
            except TypeError as e:
                logger.error(f"{operation_name} - 잘못된 타입: {e}")
                if raise_on_error:
                    # 함수 시그니처에서 파라미터 이름 추출 시도
                    param_name = extract_param_from_error(str(e))
                    raise HwpInvalidParameterError(param_name or "unknown", "잘못된 타입", "올바른 타입")
                return default_return
            
            except ValueError as e:
                logger.error(f"{operation_name} - 잘못된 값: {e}")
                if raise_on_error:
                    param_name = extract_param_from_error(str(e))
                    raise HwpInvalidParameterError(param_name or "unknown", str(e), "유효한 값")
                return default_return
            
            # 인덱스 관련 예외
            except IndexError as e:
                logger.error(f"{operation_name} - 인덱스 범위 초과: {e}")
                if raise_on_error:
                    raise HwpTableRangeError(0, 0, 0, 0)  # 실제 값은 컨텍스트에서 설정
                return default_return
            
            # API 호출 관련 예외
            except AttributeError as e:
                logger.error(f"{operation_name} - HWP API 호출 실패: {e}")
                if raise_on_error:
                    raise HwpOperationError(f"HWP API 호출 실패: {e} (HWP가 제대로 연결되지 않았거나 지원하지 않는 기능입니다)")
                return default_return
            
            # 기타 예외
            except Exception as e:
                logger.error(f"{operation_name} - 예상치 못한 오류: {e}", exc_info=True)
                if raise_on_error:
                    raise HwpOperationError(f"{operation_name} 중 예상치 못한 오류: {e}")
                return default_return
        
        return wrapper
    return decorator


def extract_param_from_error(error_msg: str) -> Optional[str]:
    """에러 메시지에서 파라미터 이름 추출 시도"""
    # 일반적인 패턴들
    patterns = [
        "argument '(.+?)'",
        "parameter '(.+?)'",
        "got an unexpected keyword argument '(.+?)'",
    ]
    
    import re
    for pattern in patterns:
        match = re.search(pattern, error_msg)
        if match:
            return match.group(1)
    return None


def validate_file_path(path: str, must_exist: bool = False) -> str:
    """
    파일 경로 검증 및 정규화
    
    Args:
        path: 검증할 경로
        must_exist: True면 파일이 존재해야 함
        
    Returns:
        정규화된 절대 경로
        
    Raises:
        HwpDocumentNotFoundError: 파일이 없을 때
        HwpInvalidParameterError: 경로가 잘못되었을 때
    """
    import os
    
    if not path:
        raise HwpInvalidParameterError("path", path, "비어있지 않은 경로")
    
    try:
        abs_path = os.path.abspath(path)
    except Exception as e:
        raise HwpInvalidParameterError("path", path, f"유효한 경로 (오류: {e})")
    
    if must_exist and not os.path.exists(abs_path):
        raise HwpDocumentNotFoundError(abs_path)
    
    return abs_path


def validate_table_coordinates(row: int, col: int, 
                              max_rows: int = TABLE_MAX_ROWS, 
                              max_cols: int = TABLE_MAX_COLS) -> None:
    """
    표 좌표 검증
    
    Args:
        row: 행 번호 (1부터 시작)
        col: 열 번호 (1부터 시작)
        max_rows: 최대 행 수
        max_cols: 최대 열 수
        
    Raises:
        HwpTableRangeError: 범위 초과
        HwpInvalidParameterError: 잘못된 값
    """
    if not isinstance(row, int) or not isinstance(col, int):
        raise HwpInvalidParameterError(
            "row/col", 
            f"{row}/{col}", 
            "정수 타입"
        )
    
    if row < 1:
        raise HwpInvalidParameterError("row", row, "1 이상의 값")
    if col < 1:
        raise HwpInvalidParameterError("col", col, "1 이상의 값")
    
    if row > max_rows or col > max_cols:
        raise HwpTableRangeError(row, col, max_rows, max_cols)


def format_error_message(operation: str, error: Exception, 
                        context: Optional[Dict[str, Any]] = None) -> str:
    """
    사용자 친화적인 에러 메시지 생성
    
    Args:
        operation: 작업 이름
        error: 발생한 예외
        context: 추가 컨텍스트 정보
        
    Returns:
        포맷된 에러 메시지
    """
    message_parts = [f"{operation} 실패"]
    
    # 에러 타입별 상세 메시지
    if isinstance(error, FileNotFoundError):
        message_parts.append("파일을 찾을 수 없습니다")
    elif isinstance(error, PermissionError):
        message_parts.append("파일 접근 권한이 없습니다")
    elif isinstance(error, AttributeError):
        message_parts.append("HWP API 호출에 실패했습니다. HWP가 제대로 연결되었는지 확인하세요")
    elif isinstance(error, ValueError):
        message_parts.append("잘못된 값이 입력되었습니다")
    elif isinstance(error, TypeError):
        message_parts.append("잘못된 타입의 데이터가 입력되었습니다")
    
    # 컨텍스트 정보 추가
    if context:
        context_str = ", ".join([f"{k}={v}" for k, v in context.items()])
        message_parts.append(f"(컨텍스트: {context_str})")
    
    # 원본 에러 메시지
    if str(error):
        message_parts.append(f"- {str(error)}")
    
    return " ".join(message_parts)