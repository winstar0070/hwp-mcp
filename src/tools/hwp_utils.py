"""
HWP MCP 프로젝트 공통 유틸리티 함수
중복된 코드를 줄이기 위한 공통 함수들을 정의합니다.
"""

import logging
from typing import Any, Dict, Optional, Callable, TypeVar
from functools import wraps

from .hwp_exceptions import HwpNotRunningError, HwpOperationError

logger = logging.getLogger(__name__)

T = TypeVar('T')


def require_hwp_connection(func: Callable[..., T]) -> Callable[..., T]:
    """
    HWP 연결이 필요한 메서드에 대한 데코레이터
    연결 상태를 확인하고 연결되지 않은 경우 예외를 발생시킵니다.
    """
    @wraps(func)
    def wrapper(self, *args, **kwargs) -> T:
        if not hasattr(self, 'is_hwp_running') or not self.is_hwp_running:
            raise HwpNotRunningError()
        if not hasattr(self, 'hwp') or not self.hwp:
            raise HwpNotRunningError()
        return func(self, *args, **kwargs)
    return wrapper


def safe_hwp_operation(operation_name: str, default_return=None):
    """
    HWP 작업을 안전하게 실행하는 데코레이터
    예외 발생 시 로깅하고 기본값을 반환합니다.
    """
    def decorator(func: Callable[..., T]) -> Callable[..., T]:
        @wraps(func)
        def wrapper(self, *args, **kwargs) -> T:
            try:
                return func(self, *args, **kwargs)
            except Exception as e:
                logger.error(f"{operation_name} 중 오류 발생: {str(e)}", exc_info=True)
                if default_return is not None:
                    return default_return
                raise HwpOperationError(f"{operation_name} 실패: {str(e)}")
        return wrapper
    return decorator


def set_font_properties(hwp, font_name: Optional[str] = None, 
                       font_size: Optional[int] = None,
                       bold: bool = False, italic: bool = False, 
                       underline: bool = False) -> bool:
    """
    텍스트 글꼴 속성을 설정하는 공통 함수
    
    Args:
        hwp: HWP COM 객체
        font_name: 글꼴 이름
        font_size: 글꼴 크기
        bold: 굵게
        italic: 기울임
        underline: 밑줄
    
    Returns:
        bool: 성공 여부
    """
    try:
        # CharShape 액션 초기화
        hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
        
        # 글꼴 설정
        if font_name:
            hwp.HParameterSet.HCharShape.FaceNameUser = font_name
            hwp.HParameterSet.HCharShape.FaceNameSymbol = font_name
            hwp.HParameterSet.HCharShape.FaceNameOther = font_name
            hwp.HParameterSet.HCharShape.FaceNameJapanese = font_name
            hwp.HParameterSet.HCharShape.FaceNameHanja = font_name
            hwp.HParameterSet.HCharShape.FaceNameLatin = font_name
            hwp.HParameterSet.HCharShape.FaceNameHangul = font_name
        
        # 글꼴 크기 설정 (포인트를 HWPUNIT으로 변환)
        if font_size:
            hwp.HParameterSet.HCharShape.Height = font_size * 100
        
        # 글꼴 스타일 설정
        if bold:
            hwp.HParameterSet.HCharShape.Bold = 1
        if italic:
            hwp.HParameterSet.HCharShape.Italic = 1
        if underline:
            hwp.HParameterSet.HCharShape.UnderlineType = 1
        
        # 설정 적용
        hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
        
        logger.debug(f"글꼴 속성 설정 완료: {font_name}, {font_size}pt")
        return True
        
    except Exception as e:
        logger.error(f"글꼴 속성 설정 실패: {str(e)}")
        return False


def move_to_table_cell(hwp, row: int, col: int) -> bool:
    """
    표의 특정 셀로 이동하는 공통 함수
    
    Args:
        hwp: HWP COM 객체
        row: 행 번호 (1부터 시작)
        col: 열 번호 (1부터 시작)
    
    Returns:
        bool: 성공 여부
    """
    try:
        # 표의 첫 번째 셀로 이동
        hwp.Run("TableColBegin")
        hwp.Run("TableRowBegin")
        
        # 목표 행으로 이동
        for _ in range(row - 1):
            hwp.Run("TableLowerCell")
        
        # 목표 열로 이동
        for _ in range(col - 1):
            hwp.Run("TableRightCell")
        
        return True
        
    except Exception as e:
        logger.error(f"셀 이동 실패: ({row}, {col}) - {str(e)}")
        return False


def parse_table_data(data: Any) -> list:
    """
    다양한 형식의 표 데이터를 2차원 리스트로 변환하는 공통 함수
    
    Args:
        data: 표 데이터 (리스트, 문자열 등)
    
    Returns:
        list: 2차원 문자열 리스트
    """
    import json
    
    # 이미 2차원 리스트인 경우
    if isinstance(data, list) and all(isinstance(row, list) for row in data):
        return [[str(cell) if cell is not None else "" for cell in row] for row in data]
    
    # JSON 문자열인 경우
    if isinstance(data, str):
        try:
            parsed = json.loads(data)
            if isinstance(parsed, list):
                return parse_table_data(parsed)  # 재귀 호출
        except json.JSONDecodeError:
            logger.warning(f"JSON 파싱 실패, 단일 셀로 처리: {data[:50]}...")
            return [[str(data)]]
    
    # 기타 경우
    return [[str(data)]]


def execute_with_retry(func: Callable[..., T], max_retries: int = 3, 
                      delay: float = 1.0, *args, **kwargs) -> T:
    """
    재시도 로직을 포함한 함수 실행
    
    Args:
        func: 실행할 함수
        max_retries: 최대 재시도 횟수
        delay: 재시도 간 대기 시간 (초)
        *args, **kwargs: 함수에 전달할 인자
    
    Returns:
        함수 실행 결과
    """
    import time
    
    last_error = None
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            last_error = e
            if attempt < max_retries - 1:
                logger.warning(f"작업 실패 (시도 {attempt + 1}/{max_retries}): {str(e)}")
                time.sleep(delay)
            else:
                logger.error(f"작업 최종 실패: {str(e)}")
    
    raise last_error


def validate_table_coordinates(row: int, col: int, max_rows: int = 100, 
                              max_cols: int = 100) -> None:
    """
    표 좌표의 유효성을 검증하는 공통 함수
    
    Args:
        row: 행 번호
        col: 열 번호
        max_rows: 최대 행 수
        max_cols: 최대 열 수
    
    Raises:
        ValueError: 좌표가 유효하지 않은 경우
    """
    if row < 1 or row > max_rows:
        raise ValueError(f"행 번호는 1과 {max_rows} 사이여야 합니다. 입력값: {row}")
    if col < 1 or col > max_cols:
        raise ValueError(f"열 번호는 1과 {max_cols} 사이여야 합니다. 입력값: {col}")


def get_hwp_action_parameter(hwp, action_name: str, parameter_set_name: str) -> Any:
    """
    HWP 액션 파라미터를 가져오는 공통 함수
    
    Args:
        hwp: HWP COM 객체
        action_name: 액션 이름
        parameter_set_name: 파라미터 세트 이름
    
    Returns:
        파라미터 세트 객체
    """
    try:
        # 파라미터 세트 가져오기
        param_set = getattr(hwp.HParameterSet, parameter_set_name)
        # 기본값 설정
        hwp.HAction.GetDefault(action_name, param_set.HSet)
        return param_set
    except Exception as e:
        logger.error(f"파라미터 세트 가져오기 실패: {action_name} - {str(e)}")
        raise


def log_operation_result(operation: str, success: bool, details: str = "") -> None:
    """
    작업 결과를 로깅하는 공통 함수
    
    Args:
        operation: 작업 이름
        success: 성공 여부
        details: 추가 세부 정보
    """
    if success:
        logger.info(f"{operation} 성공{': ' + details if details else ''}")
    else:
        logger.error(f"{operation} 실패{': ' + details if details else ''}")