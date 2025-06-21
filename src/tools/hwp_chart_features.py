"""
HWP 차트 및 그래프 기능 모듈
차트, 그래프, 수식 등 시각화 관련 기능 제공
"""
import os
import logging
from typing import List, Dict, Any, Optional, Tuple
import json

try:
    from .constants import (
        HWPUNIT_PER_MM, HWPUNIT_PER_CM,
        CHART_TYPES, CHART_DEFAULT_SIZE
    )
    from .hwp_utils import (
        safe_hwp_operation, log_operation_result,
        require_hwp_connection
    )
except ImportError:
    # 상수 기본값
    HWPUNIT_PER_MM = 2834.64
    HWPUNIT_PER_CM = 28346.4
    CHART_TYPES = {
        "column": 0, "bar": 1, "line": 2, "pie": 3,
        "area": 4, "scatter": 5, "doughnut": 6
    }
    CHART_DEFAULT_SIZE = {"width": 120, "height": 80}  # mm 단위
    
    # 더미 함수
    def safe_hwp_operation(name, default=None):
        def decorator(func):
            return func
        return decorator
    def log_operation_result(*args, **kwargs):
        pass
    def require_hwp_connection(func):
        return func

logger = logging.getLogger(__name__)

class HwpChartFeatures:
    """HWP 차트 및 그래프 기능을 제공하는 클래스"""
    
    def __init__(self, hwp_controller):
        """
        HwpChartFeatures 초기화
        
        Args:
            hwp_controller: HwpController 인스턴스
        """
        self.hwp = hwp_controller.hwp
        self.hwp_controller = hwp_controller
    
    # ============== 차트/그래프 기능 ==============
    
    @require_hwp_connection
    @safe_hwp_operation("차트 삽입")
    def insert_chart(self, chart_type: str = "column", 
                    data: List[List[Any]] = None,
                    title: str = "",
                    width: int = None,
                    height: int = None,
                    position: Optional[Dict[str, int]] = None) -> bool:
        """
        차트를 삽입합니다.
        
        Args:
            chart_type (str): 차트 유형 ("column", "bar", "line", "pie", "area", "scatter", "doughnut")
            data (List[List[Any]]): 차트 데이터 (2차원 배열)
                                   예: [["카테고리", "값"], ["A", 10], ["B", 20], ["C", 15]]
            title (str): 차트 제목
            width (int): 차트 너비 (mm 단위)
            height (int): 차트 높이 (mm 단위)
            position (dict): 차트 위치 {"x": 50, "y": 100} (mm 단위)
            
        Returns:
            bool: 성공 여부
        """
        # 기본 크기 설정
        if width is None:
            width = CHART_DEFAULT_SIZE["width"]
        if height is None:
            height = CHART_DEFAULT_SIZE["height"]
        
        # 차트 유형 확인
        if chart_type not in CHART_TYPES:
            logger.error(f"지원하지 않는 차트 유형: {chart_type}")
            return False
        
        # 데이터가 없으면 샘플 데이터 사용
        if not data:
            data = [
                ["카테고리", "값"],
                ["항목1", 25],
                ["항목2", 35],
                ["항목3", 20],
                ["항목4", 20]
            ]
        
        # 차트 생성을 위한 표 삽입
        rows = len(data)
        cols = len(data[0]) if data else 2
        
        if not self.hwp_controller.insert_table(rows, cols):
            logger.error("차트 데이터용 표 생성 실패")
            return False
        
        # 표에 데이터 입력
        for row_idx, row_data in enumerate(data):
            for col_idx, cell_value in enumerate(row_data):
                self.hwp_controller.fill_table_cell(
                    row_idx + 1, 
                    col_idx + 1, 
                    str(cell_value)
                )
        
        # 표 전체 선택
        self.hwp.Run("TableSelTable")
        
        # 차트 삽입
        try:
            # 차트 생성 명령 실행
            self.hwp.Run("InsertChart")
            
            # 차트 속성 설정 (HWP API 제한으로 기본 설정만 가능)
            # 실제 차트 유형, 크기 등은 사용자가 직접 편집해야 함
            
            # 표 선택 해제
            self.hwp.Run("Cancel")
            
            log_operation_result("차트 삽입", True, f"{chart_type} 차트")
            return True
            
        except Exception as e:
            logger.error(f"차트 삽입 중 오류: {e}")
            return False
    
    @require_hwp_connection
    @safe_hwp_operation("간단한 차트 삽입")
    def insert_simple_chart(self, values: List[float], 
                          labels: Optional[List[str]] = None,
                          chart_type: str = "column",
                          title: str = "") -> bool:
        """
        간단한 데이터로 차트를 삽입합니다.
        
        Args:
            values (List[float]): 차트 값 리스트
            labels (List[str], optional): 각 값의 레이블
            chart_type (str): 차트 유형
            title (str): 차트 제목
            
        Returns:
            bool: 성공 여부
        """
        # 레이블이 없으면 자동 생성
        if not labels:
            labels = [f"항목{i+1}" for i in range(len(values))]
        
        # 데이터 구조 변환
        data = [["항목", "값"]]
        for label, value in zip(labels, values):
            data.append([label, value])
        
        return self.insert_chart(
            chart_type=chart_type,
            data=data,
            title=title
        )
    
    # ============== 수식 편집기 기능 ==============
    
    @require_hwp_connection
    @safe_hwp_operation("수식 삽입")
    def insert_equation(self, equation_text: str = "", 
                       inline: bool = True) -> bool:
        """
        수식을 삽입합니다.
        
        Args:
            equation_text (str): 수식 텍스트 (LaTeX 형식 일부 지원)
            inline (bool): 인라인 수식 여부
            
        Returns:
            bool: 성공 여부
        """
        # 수식 편집기 열기
        self.hwp.Run("InsertEquation")
        
        # 수식 입력 (HWP 수식 편집기는 자체 문법 사용)
        # LaTeX와 완전히 호환되지는 않음
        if equation_text:
            # 기본적인 LaTeX 변환
            hwp_equation = self._convert_latex_to_hwp(equation_text)
            self.hwp_controller.insert_text(hwp_equation)
        
        # 수식 편집기 닫기
        self.hwp.Run("CloseEx")
        
        log_operation_result("수식 삽입", True, equation_text[:30] if equation_text else "빈 수식")
        return True
    
    def _convert_latex_to_hwp(self, latex_text: str) -> str:
        """
        기본적인 LaTeX 수식을 HWP 수식으로 변환합니다.
        
        Args:
            latex_text (str): LaTeX 수식 텍스트
            
        Returns:
            str: HWP 수식 텍스트
        """
        # 기본적인 변환 규칙
        conversions = {
            r"\frac": "over",
            r"\sqrt": "sqrt",
            r"\sum": "sum",
            r"\int": "int",
            r"\alpha": "alpha",
            r"\beta": "beta",
            r"\gamma": "gamma",
            r"\pi": "pi",
            r"\theta": "theta",
            r"\times": "times",
            r"\div": "div",
            r"\pm": "+-",
            r"\leq": "<=",
            r"\geq": ">=",
            r"\neq": "!=",
            r"^": "^",
            r"_": "_",
            r"{": "{",
            r"}": "}"
        }
        
        hwp_text = latex_text
        for latex, hwp in conversions.items():
            hwp_text = hwp_text.replace(latex, hwp)
        
        return hwp_text
    
    @require_hwp_connection
    @safe_hwp_operation("수식 템플릿 삽입")
    def insert_equation_template(self, template_type: str = "fraction") -> bool:
        """
        미리 정의된 수식 템플릿을 삽입합니다.
        
        Args:
            template_type (str): 템플릿 유형
                - "fraction": 분수
                - "sqrt": 제곱근
                - "sum": 합계
                - "integral": 적분
                - "matrix": 행렬
                - "quadratic": 이차방정식
                
        Returns:
            bool: 성공 여부
        """
        templates = {
            "fraction": "a over b",
            "sqrt": "sqrt {x}",
            "sum": "sum from {i=1} to {n} x_i",
            "integral": "int from {a} to {b} f(x) dx",
            "matrix": "left [ matrix { a # b ## c # d } right ]",
            "quadratic": "x = {-b +- sqrt {b^2 - 4ac}} over {2a}"
        }
        
        if template_type not in templates:
            logger.error(f"지원하지 않는 템플릿: {template_type}")
            return False
        
        return self.insert_equation(templates[template_type])
    
    # ============== 다이어그램 기능 ==============
    
    @require_hwp_connection
    @safe_hwp_operation("다이어그램 삽입")
    def insert_diagram(self, diagram_type: str = "flowchart",
                      elements: Optional[List[Dict[str, Any]]] = None) -> bool:
        """
        간단한 다이어그램을 삽입합니다.
        
        Args:
            diagram_type (str): 다이어그램 유형 ("flowchart", "org_chart", "process")
            elements (List[Dict]): 다이어그램 요소들
            
        Returns:
            bool: 성공 여부
        """
        # HWP는 복잡한 다이어그램을 직접 지원하지 않으므로
        # 도형과 연결선을 조합하여 생성
        
        if not elements:
            # 샘플 플로우차트
            elements = [
                {"type": "start", "text": "시작", "x": 100, "y": 50},
                {"type": "process", "text": "처리", "x": 100, "y": 150},
                {"type": "decision", "text": "판단", "x": 100, "y": 250},
                {"type": "end", "text": "종료", "x": 100, "y": 350}
            ]
        
        # 각 요소를 도형으로 생성
        for element in elements:
            shape_type = self._get_shape_for_element(element["type"])
            if shape_type:
                # 도형 삽입 (HWP API 제한으로 위치 지정은 제한적)
                self.hwp.Run(shape_type)
                
                # 텍스트 입력
                if "text" in element:
                    self.hwp_controller.insert_text(element["text"])
                    self.hwp.Run("Cancel")
        
        log_operation_result("다이어그램 삽입", True, diagram_type)
        return True
    
    def _get_shape_for_element(self, element_type: str) -> str:
        """다이어그램 요소 유형에 맞는 도형 명령을 반환합니다."""
        shape_map = {
            "start": "DrawObjCreEllipse",      # 타원
            "process": "DrawObjCreRectangle",   # 사각형
            "decision": "DrawObjCrePolygon",    # 다각형 (마름모)
            "end": "DrawObjCreEllipse",         # 타원
            "connector": "DrawObjCreLine"       # 연결선
        }
        return shape_map.get(element_type, "DrawObjCreRectangle")