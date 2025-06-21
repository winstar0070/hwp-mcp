"""
HwpTableTools 클래스의 고급 기능 테스트 코드
표 스타일, 정렬, 병합, 분할 등의 기능을 테스트합니다.
"""

import pytest
from unittest.mock import Mock, MagicMock, patch
import sys
import os

# 테스트를 위한 경로 설정
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tools.hwp_table_tools import HwpTableTools


class TestHwpTableAdvanced:
    """HwpTableTools 클래스의 고급 기능 테스트"""
    
    @pytest.fixture
    def mock_hwp_controller(self):
        """Mock HwpController 객체 생성"""
        controller = Mock()
        controller.hwp = Mock()
        controller.is_hwp_running = True
        return controller
    
    @pytest.fixture
    def table_tools(self, mock_hwp_controller):
        """테스트용 HwpTableTools 인스턴스 생성"""
        return HwpTableTools(mock_hwp_controller)
    
    # ============== 표 스타일 테스트 ==============
    
    def test_apply_table_style_simple(self, table_tools):
        """Simple 스타일 적용 테스트"""
        # Given
        style_name = "simple"
        
        # Mock
        table_tools.hwp_controller.hwp.Run = Mock()
        table_tools._set_table_border_style = Mock()
        table_tools._set_table_background_color = Mock()
        
        # When
        result = table_tools._apply_table_style(style_name)
        
        # Then
        assert "successfully" in result
        table_tools.hwp_controller.hwp.Run.assert_any_call("TableSelTable")
        table_tools.hwp_controller.hwp.Run.assert_any_call("Cancel")
        table_tools._set_table_border_style.assert_called_with(1, 0.5)
        table_tools._set_table_background_color.assert_called_with(header_color="#F0F0F0")
    
    def test_apply_table_style_professional(self, table_tools):
        """Professional 스타일 적용 테스트"""
        # Given
        style_name = "professional"
        
        # Mock
        table_tools.hwp_controller.hwp.Run = Mock()
        table_tools._set_table_border_style = Mock()
        table_tools._set_table_background_color = Mock()
        
        # When
        result = table_tools._apply_table_style(style_name)
        
        # Then
        assert "successfully" in result
        table_tools._set_table_border_style.assert_called_with(1, 1.0)
        table_tools._set_table_background_color.assert_called_with(
            header_color="#4472C4", 
            header_text_color="#FFFFFF"
        )
    
    def test_apply_table_style_colorful(self, table_tools):
        """Colorful 스타일 적용 테스트"""
        # Given
        style_name = "colorful"
        
        # Mock
        table_tools.hwp_controller.hwp.Run = Mock()
        table_tools._set_table_border_style = Mock()
        table_tools._set_table_alternating_rows = Mock()
        
        # When
        result = table_tools._apply_table_style(style_name)
        
        # Then
        assert "successfully" in result
        table_tools._set_table_border_style.assert_called_with(1, 0.5)
        table_tools._set_table_alternating_rows.assert_called_with("#F2F2F2", "#FFFFFF")
    
    def test_apply_table_style_dark(self, table_tools):
        """Dark 스타일 적용 테스트"""
        # Given
        style_name = "dark"
        
        # Mock
        table_tools.hwp_controller.hwp.Run = Mock()
        table_tools._set_table_border_style = Mock()
        table_tools._set_table_background_color = Mock()
        
        # When
        result = table_tools._apply_table_style(style_name)
        
        # Then
        assert "successfully" in result
        table_tools._set_table_border_style.assert_called_with(2, 1.0)
        table_tools._set_table_background_color.assert_called_with(
            header_color="#2B2B2B",
            header_text_color="#FFFFFF",
            body_color="#3A3A3A",
            body_text_color="#E0E0E0"
        )
    
    def test_apply_table_style_default(self, table_tools):
        """기본 스타일 적용 테스트"""
        # Given
        style_name = "unknown"
        
        # Mock
        table_tools.hwp_controller.hwp.Run = Mock()
        table_tools._set_table_border_style = Mock()
        
        # When
        result = table_tools._apply_table_style(style_name)
        
        # Then
        assert "successfully" in result
        table_tools._set_table_border_style.assert_called_with(1, 0.7)
    
    def test_apply_table_style_exception(self, table_tools):
        """스타일 적용 중 예외 발생 테스트"""
        # Given
        table_tools.hwp_controller.hwp.Run.side_effect = Exception("COM Error")
        
        # When
        result = table_tools._apply_table_style("simple")
        
        # Then
        assert "Error:" in result
    
    # ============== 표 정렬 테스트 ==============
    
    def test_sort_table_ascending(self, table_tools):
        """오름차순 정렬 테스트"""
        # Given
        column_index = 2
        ascending = True
        
        # Mock
        hwp = table_tools.hwp_controller.hwp
        hwp.Run = Mock()
        hwp.HAction = Mock()
        hwp.HParameterSet = Mock()
        hwp.HParameterSet.HTableSort = Mock()
        hwp.HParameterSet.HTableSort.HSet = Mock()
        
        # When
        result = table_tools._sort_table(column_index, ascending)
        
        # Then
        assert "successfully" in result
        hwp.Run.assert_any_call("TableSelTable")
        hwp.Run.assert_any_call("Cancel")
        assert hwp.HParameterSet.HTableSort.KeyColumn1 == 1  # column_index - 1
        assert hwp.HParameterSet.HTableSort.SortOrder1 == 0  # 오름차순
    
    def test_sort_table_descending(self, table_tools):
        """내림차순 정렬 테스트"""
        # Given
        column_index = 1
        ascending = False
        
        # Mock
        hwp = table_tools.hwp_controller.hwp
        hwp.Run = Mock()
        hwp.HAction = Mock()
        hwp.HParameterSet = Mock()
        hwp.HParameterSet.HTableSort = Mock()
        hwp.HParameterSet.HTableSort.HSet = Mock()
        
        # When
        result = table_tools._sort_table(column_index, ascending)
        
        # Then
        assert "successfully" in result
        assert hwp.HParameterSet.HTableSort.KeyColumn1 == 0
        assert hwp.HParameterSet.HTableSort.SortOrder1 == 1  # 내림차순
    
    def test_sort_table_exception(self, table_tools):
        """정렬 중 예외 발생 테스트"""
        # Given
        table_tools.hwp_controller.hwp.Run.side_effect = Exception("COM Error")
        
        # When
        result = table_tools._sort_table(1)
        
        # Then
        assert "Error:" in result
    
    # ============== 셀 병합 테스트 ==============
    
    def test_merge_cells_success(self, table_tools):
        """셀 병합 성공 테스트"""
        # Given
        start_row, start_col = 1, 1
        end_row, end_col = 2, 3
        
        # Mock
        hwp = table_tools.hwp_controller.hwp
        hwp.Run = Mock()
        table_tools._move_to_cell = Mock()
        
        # When
        result = table_tools._merge_cells(start_row, start_col, end_row, end_col)
        
        # Then
        assert "successfully" in result
        table_tools._move_to_cell.assert_any_call(start_row, start_col)
        table_tools._move_to_cell.assert_any_call(end_row, end_col)
        hwp.Run.assert_any_call("TableCellBlock")
        hwp.Run.assert_any_call("TableCellBlockExtend")
        hwp.Run.assert_any_call("TableMergeCell")
    
    def test_merge_cells_exception(self, table_tools):
        """셀 병합 중 예외 발생 테스트"""
        # Given
        table_tools._move_to_cell = Mock(side_effect=Exception("Move Error"))
        
        # When
        result = table_tools._merge_cells(1, 1, 2, 2)
        
        # Then
        assert "Error:" in result
    
    # ============== 셀 분할 테스트 ==============
    
    def test_split_cell_success(self, table_tools):
        """셀 분할 성공 테스트"""
        # Given
        rows = 2
        cols = 3
        
        # Mock
        hwp = table_tools.hwp_controller.hwp
        hwp.Run = Mock()
        hwp.HAction = Mock()
        hwp.HParameterSet = Mock()
        hwp.HParameterSet.HTableSplitCell = Mock()
        hwp.HParameterSet.HTableSplitCell.HSet = Mock()
        
        # When
        result = table_tools._split_cell(rows, cols)
        
        # Then
        assert "successfully" in result
        hwp.Run.assert_called_with("TableSelCell")
        assert hwp.HParameterSet.HTableSplitCell.Rows == rows
        assert hwp.HParameterSet.HTableSplitCell.Cols == cols
    
    def test_split_cell_exception(self, table_tools):
        """셀 분할 중 예외 발생 테스트"""
        # Given
        table_tools.hwp_controller.hwp.Run.side_effect = Exception("COM Error")
        
        # When
        result = table_tools._split_cell(2, 2)
        
        # Then
        assert "Error:" in result
    
    # ============== 헬퍼 메서드 테스트 ==============
    
    def test_move_to_cell(self, table_tools):
        """특정 셀로 이동 테스트"""
        # Given
        row, col = 3, 2
        hwp = table_tools.hwp_controller.hwp
        hwp.Run = Mock()
        
        # When
        table_tools._move_to_cell(row, col)
        
        # Then
        # 첫 번째 셀로 이동
        hwp.Run.assert_any_call("TableColBegin")
        hwp.Run.assert_any_call("TableRowBegin")
        
        # 목표 행으로 이동 (3행 = 2번 아래로)
        assert hwp.Run.call_args_list.count(Mock(call("TableLowerCell"))) == 2
        
        # 목표 열로 이동 (2열 = 1번 오른쪽으로)
        assert hwp.Run.call_args_list.count(Mock(call("TableRightCell"))) == 1
    
    def test_set_table_border_style(self, table_tools):
        """테이블 테두리 스타일 설정 테스트"""
        # Given
        border_type = 1
        width = 1.5
        
        # Mock
        hwp = table_tools.hwp_controller.hwp
        hwp.HAction = Mock()
        hwp.HParameterSet = Mock()
        hwp.HParameterSet.HCellBorderFill = Mock()
        hwp.HParameterSet.HCellBorderFill.HSet = Mock()
        
        # When
        table_tools._set_table_border_style(border_type, width)
        
        # Then
        assert hwp.HParameterSet.HCellBorderFill.BorderType == border_type
        assert hwp.HParameterSet.HCellBorderFill.BorderWidth == 15  # width * 10
    
    # ============== 통합 시나리오 테스트 ==============
    
    def test_create_styled_sorted_table(self, table_tools):
        """표 생성 -> 스타일 적용 -> 정렬 통합 테스트"""
        # Given
        mock_controller = table_tools.hwp_controller
        mock_controller.insert_table = Mock(return_value=True)
        mock_controller.fill_table_with_data = Mock(return_value=True)
        
        # Mock 메서드들
        table_tools._apply_table_style = Mock(return_value="Style applied")
        table_tools._sort_table = Mock(return_value="Table sorted")
        
        # When
        # 1. 표 생성
        create_result = table_tools.insert_table(3, 3)
        
        # 2. 데이터 입력
        data = [["Name", "Age", "Score"], ["Alice", "25", "90"], ["Bob", "22", "85"]]
        fill_result = table_tools.fill_table_with_data(data, has_header=True)
        
        # 3. 스타일 적용
        style_result = table_tools._apply_table_style("professional")
        
        # 4. 정렬
        sort_result = table_tools._sort_table(3, False)  # Score 열 내림차순
        
        # Then
        assert "inserted" in create_result
        assert "완료" in fill_result
        assert style_result == "Style applied"
        assert sort_result == "Table sorted"