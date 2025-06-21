"""
HWP 유틸리티 함수 테스트
hwp_utils.py의 모든 함수에 대한 단위 테스트를 포함합니다.
"""

import pytest
from unittest.mock import Mock, MagicMock, patch, call
import sys
import os
import json
import time

# 테스트를 위한 경로 설정
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tools.hwp_utils import (
    require_hwp_connection,
    safe_hwp_operation,
    set_font_properties,
    move_to_table_cell,
    move_to_table_cell_optimized,
    TablePosition,
    parse_table_data,
    execute_with_retry,
    validate_table_coordinates,
    get_hwp_action_parameter,
    log_operation_result
)
from tools.hwp_exceptions import HwpNotRunningError, HwpOperationError


class TestRequireHwpConnection:
    """require_hwp_connection 데코레이터 테스트"""
    
    def test_successful_connection(self):
        """정상적인 HWP 연결 상태에서의 테스트"""
        # Given
        class TestClass:
            is_hwp_running = True
            hwp = Mock()
            
            @require_hwp_connection
            def test_method(self):
                return "success"
        
        # When
        obj = TestClass()
        result = obj.test_method()
        
        # Then
        assert result == "success"
    
    def test_no_hwp_running_attribute(self):
        """is_hwp_running 속성이 없을 때"""
        # Given
        class TestClass:
            hwp = Mock()
            
            @require_hwp_connection
            def test_method(self):
                return "success"
        
        # When/Then
        obj = TestClass()
        with pytest.raises(HwpNotRunningError):
            obj.test_method()
    
    def test_hwp_not_running(self):
        """HWP가 실행되지 않은 상태"""
        # Given
        class TestClass:
            is_hwp_running = False
            hwp = Mock()
            
            @require_hwp_connection
            def test_method(self):
                return "success"
        
        # When/Then
        obj = TestClass()
        with pytest.raises(HwpNotRunningError):
            obj.test_method()
    
    def test_no_hwp_object(self):
        """HWP 객체가 없을 때"""
        # Given
        class TestClass:
            is_hwp_running = True
            hwp = None
            
            @require_hwp_connection
            def test_method(self):
                return "success"
        
        # When/Then
        obj = TestClass()
        with pytest.raises(HwpNotRunningError):
            obj.test_method()


class TestSafeHwpOperation:
    """safe_hwp_operation 데코레이터 테스트"""
    
    def test_successful_operation(self):
        """정상적인 작업 실행"""
        # Given
        class TestClass:
            @safe_hwp_operation("테스트 작업")
            def test_method(self):
                return "success"
        
        # When
        obj = TestClass()
        result = obj.test_method()
        
        # Then
        assert result == "success"
    
    def test_operation_with_exception_and_default(self):
        """예외 발생 시 기본값 반환"""
        # Given
        class TestClass:
            @safe_hwp_operation("테스트 작업", default_return="default")
            def test_method(self):
                raise Exception("Test error")
        
        # When
        obj = TestClass()
        result = obj.test_method()
        
        # Then
        assert result == "default"
    
    def test_operation_with_exception_no_default(self):
        """예외 발생 시 기본값이 없으면 HwpOperationError 발생"""
        # Given
        class TestClass:
            @safe_hwp_operation("테스트 작업")
            def test_method(self):
                raise Exception("Test error")
        
        # When/Then
        obj = TestClass()
        with pytest.raises(HwpOperationError) as exc_info:
            obj.test_method()
        assert "테스트 작업 실패" in str(exc_info.value)


class TestSetFontProperties:
    """set_font_properties 함수 테스트"""
    
    @pytest.fixture
    def mock_hwp(self):
        """Mock HWP 객체"""
        hwp = Mock()
        hwp.HAction = Mock()
        hwp.HParameterSet = Mock()
        hwp.HParameterSet.HCharShape = Mock()
        hwp.HParameterSet.HCharShape.HSet = Mock()
        return hwp
    
    def test_set_all_properties(self, mock_hwp):
        """모든 글꼴 속성 설정"""
        # When
        result = set_font_properties(
            mock_hwp, 
            font_name="맑은 고딕",
            font_size=12,
            bold=True,
            italic=True,
            underline=True
        )
        
        # Then
        assert result is True
        assert mock_hwp.HParameterSet.HCharShape.FaceNameHangul == "맑은 고딕"
        assert mock_hwp.HParameterSet.HCharShape.Height == 1200  # 12 * 100
        assert mock_hwp.HParameterSet.HCharShape.Bold == 1
        assert mock_hwp.HParameterSet.HCharShape.Italic == 1
        assert mock_hwp.HParameterSet.HCharShape.UnderlineType == 1
        mock_hwp.HAction.Execute.assert_called_once()
    
    def test_set_partial_properties(self, mock_hwp):
        """일부 속성만 설정"""
        # When
        result = set_font_properties(mock_hwp, font_size=10)
        
        # Then
        assert result is True
        assert mock_hwp.HParameterSet.HCharShape.Height == 1000
    
    def test_set_font_properties_with_exception(self, mock_hwp):
        """예외 발생 시 처리"""
        # Given
        mock_hwp.HAction.GetDefault.side_effect = Exception("Test error")
        
        # When
        result = set_font_properties(mock_hwp, font_name="Arial")
        
        # Then
        assert result is False


class TestMoveToTableCell:
    """move_to_table_cell 함수 테스트"""
    
    @pytest.fixture
    def mock_hwp(self):
        """Mock HWP 객체"""
        return Mock()
    
    def test_move_to_first_cell(self, mock_hwp):
        """첫 번째 셀로 이동"""
        # When
        result = move_to_table_cell(mock_hwp, 1, 1)
        
        # Then
        assert result is True
        mock_hwp.Run.assert_any_call("TableColBegin")
        mock_hwp.Run.assert_any_call("TableRowBegin")
        # 추가 이동이 없어야 함
        assert mock_hwp.Run.call_count == 2
    
    def test_move_to_specific_cell(self, mock_hwp):
        """특정 셀로 이동"""
        # When
        result = move_to_table_cell(mock_hwp, 3, 4)
        
        # Then
        assert result is True
        # 기본 이동 + 행 이동 2번 + 열 이동 3번
        calls = mock_hwp.Run.call_args_list
        assert call("TableColBegin") in calls
        assert call("TableRowBegin") in calls
        assert calls.count(call("TableLowerCell")) == 2
        assert calls.count(call("TableRightCell")) == 3
    
    def test_move_to_cell_with_exception(self, mock_hwp):
        """예외 발생 시 처리"""
        # Given
        mock_hwp.Run.side_effect = Exception("Test error")
        
        # When
        result = move_to_table_cell(mock_hwp, 2, 2)
        
        # Then
        assert result is False


class TestMoveToTableCellOptimized:
    """move_to_table_cell_optimized 함수 테스트"""
    
    @pytest.fixture
    def mock_hwp(self):
        """Mock HWP 객체"""
        return Mock()
    
    def test_move_same_position(self, mock_hwp):
        """같은 위치로 이동 (이동하지 않음)"""
        # When
        result = move_to_table_cell_optimized(mock_hwp, 2, 3, 2, 3)
        
        # Then
        assert result is True
        mock_hwp.Run.assert_not_called()
    
    def test_move_from_first_cell(self, mock_hwp):
        """첫 번째 셀에서 다른 셀로 이동"""
        # When
        result = move_to_table_cell_optimized(mock_hwp, 3, 4, 1, 1)
        
        # Then
        assert result is True
        calls = mock_hwp.Run.call_args_list
        assert call("TableColBegin") in calls
        assert call("TableRowBegin") in calls
        assert calls.count(call("TableLowerCell")) == 2
        assert calls.count(call("TableRightCell")) == 3
    
    def test_move_relative_down_right(self, mock_hwp):
        """상대적 이동: 아래쪽, 오른쪽"""
        # When
        result = move_to_table_cell_optimized(mock_hwp, 4, 5, 2, 3)
        
        # Then
        assert result is True
        calls = mock_hwp.Run.call_args_list
        assert calls.count(call("TableLowerCell")) == 2  # 4-2 = 2
        assert calls.count(call("TableRightCell")) == 2  # 5-3 = 2
    
    def test_move_relative_up_left(self, mock_hwp):
        """상대적 이동: 위쪽, 왼쪽"""
        # When
        result = move_to_table_cell_optimized(mock_hwp, 2, 3, 4, 5)
        
        # Then
        assert result is True
        calls = mock_hwp.Run.call_args_list
        assert calls.count(call("TableUpperCell")) == 2  # 2-4 = -2
        assert calls.count(call("TableLeftCell")) == 2   # 3-5 = -2
    
    def test_move_optimized_with_exception(self, mock_hwp):
        """예외 발생 시 처리"""
        # Given
        mock_hwp.Run.side_effect = Exception("Test error")
        
        # When
        result = move_to_table_cell_optimized(mock_hwp, 2, 2, 1, 1)
        
        # Then
        assert result is False


class TestTablePosition:
    """TablePosition 클래스 테스트"""
    
    def test_initial_position(self):
        """초기 위치 설정"""
        # When
        pos = TablePosition()
        
        # Then
        assert pos.get_position() == (1, 1)
    
    def test_custom_initial_position(self):
        """사용자 정의 초기 위치"""
        # When
        pos = TablePosition(3, 4)
        
        # Then
        assert pos.get_position() == (3, 4)
    
    def test_move_to_position(self):
        """위치 이동"""
        # Given
        pos = TablePosition(1, 1)
        
        # When
        pos.move_to(5, 6)
        
        # Then
        assert pos.get_position() == (5, 6)


class TestParseTableData:
    """parse_table_data 함수 테스트"""
    
    def test_parse_2d_list(self):
        """이미 2차원 리스트인 경우"""
        # Given
        data = [["A", "B"], ["C", "D"]]
        
        # When
        result = parse_table_data(data)
        
        # Then
        assert result == [["A", "B"], ["C", "D"]]
    
    def test_parse_2d_list_with_none(self):
        """None 값을 포함한 2차원 리스트"""
        # Given
        data = [["A", None], [None, "D"]]
        
        # When
        result = parse_table_data(data)
        
        # Then
        assert result == [["A", ""], ["", "D"]]
    
    def test_parse_json_string(self):
        """JSON 문자열 파싱"""
        # Given
        data = '[["A", "B"], ["C", "D"]]'
        
        # When
        result = parse_table_data(data)
        
        # Then
        assert result == [["A", "B"], ["C", "D"]]
    
    def test_parse_invalid_json_string(self):
        """잘못된 JSON 문자열"""
        # Given
        data = "This is not JSON"
        
        # When
        result = parse_table_data(data)
        
        # Then
        assert result == [["This is not JSON"]]
    
    def test_parse_single_value(self):
        """단일 값"""
        # Given
        data = 123
        
        # When
        result = parse_table_data(data)
        
        # Then
        assert result == [["123"]]


class TestExecuteWithRetry:
    """execute_with_retry 함수 테스트"""
    
    def test_successful_execution(self):
        """첫 번째 시도에서 성공"""
        # Given
        mock_func = Mock(return_value="success")
        
        # When
        result = execute_with_retry(mock_func, max_retries=3)
        
        # Then
        assert result == "success"
        assert mock_func.call_count == 1
    
    def test_retry_and_success(self):
        """재시도 후 성공"""
        # Given
        mock_func = Mock(side_effect=[Exception("Error 1"), Exception("Error 2"), "success"])
        
        # When
        with patch('time.sleep'):  # 대기 시간 건너뛰기
            result = execute_with_retry(mock_func, max_retries=3, delay=0.1)
        
        # Then
        assert result == "success"
        assert mock_func.call_count == 3
    
    def test_all_retries_fail(self):
        """모든 재시도 실패"""
        # Given
        mock_func = Mock(side_effect=Exception("Persistent error"))
        
        # When/Then
        with patch('time.sleep'):
            with pytest.raises(Exception) as exc_info:
                execute_with_retry(mock_func, max_retries=3, delay=0.1)
        
        assert str(exc_info.value) == "Persistent error"
        assert mock_func.call_count == 3
    
    def test_with_args_and_kwargs(self):
        """인자와 키워드 인자 전달"""
        # Given
        mock_func = Mock(return_value="result")
        
        # When
        result = execute_with_retry(mock_func, 3, 0.1, "arg1", "arg2", key1="value1")
        
        # Then
        assert result == "result"
        mock_func.assert_called_with("arg1", "arg2", key1="value1")


class TestValidateTableCoordinates:
    """validate_table_coordinates 함수 테스트"""
    
    def test_valid_coordinates(self):
        """유효한 좌표"""
        # When/Then - 예외가 발생하지 않아야 함
        validate_table_coordinates(1, 1)
        validate_table_coordinates(50, 50)
        validate_table_coordinates(100, 100)
    
    def test_invalid_row_too_small(self):
        """행 번호가 너무 작음"""
        # When/Then
        with pytest.raises(ValueError) as exc_info:
            validate_table_coordinates(0, 1)
        assert "행 번호는 1과 100 사이여야 합니다" in str(exc_info.value)
    
    def test_invalid_row_too_large(self):
        """행 번호가 너무 큼"""
        # When/Then
        with pytest.raises(ValueError) as exc_info:
            validate_table_coordinates(101, 1)
        assert "행 번호는 1과 100 사이여야 합니다" in str(exc_info.value)
    
    def test_invalid_col_too_small(self):
        """열 번호가 너무 작음"""
        # When/Then
        with pytest.raises(ValueError) as exc_info:
            validate_table_coordinates(1, 0)
        assert "열 번호는 1과 100 사이여야 합니다" in str(exc_info.value)
    
    def test_custom_max_values(self):
        """사용자 정의 최대값"""
        # When/Then
        validate_table_coordinates(5, 5, max_rows=10, max_cols=10)
        
        with pytest.raises(ValueError):
            validate_table_coordinates(11, 5, max_rows=10, max_cols=10)


class TestGetHwpActionParameter:
    """get_hwp_action_parameter 함수 테스트"""
    
    @pytest.fixture
    def mock_hwp(self):
        """Mock HWP 객체"""
        hwp = Mock()
        hwp.HParameterSet = Mock()
        hwp.HAction = Mock()
        return hwp
    
    def test_get_parameter_success(self, mock_hwp):
        """파라미터 가져오기 성공"""
        # Given
        mock_param_set = Mock()
        mock_param_set.HSet = Mock()
        setattr(mock_hwp.HParameterSet, "TestParam", mock_param_set)
        
        # When
        result = get_hwp_action_parameter(mock_hwp, "TestAction", "TestParam")
        
        # Then
        assert result == mock_param_set
        mock_hwp.HAction.GetDefault.assert_called_once_with("TestAction", mock_param_set.HSet)
    
    def test_get_parameter_failure(self, mock_hwp):
        """파라미터 가져오기 실패"""
        # Given
        mock_hwp.HParameterSet = Mock(spec=[])  # TestParam 속성이 없음
        
        # When/Then
        with pytest.raises(AttributeError):
            get_hwp_action_parameter(mock_hwp, "TestAction", "TestParam")


class TestLogOperationResult:
    """log_operation_result 함수 테스트"""
    
    @patch('tools.hwp_utils.logger')
    def test_log_success_without_details(self, mock_logger):
        """성공 로그 (세부사항 없음)"""
        # When
        log_operation_result("테스트 작업", True)
        
        # Then
        mock_logger.info.assert_called_once_with("테스트 작업 성공")
    
    @patch('tools.hwp_utils.logger')
    def test_log_success_with_details(self, mock_logger):
        """성공 로그 (세부사항 포함)"""
        # When
        log_operation_result("테스트 작업", True, "추가 정보")
        
        # Then
        mock_logger.info.assert_called_once_with("테스트 작업 성공: 추가 정보")
    
    @patch('tools.hwp_utils.logger')
    def test_log_failure_without_details(self, mock_logger):
        """실패 로그 (세부사항 없음)"""
        # When
        log_operation_result("테스트 작업", False)
        
        # Then
        mock_logger.error.assert_called_once_with("테스트 작업 실패")
    
    @patch('tools.hwp_utils.logger')
    def test_log_failure_with_details(self, mock_logger):
        """실패 로그 (세부사항 포함)"""
        # When
        log_operation_result("테스트 작업", False, "오류 정보")
        
        # Then
        mock_logger.error.assert_called_once_with("테스트 작업 실패: 오류 정보")