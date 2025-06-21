"""
차트 및 배치 기능 테스트
"""

import pytest
from unittest.mock import Mock, MagicMock, patch
import sys
import os
import json

# 테스트를 위한 경로 설정
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tools.hwp_chart_features import HwpChartFeatures
from tools.hwp_batch_processor import HwpBatchProcessor
from tools.hwp_exceptions import HwpBatchError, HwpTransactionError


class TestChartFeatures:
    """차트 기능 테스트"""
    
    @pytest.fixture
    def mock_hwp_controller(self):
        """Mock HWP controller"""
        controller = Mock()
        controller.hwp = Mock()
        controller.insert_table = Mock(return_value=True)
        controller.fill_table_cell = Mock(return_value=True)
        controller.insert_text = Mock(return_value=True)
        return controller
    
    @pytest.fixture
    def chart_features(self, mock_hwp_controller):
        """차트 기능 인스턴스"""
        return HwpChartFeatures(mock_hwp_controller)
    
    def test_insert_chart_default(self, chart_features, mock_hwp_controller):
        """기본 차트 삽입 테스트"""
        # When
        result = chart_features.insert_chart()
        
        # Then
        assert result is True
        mock_hwp_controller.insert_table.assert_called_once()
        assert mock_hwp_controller.fill_table_cell.call_count > 0
    
    def test_insert_chart_with_data(self, chart_features, mock_hwp_controller):
        """데이터를 포함한 차트 삽입 테스트"""
        # Given
        data = [
            ["Category", "Value"],
            ["A", 10],
            ["B", 20],
            ["C", 15]
        ]
        
        # When
        result = chart_features.insert_chart("bar", data, "Test Chart")
        
        # Then
        assert result is True
        mock_hwp_controller.insert_table.assert_called_with(4, 2)
        assert mock_hwp_controller.fill_table_cell.call_count == 8  # 4행 x 2열
    
    def test_insert_simple_chart(self, chart_features):
        """간단한 차트 삽입 테스트"""
        # Given
        values = [25, 35, 20, 20]
        labels = ["Q1", "Q2", "Q3", "Q4"]
        
        # When
        result = chart_features.insert_simple_chart(values, labels, "pie", "Sales")
        
        # Then
        assert result is True
    
    def test_insert_equation(self, chart_features, mock_hwp_controller):
        """수식 삽입 테스트"""
        # Given
        equation = "x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}"
        
        # When
        result = chart_features.insert_equation(equation)
        
        # Then
        assert result is True
        mock_hwp_controller.hwp.Run.assert_any_call("InsertEquation")
        mock_hwp_controller.hwp.Run.assert_any_call("CloseEx")
    
    def test_insert_equation_template(self, chart_features):
        """수식 템플릿 삽입 테스트"""
        # When
        result = chart_features.insert_equation_template("quadratic")
        
        # Then
        assert result is True
    
    def test_latex_to_hwp_conversion(self, chart_features):
        """LaTeX를 HWP 수식으로 변환 테스트"""
        # Given
        latex = "\\frac{a}{b} + \\sqrt{x}"
        
        # When
        hwp_text = chart_features._convert_latex_to_hwp(latex)
        
        # Then
        assert "over" in hwp_text
        assert "sqrt" in hwp_text
        assert "\\frac" not in hwp_text


class TestBatchProcessor:
    """배치 처리 기능 테스트"""
    
    @pytest.fixture
    def mock_hwp_controller(self):
        """Mock HWP controller"""
        controller = Mock()
        controller.hwp = Mock()
        controller.insert_text = Mock(return_value=True)
        controller.insert_table = Mock(return_value=True)
        controller.insert_paragraph = Mock(return_value=True)
        controller.save_document = Mock(return_value=True)
        controller.open_document = Mock(return_value=True)
        controller.create_new_document = Mock(return_value=True)
        controller.fill_table_cell = Mock(return_value=True)
        return controller
    
    @pytest.fixture
    def batch_processor(self, mock_hwp_controller):
        """배치 프로세서 인스턴스"""
        return HwpBatchProcessor(mock_hwp_controller)
    
    def test_transaction_success(self, batch_processor, mock_hwp_controller):
        """트랜잭션 성공 테스트"""
        # Given
        operations_executed = []
        
        # When
        with batch_processor.transaction() as txn_id:
            operations_executed.append("op1")
            operations_executed.append("op2")
        
        # Then
        assert len(operations_executed) == 2
        assert mock_hwp_controller.save_document.called  # 저장점 생성
    
    def test_transaction_rollback(self, batch_processor, mock_hwp_controller):
        """트랜잭션 롤백 테스트"""
        # When/Then
        with pytest.raises(HwpBatchError):
            with batch_processor.transaction():
                raise Exception("Test error")
        
        # 롤백 확인
        assert mock_hwp_controller.open_document.called
    
    def test_execute_batch_success(self, batch_processor):
        """배치 작업 성공 테스트"""
        # Given
        operations = [
            {"action": "insert_text", "params": {"text": "Hello"}},
            {"action": "insert_paragraph"},
            {"action": "insert_text", "params": {"text": "World"}}
        ]
        
        # When
        result = batch_processor.execute_batch(operations, use_transaction=False)
        
        # Then
        assert result["success"] is True
        assert result["executed"] == 3
        assert result["failed"] == 0
    
    def test_execute_batch_with_error(self, batch_processor, mock_hwp_controller):
        """배치 작업 오류 처리 테스트"""
        # Given
        mock_hwp_controller.insert_table.side_effect = Exception("Table error")
        operations = [
            {"action": "insert_text", "params": {"text": "Hello"}},
            {"action": "insert_table", "params": {"rows": 3, "cols": 3}},
            {"action": "insert_text", "params": {"text": "World"}}
        ]
        
        # When
        result = batch_processor.execute_batch(operations, use_transaction=False, stop_on_error=True)
        
        # Then
        assert result["success"] is False
        assert result["executed"] == 1
        assert result["failed"] == 1
        assert len(result["errors"]) == 1
    
    def test_insert_large_table_data(self, batch_processor, mock_hwp_controller):
        """대용량 표 데이터 삽입 테스트"""
        # Given
        data = [[f"Row{i}", f"Data{i}"] for i in range(200)]
        progress_values = []
        
        def progress_callback(progress, current, total):
            progress_values.append(progress)
        
        # When
        result = batch_processor.insert_large_table_data(
            data, 
            chunk_size=50,
            progress_callback=progress_callback
        )
        
        # Then
        assert result is True
        assert mock_hwp_controller.insert_table.called
        assert len(progress_values) > 0
        assert progress_values[-1] == 100.0
    
    def test_process_multiple_documents(self, batch_processor):
        """여러 문서 처리 테스트"""
        # Given
        document_tasks = [
            {
                "filename": "doc1.hwp",
                "operations": [
                    {"action": "insert_text", "params": {"text": "Document 1"}}
                ]
            },
            {
                "filename": "doc2.hwp",
                "operations": [
                    {"action": "insert_text", "params": {"text": "Document 2"}}
                ]
            }
        ]
        
        # When
        result = batch_processor.process_multiple_documents(document_tasks)
        
        # Then
        assert result["total"] == 2
        assert result["success"] == 2
        assert result["failed"] == 0
    
    def test_split_data_for_parallel(self, batch_processor):
        """병렬 처리를 위한 데이터 분할 테스트"""
        # Given
        data = [[i, f"data{i}"] for i in range(100)]
        
        # When
        chunks = batch_processor.split_table_data_for_parallel(data, num_workers=4)
        
        # Then
        assert len(chunks) == 4
        assert sum(len(chunk) for chunk in chunks) == 100
        assert len(chunks[0]) == 25