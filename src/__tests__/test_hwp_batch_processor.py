"""
HWP 배치 프로세서 전체 기능 테스트
hwp_batch_processor.py의 모든 기능에 대한 단위 테스트를 포함합니다.
"""

import pytest
from unittest.mock import Mock, MagicMock, patch, call
import sys
import os
import tempfile
import time

# 테스트를 위한 경로 설정
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tools.hwp_batch_processor import HwpBatchProcessor
from tools.hwp_exceptions import HwpBatchError, HwpOperationError


class TestBatchProcessorInit:
    """HwpBatchProcessor 초기화 테스트"""
    
    def test_initialization(self):
        """정상적인 초기화"""
        # Given
        mock_controller = Mock()
        mock_controller.hwp = Mock()
        
        # When
        processor = HwpBatchProcessor(mock_controller)
        
        # Then
        assert processor.hwp == mock_controller.hwp
        assert processor.hwp_controller == mock_controller
        assert processor._transaction_stack == []
        assert processor._in_transaction is False


class TestTransactionManagement:
    """트랜잭션 관리 기능 테스트"""
    
    @pytest.fixture
    def processor(self):
        """테스트용 프로세서 생성"""
        mock_controller = Mock()
        mock_controller.hwp = Mock()
        mock_controller.save_document = Mock(return_value=True)
        mock_controller.open_document = Mock(return_value=True)
        return HwpBatchProcessor(mock_controller)
    
    def test_transaction_success(self, processor):
        """트랜잭션 성공 케이스"""
        # Given
        with patch('tempfile.mktemp', return_value='temp.hwp'):
            with patch('os.remove') as mock_remove:
                # When
                with processor.transaction() as txn_id:
                    assert processor._in_transaction is True
                    assert len(processor._transaction_stack) == 1
                    assert txn_id.startswith('txn_')
                
                # Then
                assert processor._in_transaction is False
                assert len(processor._transaction_stack) == 0
                mock_remove.assert_called_once_with('temp.hwp')
    
    def test_transaction_failure_and_rollback(self, processor):
        """트랜잭션 실패 및 롤백"""
        # Given
        with patch('tempfile.mktemp', return_value='temp.hwp'):
            with patch('os.path.exists', return_value=True):
                with patch('os.remove') as mock_remove:
                    # When/Then
                    with pytest.raises(HwpBatchError):
                        with processor.transaction():
                            assert processor._in_transaction is True
                            raise Exception("Test error")
                    
                    # 롤백 확인
                    assert processor._in_transaction is False
                    assert len(processor._transaction_stack) == 0
                    processor.hwp_controller.open_document.assert_called_with('temp.hwp')
                    assert mock_remove.call_count == 1  # savepoint 파일 삭제
    
    def test_transaction_without_savepoint(self, processor):
        """저장점 없는 트랜잭션"""
        # When
        with processor.transaction(save_point=False) as txn_id:
            assert processor._in_transaction is True
            assert len(processor._transaction_stack) == 0  # 저장점 없음
        
        # Then
        assert processor._in_transaction is False
    
    def test_nested_transactions(self, processor):
        """중첩 트랜잭션"""
        # Given
        with patch('tempfile.mktemp', side_effect=['temp1.hwp', 'temp2.hwp']):
            with patch('os.remove'):
                # When
                with processor.transaction() as txn1:
                    assert len(processor._transaction_stack) == 1
                    
                    with processor.transaction() as txn2:
                        assert len(processor._transaction_stack) == 2
                        assert txn1 != txn2
                    
                    assert len(processor._transaction_stack) == 1
                
                assert len(processor._transaction_stack) == 0


class TestBatchExecution:
    """배치 작업 실행 테스트"""
    
    @pytest.fixture
    def processor(self):
        """테스트용 프로세서 생성"""
        mock_controller = Mock()
        mock_controller.hwp = Mock()
        mock_controller.insert_text = Mock(return_value=True)
        mock_controller.insert_table = Mock(return_value=True)
        mock_controller.insert_paragraph = Mock(return_value=True)
        mock_controller.save_document = Mock(return_value=True)
        mock_controller.open_document = Mock(return_value=True)
        
        # require_hwp_connection 데코레이터를 위한 설정
        mock_controller.is_hwp_running = True
        
        processor = HwpBatchProcessor(mock_controller)
        processor.is_hwp_running = True  # 데코레이터 통과를 위해
        return processor
    
    def test_execute_batch_success(self, processor):
        """배치 작업 성공"""
        # Given
        operations = [
            {"action": "insert_text", "params": {"text": "Hello"}},
            {"action": "insert_paragraph"},
            {"action": "insert_text", "params": {"text": "World"}}
        ]
        
        # When
        with patch.object(processor, 'transaction', return_value=processor._dummy_context()):
            result = processor.execute_batch(operations, use_transaction=False)
        
        # Then
        assert result["success"] is True
        assert result["executed"] == 3
        assert result["failed"] == 0
        assert len(result["results"]) == 3
        assert all(r["success"] for r in result["results"])
    
    def test_execute_batch_with_error_stop_on_error(self, processor):
        """배치 작업 중 오류 발생 (stop_on_error=True)"""
        # Given
        processor.hwp_controller.insert_table.side_effect = Exception("Table error")
        operations = [
            {"action": "insert_text", "params": {"text": "Hello"}},
            {"action": "insert_table", "params": {"rows": 3, "cols": 3}},
            {"action": "insert_text", "params": {"text": "World"}}
        ]
        
        # When
        with patch.object(processor, 'transaction', return_value=processor._dummy_context()):
            result = processor.execute_batch(operations, use_transaction=False, stop_on_error=True)
        
        # Then
        assert result["success"] is False
        assert result["executed"] == 1
        assert result["failed"] == 1
        assert len(result["results"]) == 2  # 3번째 작업은 실행되지 않음
    
    def test_execute_batch_with_error_continue(self, processor):
        """배치 작업 중 오류 발생 (stop_on_error=False)"""
        # Given
        processor.hwp_controller.insert_table.side_effect = Exception("Table error")
        operations = [
            {"action": "insert_text", "params": {"text": "Hello"}},
            {"action": "insert_table", "params": {"rows": 3, "cols": 3}},
            {"action": "insert_text", "params": {"text": "World"}}
        ]
        
        # When
        with patch.object(processor, 'transaction', return_value=processor._dummy_context()):
            result = processor.execute_batch(operations, use_transaction=False, stop_on_error=False)
        
        # Then
        assert result["success"] is False
        assert result["executed"] == 2
        assert result["failed"] == 1
        assert len(result["results"]) == 3
    
    def test_execute_batch_with_transaction(self, processor):
        """트랜잭션을 사용한 배치 작업"""
        # Given
        operations = [
            {"action": "insert_text", "params": {"text": "Test"}}
        ]
        
        # When
        with patch('tempfile.mktemp', return_value='temp.hwp'):
            with patch('os.remove'):
                result = processor.execute_batch(operations, use_transaction=True)
        
        # Then
        assert result["success"] is True
        processor.hwp_controller.save_document.assert_called()  # 저장점 생성
    
    def test_execute_unknown_operation(self, processor):
        """알 수 없는 작업 실행"""
        # Given
        operations = [
            {"action": "unknown_action", "params": {}}
        ]
        
        # When
        with patch.object(processor, 'transaction', return_value=processor._dummy_context()):
            result = processor.execute_batch(operations, use_transaction=False)
        
        # Then
        assert result["success"] is False
        assert result["failed"] == 1
        assert "알 수 없는 작업" in result["errors"][0]


class TestLargeDataProcessing:
    """대용량 데이터 처리 테스트"""
    
    @pytest.fixture
    def processor(self):
        """테스트용 프로세서 생성"""
        mock_controller = Mock()
        mock_controller.hwp = Mock()
        mock_controller.insert_table = Mock(return_value=True)
        mock_controller.fill_table_cell = Mock(return_value=True)
        mock_controller.is_hwp_running = True
        
        processor = HwpBatchProcessor(mock_controller)
        processor.is_hwp_running = True
        return processor
    
    def test_insert_large_table_data_success(self, processor):
        """대용량 표 데이터 삽입 성공"""
        # Given
        data = [[f"Row{i}", f"Data{i}"] for i in range(200)]
        progress_values = []
        
        def progress_callback(progress, current, total):
            progress_values.append((progress, current, total))
        
        # When
        with patch('time.sleep'):  # sleep 호출 무시
            result = processor.insert_large_table_data(
                data, 
                chunk_size=50,
                progress_callback=progress_callback
            )
        
        # Then
        assert result is True
        processor.hwp_controller.insert_table.assert_called_once_with(200, 2)
        assert processor.hwp_controller.fill_table_cell.call_count == 400  # 200행 x 2열
        assert len(progress_values) == 4  # 200/50 = 4 청크
        assert progress_values[-1][0] == 100.0  # 마지막 진행률 100%
    
    def test_insert_large_table_data_empty(self, processor):
        """빈 데이터 처리"""
        # When
        result = processor.insert_large_table_data([])
        
        # Then
        assert result is True
        processor.hwp_controller.insert_table.assert_not_called()
    
    def test_insert_large_table_data_table_creation_failure(self, processor):
        """표 생성 실패"""
        # Given
        processor.hwp_controller.insert_table.return_value = False
        data = [["A", "B"], ["C", "D"]]
        
        # When
        result = processor.insert_large_table_data(data)
        
        # Then
        assert result is False
        processor.hwp_controller.fill_table_cell.assert_not_called()
    
    def test_insert_large_table_data_chunk_failure(self, processor):
        """청크 처리 중 실패"""
        # Given
        data = [[f"Row{i}", f"Data{i}"] for i in range(10)]
        processor.hwp_controller.fill_table_cell.side_effect = [True] * 5 + [Exception("Cell error")] + [True] * 14
        
        # When
        with patch('time.sleep'):
            result = processor.insert_large_table_data(data, chunk_size=5)
        
        # Then
        assert result is False


class TestMultipleDocumentProcessing:
    """여러 문서 처리 테스트"""
    
    @pytest.fixture
    def processor(self):
        """테스트용 프로세서 생성"""
        mock_controller = Mock()
        mock_controller.hwp = Mock()
        mock_controller.create_new_document = Mock(return_value=True)
        mock_controller.save_document = Mock(return_value=True)
        mock_controller.insert_text = Mock(return_value=True)
        mock_controller.is_hwp_running = True
        
        processor = HwpBatchProcessor(mock_controller)
        processor.is_hwp_running = True
        return processor
    
    def test_process_multiple_documents_success(self, processor):
        """여러 문서 성공적으로 처리"""
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
        with patch.object(processor, 'execute_batch', return_value={"success": True, "executed": 1}):
            result = processor.process_multiple_documents(document_tasks)
        
        # Then
        assert result["total"] == 2
        assert result["success"] == 2
        assert result["failed"] == 0
        assert len(result["documents"]) == 2
        assert processor.hwp_controller.create_new_document.call_count == 2
        assert processor.hwp_controller.save_document.call_count == 2
    
    def test_process_multiple_documents_with_output_dir(self, processor):
        """출력 디렉토리 지정하여 문서 처리"""
        # Given
        document_tasks = [
            {"filename": "test.hwp", "operations": []}
        ]
        
        # When
        with tempfile.TemporaryDirectory() as temp_dir:
            with patch.object(processor, 'execute_batch', return_value={"success": True, "executed": 0}):
                result = processor.process_multiple_documents(document_tasks, output_dir=temp_dir)
            
            # Then
            expected_path = os.path.join(temp_dir, "test.hwp")
            processor.hwp_controller.save_document.assert_called_with(expected_path)
            assert result["documents"][0]["path"] == expected_path
    
    def test_process_multiple_documents_with_failure(self, processor):
        """일부 문서 처리 실패"""
        # Given
        document_tasks = [
            {"filename": "success.hwp", "operations": []},
            {"filename": "fail.hwp", "operations": []}
        ]
        
        # execute_batch를 첫 번째는 성공, 두 번째는 실패로 설정
        with patch.object(processor, 'execute_batch', side_effect=[
            {"success": True, "executed": 0},
            {"success": False, "executed": 0}
        ]):
            # When
            result = processor.process_multiple_documents(document_tasks)
            
            # Then
            assert result["total"] == 2
            assert result["success"] == 1
            assert result["failed"] == 1
    
    def test_process_documents_with_exception(self, processor):
        """문서 처리 중 예외 발생"""
        # Given
        processor.hwp_controller.create_new_document.side_effect = Exception("Create error")
        document_tasks = [{"filename": "error.hwp", "operations": []}]
        
        # When
        result = processor.process_multiple_documents(document_tasks)
        
        # Then
        assert result["failed"] == 1
        assert result["documents"][0]["success"] is False
        assert "Create error" in result["documents"][0]["error"]


class TestParallelProcessing:
    """병렬 처리 시뮬레이션 테스트"""
    
    @pytest.fixture
    def processor(self):
        """테스트용 프로세서 생성"""
        mock_controller = Mock()
        return HwpBatchProcessor(mock_controller)
    
    def test_split_table_data_even_division(self, processor):
        """균등 분할"""
        # Given
        data = [[i, f"data{i}"] for i in range(100)]
        
        # When
        chunks = processor.split_table_data_for_parallel(data, num_workers=4)
        
        # Then
        assert len(chunks) == 4
        assert all(len(chunk) == 25 for chunk in chunks)
        assert sum(len(chunk) for chunk in chunks) == 100
    
    def test_split_table_data_uneven_division(self, processor):
        """불균등 분할"""
        # Given
        data = [[i, f"data{i}"] for i in range(97)]
        
        # When
        chunks = processor.split_table_data_for_parallel(data, num_workers=4)
        
        # Then
        assert len(chunks) == 4
        assert len(chunks[0]) == 25  # 97/4 = 24.25 -> 올림해서 25
        assert len(chunks[1]) == 25
        assert len(chunks[2]) == 25
        assert len(chunks[3]) == 22  # 나머지
        assert sum(len(chunk) for chunk in chunks) == 97
    
    def test_split_table_data_small_data(self, processor):
        """작은 데이터 분할"""
        # Given
        data = [[1, "a"], [2, "b"]]
        
        # When
        chunks = processor.split_table_data_for_parallel(data, num_workers=4)
        
        # Then
        assert len(chunks) == 2  # 데이터가 작으면 청크 수도 작음
        assert chunks[0] == [[1, "a"]]
        assert chunks[1] == [[2, "b"]]
    
    def test_split_empty_data(self, processor):
        """빈 데이터 분할"""
        # When
        chunks = processor.split_table_data_for_parallel([], num_workers=4)
        
        # Then
        assert chunks == []


class TestPrivateMethods:
    """내부 메서드 테스트"""
    
    @pytest.fixture
    def processor(self):
        """테스트용 프로세서 생성"""
        mock_controller = Mock()
        mock_controller.save_document = Mock(return_value=True)
        mock_controller.open_document = Mock(return_value=True)
        return HwpBatchProcessor(mock_controller)
    
    def test_create_savepoint(self, processor):
        """저장점 생성"""
        # Given
        with patch('tempfile.mktemp', return_value='savepoint.hwp'):
            # When
            result = processor._create_savepoint()
            
            # Then
            assert result == 'savepoint.hwp'
            processor.hwp_controller.save_document.assert_called_once_with('savepoint.hwp')
    
    def test_rollback_to_savepoint(self, processor):
        """저장점으로 롤백"""
        # Given
        with patch('os.path.exists', return_value=True):
            with patch('os.remove') as mock_remove:
                # When
                processor._rollback_to_savepoint('savepoint.hwp')
                
                # Then
                processor.hwp_controller.open_document.assert_called_once_with('savepoint.hwp')
                mock_remove.assert_called_once_with('savepoint.hwp')
    
    def test_rollback_to_nonexistent_savepoint(self, processor):
        """존재하지 않는 저장점으로 롤백"""
        # Given
        with patch('os.path.exists', return_value=False):
            # When
            processor._rollback_to_savepoint('nonexistent.hwp')
            
            # Then
            processor.hwp_controller.open_document.assert_not_called()