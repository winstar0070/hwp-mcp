"""
HWP 배치 작업 처리 모듈
트랜잭션 처리, 대용량 데이터 처리 등 배치 작업 관련 기능 제공
"""
import os
import logging
import time
from typing import List, Dict, Any, Optional, Callable
from contextlib import contextmanager
import json
import tempfile

try:
    from .constants import (
        BATCH_CHUNK_SIZE, MAX_RETRY_COUNT, 
        RETRY_DELAY, BATCH_OPERATION_TIMEOUT
    )
    from .hwp_utils import (
        execute_with_retry, log_operation_result,
        require_hwp_connection
    )
    from .hwp_exceptions import (
        HwpBatchError, HwpOperationError
    )
except ImportError:
    # 상수 기본값
    BATCH_CHUNK_SIZE = 100
    MAX_RETRY_COUNT = 3
    RETRY_DELAY = 1
    BATCH_OPERATION_TIMEOUT = 300
    
    # 더미 클래스/함수
    class HwpBatchError(Exception):
        pass
    class HwpOperationError(Exception):
        pass
    
    def execute_with_retry(func, *args, **kwargs):
        return func(*args, **kwargs)
    def log_operation_result(*args, **kwargs):
        pass
    def require_hwp_connection(func):
        return func

logger = logging.getLogger(__name__)

class HwpBatchProcessor:
    """HWP 배치 작업 처리를 위한 클래스"""
    
    def __init__(self, hwp_controller):
        """
        HwpBatchProcessor 초기화
        
        Args:
            hwp_controller: HwpController 인스턴스
        """
        self.hwp = hwp_controller.hwp
        self.hwp_controller = hwp_controller
        self._transaction_stack = []
        self._in_transaction = False
    
    # ============== 트랜잭션 처리 ==============
    
    @contextmanager
    def transaction(self, save_point: bool = True):
        """
        트랜잭션 컨텍스트 매니저
        
        Args:
            save_point (bool): 저장 지점 생성 여부
            
        Usage:
            with batch_processor.transaction():
                # 여러 작업 수행
                # 오류 발생 시 자동 롤백
        """
        transaction_id = f"txn_{time.time()}"
        temp_file = None
        
        try:
            # 트랜잭션 시작
            self._in_transaction = True
            
            # 저장 지점 생성
            if save_point:
                temp_file = self._create_savepoint()
                self._transaction_stack.append({
                    "id": transaction_id,
                    "savepoint": temp_file,
                    "start_time": time.time()
                })
            
            logger.info(f"트랜잭션 시작: {transaction_id}")
            yield transaction_id
            
            # 트랜잭션 성공 - 커밋
            if save_point and temp_file:
                os.remove(temp_file)  # 임시 파일 정리
            self._transaction_stack.pop() if self._transaction_stack else None
            
            logger.info(f"트랜잭션 커밋: {transaction_id}")
            
        except Exception as e:
            # 트랜잭션 실패 - 롤백
            logger.error(f"트랜잭션 실패, 롤백 수행: {transaction_id} - {e}")
            
            if save_point and self._transaction_stack:
                self._rollback_to_savepoint(self._transaction_stack[-1]["savepoint"])
                self._transaction_stack.pop()
            
            raise HwpBatchError(f"트랜잭션 실패: {e}")
            
        finally:
            self._in_transaction = False
    
    def _create_savepoint(self) -> str:
        """현재 문서 상태를 임시 파일로 저장합니다."""
        temp_file = tempfile.mktemp(suffix=".hwp")
        self.hwp_controller.save_document(temp_file)
        return temp_file
    
    def _rollback_to_savepoint(self, savepoint_file: str):
        """저장 지점으로 문서를 복원합니다."""
        if os.path.exists(savepoint_file):
            self.hwp_controller.open_document(savepoint_file)
            os.remove(savepoint_file)
    
    # ============== 배치 작업 실행 ==============
    
    @require_hwp_connection
    def execute_batch(self, operations: List[Dict[str, Any]], 
                     use_transaction: bool = True,
                     stop_on_error: bool = True) -> Dict[str, Any]:
        """
        배치 작업을 실행합니다.
        
        Args:
            operations (List[Dict]): 실행할 작업 목록
                예: [
                    {"action": "insert_text", "params": {"text": "Hello"}},
                    {"action": "insert_table", "params": {"rows": 3, "cols": 3}}
                ]
            use_transaction (bool): 트랜잭션 사용 여부
            stop_on_error (bool): 오류 발생 시 중단 여부
            
        Returns:
            Dict: 실행 결과
                {
                    "success": bool,
                    "executed": int,
                    "failed": int,
                    "results": List[Dict],
                    "errors": List[str]
                }
        """
        results = {
            "success": True,
            "executed": 0,
            "failed": 0,
            "results": [],
            "errors": []
        }
        
        context_manager = self.transaction() if use_transaction else self._dummy_context()
        
        with context_manager:
            for idx, operation in enumerate(operations):
                try:
                    # 작업 실행
                    result = self._execute_operation(operation)
                    results["results"].append({
                        "index": idx,
                        "operation": operation["action"],
                        "success": True,
                        "result": result
                    })
                    results["executed"] += 1
                    
                except Exception as e:
                    error_msg = f"작업 {idx} 실패: {operation['action']} - {str(e)}"
                    logger.error(error_msg)
                    
                    results["results"].append({
                        "index": idx,
                        "operation": operation["action"],
                        "success": False,
                        "error": str(e)
                    })
                    results["failed"] += 1
                    results["errors"].append(error_msg)
                    
                    if stop_on_error:
                        results["success"] = False
                        break
        
        results["success"] = results["failed"] == 0
        return results
    
    def _execute_operation(self, operation: Dict[str, Any]) -> Any:
        """단일 작업을 실행합니다."""
        action = operation.get("action")
        params = operation.get("params", {})
        
        # 작업 매핑
        action_map = {
            "insert_text": self.hwp_controller.insert_text,
            "insert_table": self.hwp_controller.insert_table,
            "insert_image": self.hwp_controller.insert_image,
            "insert_paragraph": self.hwp_controller.insert_paragraph,
            "save_document": self.hwp_controller.save_document,
            "fill_table_cell": self.hwp_controller.fill_table_cell,
            "set_font_style": self.hwp_controller.set_font_style,
            # 추가 작업들...
        }
        
        if action not in action_map:
            raise HwpOperationError(f"알 수 없는 작업: {action}")
        
        # 파라미터 언패킹하여 함수 호출
        return action_map[action](**params)
    
    @contextmanager
    def _dummy_context(self):
        """트랜잭션을 사용하지 않을 때의 더미 컨텍스트 매니저"""
        yield None
    
    # ============== 대용량 데이터 처리 ==============
    
    @require_hwp_connection
    def insert_large_table_data(self, data: List[List[Any]], 
                              chunk_size: int = None,
                              progress_callback: Optional[Callable] = None) -> bool:
        """
        대용량 표 데이터를 청크 단위로 처리합니다.
        
        Args:
            data (List[List[Any]]): 표 데이터
            chunk_size (int): 청크 크기
            progress_callback (Callable): 진행률 콜백 함수
            
        Returns:
            bool: 성공 여부
        """
        if chunk_size is None:
            chunk_size = BATCH_CHUNK_SIZE
        
        total_rows = len(data)
        if total_rows == 0:
            return True
        
        # 첫 번째 청크로 표 생성
        first_chunk = data[:chunk_size]
        cols = len(first_chunk[0]) if first_chunk else 0
        
        # 초기 표 생성 (전체 크기로)
        if not self.hwp_controller.insert_table(total_rows, cols):
            logger.error("표 생성 실패")
            return False
        
        # 청크 단위로 데이터 입력
        for chunk_start in range(0, total_rows, chunk_size):
            chunk_end = min(chunk_start + chunk_size, total_rows)
            chunk_data = data[chunk_start:chunk_end]
            
            try:
                # 청크 데이터 입력
                for row_idx, row_data in enumerate(chunk_data):
                    actual_row = chunk_start + row_idx + 1
                    
                    for col_idx, cell_value in enumerate(row_data):
                        self.hwp_controller.fill_table_cell(
                            actual_row,
                            col_idx + 1,
                            str(cell_value)
                        )
                
                # 진행률 콜백
                if progress_callback:
                    progress = (chunk_end / total_rows) * 100
                    progress_callback(progress, chunk_end, total_rows)
                
                # 메모리 관리를 위한 짧은 대기
                time.sleep(0.01)
                
            except Exception as e:
                logger.error(f"청크 처리 실패: 행 {chunk_start}-{chunk_end} - {e}")
                return False
        
        log_operation_result("대용량 표 데이터 입력", True, f"{total_rows}행 처리 완료")
        return True
    
    @require_hwp_connection
    def process_multiple_documents(self, 
                                 document_tasks: List[Dict[str, Any]],
                                 output_dir: str = None) -> Dict[str, Any]:
        """
        여러 문서를 순차적으로 처리합니다.
        
        Args:
            document_tasks (List[Dict]): 문서별 작업 목록
                예: [
                    {
                        "filename": "report1.hwp",
                        "operations": [작업 목록]
                    }
                ]
            output_dir (str): 출력 디렉토리
            
        Returns:
            Dict: 처리 결과
        """
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        results = {
            "total": len(document_tasks),
            "success": 0,
            "failed": 0,
            "documents": []
        }
        
        for doc_task in document_tasks:
            filename = doc_task.get("filename", f"document_{time.time()}.hwp")
            operations = doc_task.get("operations", [])
            
            try:
                # 새 문서 생성
                self.hwp_controller.create_new_document()
                
                # 작업 실행
                batch_result = self.execute_batch(operations, use_transaction=True)
                
                # 문서 저장
                if output_dir:
                    output_path = os.path.join(output_dir, filename)
                else:
                    output_path = filename
                
                self.hwp_controller.save_document(output_path)
                
                results["documents"].append({
                    "filename": filename,
                    "path": output_path,
                    "success": batch_result["success"],
                    "operations": batch_result["executed"]
                })
                
                if batch_result["success"]:
                    results["success"] += 1
                else:
                    results["failed"] += 1
                    
            except Exception as e:
                logger.error(f"문서 처리 실패: {filename} - {e}")
                results["failed"] += 1
                results["documents"].append({
                    "filename": filename,
                    "success": False,
                    "error": str(e)
                })
        
        return results
    
    # ============== 병렬 처리 시뮬레이션 ==============
    
    def split_table_data_for_parallel(self, data: List[List[Any]], 
                                     num_workers: int = 4) -> List[List[List[Any]]]:
        """
        표 데이터를 병렬 처리를 위해 분할합니다.
        
        Args:
            data (List[List[Any]]): 원본 데이터
            num_workers (int): 작업자 수
            
        Returns:
            List[List[List[Any]]]: 분할된 데이터 청크
        """
        total_rows = len(data)
        chunk_size = (total_rows + num_workers - 1) // num_workers
        
        chunks = []
        for i in range(0, total_rows, chunk_size):
            chunk = data[i:i + chunk_size]
            chunks.append(chunk)
        
        return chunks