"""
HWP 예외 클래스 테스트
hwp_exceptions.py의 모든 예외 클래스에 대한 단위 테스트를 포함합니다.
"""

import pytest
import sys
import os

# 테스트를 위한 경로 설정
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tools.hwp_exceptions import (
    HwpError,
    HwpConnectionError,
    HwpNotRunningError,
    HwpDocumentError,
    HwpDocumentNotFoundError,
    HwpDocumentAccessError,
    HwpDocumentSaveError,
    HwpTableError,
    HwpTableNotFoundError,
    HwpTableCellError,
    HwpTableRangeError,
    HwpImageError,
    HwpImageNotFoundError,
    HwpImageFormatError,
    HwpPDFError,
    HwpPDFExportError,
    HwpTemplateError,
    HwpTemplateNotFoundError,
    HwpFieldError,
    HwpFieldNotFoundError,
    HwpParameterError,
    HwpInvalidParameterError,
    HwpOperationError,
    HwpOperationNotAllowedError,
    HwpBatchError,
    HwpTransactionError,
    HwpChunkProcessingError,
    HwpTimeoutError,
    handle_hwp_error
)


class TestBasicExceptions:
    """기본 예외 클래스 테스트"""
    
    def test_hwp_error(self):
        """기본 HwpError 예외"""
        # When
        error = HwpError("테스트 오류")
        
        # Then
        assert str(error) == "테스트 오류"
        assert isinstance(error, Exception)


class TestConnectionExceptions:
    """연결 관련 예외 클래스 테스트"""
    
    def test_hwp_connection_error_default(self):
        """HwpConnectionError 기본 메시지"""
        # When
        error = HwpConnectionError()
        
        # Then
        assert str(error) == "HWP 프로그램에 연결할 수 없습니다."
        assert isinstance(error, HwpError)
    
    def test_hwp_connection_error_custom(self):
        """HwpConnectionError 사용자 정의 메시지"""
        # When
        error = HwpConnectionError("사용자 정의 메시지")
        
        # Then
        assert str(error) == "사용자 정의 메시지"
    
    def test_hwp_not_running_error(self):
        """HwpNotRunningError"""
        # When
        error = HwpNotRunningError()
        
        # Then
        assert str(error) == "HWP 프로그램이 실행 중이지 않습니다. 먼저 한글을 실행해주세요."
        assert isinstance(error, HwpConnectionError)


class TestDocumentExceptions:
    """문서 관련 예외 클래스 테스트"""
    
    def test_hwp_document_error(self):
        """HwpDocumentError"""
        # When
        error = HwpDocumentError("문서 오류")
        
        # Then
        assert str(error) == "문서 오류"
        assert isinstance(error, HwpError)
    
    def test_hwp_document_not_found_error(self):
        """HwpDocumentNotFoundError"""
        # When
        error = HwpDocumentNotFoundError("/path/to/file.hwp")
        
        # Then
        assert str(error) == "문서를 찾을 수 없습니다: /path/to/file.hwp"
        assert error.path == "/path/to/file.hwp"
        assert isinstance(error, HwpDocumentError)
    
    def test_hwp_document_access_error(self):
        """HwpDocumentAccessError"""
        # When
        error = HwpDocumentAccessError("/path/to/protected.hwp")
        
        # Then
        assert str(error) == "문서에 접근할 수 없습니다: /path/to/protected.hwp"
        assert error.path == "/path/to/protected.hwp"
    
    def test_hwp_document_save_error_without_reason(self):
        """HwpDocumentSaveError (이유 없음)"""
        # When
        error = HwpDocumentSaveError("/path/to/save.hwp")
        
        # Then
        assert str(error) == "문서를 저장할 수 없습니다: /path/to/save.hwp"
        assert error.path == "/path/to/save.hwp"
    
    def test_hwp_document_save_error_with_reason(self):
        """HwpDocumentSaveError (이유 포함)"""
        # When
        error = HwpDocumentSaveError("/path/to/save.hwp", "디스크 공간 부족")
        
        # Then
        assert str(error) == "문서를 저장할 수 없습니다: /path/to/save.hwp (디스크 공간 부족)"


class TestTableExceptions:
    """표 관련 예외 클래스 테스트"""
    
    def test_hwp_table_error(self):
        """HwpTableError"""
        # When
        error = HwpTableError("표 오류")
        
        # Then
        assert str(error) == "표 오류"
        assert isinstance(error, HwpError)
    
    def test_hwp_table_not_found_error(self):
        """HwpTableNotFoundError"""
        # When
        error = HwpTableNotFoundError()
        
        # Then
        assert str(error) == "현재 위치에 표가 없습니다."
    
    def test_hwp_table_cell_error_without_message(self):
        """HwpTableCellError (메시지 없음)"""
        # When
        error = HwpTableCellError(2, 3)
        
        # Then
        assert str(error) == "셀(2, 3) 오류"
        assert error.row == 2
        assert error.col == 3
    
    def test_hwp_table_cell_error_with_message(self):
        """HwpTableCellError (메시지 포함)"""
        # When
        error = HwpTableCellError(2, 3, "병합된 셀")
        
        # Then
        assert str(error) == "셀(2, 3) 오류: 병합된 셀"
    
    def test_hwp_table_range_error(self):
        """HwpTableRangeError"""
        # When
        error = HwpTableRangeError(11, 5, 10, 4)
        
        # Then
        assert str(error) == "셀 범위를 초과했습니다. 요청: (11, 5), 최대: (10, 4)"
        assert error.row == 11
        assert error.col == 5
        assert error.max_row == 10
        assert error.max_col == 4


class TestImageExceptions:
    """이미지 관련 예외 클래스 테스트"""
    
    def test_hwp_image_error(self):
        """HwpImageError"""
        # When
        error = HwpImageError("이미지 오류")
        
        # Then
        assert str(error) == "이미지 오류"
        assert isinstance(error, HwpError)
    
    def test_hwp_image_not_found_error(self):
        """HwpImageNotFoundError"""
        # When
        error = HwpImageNotFoundError("/path/to/image.jpg")
        
        # Then
        assert str(error) == "이미지 파일을 찾을 수 없습니다: /path/to/image.jpg"
        assert error.path == "/path/to/image.jpg"
    
    def test_hwp_image_format_error(self):
        """HwpImageFormatError"""
        # When
        error = HwpImageFormatError("BMP")
        
        # Then
        assert str(error) == "지원하지 않는 이미지 형식입니다: BMP"
        assert error.format == "BMP"


class TestPDFExceptions:
    """PDF 관련 예외 클래스 테스트"""
    
    def test_hwp_pdf_error(self):
        """HwpPDFError"""
        # When
        error = HwpPDFError("PDF 오류")
        
        # Then
        assert str(error) == "PDF 오류"
        assert isinstance(error, HwpError)
    
    def test_hwp_pdf_export_error_without_reason(self):
        """HwpPDFExportError (이유 없음)"""
        # When
        error = HwpPDFExportError("/path/to/output.pdf")
        
        # Then
        assert str(error) == "PDF 변환에 실패했습니다: /path/to/output.pdf"
        assert error.path == "/path/to/output.pdf"
    
    def test_hwp_pdf_export_error_with_reason(self):
        """HwpPDFExportError (이유 포함)"""
        # When
        error = HwpPDFExportError("/path/to/output.pdf", "글꼴 오류")
        
        # Then
        assert str(error) == "PDF 변환에 실패했습니다: /path/to/output.pdf (글꼴 오류)"


class TestTemplateExceptions:
    """템플릿 관련 예외 클래스 테스트"""
    
    def test_hwp_template_error(self):
        """HwpTemplateError"""
        # When
        error = HwpTemplateError("템플릿 오류")
        
        # Then
        assert str(error) == "템플릿 오류"
        assert isinstance(error, HwpError)
    
    def test_hwp_template_not_found_error(self):
        """HwpTemplateNotFoundError"""
        # When
        error = HwpTemplateNotFoundError("report_template")
        
        # Then
        assert str(error) == "템플릿을 찾을 수 없습니다: report_template"
        assert error.template_name == "report_template"


class TestFieldExceptions:
    """필드 관련 예외 클래스 테스트"""
    
    def test_hwp_field_error(self):
        """HwpFieldError"""
        # When
        error = HwpFieldError("필드 오류")
        
        # Then
        assert str(error) == "필드 오류"
        assert isinstance(error, HwpError)
    
    def test_hwp_field_not_found_error(self):
        """HwpFieldNotFoundError"""
        # When
        error = HwpFieldNotFoundError("{{customer_name}}")
        
        # Then
        assert str(error) == "필드를 찾을 수 없습니다: {{customer_name}}"
        assert error.field_name == "{{customer_name}}"


class TestParameterExceptions:
    """매개변수 관련 예외 클래스 테스트"""
    
    def test_hwp_parameter_error(self):
        """HwpParameterError"""
        # When
        error = HwpParameterError("매개변수 오류")
        
        # Then
        assert str(error) == "매개변수 오류"
        assert isinstance(error, HwpError)
    
    def test_hwp_invalid_parameter_error(self):
        """HwpInvalidParameterError"""
        # When
        error = HwpInvalidParameterError("font_size", -10, "양수")
        
        # Then
        assert str(error) == "잘못된 매개변수 'font_size': -10 (예상: 양수)"
        assert error.param_name == "font_size"
        assert error.value == -10
        assert error.expected == "양수"


class TestOperationExceptions:
    """작업 관련 예외 클래스 테스트"""
    
    def test_hwp_operation_error(self):
        """HwpOperationError"""
        # When
        error = HwpOperationError("작업 오류")
        
        # Then
        assert str(error) == "작업 오류"
        assert isinstance(error, HwpError)
    
    def test_hwp_operation_not_allowed_error_without_reason(self):
        """HwpOperationNotAllowedError (이유 없음)"""
        # When
        error = HwpOperationNotAllowedError("문서 병합")
        
        # Then
        assert str(error) == "허용되지 않는 작업입니다: 문서 병합"
        assert error.operation == "문서 병합"
    
    def test_hwp_operation_not_allowed_error_with_reason(self):
        """HwpOperationNotAllowedError (이유 포함)"""
        # When
        error = HwpOperationNotAllowedError("문서 병합", "보호된 문서")
        
        # Then
        assert str(error) == "허용되지 않는 작업입니다: 문서 병합 (보호된 문서)"


class TestBatchExceptions:
    """배치 작업 관련 예외 클래스 테스트"""
    
    def test_hwp_batch_error(self):
        """HwpBatchError"""
        # When
        error = HwpBatchError("배치 오류")
        
        # Then
        assert str(error) == "배치 오류"
        assert isinstance(error, HwpError)
    
    def test_hwp_transaction_error(self):
        """HwpTransactionError"""
        # When
        error = HwpTransactionError("TX123", "롤백 실패")
        
        # Then
        assert str(error) == "트랜잭션 TX123 오류: 롤백 실패"
        assert error.transaction_id == "TX123"
        assert isinstance(error, HwpBatchError)
    
    def test_hwp_chunk_processing_error(self):
        """HwpChunkProcessingError"""
        # When
        error = HwpChunkProcessingError(3, 10, "데이터 형식 오류")
        
        # Then
        assert str(error) == "청크 3/10 처리 오류: 데이터 형식 오류"
        assert error.chunk_index == 3
        assert error.total_chunks == 10


class TestTimeoutException:
    """시간 초과 예외 클래스 테스트"""
    
    def test_hwp_timeout_error(self):
        """HwpTimeoutError"""
        # When
        error = HwpTimeoutError("PDF 변환", 30)
        
        # Then
        assert str(error) == "작업 시간이 초과되었습니다: PDF 변환 (30초)"
        assert error.operation == "PDF 변환"
        assert error.timeout == 30


class TestExceptionHandler:
    """예외 처리 헬퍼 함수 테스트"""
    
    def test_handle_hwp_error_success(self):
        """정상 실행"""
        # Given
        @handle_hwp_error
        def test_func():
            return "success"
        
        # When
        result = test_func()
        
        # Then
        assert result == "success"
    
    def test_handle_hwp_error_with_hwp_exception(self):
        """HWP 예외는 그대로 발생"""
        # Given
        @handle_hwp_error
        def test_func():
            raise HwpDocumentError("문서 오류")
        
        # When/Then
        with pytest.raises(HwpDocumentError) as exc_info:
            test_func()
        assert str(exc_info.value) == "문서 오류"
    
    def test_handle_hwp_error_with_file_not_found(self):
        """FileNotFoundError를 HwpDocumentNotFoundError로 변환"""
        # Given
        @handle_hwp_error
        def test_func():
            raise FileNotFoundError("test.hwp")
        
        # When/Then
        with pytest.raises(HwpDocumentNotFoundError) as exc_info:
            test_func()
        assert "test.hwp" in str(exc_info.value)
    
    def test_handle_hwp_error_with_permission_error(self):
        """PermissionError를 HwpDocumentAccessError로 변환"""
        # Given
        @handle_hwp_error
        def test_func():
            raise PermissionError("Access denied")
        
        # When/Then
        with pytest.raises(HwpDocumentAccessError) as exc_info:
            test_func()
        assert "Access denied" in str(exc_info.value)
    
    def test_handle_hwp_error_with_generic_exception(self):
        """일반 예외를 HwpError로 변환"""
        # Given
        @handle_hwp_error
        def test_func():
            raise ValueError("Invalid value")
        
        # When/Then
        with pytest.raises(HwpError) as exc_info:
            test_func()
        assert "예상치 못한 오류: Invalid value" in str(exc_info.value)