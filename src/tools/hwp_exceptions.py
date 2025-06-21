"""
HWP MCP 전용 예외 클래스 정의
각 기능별로 구체적인 예외를 정의하여 더 나은 에러 처리를 제공합니다.
"""

class HwpError(Exception):
    """HWP 관련 기본 예외 클래스"""
    pass

class HwpConnectionError(HwpError):
    """HWP 프로그램 연결 실패 예외"""
    def __init__(self, message="HWP 프로그램에 연결할 수 없습니다."):
        self.message = message
        super().__init__(self.message)

class HwpNotRunningError(HwpConnectionError):
    """HWP 프로그램이 실행되지 않음"""
    def __init__(self):
        super().__init__("HWP 프로그램이 실행 중이지 않습니다. 먼저 한글을 실행해주세요.")

class HwpDocumentError(HwpError):
    """문서 관련 예외"""
    pass

class HwpDocumentNotFoundError(HwpDocumentError):
    """문서를 찾을 수 없음"""
    def __init__(self, path):
        self.path = path
        super().__init__(f"문서를 찾을 수 없습니다: {path}")

class HwpDocumentAccessError(HwpDocumentError):
    """문서 접근 권한 없음"""
    def __init__(self, path):
        self.path = path
        super().__init__(f"문서에 접근할 수 없습니다: {path}")

class HwpDocumentSaveError(HwpDocumentError):
    """문서 저장 실패"""
    def __init__(self, path, reason=""):
        self.path = path
        msg = f"문서를 저장할 수 없습니다: {path}"
        if reason:
            msg += f" ({reason})"
        super().__init__(msg)

class HwpTableError(HwpError):
    """표 관련 예외"""
    pass

class HwpTableNotFoundError(HwpTableError):
    """표를 찾을 수 없음"""
    def __init__(self):
        super().__init__("현재 위치에 표가 없습니다.")

class HwpTableCellError(HwpTableError):
    """표 셀 관련 예외"""
    def __init__(self, row, col, message=""):
        self.row = row
        self.col = col
        msg = f"셀({row}, {col}) 오류"
        if message:
            msg += f": {message}"
        super().__init__(msg)

class HwpTableRangeError(HwpTableError):
    """표 범위 초과 예외"""
    def __init__(self, row, col, max_row, max_col):
        self.row = row
        self.col = col
        self.max_row = max_row
        self.max_col = max_col
        super().__init__(
            f"셀 범위를 초과했습니다. 요청: ({row}, {col}), 최대: ({max_row}, {max_col})"
        )

class HwpImageError(HwpError):
    """이미지 관련 예외"""
    pass

class HwpImageNotFoundError(HwpImageError):
    """이미지 파일을 찾을 수 없음"""
    def __init__(self, path):
        self.path = path
        super().__init__(f"이미지 파일을 찾을 수 없습니다: {path}")

class HwpImageFormatError(HwpImageError):
    """지원하지 않는 이미지 형식"""
    def __init__(self, format):
        self.format = format
        super().__init__(f"지원하지 않는 이미지 형식입니다: {format}")

class HwpPDFError(HwpError):
    """PDF 변환 관련 예외"""
    pass

class HwpPDFExportError(HwpPDFError):
    """PDF 변환 실패"""
    def __init__(self, path, reason=""):
        self.path = path
        msg = f"PDF 변환에 실패했습니다: {path}"
        if reason:
            msg += f" ({reason})"
        super().__init__(msg)

class HwpTemplateError(HwpError):
    """템플릿 관련 예외"""
    pass

class HwpTemplateNotFoundError(HwpTemplateError):
    """템플릿을 찾을 수 없음"""
    def __init__(self, template_name):
        self.template_name = template_name
        super().__init__(f"템플릿을 찾을 수 없습니다: {template_name}")

class HwpFieldError(HwpError):
    """필드 관련 예외"""
    pass

class HwpFieldNotFoundError(HwpFieldError):
    """필드를 찾을 수 없음"""
    def __init__(self, field_name):
        self.field_name = field_name
        super().__init__(f"필드를 찾을 수 없습니다: {field_name}")

class HwpParameterError(HwpError):
    """매개변수 오류"""
    pass

class HwpInvalidParameterError(HwpParameterError):
    """잘못된 매개변수"""
    def __init__(self, param_name, value, expected):
        self.param_name = param_name
        self.value = value
        self.expected = expected
        super().__init__(
            f"잘못된 매개변수 '{param_name}': {value} (예상: {expected})"
        )

class HwpOperationError(HwpError):
    """작업 실행 오류"""
    pass

class HwpOperationNotAllowedError(HwpOperationError):
    """허용되지 않는 작업"""
    def __init__(self, operation, reason=""):
        self.operation = operation
        msg = f"허용되지 않는 작업입니다: {operation}"
        if reason:
            msg += f" ({reason})"
        super().__init__(msg)

class HwpBatchError(HwpError):
    """배치 작업 관련 예외의 기본 클래스"""
    pass

class HwpTransactionError(HwpBatchError):
    """트랜잭션 처리 중 발생하는 예외"""
    def __init__(self, transaction_id: str, message: str = ""):
        self.transaction_id = transaction_id
        super().__init__(f"트랜잭션 {transaction_id} 오류: {message}")

class HwpChunkProcessingError(HwpBatchError):
    """청크 단위 처리 중 발생하는 예외"""
    def __init__(self, chunk_index: int, total_chunks: int, message: str = ""):
        self.chunk_index = chunk_index
        self.total_chunks = total_chunks
        super().__init__(f"청크 {chunk_index}/{total_chunks} 처리 오류: {message}")

class HwpTimeoutError(HwpError):
    """작업 시간 초과"""
    def __init__(self, operation, timeout):
        self.operation = operation
        self.timeout = timeout
        super().__init__(f"작업 시간이 초과되었습니다: {operation} ({timeout}초)")

# 예외 처리 헬퍼 함수
def handle_hwp_error(func):
    """HWP 관련 예외를 처리하는 데코레이터"""
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except HwpError:
            # HWP 전용 예외는 그대로 발생
            raise
        except FileNotFoundError as e:
            # 파일 관련 예외를 HWP 예외로 변환
            raise HwpDocumentNotFoundError(str(e))
        except PermissionError as e:
            # 권한 관련 예외를 HWP 예외로 변환
            raise HwpDocumentAccessError(str(e))
        except Exception as e:
            # 기타 예외는 일반 HWP 오류로 변환
            raise HwpError(f"예상치 못한 오류: {str(e)}")
    return wrapper