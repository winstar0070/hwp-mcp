"""
HWP-MCP 통합 테스트
여러 모듈의 상호작용을 테스트합니다.
"""

import pytest
from unittest.mock import Mock, MagicMock, patch
import sys
import os
import json

# 테스트를 위한 경로 설정
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from tools.hwp_controller import HwpController
from tools.hwp_table_tools import HwpTableTools
from tools.hwp_document_features import HwpDocumentFeatures
from tools.hwp_advanced_features import HwpAdvancedFeatures
from tools.hwp_exceptions import *
from tools.constants import *
from tools.config import HwpConfig, ConfigManager


class TestIntegration:
    """통합 테스트 클래스"""
    
    @pytest.fixture
    def mock_com_object(self):
        """Mock COM 객체 생성"""
        mock_hwp = Mock()
        mock_hwp.Version = "2022"
        mock_hwp.XHwpWindows = Mock()
        mock_hwp.XHwpWindows.Item = Mock(return_value=Mock(Visible=True))
        mock_hwp.Run = Mock(return_value=True)
        mock_hwp.Open = Mock(return_value=True)
        mock_hwp.SaveAs = Mock(return_value=True)
        mock_hwp.Save = Mock(return_value=True)
        mock_hwp.HAction = Mock()
        mock_hwp.HParameterSet = Mock()
        return mock_hwp
    
    @pytest.fixture
    def hwp_controller(self, mock_com_object):
        """테스트용 HwpController 생성"""
        with patch('win32com.client.Dispatch', return_value=mock_com_object):
            controller = HwpController()
            controller.connect()
            return controller
    
    # ============== 문서 생성 및 편집 통합 테스트 ==============
    
    def test_create_document_with_advanced_features(self, hwp_controller):
        """고급 기능을 사용한 문서 생성 통합 테스트"""
        # Given
        doc_features = hwp_controller.get_document_features()
        table_tools = HwpTableTools(hwp_controller)
        
        # Mock 설정
        hwp_controller.hwp.HAction.GetDefault = Mock()
        hwp_controller.hwp.HAction.Execute = Mock(return_value=True)
        hwp_controller.hwp.HParameterSet = Mock()
        
        # When
        # 1. 새 문서 생성
        assert hwp_controller.create_new_document() is True
        
        # 2. 제목 삽입 (서식 포함)
        assert hwp_controller.insert_text_with_font(
            "통합 테스트 문서",
            font_name="맑은 고딕",
            font_size=20,
            bold=True
        ) is True
        
        # 3. 단락 추가
        assert hwp_controller.insert_paragraph() is True
        
        # 4. 하이퍼링크 삽입
        mock_hyperlink = Mock()
        hwp_controller.hwp.HParameterSet.HHyperLink = mock_hyperlink
        assert doc_features.insert_hyperlink(
            "테스트 링크",
            "https://example.com",
            "예제 사이트"
        ) is True
        
        # 5. 표 삽입 및 데이터 입력
        assert table_tools.insert_table(3, 3) == "Table inserted with 3 rows and 3 columns"
        
        # 6. 북마크 삽입
        mock_bookmark = Mock()
        hwp_controller.hwp.HParameterSet.HBookmark = mock_bookmark
        assert doc_features.insert_bookmark("section1") is True
        
        # 7. 문서 저장
        assert hwp_controller.save_document("test_document.hwp") is True
    
    def test_batch_operations_integration(self, hwp_controller):
        """배치 작업 통합 테스트"""
        # Given
        operations = [
            {"operation": "create_new_document"},
            {"operation": "insert_text", "params": {"text": "배치 테스트"}},
            {"operation": "insert_paragraph"},
            {"operation": "insert_table", "params": {"rows": 2, "cols": 2}},
            {"operation": "save_document", "params": {"file_path": "batch_test.hwp"}}
        ]
        
        # When
        results = []
        for op in operations:
            operation = op["operation"]
            params = op.get("params", {})
            
            if hasattr(hwp_controller, operation):
                method = getattr(hwp_controller, operation)
                result = method(**params) if params else method()
                results.append(result)
        
        # Then
        assert all(results)  # 모든 작업이 성공해야 함
    
    # ============== 예외 처리 통합 테스트 ==============
    
    def test_exception_handling_integration(self, hwp_controller):
        """예외 처리 통합 테스트"""
        # Given
        hwp_controller.is_hwp_running = False
        
        # When & Then
        # HWP 연결되지 않은 상태에서 작업 시도
        with pytest.raises(HwpNotRunningError):
            hwp_controller.insert_table(5, 5)
        
        # 잘못된 매개변수로 작업 시도
        hwp_controller.is_hwp_running = True
        with pytest.raises(HwpInvalidParameterError):
            hwp_controller.insert_table(-1, 5)
        
        # 파일을 찾을 수 없는 경우
        with pytest.raises(HwpDocumentNotFoundError):
            hwp_controller.open_document("non_existent_file.hwp")
    
    # ============== 설정 관리 통합 테스트 ==============
    
    def test_config_integration(self):
        """설정 관리 통합 테스트"""
        # Given
        config = HwpConfig()
        
        # When
        # 설정 변경
        config.default_font = "바탕"
        config.default_font_size = 12
        config.table_max_rows = 50
        
        # Then
        assert config.default_font == "바탕"
        assert config.default_font_size == 12
        assert config.table_max_rows == 50
        
        # 설정을 딕셔너리로 변환
        config_dict = config.to_dict()
        assert config_dict["default_font"] == "바탕"
        
        # 딕셔너리에서 설정 복원
        new_config = HwpConfig.from_dict(config_dict)
        assert new_config.default_font == "바탕"
    
    # ============== 표 작업 통합 테스트 ==============
    
    def test_table_operations_integration(self, hwp_controller):
        """표 작업 통합 테스트"""
        # Given
        table_tools = HwpTableTools(hwp_controller)
        hwp_controller.hwp.PositionToFieldEx = Mock(return_value=False)  # 표 밖에 있음
        
        # When
        # 1. 표 생성
        result1 = table_tools.create_table_with_data(
            rows=4,
            cols=3,
            data='[["이름", "나이", "점수"], ["김철수", "25", "90"], ["이영희", "23", "85"], ["박민수", "27", "88"]]',
            has_header=True
        )
        
        # 2. 셀 텍스트 가져오기
        hwp_controller.get_table_cell_text = Mock(return_value="김철수")
        result2 = table_tools.get_cell_text(2, 1)
        
        # 3. 셀 병합
        hwp_controller.merge_table_cells = Mock(return_value=True)
        result3 = table_tools.merge_cells(1, 1, 1, 3)
        
        # Then
        assert "생성" in result1
        assert result2 == "김철수"
        assert "완료" in result3
    
    # ============== 문서 기능 통합 테스트 ==============
    
    def test_document_features_integration(self, hwp_controller):
        """문서 기능 통합 테스트"""
        # Given
        doc_features = hwp_controller.get_document_features()
        
        # Mock 설정
        hwp_controller.hwp.HAction.GetDefault = Mock()
        hwp_controller.hwp.HAction.Execute = Mock(side_effect=[True, True, False])  # 2번 찾고 종료
        hwp_controller.hwp.HParameterSet = Mock()
        
        # 필요한 Mock 객체들
        mock_find_replace = Mock()
        mock_find_replace.HSet = Mock()
        hwp_controller.hwp.HParameterSet.HFindReplace = mock_find_replace
        
        mock_char_shape = Mock()
        mock_char_shape.HSet = Mock()
        hwp_controller.hwp.HParameterSet.HCharShape = mock_char_shape
        
        # When
        # 1. 텍스트 삽입
        hwp_controller.insert_text("이것은 중요한 내용입니다. 중요한 점을 기억하세요.")
        
        # 2. 검색 및 하이라이트
        count = doc_features.search_and_highlight("중요한", "yellow")
        
        # 3. 각주 삽입
        result = doc_features.insert_footnote("참고", "이것은 각주입니다")
        
        # 4. 필드 삽입
        field_result = doc_features.insert_field("date")
        
        # Then
        assert count == 2  # "중요한" 2번 발견
        assert result is True
        assert field_result is True
    
    # ============== 고급 기능 통합 테스트 ==============
    
    def test_advanced_features_integration(self, hwp_controller):
        """고급 기능 통합 테스트"""
        # Given
        advanced = hwp_controller.get_advanced_features()
        
        # Mock 설정
        hwp_controller.hwp.HAction.GetDefault = Mock()
        hwp_controller.hwp.HAction.Execute = Mock(return_value=True)
        hwp_controller.hwp.HParameterSet = Mock()
        
        # 페이지 설정 Mock
        mock_page_def = Mock()
        mock_page_def.HSet = Mock()
        hwp_controller.hwp.HParameterSet.HSecDef = mock_page_def
        
        # When
        # 1. 페이지 설정
        page_result = advanced.set_page(
            paper_size="A4",
            orientation="portrait",
            top_margin=20,
            bottom_margin=20
        )
        
        # 2. 머리말/꼬리말 설정
        header_result = advanced.set_header_footer(
            header_text="문서 제목",
            footer_text="페이지 ",
            show_page_number=True
        )
        
        # 3. 문단 서식 설정
        para_result = advanced.set_paragraph(
            alignment="center",
            line_spacing=1.5
        )
        
        # Then
        assert page_result is True
        assert header_result is True
        assert para_result is True
    
    # ============== 전체 문서 생성 시나리오 테스트 ==============
    
    def test_complete_document_scenario(self, hwp_controller):
        """완전한 문서 생성 시나리오 테스트"""
        # Given
        doc_features = hwp_controller.get_document_features()
        table_tools = HwpTableTools(hwp_controller)
        advanced = hwp_controller.get_advanced_features()
        
        # Mock 설정
        hwp_controller.hwp.HAction.GetDefault = Mock()
        hwp_controller.hwp.HAction.Execute = Mock(return_value=True)
        hwp_controller.hwp.HParameterSet = Mock()
        hwp_controller.hwp.PositionToFieldEx = Mock(return_value=False)
        
        # When - 보고서 스타일 문서 생성
        # 1. 문서 생성 및 페이지 설정
        assert hwp_controller.create_new_document() is True
        
        # 2. 제목 페이지
        assert hwp_controller.insert_text_with_font(
            "2024년 연간 보고서",
            font_size=24,
            bold=True
        ) is True
        
        # 3. 목차 북마크
        mock_bookmark = Mock()
        hwp_controller.hwp.HParameterSet.HBookmark = mock_bookmark
        assert doc_features.insert_bookmark("toc") is True
        
        # 4. 본문 시작
        assert hwp_controller.insert_paragraph() is True
        assert hwp_controller.insert_text("1. 개요") is True
        assert hwp_controller.insert_paragraph() is True
        
        # 5. 표 삽입
        assert table_tools.insert_table(5, 4) == "Table inserted with 5 rows and 4 columns"
        
        # 6. 하이퍼링크
        mock_hyperlink = Mock()
        hwp_controller.hwp.HParameterSet.HHyperLink = mock_hyperlink
        assert doc_features.insert_hyperlink(
            "자세한 내용",
            "https://example.com/details",
            "상세 정보 보기"
        ) is True
        
        # 7. 문서 저장
        assert hwp_controller.save_document("annual_report_2024.hwp") is True
        
        # Then - 모든 작업이 성공적으로 완료되어야 함
        # 실제 테스트에서는 생성된 문서의 내용을 검증할 수 있음