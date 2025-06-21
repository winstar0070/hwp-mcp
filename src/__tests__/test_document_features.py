"""
HwpDocumentFeatures 클래스의 테스트 코드
문서 편집 고급 기능들을 테스트합니다.
"""

import pytest
from unittest.mock import Mock, MagicMock, patch
import sys
import os

# 테스트를 위한 경로 설정
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tools.hwp_document_features import HwpDocumentFeatures
from tools.hwp_exceptions import HwpError


class TestHwpDocumentFeatures:
    """HwpDocumentFeatures 클래스 테스트"""
    
    @pytest.fixture
    def mock_hwp_controller(self):
        """Mock HwpController 객체 생성"""
        controller = Mock()
        controller.hwp = Mock()
        controller.insert_text = Mock(return_value=True)
        return controller
    
    @pytest.fixture
    def document_features(self, mock_hwp_controller):
        """테스트용 HwpDocumentFeatures 인스턴스 생성"""
        return HwpDocumentFeatures(mock_hwp_controller)
    
    # ============== 각주/미주 테스트 ==============
    
    def test_insert_footnote_success(self, document_features, mock_hwp_controller):
        """각주 삽입 성공 테스트"""
        # Given
        text = "본문 텍스트"
        note_text = "각주 내용"
        
        # When
        result = document_features.insert_footnote(text, note_text)
        
        # Then
        assert result is True
        mock_hwp_controller.insert_text.assert_called_with(text)
        document_features.hwp.Run.assert_any_call("InsertFootnote")
        document_features.hwp.Run.assert_any_call("CloseEx")
    
    def test_insert_footnote_without_text(self, document_features):
        """본문 텍스트 없이 각주 삽입 테스트"""
        # When
        result = document_features.insert_footnote("", "각주 내용")
        
        # Then
        assert result is True
        document_features.hwp.Run.assert_any_call("InsertFootnote")
    
    def test_insert_endnote_success(self, document_features, mock_hwp_controller):
        """미주 삽입 성공 테스트"""
        # Given
        text = "본문 텍스트"
        note_text = "미주 내용"
        
        # When
        result = document_features.insert_endnote(text, note_text)
        
        # Then
        assert result is True
        mock_hwp_controller.insert_text.assert_called_with(text)
        document_features.hwp.Run.assert_any_call("InsertEndnote")
        document_features.hwp.Run.assert_any_call("CloseEx")
    
    # ============== 하이퍼링크 테스트 ==============
    
    def test_insert_hyperlink_success(self, document_features):
        """하이퍼링크 삽입 성공 테스트"""
        # Given
        text = "링크 텍스트"
        url = "https://example.com"
        tooltip = "예제 사이트"
        
        # Mock HAction
        document_features.hwp.HAction = Mock()
        document_features.hwp.HParameterSet = Mock()
        document_features.hwp.HParameterSet.HHyperLink = Mock()
        document_features.hwp.HParameterSet.HHyperLink.HSet = Mock()
        
        # When
        result = document_features.insert_hyperlink(text, url, tooltip)
        
        # Then
        assert result is True
        assert document_features.hwp.HParameterSet.HHyperLink.Text == text
        assert document_features.hwp.HParameterSet.HHyperLink.Href == url
        assert document_features.hwp.HParameterSet.HHyperLink.ToolTip == tooltip
    
    def test_insert_hyperlink_without_tooltip(self, document_features):
        """도구 설명 없이 하이퍼링크 삽입 테스트"""
        # Given
        text = "링크"
        url = "https://example.com"
        
        # Mock
        document_features.hwp.HAction = Mock()
        document_features.hwp.HParameterSet = Mock()
        document_features.hwp.HParameterSet.HHyperLink = Mock()
        
        # When
        result = document_features.insert_hyperlink(text, url)
        
        # Then
        assert result is True
    
    # ============== 북마크 테스트 ==============
    
    def test_insert_bookmark_success(self, document_features):
        """북마크 삽입 성공 테스트"""
        # Given
        bookmark_name = "chapter1"
        
        # Mock
        document_features.hwp.HAction = Mock()
        document_features.hwp.HParameterSet = Mock()
        document_features.hwp.HParameterSet.HBookmark = Mock()
        document_features.hwp.HParameterSet.HBookmark.HSet = Mock()
        
        # When
        result = document_features.insert_bookmark(bookmark_name)
        
        # Then
        assert result is True
        assert document_features.hwp.HParameterSet.HBookmark.Name == bookmark_name
    
    def test_goto_bookmark_success(self, document_features):
        """북마크로 이동 성공 테스트"""
        # Given
        bookmark_name = "chapter1"
        
        # Mock
        document_features.hwp.HAction = Mock()
        document_features.hwp.HParameterSet = Mock()
        document_features.hwp.HParameterSet.HGotoBookmark = Mock()
        document_features.hwp.HParameterSet.HGotoBookmark.HSet = Mock()
        
        # When
        result = document_features.goto_bookmark(bookmark_name)
        
        # Then
        assert result is True
        assert document_features.hwp.HParameterSet.HGotoBookmark.Name == bookmark_name
    
    # ============== 주석 테스트 ==============
    
    def test_insert_comment_success(self, document_features, mock_hwp_controller):
        """주석 삽입 성공 테스트"""
        # Given
        comment_text = "검토 필요"
        author = "검토자"
        
        # When
        result = document_features.insert_comment(comment_text, author)
        
        # Then
        assert result is True
        document_features.hwp.Run.assert_any_call("InsertFieldMemo")
        mock_hwp_controller.insert_text.assert_called_with(f"[{author}] {comment_text}")
    
    def test_insert_comment_default_author(self, document_features, mock_hwp_controller):
        """기본 작성자로 주석 삽입 테스트"""
        # Given
        comment_text = "메모"
        
        # When
        result = document_features.insert_comment(comment_text)
        
        # Then
        assert result is True
        mock_hwp_controller.insert_text.assert_called_with("[사용자] 메모")
    
    # ============== 검색 및 하이라이트 테스트 ==============
    
    def test_search_and_highlight_success(self, document_features):
        """검색 및 하이라이트 성공 테스트"""
        # Given
        search_text = "중요"
        
        # Mock
        document_features.hwp.HAction = Mock()
        document_features.hwp.HParameterSet = Mock()
        document_features.hwp.HParameterSet.HFindReplace = Mock()
        document_features.hwp.HParameterSet.HFindReplace.HSet = Mock()
        document_features.hwp.HParameterSet.HCharShape = Mock()
        document_features.hwp.HParameterSet.HCharShape.HSet = Mock()
        
        # 찾기 결과 시뮬레이션 (3번 찾음)
        document_features.hwp.HAction.Execute = Mock(side_effect=[True, True, True, False])
        
        # When
        count = document_features.search_and_highlight(search_text, "yellow")
        
        # Then
        assert count == 3
        document_features.hwp.Run.assert_called_with("MoveDocBegin")
    
    def test_search_and_highlight_with_options(self, document_features):
        """옵션을 사용한 검색 및 하이라이트 테스트"""
        # Given
        search_text = "Test"
        
        # Mock
        document_features.hwp.HAction = Mock()
        document_features.hwp.HParameterSet = Mock()
        document_features.hwp.HParameterSet.HFindReplace = Mock()
        document_features.hwp.HParameterSet.HFindReplace.HSet = Mock()
        document_features.hwp.HParameterSet.HCharShape = Mock()
        
        # 찾기 실패
        document_features.hwp.HAction.Execute = Mock(return_value=False)
        
        # When
        count = document_features.search_and_highlight(
            search_text, 
            "green",
            case_sensitive=True,
            whole_word=True
        )
        
        # Then
        assert count == 0
        assert document_features.hwp.HParameterSet.HFindReplace.IgnoreCase == 0
        assert document_features.hwp.HParameterSet.HFindReplace.WholeWordOnly == 1
    
    # ============== 워터마크 테스트 ==============
    
    def test_insert_watermark_success(self, document_features, mock_hwp_controller):
        """워터마크 삽입 성공 테스트"""
        # Given
        text = "기밀"
        
        # When
        result = document_features.insert_watermark(text, font_size=100, color="red")
        
        # Then
        assert result is True
        document_features.hwp.Run.assert_any_call("HeaderFooter")
        document_features.hwp.Run.assert_any_call("DrawObjCreTextBox")
        document_features.hwp.Run.assert_any_call("CloseEx")
    
    # ============== 필드 코드 테스트 ==============
    
    def test_insert_field_date(self, document_features):
        """날짜 필드 삽입 테스트"""
        # When
        result = document_features.insert_field("date")
        
        # Then
        assert result is True
        document_features.hwp.Run.assert_called_with("InsertFieldDate")
    
    def test_insert_field_page_number(self, document_features):
        """페이지 번호 필드 삽입 테스트"""
        # When
        result = document_features.insert_field("page")
        
        # Then
        assert result is True
        document_features.hwp.Run.assert_called_with("InsertPageNumber")
    
    def test_insert_field_invalid_type(self, document_features):
        """잘못된 필드 타입 테스트"""
        # When
        result = document_features.insert_field("invalid_type")
        
        # Then
        assert result is False
    
    # ============== 문서 보안 테스트 ==============
    
    def test_set_document_password_both(self, document_features):
        """읽기/쓰기 암호 설정 테스트"""
        # Given
        read_pwd = "read123"
        write_pwd = "write456"
        
        # Mock
        document_features.hwp.HAction = Mock()
        document_features.hwp.HParameterSet = Mock()
        document_features.hwp.HParameterSet.HFileSecurity = Mock()
        document_features.hwp.HParameterSet.HFileSecurity.HSet = Mock()
        
        # When
        result = document_features.set_document_password(read_pwd, write_pwd)
        
        # Then
        assert result is True
        assert document_features.hwp.HParameterSet.HFileSecurity.ReadPassword == read_pwd
        assert document_features.hwp.HParameterSet.HFileSecurity.WritePassword == write_pwd
    
    def test_set_document_password_read_only(self, document_features):
        """읽기 암호만 설정 테스트"""
        # Given
        read_pwd = "read123"
        
        # Mock
        document_features.hwp.HAction = Mock()
        document_features.hwp.HParameterSet = Mock()
        document_features.hwp.HParameterSet.HFileSecurity = Mock()
        
        # When
        result = document_features.set_document_password(read_password=read_pwd)
        
        # Then
        assert result is True
        assert document_features.hwp.HParameterSet.HFileSecurity.ReadPassword == read_pwd
    
    # ============== 예외 처리 테스트 ==============
    
    def test_insert_footnote_exception(self, document_features):
        """각주 삽입 중 예외 발생 테스트"""
        # Given
        document_features.hwp.Run.side_effect = Exception("COM Error")
        
        # When
        result = document_features.insert_footnote("text", "note")
        
        # Then
        assert result is False
    
    def test_search_and_highlight_exception(self, document_features):
        """검색 중 예외 발생 테스트"""
        # Given
        document_features.hwp.Run.side_effect = Exception("COM Error")
        
        # When
        count = document_features.search_and_highlight("text")
        
        # Then
        assert count == 0