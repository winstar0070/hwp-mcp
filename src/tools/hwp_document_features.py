"""
HWP 문서 편집 고급 기능 모듈
각주, 하이퍼링크, 북마크, 주석 등 문서 편집 관련 고급 기능 제공
"""
import os
import logging
from typing import Optional, Dict, List, Union
from datetime import datetime

try:
    from .constants import HIGHLIGHT_COLORS, FIELD_TYPES
except ImportError:
    # 상수 기본값
    HIGHLIGHT_COLORS = {
        "yellow": 4, "green": 3, "cyan": 6, "magenta": 5,
        "red": 2, "blue": 1, "gray": 7
    }
    FIELD_TYPES = {
        "date": "InsertFieldDate", "time": "InsertFieldTime",
        "page": "InsertPageNumber", "total_pages": "InsertPageCount",
        "filename": "InsertFieldFileName", "path": "InsertFieldPath",
        "author": "InsertFieldAuthor"
    }

logger = logging.getLogger(__name__)

class HwpDocumentFeatures:
    """HWP 문서 편집 고급 기능을 제공하는 클래스"""
    
    def __init__(self, hwp_controller):
        """
        HwpDocumentFeatures 초기화
        
        Args:
            hwp_controller: HwpController 인스턴스
        """
        self.hwp = hwp_controller.hwp
        self.hwp_controller = hwp_controller
    
    # ============== 각주/미주 기능 ==============
    
    def insert_footnote(self, text: str, note_text: str) -> bool:
        """
        현재 위치에 각주를 삽입합니다.
        
        Args:
            text (str): 본문에 표시할 텍스트
            note_text (str): 각주 내용
            
        Returns:
            bool: 성공 여부
        """
        try:
            # 본문 텍스트 삽입
            if text:
                self.hwp_controller.insert_text(text)
            
            # 각주 삽입
            self.hwp.Run("InsertFootnote")
            
            # 각주 내용 입력
            self.hwp_controller.insert_text(note_text)
            
            # 각주 편집 모드 종료 (본문으로 돌아가기)
            self.hwp.Run("CloseEx")
            
            logger.info(f"각주 삽입 성공: {text[:20]}...")
            return True
            
        except AttributeError as e:
            logger.error(f"각주 삽입 API 호출 실패: {e} (HWP 연결 상태를 확인하세요)")
            return False
        except Exception as e:
            logger.error(f"각주 삽입 중 예상치 못한 오류: {e}")
            return False
    
    def insert_endnote(self, text: str, note_text: str) -> bool:
        """
        현재 위치에 미주를 삽입합니다.
        
        Args:
            text (str): 본문에 표시할 텍스트
            note_text (str): 미주 내용
            
        Returns:
            bool: 성공 여부
        """
        try:
            # 본문 텍스트 삽입
            if text:
                self.hwp_controller.insert_text(text)
            
            # 미주 삽입
            self.hwp.Run("InsertEndnote")
            
            # 미주 내용 입력
            self.hwp_controller.insert_text(note_text)
            
            # 미주 편집 모드 종료
            self.hwp.Run("CloseEx")
            
            logger.info(f"미주 삽입 성공: {text[:20]}...")
            return True
            
        except AttributeError as e:
            logger.error(f"미주 삽입 API 호출 실패: {e} (HWP 연결 상태를 확인하세요)")
            return False
        except Exception as e:
            logger.error(f"미주 삽입 중 예상치 못한 오류: {e}")
            return False
    
    # ============== 하이퍼링크 기능 ==============
    
    def insert_hyperlink(self, text: str, url: str, tooltip: str = "") -> bool:
        """
        하이퍼링크를 삽입합니다.
        
        Args:
            text (str): 링크로 표시할 텍스트
            url (str): 링크 주소 (웹 주소 또는 문서 내 위치)
            tooltip (str): 마우스 오버 시 표시할 도구 설명
            
        Returns:
            bool: 성공 여부
        """
        try:
            # 하이퍼링크 삽입 액션 초기화
            self.hwp.HAction.GetDefault("InsertHyperlink", self.hwp.HParameterSet.HHyperLink.HSet)
            
            # 링크 텍스트 설정
            self.hwp.HParameterSet.HHyperLink.Text = text
            
            # URL 설정
            self.hwp.HParameterSet.HHyperLink.Href = url
            
            # 도구 설명 설정
            if tooltip:
                self.hwp.HParameterSet.HHyperLink.ToolTip = tooltip
            
            # 하이퍼링크 삽입 실행
            self.hwp.HAction.Execute("InsertHyperlink", self.hwp.HParameterSet.HHyperLink.HSet)
            
            logger.info(f"하이퍼링크 삽입 성공: {text} -> {url}")
            return True
            
        except AttributeError as e:
            logger.error(f"하이퍼링크 API 호출 실패: {e} (HWP 연결 상태를 확인하세요)")
            return False
        except ValueError as e:
            logger.error(f"하이퍼링크 URL 형식 오류: {e} (유효한 URL을 입력하세요)")
            return False
        except Exception as e:
            logger.error(f"하이퍼링크 삽입 중 예상치 못한 오류: {e}")
            return False
    
    # ============== 북마크(책갈피) 기능 ==============
    
    def insert_bookmark(self, bookmark_name: str) -> bool:
        """
        현재 위치에 북마크를 삽입합니다.
        
        Args:
            bookmark_name (str): 북마크 이름
            
        Returns:
            bool: 성공 여부
        """
        try:
            # 북마크 삽입 액션 초기화
            self.hwp.HAction.GetDefault("InsertBookmark", self.hwp.HParameterSet.HBookmark.HSet)
            
            # 북마크 이름 설정
            self.hwp.HParameterSet.HBookmark.Name = bookmark_name
            
            # 북마크 삽입 실행
            self.hwp.HAction.Execute("InsertBookmark", self.hwp.HParameterSet.HBookmark.HSet)
            
            logger.info(f"북마크 삽입 성공: {bookmark_name}")
            return True
            
        except AttributeError as e:
            logger.error(f"북마크 API 호출 실패: {e} (HWP 연결 상태를 확인하세요)")
            return False
        except ValueError as e:
            logger.error(f"북마크 이름 오류: {e} (북마크 이름이 비어있을 수 없습니다)")
            return False
        except Exception as e:
            logger.error(f"북마크 삽입 중 예상치 못한 오류: {e}")
            return False
    
    def goto_bookmark(self, bookmark_name: str) -> bool:
        """
        특정 북마크로 이동합니다.
        
        Args:
            bookmark_name (str): 이동할 북마크 이름
            
        Returns:
            bool: 성공 여부
        """
        try:
            # 북마크로 이동
            self.hwp.HAction.GetDefault("GotoBookmark", self.hwp.HParameterSet.HGotoBookmark.HSet)
            self.hwp.HParameterSet.HGotoBookmark.Name = bookmark_name
            self.hwp.HAction.Execute("GotoBookmark", self.hwp.HParameterSet.HGotoBookmark.HSet)
            
            logger.info(f"북마크로 이동 성공: {bookmark_name}")
            return True
            
        except AttributeError as e:
            logger.error(f"북마크 이동 API 호출 실패: {e} (HWP 연결 상태를 확인하세요)")
            return False
        except KeyError as e:
            logger.error(f"북마크를 찾을 수 없습니다: {e} (북마크 이름을 확인하세요)")
            return False
        except Exception as e:
            logger.error(f"북마크로 이동 중 예상치 못한 오류: {e}")
            return False
    
    # ============== 주석(Comment) 기능 ==============
    
    def insert_comment(self, comment_text: str, author: str = "사용자") -> bool:
        """
        현재 선택된 텍스트에 주석을 추가합니다.
        
        Args:
            comment_text (str): 주석 내용
            author (str): 작성자 이름
            
        Returns:
            bool: 성공 여부
        """
        try:
            # 주석 삽입
            self.hwp.Run("InsertFieldMemo")
            
            # 주석 내용 입력
            self.hwp_controller.insert_text(f"[{author}] {comment_text}")
            
            # 주석 편집 모드 종료
            self.hwp.Run("CloseEx")
            
            logger.info(f"주석 삽입 성공: {comment_text[:20]}...")
            return True
            
        except AttributeError as e:
            logger.error(f"주석 API 호출 실패: {e} (HWP 연결 상태를 확인하세요)")
            return False
        except Exception as e:
            logger.error(f"주석 삽입 중 예상치 못한 오류: {e}")
            return False
    
    # ============== 검색 및 하이라이트 기능 ==============
    
    def search_and_highlight(self, search_text: str, highlight_color: str = "yellow",
                           case_sensitive: bool = False, whole_word: bool = False) -> int:
        """
        문서 전체에서 텍스트를 검색하고 하이라이트합니다.
        
        Args:
            search_text (str): 검색할 텍스트
            highlight_color (str): 하이라이트 색상 ("yellow", "green", "cyan", "magenta", "red")
            case_sensitive (bool): 대소문자 구분
            whole_word (bool): 전체 단어 일치
            
        Returns:
            int: 하이라이트된 개수
        """
        try:
            # 색상 맵핑
            color_value = HIGHLIGHT_COLORS.get(highlight_color.lower(), 4)  # 기본값: 노란색
            
            # 문서 처음으로 이동
            self.hwp.Run("MoveDocBegin")
            
            found_count = 0
            
            # 찾기 설정
            self.hwp.HAction.GetDefault("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)
            self.hwp.HParameterSet.HFindReplace.FindString = search_text
            self.hwp.HParameterSet.HFindReplace.IgnoreCase = 0 if case_sensitive else 1
            self.hwp.HParameterSet.HFindReplace.WholeWordOnly = 1 if whole_word else 0
            self.hwp.HParameterSet.HFindReplace.Direction = 0  # 앞으로 찾기
            
            # 모든 텍스트 찾아서 하이라이트
            while True:
                # 다음 찾기
                result = self.hwp.HAction.Execute("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)
                
                if not result:
                    break
                
                # 찾은 텍스트에 하이라이트 적용
                self.hwp.HAction.GetDefault("CharShapeBackgroundColor", 
                                          self.hwp.HParameterSet.HCharShape.HSet)
                self.hwp.HParameterSet.HCharShape.BackColor = color_value
                self.hwp.HAction.Execute("CharShapeBackgroundColor", 
                                       self.hwp.HParameterSet.HCharShape.HSet)
                
                found_count += 1
            
            logger.info(f"검색 및 하이라이트 완료: '{search_text}' {found_count}개 발견")
            return found_count
            
        except AttributeError as e:
            logger.error(f"검색 및 하이라이트 API 호출 실패: {e} (HWP 연결 상태를 확인하세요)")
            return 0
        except KeyError as e:
            logger.error(f"지원하지 않는 하이라이트 색상: {e}")
            return 0
        except Exception as e:
            logger.error(f"검색 및 하이라이트 중 예상치 못한 오류: {e}")
            return 0
    
    # ============== 워터마크 기능 ==============
    
    def insert_watermark(self, text: str, font_size: int = 72, 
                        color: str = "gray", opacity: int = 30,
                        angle: int = -45) -> bool:
        """
        문서에 워터마크를 삽입합니다.
        
        Args:
            text (str): 워터마크 텍스트
            font_size (int): 글자 크기
            color (str): 색상
            opacity (int): 투명도 (0-100)
            angle (int): 회전 각도
            
        Returns:
            bool: 성공 여부
        """
        try:
            # 머리말 영역에 워터마크 삽입
            self.hwp.Run("HeaderFooter")
            
            # 워터마크 텍스트 삽입 (텍스트 상자 활용)
            self.hwp.Run("DrawObjCreTextBox")
            
            # 텍스트 입력
            self.hwp_controller.insert_text_with_font(
                text=text,
                font_size=font_size,
                bold=True
            )
            
            # 텍스트 상자 속성 설정 (회전, 투명도 등)
            # 실제 구현은 HWP API의 제한으로 단순화
            
            # 머리말 편집 종료
            self.hwp.Run("CloseEx")
            
            logger.info(f"워터마크 삽입 성공: {text}")
            return True
            
        except AttributeError as e:
            logger.error(f"워터마크 API 호출 실패: {e} (HWP 연결 상태를 확인하세요)")
            return False
        except Exception as e:
            logger.error(f"워터마크 삽입 중 예상치 못한 오류: {e}")
            return False
    
    # ============== 필드 코드 기능 ==============
    
    def insert_field(self, field_type: str, format: str = "") -> bool:
        """
        필드 코드를 삽입합니다.
        
        Args:
            field_type (str): 필드 유형 ("date", "time", "page", "filename", "author")
            format (str): 날짜/시간 형식 (선택사항)
            
        Returns:
            bool: 성공 여부
        """
        try:
            if field_type not in FIELD_TYPES:
                logger.error(f"지원하지 않는 필드 유형: {field_type}")
                return False
            
            # 필드 삽입
            self.hwp.Run(FIELD_TYPES[field_type])
            
            logger.info(f"필드 삽입 성공: {field_type}")
            return True
            
        except AttributeError as e:
            logger.error(f"필드 API 호출 실패: {e} (HWP 연결 상태를 확인하세요)")
            return False
        except KeyError as e:
            logger.error(f"지원하지 않는 필드 유형: {e}")
            return False
        except Exception as e:
            logger.error(f"필드 삽입 중 예상치 못한 오류: {e}")
            return False
    
    # ============== 문서 보안 기능 ==============
    
    def set_document_password(self, read_password: str = "", 
                            write_password: str = "") -> bool:
        """
        문서에 암호를 설정합니다.
        
        Args:
            read_password (str): 읽기 암호
            write_password (str): 쓰기 암호
            
        Returns:
            bool: 성공 여부
        """
        try:
            # 문서 보안 설정
            self.hwp.HAction.GetDefault("FileSaveAsSecurity", 
                                      self.hwp.HParameterSet.HFileSecurity.HSet)
            
            if read_password:
                self.hwp.HParameterSet.HFileSecurity.ReadPassword = read_password
            
            if write_password:
                self.hwp.HParameterSet.HFileSecurity.WritePassword = write_password
            
            # 보안 설정 적용
            self.hwp.HAction.Execute("FileSaveAsSecurity", 
                                   self.hwp.HParameterSet.HFileSecurity.HSet)
            
            logger.info("문서 암호 설정 성공")
            return True
            
        except AttributeError as e:
            logger.error(f"문서 암호 설정 API 호출 실패: {e} (HWP 연결 상태를 확인하세요)")
            return False
        except ValueError as e:
            logger.error(f"암호 형식 오류: {e} (암호는 비어있을 수 없습니다)")
            return False
        except Exception as e:
            logger.error(f"문서 암호 설정 중 예상치 못한 오류: {e}")
            return False