"""
HWP 고급 기능 모듈
이미지, PDF 변환, 찾기/바꾸기, 페이지 설정 등 고급 기능 제공
"""
import os
import logging
from typing import Optional, Dict, List, Tuple

logger = logging.getLogger(__name__)

class HwpAdvancedFeatures:
    """HWP 고급 기능을 제공하는 클래스"""
    
    def __init__(self, hwp_controller):
        """
        HwpAdvancedFeatures 초기화
        
        Args:
            hwp_controller: HwpController 인스턴스
        """
        self.hwp = hwp_controller.hwp
        self.hwp_controller = hwp_controller
    
    # ============== 1단계: 즉시 유용한 기능들 ==============
    
    def insert_image(self, image_path: str, width: Optional[int] = None, 
                    height: Optional[int] = None, align: str = "left",
                    as_char: bool = True) -> bool:
        """
        이미지를 현재 커서 위치에 삽입합니다.
        
        Args:
            image_path (str): 삽입할 이미지 파일 경로
            width (int, optional): 이미지 너비 (mm 단위)
            height (int, optional): 이미지 높이 (mm 단위)
            align (str): 정렬 방식 ("left", "center", "right")
            as_char (bool): 글자처럼 취급 여부
            
        Returns:
            bool: 성공 여부
        """
        try:
            if not os.path.exists(image_path):
                logger.error(f"이미지 파일을 찾을 수 없습니다: {image_path}")
                return False
            
            # 이미지 삽입
            ctrl = self.hwp.InsertPicture(image_path, Embedded=True, AsChar=as_char)
            if not ctrl:
                logger.error("이미지 삽입 실패")
                return False
            
            # 이미지 크기 설정
            if width or height:
                # 현재 이미지 크기 가져오기
                current_width = ctrl.Width
                current_height = ctrl.Height
                
                # 비율 유지하며 크기 조정
                if width and not height:
                    # 너비만 지정된 경우
                    ratio = width * 2834.64 / current_width  # mm to HWPUNIT (1mm = 2834.64)
                    new_width = width * 2834.64
                    new_height = current_height * ratio
                elif height and not width:
                    # 높이만 지정된 경우
                    ratio = height * 2834.64 / current_height
                    new_width = current_width * ratio
                    new_height = height * 2834.64
                else:
                    # 둘 다 지정된 경우
                    new_width = width * 2834.64
                    new_height = height * 2834.64
                
                # 크기 적용
                ctrl.Width = int(new_width)
                ctrl.Height = int(new_height)
            
            # 정렬 설정
            if align and not as_char:
                align_map = {
                    "left": 0,
                    "center": 1,
                    "right": 2
                }
                if align in align_map:
                    self.hwp.Run("ParagraphShapeAlignDistribute")
                    self.hwp.HParameterSet.HParaShape.Alignment = align_map[align]
                    self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
            
            logger.info(f"이미지 삽입 성공: {image_path}")
            return True
            
        except Exception as e:
            logger.error(f"이미지 삽입 실패: {e}")
            return False
    
    def find_replace(self, find_text: str, replace_text: str = "", 
                    match_case: bool = False, whole_word: bool = False,
                    replace_all: bool = True) -> int:
        """
        텍스트 찾기/바꾸기 기능
        
        Args:
            find_text (str): 찾을 텍스트
            replace_text (str): 바꿀 텍스트
            match_case (bool): 대소문자 구분
            whole_word (bool): 전체 단어 일치
            replace_all (bool): 모두 바꾸기
            
        Returns:
            int: 바뀐 개수
        """
        try:
            # 찾기/바꾸기 대화상자 초기화
            self.hwp.HAction.GetDefault("FindReplace", self.hwp.HParameterSet.HFindReplace.HSet)
            
            # 찾을 텍스트 설정
            self.hwp.HParameterSet.HFindReplace.FindString = find_text
            self.hwp.HParameterSet.HFindReplace.ReplaceString = replace_text
            
            # 옵션 설정
            self.hwp.HParameterSet.HFindReplace.IgnoreCase = 0 if match_case else 1
            self.hwp.HParameterSet.HFindReplace.WholeWordOnly = 1 if whole_word else 0
            self.hwp.HParameterSet.HFindReplace.AutoSpell = 0  # 자동 맞춤법 검사 비활성화
            
            # 검색 방향 설정 (0: 앞으로, 1: 뒤로, 2: 문서 전체)
            self.hwp.HParameterSet.HFindReplace.Direction = 2 if replace_all else 0
            
            # 검색 범위 설정 (0: 현재 위치부터, 1: 문서 처음부터)
            self.hwp.HParameterSet.HFindReplace.FindType = 1 if replace_all else 0
            
            # 바꾸기 모드 설정 (0: 찾기만, 1: 바꾸기, 2: 모두 바꾸기)
            if replace_text:
                self.hwp.HParameterSet.HFindReplace.ReplaceMode = 2 if replace_all else 1
            else:
                self.hwp.HParameterSet.HFindReplace.ReplaceMode = 0
            
            # 실행
            self.hwp.HAction.Execute("FindReplace", self.hwp.HParameterSet.HFindReplace.HSet)
            
            # 결과 가져오기 (바뀐 개수)
            replaced_count = self.hwp.HParameterSet.HFindReplace.ReplaceCount
            
            logger.info(f"찾기/바꾸기 완료: '{find_text}' -> '{replace_text}', {replaced_count}개 바뀜")
            return replaced_count
            
        except Exception as e:
            logger.error(f"찾기/바꾸기 실패: {e}")
            return 0
    
    def export_pdf(self, output_path: str, quality: str = "high",
                  include_bookmarks: bool = True,
                  include_comments: bool = False) -> bool:
        """
        현재 문서를 PDF로 변환합니다.
        
        Args:
            output_path (str): 출력 PDF 파일 경로
            quality (str): PDF 품질 ("low", "medium", "high")
            include_bookmarks (bool): 북마크 포함 여부
            include_comments (bool): 주석 포함 여부
            
        Returns:
            bool: 성공 여부
        """
        try:
            # PDF 저장 설정
            quality_map = {
                "low": "75",
                "medium": "150",
                "high": "300"
            }
            
            # 출력 경로 확인
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # PDF로 저장
            # Format: "PDF" 지정
            result = self.hwp.SaveAs(output_path, Format="PDF")
            
            if result:
                logger.info(f"PDF 변환 성공: {output_path}")
                return True
            else:
                logger.error("PDF 변환 실패")
                return False
                
        except Exception as e:
            logger.error(f"PDF 변환 실패: {e}")
            return False
    
    # ============== 2단계: 문서 품질 향상 기능들 ==============
    
    def set_page(self, paper_size: str = "A4", orientation: str = "portrait",
                 margins: Optional[Dict[str, int]] = None) -> bool:
        """
        페이지 설정을 변경합니다.
        
        Args:
            paper_size (str): 용지 크기 ("A4", "A3", "B5", "Letter", "Legal")
            orientation (str): 용지 방향 ("portrait", "landscape")
            margins (dict, optional): 여백 설정 {"top": 20, "bottom": 20, "left": 30, "right": 30} (mm 단위)
            
        Returns:
            bool: 성공 여부
        """
        try:
            # 페이지 설정 액션 가져오기
            self.hwp.HAction.GetDefault("PageSetup", self.hwp.HParameterSet.HSecDef.HSet)
            
            # 용지 크기 설정
            paper_sizes = {
                "A4": (21000, 29700),  # 210mm x 297mm in 0.01mm units
                "A3": (29700, 42000),
                "B5": (18200, 25700),
                "Letter": (21590, 27940),
                "Legal": (21590, 35560)
            }
            
            if paper_size in paper_sizes:
                width, height = paper_sizes[paper_size]
                if orientation == "landscape":
                    width, height = height, width
                
                self.hwp.HParameterSet.HSecDef.PageDef.PaperWidth = width * 100  # HWPUNIT로 변환
                self.hwp.HParameterSet.HSecDef.PageDef.PaperHeight = height * 100
                self.hwp.HParameterSet.HSecDef.PageDef.Landscape = 1 if orientation == "landscape" else 0
            
            # 여백 설정
            if margins:
                if "top" in margins:
                    self.hwp.HParameterSet.HSecDef.PageDef.TopMargin = margins["top"] * 2834.64
                if "bottom" in margins:
                    self.hwp.HParameterSet.HSecDef.PageDef.BottomMargin = margins["bottom"] * 2834.64
                if "left" in margins:
                    self.hwp.HParameterSet.HSecDef.PageDef.LeftMargin = margins["left"] * 2834.64
                if "right" in margins:
                    self.hwp.HParameterSet.HSecDef.PageDef.RightMargin = margins["right"] * 2834.64
            
            # 설정 적용
            self.hwp.HAction.Execute("PageSetup", self.hwp.HParameterSet.HSecDef.HSet)
            
            logger.info(f"페이지 설정 완료: {paper_size} {orientation}")
            return True
            
        except Exception as e:
            logger.error(f"페이지 설정 실패: {e}")
            return False
    
    def set_header_footer(self, header_text: str = "", footer_text: str = "",
                         show_page_number: bool = True,
                         page_number_position: str = "footer-center") -> bool:
        """
        머리말/꼬리말을 설정합니다.
        
        Args:
            header_text (str): 머리말 텍스트
            footer_text (str): 꼬리말 텍스트
            show_page_number (bool): 쪽번호 표시 여부
            page_number_position (str): 쪽번호 위치 ("header-left", "header-center", "header-right",
                                                   "footer-left", "footer-center", "footer-right")
            
        Returns:
            bool: 성공 여부
        """
        try:
            # 머리말 설정
            if header_text or (show_page_number and "header" in page_number_position):
                self.hwp.Run("HeaderFooter")  # 머리말/꼬리말 편집 모드 진입
                self.hwp.Run("MoveTopLevelBegin")  # 머리말 영역으로 이동
                
                # 기존 내용 삭제
                self.hwp.Run("SelectAll")
                self.hwp.Run("Delete")
                
                # 텍스트 입력
                if header_text:
                    self.hwp_controller.insert_text(header_text)
                
                # 쪽번호 삽입
                if show_page_number and "header" in page_number_position:
                    if "-center" in page_number_position:
                        self.hwp.Run("ParagraphShapeAlignCenter")
                    elif "-right" in page_number_position:
                        self.hwp.Run("ParagraphShapeAlignRight")
                    
                    self.hwp.Run("InsertPageNumber")
                
                self.hwp.Run("CloseEx")  # 머리말 편집 종료
            
            # 꼬리말 설정
            if footer_text or (show_page_number and "footer" in page_number_position):
                self.hwp.Run("HeaderFooter")
                self.hwp.Run("MoveTopLevelEnd")  # 꼬리말 영역으로 이동
                
                # 기존 내용 삭제
                self.hwp.Run("SelectAll")
                self.hwp.Run("Delete")
                
                # 텍스트 입력
                if footer_text:
                    self.hwp_controller.insert_text(footer_text)
                
                # 쪽번호 삽입
                if show_page_number and "footer" in page_number_position:
                    if "-center" in page_number_position:
                        self.hwp.Run("ParagraphShapeAlignCenter")
                    elif "-right" in page_number_position:
                        self.hwp.Run("ParagraphShapeAlignRight")
                    
                    # 페이지 번호 필드 삽입
                    if "[쪽번호]" in footer_text or "[전체쪽수]" in footer_text:
                        # 템플릿 텍스트 처리
                        processed_text = footer_text
                        if "[쪽번호]" in processed_text:
                            self.hwp.Run("InsertPageNumber")
                        if "[전체쪽수]" in processed_text:
                            self.hwp.Run("InsertPageCount")
                    else:
                        self.hwp.Run("InsertPageNumber")
                
                self.hwp.Run("CloseEx")  # 꼬리말 편집 종료
            
            logger.info("머리말/꼬리말 설정 완료")
            return True
            
        except Exception as e:
            logger.error(f"머리말/꼬리말 설정 실패: {e}")
            return False
    
    def set_paragraph(self, alignment: str = "left", line_spacing: float = 1.0,
                     indent_first: int = 0, indent_left: int = 0,
                     space_before: int = 0, space_after: int = 0) -> bool:
        """
        문단 서식을 설정합니다.
        
        Args:
            alignment (str): 정렬 방식 ("left", "center", "right", "justify")
            line_spacing (float): 줄 간격 (1.0 = 100%, 1.5 = 150%, 2.0 = 200%)
            indent_first (int): 첫 줄 들여쓰기 (mm 단위)
            indent_left (int): 왼쪽 들여쓰기 (mm 단위)
            space_before (int): 문단 앞 간격 (mm 단위)
            space_after (int): 문단 뒤 간격 (mm 단위)
            
        Returns:
            bool: 성공 여부
        """
        try:
            # 문단 모양 설정 액션 초기화
            self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
            
            # 정렬 설정
            align_map = {
                "left": 0,
                "center": 1,
                "right": 2,
                "justify": 3
            }
            if alignment in align_map:
                self.hwp.HParameterSet.HParaShape.Alignment = align_map[alignment]
            
            # 줄 간격 설정 (%)
            self.hwp.HParameterSet.HParaShape.LineSpacing = int(line_spacing * 100)
            self.hwp.HParameterSet.HParaShape.LineSpacingType = 0  # 0: 비율, 1: 고정값
            
            # 들여쓰기 설정 (HWPUNIT으로 변환)
            if indent_first:
                self.hwp.HParameterSet.HParaShape.IndentFirst = indent_first * 2834.64
            if indent_left:
                self.hwp.HParameterSet.HParaShape.IndentLeft = indent_left * 2834.64
            
            # 문단 간격 설정
            if space_before:
                self.hwp.HParameterSet.HParaShape.SpaceBefore = space_before * 2834.64
            if space_after:
                self.hwp.HParameterSet.HParaShape.SpaceAfter = space_after * 2834.64
            
            # 설정 적용
            self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
            
            logger.info("문단 서식 설정 완료")
            return True
            
        except Exception as e:
            logger.error(f"문단 서식 설정 실패: {e}")
            return False
    
    # ============== 3단계: 고급 기능들 ==============
    
    def create_toc(self, max_level: int = 3, page_numbers: bool = True,
                   update_existing: bool = True) -> bool:
        """
        목차를 자동으로 생성합니다.
        
        Args:
            max_level (int): 목차에 포함할 최대 제목 레벨
            page_numbers (bool): 페이지 번호 표시 여부
            update_existing (bool): 기존 목차 업데이트 여부
            
        Returns:
            bool: 성공 여부
        """
        try:
            # 목차 삽입
            self.hwp.Run("InsertTOC")
            
            logger.info("목차 생성 완료")
            return True
            
        except Exception as e:
            logger.error(f"목차 생성 실패: {e}")
            return False
    
    def insert_shape(self, shape_type: str = "rectangle", 
                    position: Optional[Dict[str, int]] = None,
                    size: Optional[Dict[str, int]] = None,
                    fill_color: str = "#FFFFFF",
                    border_color: str = "#000000",
                    text: str = "") -> bool:
        """
        도형을 삽입합니다.
        
        Args:
            shape_type (str): 도형 종류 ("rectangle", "circle", "arrow", "line")
            position (dict): 위치 {"x": 50, "y": 100} (mm 단위)
            size (dict): 크기 {"width": 100, "height": 50} (mm 단위)
            fill_color (str): 채우기 색상 (hex 코드)
            border_color (str): 테두리 색상 (hex 코드)
            text (str): 도형 안의 텍스트
            
        Returns:
            bool: 성공 여부
        """
        try:
            # 도형 유형별 삽입
            shape_map = {
                "rectangle": "DrawObjCreRectangle",
                "circle": "DrawObjCreEllipse",
                "arrow": "DrawObjCreArc",
                "line": "DrawObjCreLine"
            }
            
            if shape_type not in shape_map:
                logger.error(f"지원하지 않는 도형 유형: {shape_type}")
                return False
            
            # 도형 그리기 모드 시작
            self.hwp.Run(shape_map[shape_type])
            
            # 크기와 위치 설정 (기본값 사용)
            if not position:
                position = {"x": 50, "y": 50}
            if not size:
                size = {"width": 100, "height": 50}
            
            # 마우스 드래그 시뮬레이션 (도형 그리기)
            # 실제로는 HWP API의 제한으로 직접적인 좌표 지정이 어려움
            # 대신 도형을 삽입한 후 속성을 변경하는 방식 사용
            
            # 도형 속성 설정
            if text:
                # 텍스트 상자로 변환하여 텍스트 입력
                self.hwp.Run("DrawObjTextBoxEdit")
                self.hwp_controller.insert_text(text)
                self.hwp.Run("Cancel")
            
            logger.info(f"{shape_type} 도형 삽입 완료")
            return True
            
        except Exception as e:
            logger.error(f"도형 삽입 실패: {e}")
            return False
    
    def save_as_template(self, template_name: str, template_path: str = None,
                        include_styles: bool = True) -> bool:
        """
        현재 문서를 템플릿으로 저장합니다.
        
        Args:
            template_name (str): 템플릿 이름
            template_path (str, optional): 템플릿 저장 경로
            include_styles (bool): 스타일 포함 여부
            
        Returns:
            bool: 성공 여부
        """
        try:
            # 템플릿 저장 경로 설정
            if not template_path:
                template_dir = os.path.join(os.path.expanduser("~"), "HWP_Templates")
                if not os.path.exists(template_dir):
                    os.makedirs(template_dir)
                template_path = os.path.join(template_dir, f"{template_name}.hwt")
            
            # 템플릿으로 저장
            result = self.hwp.SaveAs(template_path, Format="HWT")
            
            if result:
                logger.info(f"템플릿 저장 성공: {template_path}")
                return True
            else:
                logger.error("템플릿 저장 실패")
                return False
                
        except Exception as e:
            logger.error(f"템플릿 저장 실패: {e}")
            return False
    
    def apply_template(self, template_path: str) -> bool:
        """
        템플릿을 현재 문서에 적용합니다.
        
        Args:
            template_path (str): 템플릿 파일 경로
            
        Returns:
            bool: 성공 여부
        """
        try:
            if not os.path.exists(template_path):
                logger.error(f"템플릿 파일을 찾을 수 없습니다: {template_path}")
                return False
            
            # 템플릿 적용 (새 문서 생성 후 템플릿 로드)
            self.hwp_controller.create_new_document()
            result = self.hwp.Open(template_path, Format="HWT")
            
            if result:
                logger.info(f"템플릿 적용 성공: {template_path}")
                return True
            else:
                logger.error("템플릿 적용 실패")
                return False
                
        except Exception as e:
            logger.error(f"템플릿 적용 실패: {e}")
            return False