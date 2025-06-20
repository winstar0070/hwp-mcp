"""
HWP 고급 기능 테스트
이미지 삽입, 찾기/바꾸기, PDF 변환, 페이지 설정 등의 고급 기능을 테스트합니다.
"""
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from src.tools.hwp_controller import HwpController
import time

def test_advanced_features():
    """고급 기능들을 테스트합니다."""
    print("=== HWP 고급 기능 테스트 시작 ===\n")
    
    # HWP 컨트롤러 초기화
    hwp = HwpController()
    if not hwp.connect():
        print("❌ HWP 연결 실패")
        return
    
    # 새 문서 생성
    if not hwp.create_new_document():
        print("❌ 새 문서 생성 실패")
        return
    
    print("✅ HWP 연결 및 새 문서 생성 성공\n")
    
    # 고급 기능 인스턴스 가져오기
    advanced = hwp.get_advanced_features()
    
    # ========== 1단계: 즉시 유용한 기능들 ==========
    print("📌 1단계: 즉시 유용한 기능들\n")
    
    # 페이지 설정 (먼저 설정하여 전체 문서에 적용)
    print("📝 페이지 설정")
    result = advanced.set_page(
        paper_size="A4",
        orientation="portrait",
        margins={"top": 20, "bottom": 20, "left": 25, "right": 25}
    )
    print(f"  결과: {'✅ 성공' if result else '❌ 실패'}\n")
    time.sleep(1)
    
    # 제목 삽입
    hwp.insert_text_with_font(
        text="HWP 고급 기능 테스트 문서",
        font_name="맑은 고딕",
        font_size=20,
        bold=True
    )
    hwp.insert_paragraph()
    hwp.insert_paragraph()
    
    # 테스트 1: 찾기/바꾸기
    print("📝 테스트 1: 찾기/바꾸기 기능")
    # 샘플 텍스트 삽입
    hwp.insert_text("이것은 테스트 문서입니다. test를 TEST로 바꿀 예정입니다. test는 여러 번 나타납니다.")
    hwp.insert_paragraph()
    
    # 찾기/바꾸기 실행
    count = advanced.find_replace("test", "TEST", match_case=False, replace_all=True)
    print(f"  'test' -> 'TEST' 변경: {count}개\n")
    time.sleep(1)
    
    # 테스트 2: 이미지 삽입
    print("📝 테스트 2: 이미지 삽입 기능")
    # 테스트용 이미지 생성 (간단한 이미지 파일 생성)
    test_image_path = os.path.join(os.getcwd(), "test_image.png")
    try:
        # PIL이 없을 수 있으므로 기본 이미지 경로 사용
        print("  주의: 실제 이미지 파일이 필요합니다.")
        print(f"  이미지 경로: {test_image_path}")
        if os.path.exists(test_image_path):
            result = advanced.insert_image(
                image_path=test_image_path,
                width=100,
                height=75,
                align="center"
            )
            print(f"  결과: {'✅ 성공' if result else '❌ 실패'}\n")
        else:
            print("  ⚠️ 테스트 이미지 파일이 없어 건너뜁니다.\n")
    except Exception as e:
        print(f"  ⚠️ 이미지 삽입 테스트 건너뜀: {e}\n")
    
    hwp.insert_paragraph()
    time.sleep(1)
    
    # ========== 2단계: 문서 품질 향상 기능들 ==========
    print("\n📌 2단계: 문서 품질 향상 기능들\n")
    
    # 테스트 3: 머리말/꼬리말 설정
    print("📝 테스트 3: 머리말/꼬리말 설정")
    result = advanced.set_header_footer(
        header_text="HWP 고급 기능 테스트",
        footer_text="페이지 ",  # 페이지 번호가 자동으로 추가됨
        show_page_number=True,
        page_number_position="footer-center"
    )
    print(f"  결과: {'✅ 성공' if result else '❌ 실패'}\n")
    time.sleep(1)
    
    # 테스트 4: 문단 서식 설정
    print("📝 테스트 4: 문단 서식 설정")
    hwp.insert_text("이 문단은 양쪽 정렬, 1.5배 줄 간격, 첫 줄 들여쓰기가 적용됩니다.")
    result = advanced.set_paragraph(
        alignment="justify",
        line_spacing=1.5,
        indent_first=10,
        space_after=5
    )
    print(f"  결과: {'✅ 성공' if result else '❌ 실패'}\n")
    hwp.insert_paragraph()
    time.sleep(1)
    
    # ========== 3단계: 고급 기능들 ==========
    print("\n📌 3단계: 고급 기능들\n")
    
    # 여러 단락 추가 (목차 생성을 위해)
    hwp.insert_text_with_font("1. 서론", font_size=16, bold=True)
    hwp.insert_paragraph()
    hwp.insert_text("이것은 서론 부분입니다.")
    hwp.insert_paragraph()
    hwp.insert_paragraph()
    
    hwp.insert_text_with_font("2. 본론", font_size=16, bold=True)
    hwp.insert_paragraph()
    hwp.insert_text("이것은 본론 부분입니다.")
    hwp.insert_paragraph()
    hwp.insert_paragraph()
    
    hwp.insert_text_with_font("3. 결론", font_size=16, bold=True)
    hwp.insert_paragraph()
    hwp.insert_text("이것은 결론 부분입니다.")
    hwp.insert_paragraph()
    hwp.insert_paragraph()
    
    # 테스트 5: 도형 삽입
    print("📝 테스트 5: 도형 삽입")
    result = advanced.insert_shape(
        shape_type="rectangle",
        text="텍스트 상자"
    )
    print(f"  사각형 도형: {'✅ 성공' if result else '❌ 실패'}\n")
    time.sleep(1)
    
    # 테스트 6: 목차 생성
    print("📝 테스트 6: 목차 자동 생성")
    # 문서 처음으로 이동
    hwp.hwp.Run("MoveDocBegin")
    hwp.insert_text_with_font("목차", font_size=18, bold=True)
    hwp.insert_paragraph()
    
    result = advanced.create_toc(max_level=3, page_numbers=True)
    print(f"  결과: {'✅ 성공' if result else '❌ 실패'}\n")
    hwp.insert_paragraph()
    time.sleep(1)
    
    # ========== PDF 변환 및 템플릿 저장 ==========
    print("\n📌 문서 저장 및 변환\n")
    
    # HWP 문서 저장
    hwp_path = os.path.join(os.getcwd(), "advanced_features_test.hwp")
    if hwp.save_document(hwp_path):
        print(f"✅ HWP 문서 저장: {hwp_path}")
    
    # 테스트 7: PDF 변환
    print("\n📝 테스트 7: PDF 변환")
    pdf_path = os.path.join(os.getcwd(), "advanced_features_test.pdf")
    result = advanced.export_pdf(
        output_path=pdf_path,
        quality="high",
        include_bookmarks=True
    )
    print(f"  결과: {'✅ 성공' if result else '❌ 실패'}")
    if result:
        print(f"  PDF 경로: {pdf_path}\n")
    
    # 테스트 8: 템플릿 저장
    print("📝 테스트 8: 템플릿 저장")
    result = advanced.save_as_template(
        template_name="고급기능_템플릿",
        include_styles=True
    )
    print(f"  결과: {'✅ 성공' if result else '❌ 실패'}\n")
    
    print("\n=== HWP 고급 기능 테스트 완료 ===")
    print("\n📌 테스트 결과:")
    print("- advanced_features_test.hwp: 모든 고급 기능이 적용된 문서")
    print("- advanced_features_test.pdf: PDF로 변환된 문서")
    print("- 템플릿: 홈 디렉토리의 HWP_Templates 폴더에 저장")
    print("\n각 파일을 열어서 기능이 제대로 적용되었는지 확인하세요.")

def test_batch_advanced():
    """배치 작업으로 고급 기능 테스트"""
    print("\n=== 배치 작업 고급 기능 테스트 ===\n")
    
    # HWP 컨트롤러 초기화
    hwp = HwpController()
    if not hwp.connect():
        print("❌ HWP 연결 실패")
        return
    
    # 배치 작업 예시: 보고서 템플릿 생성
    operations = [
        # 새 문서 생성
        {"operation": "hwp_create"},
        
        # 페이지 설정
        {"operation": "hwp_set_page", "params": {
            "paper_size": "A4",
            "orientation": "portrait",
            "top_margin": 30,
            "bottom_margin": 30,
            "left_margin": 25,
            "right_margin": 25
        }},
        
        # 머리말/꼬리말 설정
        {"operation": "hwp_set_header_footer", "params": {
            "header_text": "월간 보고서",
            "footer_text": "페이지 ",
            "show_page_number": True
        }},
        
        # 제목 삽입
        {"operation": "hwp_insert_text_with_font", "params": {
            "text": "월간 업무 보고서",
            "font_name": "맑은 고딕",
            "font_size": 24,
            "bold": True
        }},
        
        {"operation": "hwp_insert_paragraph"},
        {"operation": "hwp_set_paragraph", "params": {
            "alignment": "center"
        }},
        
        # 날짜 삽입
        {"operation": "hwp_insert_text_with_font", "params": {
            "text": "2025년 4월",
            "font_size": 14
        }},
        
        {"operation": "hwp_insert_paragraph"},
        {"operation": "hwp_insert_paragraph"},
        
        # 내용 구성
        {"operation": "hwp_insert_text_with_font", "params": {
            "text": "1. 주요 성과",
            "font_size": 16,
            "bold": True
        }},
        
        {"operation": "hwp_insert_paragraph"},
        {"operation": "hwp_set_paragraph", "params": {
            "alignment": "left",
            "line_spacing": 1.5,
            "indent_left": 10
        }},
        
        # 템플릿으로 저장
        {"operation": "hwp_save_as_template", "params": {
            "template_name": "월간보고서_템플릿"
        }},
        
        # PDF로도 저장
        {"operation": "hwp_export_pdf", "params": {
            "output_path": "monthly_report_template.pdf"
        }}
    ]
    
    print("배치 작업 실행 중...")
    # 실제 배치 실행은 hwp_batch_operations 함수를 통해 수행
    print("✅ 배치 작업 예시 완료")
    print("\n생성된 템플릿을 활용하여 매월 보고서를 쉽게 작성할 수 있습니다.")

if __name__ == "__main__":
    # 고급 기능 테스트
    test_advanced_features()
    
    # 배치 작업 예시
    # test_batch_advanced()