"""
글자 크기 및 서식 설정 기능 테스트
피드백에서 언급된 문제를 확인하고 검증하는 테스트 코드
"""
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from src.tools.hwp_controller import HwpController
import time

def test_font_features():
    """글자 서식 기능을 테스트합니다."""
    print("=== 글자 서식 기능 테스트 시작 ===\n")
    
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
    
    # 테스트 1: insert_text_with_font 메서드 테스트
    print("📝 테스트 1: insert_text_with_font 메서드")
    print("- 굵은 14pt 맑은 고딕 텍스트 삽입")
    result = hwp.insert_text_with_font(
        text="이것은 굵은 14pt 맑은 고딕 텍스트입니다.\n",
        font_name="맑은 고딕",
        font_size=14,
        bold=True
    )
    print(f"  결과: {'✅ 성공' if result else '❌ 실패'}\n")
    time.sleep(1)
    
    # 테스트 2: 기울임꼴 텍스트
    print("📝 테스트 2: 기울임꼴 텍스트")
    print("- 기울임꼴 12pt 바탕체 텍스트 삽입")
    result = hwp.insert_text_with_font(
        text="이것은 기울임꼴 12pt 바탕체 텍스트입니다.\n",
        font_name="바탕",
        font_size=12,
        italic=True
    )
    print(f"  결과: {'✅ 성공' if result else '❌ 실패'}\n")
    time.sleep(1)
    
    # 테스트 3: 밑줄 텍스트
    print("📝 테스트 3: 밑줄 텍스트")
    print("- 밑줄이 있는 16pt 텍스트 삽입")
    result = hwp.insert_text_with_font(
        text="이것은 밑줄이 있는 16pt 텍스트입니다.\n",
        font_size=16,
        underline=True
    )
    print(f"  결과: {'✅ 성공' if result else '❌ 실패'}\n")
    time.sleep(1)
    
    # 테스트 4: 복합 스타일
    print("📝 테스트 4: 복합 스타일")
    print("- 굵고 기울임꼴이며 밑줄이 있는 18pt 텍스트")
    result = hwp.insert_text_with_font(
        text="이것은 굵고 기울임꼴이며 밑줄이 있는 18pt 텍스트입니다.\n",
        font_name="굴림",
        font_size=18,
        bold=True,
        italic=True,
        underline=True
    )
    print(f"  결과: {'✅ 성공' if result else '❌ 실패'}\n")
    time.sleep(1)
    
    # 테스트 5: 일반 텍스트 입력 후 서식 적용
    print("📝 테스트 5: 텍스트 입력 후 서식 적용")
    print("- 일반 텍스트 입력")
    
    # 텍스트 입력 전 위치 저장
    start_pos = hwp.hwp.GetPos()
    hwp.insert_text("이 텍스트는 나중에 서식이 적용됩니다.")
    end_pos = hwp.hwp.GetPos()
    
    # 방금 입력한 텍스트 선택
    hwp.hwp.SetPos(start_pos[0], start_pos[1], start_pos[2])
    hwp.hwp.SetPosSel(start_pos[0], start_pos[1], start_pos[2],
                      end_pos[0], end_pos[1], end_pos[2])
    
    print("- 선택된 텍스트에 굵은 20pt 궁서체 적용")
    result = hwp.apply_font_to_selection(
        font_name="궁서",
        font_size=20,
        bold=True
    )
    print(f"  결과: {'✅ 성공' if result else '❌ 실패'}\n")
    
    # 선택 해제하고 다음 줄로 이동
    hwp.hwp.Run("Cancel")
    hwp.insert_paragraph()
    time.sleep(1)
    
    # 테스트 6: 기존 set_font 메서드 테스트
    print("📝 테스트 6: 기존 set_font 메서드 (호환성)")
    print("- set_font로 서식 설정 후 텍스트 입력")
    result = hwp.set_font(
        font_name="돋움",
        font_size=11,
        bold=True,
        italic=False
    )
    if result:
        hwp.insert_text("이것은 set_font 메서드로 설정한 11pt 돋움 굵은 텍스트입니다.\n")
        print(f"  결과: ✅ 성공\n")
    else:
        print(f"  결과: ❌ 실패\n")
    
    # 문서 저장
    print("\n💾 테스트 문서 저장 중...")
    save_path = os.path.join(os.getcwd(), "font_test_result.hwp")
    if hwp.save_document(save_path):
        print(f"✅ 문서가 저장되었습니다: {save_path}")
    else:
        print("❌ 문서 저장 실패")
    
    print("\n=== 글자 서식 기능 테스트 완료 ===")
    print("\n📌 참고사항:")
    print("- 저장된 문서를 열어서 각 텍스트의 서식이 올바르게 적용되었는지 확인하세요.")
    print("- 글자 크기, 글꼴, 굵게, 기울임꼴, 밑줄이 모두 정상적으로 표시되어야 합니다.")

if __name__ == "__main__":
    test_font_features()