import io
import copy
import streamlit as st
from pptx import Presentation

# 페이지 기본 설정
st.set_page_config(
    page_config_title="Meeting App - PPT 취합 & 자동 정화 도구",
    page_icon="📑",
    layout="wide"
)

def remove_empty_placeholders(slide):
    """
    슬라이드 내 빈 개체 틀(Ghost Placeholders) 및 불필요한 기본 안내 문구 상자 완전 제거
    """
    removed_count = 0
    shapes_to_remove = []

    for shape in slide.shapes:
        if shape.is_placeholder:
            # 텍스트 프레임이 있는 경우 내용 유무 검사
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                # 비어있거나 마스터 상속 기본 안내 문구인 경우 삭제 대상 처리
                if not text or text in ["텍스트를 입력하십시오", "제목을 입력하십시오", "Click to edit Master title style"]:
                    shapes_to_remove.append(shape)
            else:
                # 내용이 없는 미디어/이미지 개체 틀
                shapes_to_remove.append(shape)

    for shape in shapes_to_remove:
        try:
            sp = shape._element
            sp.getparent().remove(sp)
            removed_count += 1
        except Exception:
            pass

    return removed_count


def copy_slide_content(dest_prs, source_slide):
    """
    소스 슬라이드의 개체, 서식, 미디어 관계(rId)를 안전하게 복제하여 XML 참조 충돌 방지
    """
    # 빈 레이아웃(Blank Layout)으로 대상 슬라이드 생성
    try:
        blank_layout = dest_prs.slide_layouts[6]
    except IndexError:
        blank_layout = dest_prs.slide_layouts[len(dest_prs.slide_layouts) - 1]

    new_slide = dest_prs.slides.add_slide(blank_layout)

    # 1. 새 슬라이드의 기본 개체 틀 요소 말끔히 삭제 (기본 마스터 상속 차단)
    for shape in list(new_slide.shapes):
        sp = shape._element
        sp.getparent().remove(sp)

    # 2. 소스 슬라이드의 모든 도형/개체 XML 깊은 복사(Deep Copy)
    for shape in source_slide.shapes:
        new_element = copy.deepcopy(shape._element)
        new_slide.shapes._spTree.append(new_element)

    # 3. 미디어, 이미지, 기타 관계(Relationships/rId) 참조 재연결 (파일 손상 복구 경고 예방)
    for rel in source_slide.part.rels.values():
        if "notesSlide" in rel.reltype or "slideLayout" in rel.reltype:
            continue
        try:
            new_slide.part.rels.get_or_add_rel(
                rel.reltype,
                rel.target_ref if rel.is_external else rel.target_part
            )
        except Exception:
            pass

    return new_slide


def merge_presentations(uploaded_files, clean_placeholders=True):
    """
    여러 PPTX 파일을 순서대로 병합하고 유령 개체 틀 정화 작업을 수행
    """
    if not uploaded_files:
        return None, 0, 0

    # 기준 슬라이드 크기 설정을 위해 첫 번째 파일 참조
    first_bytes = io.BytesIO(uploaded_files[0].getvalue())
    first_prs = Presentation(first_bytes)

    # 병합 결과를 담을 새 프레젠테이션 객체
    merged_prs = Presentation()
    merged_prs.slide_width = first_prs.slide_width
    merged_prs.slide_height = first_prs.slide_height

    # 새 프레젠테이션 생성 시 기본 추가되는 첫 슬라이드 제거
    for i in range(len(merged_prs.slides) - 1, -1, -1):
        r_id = merged_prs.slides._sldIdLst[i].rId
        merged_prs.part.drop_rel(r_id)
        del merged_prs.slides._sldIdLst[i]

    total_slides = 0
    total_cleaned_placeholders = 0

    for file in uploaded_files:
        file_bytes = io.BytesIO(file.getvalue())
        prs = Presentation(file_bytes)

        for slide in prs.slides:
            # 1. 원본 슬라이드 내 유령 개체 틀 제거
            if clean_placeholders:
                cleaned = remove_empty_placeholders(slide)
                total_cleaned_placeholders += cleaned

            # 2. 슬라이드 요소 및 서식 안전 복사
            new_slide = copy_slide_content(merged_prs, slide)

            # 3. 복사본 재검증 정화
            if clean_placeholders:
                cleaned_post = remove_empty_placeholders(new_slide)
                total_cleaned_placeholders += cleaned_post

            total_slides += 1

    # 메모리 스트림에 결과 저장
    output_stream = io.BytesIO()
    merged_prs.save(output_stream)
    output_stream.seek(0)

    return output_stream, total_slides, total_cleaned_placeholders


# --- Streamlit 웹 대시보드 UI ---
def main():
    st.title("📑 Meeting App - PPT 취합 & 오류 해결 자동화")
    st.markdown(
        """
        제출된 여러 회의 자료(`.pptx`)를 취합할 때 발생하는 **유령 글상자(`텍스트를 입력하십시오`) 노출**과 
        **'내용에 문제가 있습니다' 파일 손상 경고**를 자동으로 정화하여 완벽한 하나의 통합 발표 자료로 만들어 드립니다.
        """
    )
    st.divider()

    # 사이드바 옵션 설정
    st.sidebar.header("🛠️ 정화 및 병합 옵션")
    clean_option = st.sidebar.checkbox("유령 개체 틀(빈 글상자) 자동 정화", value=True)
    output_filename = st.sidebar.text_input("출력 파일명 설정", value="회의자료_통합본_정화완료.pptx")

    # 파일 업로더
    uploaded_files = st.file_uploader(
        "취합할 PPTX 파일들을 선택하거나 드래그앤드롭하세요 (여러 파일 선택 가능)",
        type=["pptx"],
        accept_multiple_files=True
    )

    if uploaded_files:
        st.subheader(f"📋 업로드된 파일 목록 ({len(uploaded_files)}개)")
        
        # 업로드 파일 순서 확인 UI
        for idx, file in enumerate(uploaded_files, start=1):
            st.text(f"{idx}. {file.name} ({file.size / 1024:.1f} KB)")

        st.write("")
        if st.button("🚀 취합 & 정화 시작하기", type="primary", use_container_width=True):
            with st.spinner("슬라이드를 병합하고 유령 개체 틀을 정화하는 중입니다..."):
                try:
                    result_stream, total_slides, cleaned_count = merge_presentations(
                        uploaded_files, clean_placeholders=clean_option
                    )

                    if result_stream:
                        st.success("✅ PPTX 취합 및 자동 정화가 완료되었습니다!")
                        
                        # 지표 요약 카드
                        col1, col2, col3 = st.columns(3)
                        col1.metric("통합된 슬라이드 수", f"{total_slides}장")
                        col2.metric("제거된 유령 글상자", f"{cleaned_count}개")
                        col3.metric("최종 파일 크기", f"{len(result_stream.getvalue()) / 1024:.1f} KB")

                        st.divider()
                        
                        # 다운로드 버튼
                        st.download_button(
                            label="📥 정화 완료된 통합 PPTX 다운로드",
                            data=result_stream,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"❌ 취합 중 오류가 발생했습니다: {str(e)}")
                    st.info("💡 팁: 구버전 `.ppt` 파일이 포함되어 있다면 PowerPoint에서 `.pptx`로 재저장 후 다시 시도해 보세요.")

if __name__ == "__main__":
    main()