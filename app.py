import io
import copy
import streamlit as st
from pptx import Presentation

# 페이지 기본 설정
st.set_page_config(
    page_title="Meeting App - PPT 취합 & 자동 정화 도구",
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


def copy_slide_content(dest_prs, source_slide, base_layouts_by_name):
    """
    소스 슬라이드의 원본 레이아웃 이름/순서를 기준(Base) PPT 마스터에서 찾아 매핑 후 복제
    """
    source_layout_name = source_slide.slide_layout.name
    target_layout = None

    # 1. 원본 슬라이드 레이아웃 이름과 일치하는 마스터 레이아웃 탐색
    if source_layout_name in base_layouts_by_name:
        target_layout = base_layouts_by_name[source_layout_name]
    else:
        # 2. 이름 매핑 실패 시 레이아웃 인덱스 순서로 대체 탐색
        try:
            layout_idx = source_slide.part.package.presentations[0].slide_layouts.index(source_slide.slide_layout)
            if layout_idx < len(dest_prs.slide_layouts):
                target_layout = dest_prs.slide_layouts[layout_idx]
        except Exception:
            pass

    # 3. 매핑 실패 시 기본 빈 레이아웃 지정
    if not target_layout:
        target_layout = dest_prs.slide_layouts[6] if len(dest_prs.slide_layouts) > 6 else dest_prs.slide_layouts[0]

    # 대상 마스터 레이아웃을 반영하여 새 슬라이드 생성
    new_slide = dest_prs.slides.add_slide(target_layout)

    # 마스터 레이아웃 생성 시 포함되는 중복 기본 개체 틀 완전 제거
    for shape in list(new_slide.placeholders):
        sp = shape._element
        sp.getparent().remove(sp)

    # 소스 슬라이드의 모든 도형/개체 깊은 복사(Deep Copy)
    for shape in source_slide.shapes:
        new_element = copy.deepcopy(shape._element)
        new_slide.shapes._spTree.append(new_element)

    # 미디어/이미지/관계 참조 재연결
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
    여러 PPTX 파일을 첫 번째 PPT 마스터 기반 레이아웃 매핑으로 병합
    """
    if not uploaded_files:
        return None, 0, 0

    # 첫 번째 PPT를 병합의 Base 기준 프레젠테이션으로 지정
    first_bytes = io.BytesIO(uploaded_files[0].getvalue())
    merged_prs = Presentation(first_bytes)

    # 기준(Base) 마스터 레이아웃 사전 구축
    base_layouts_by_name = {layout.name: layout for layout in merged_prs.slide_layouts}

    # 첫 번째 파일의 슬라이드 개수 기록
    total_slides = len(merged_prs.slides)
    total_cleaned_placeholders = 0

    # 첫 번째 파일 내 유령 글상자 정화
    if clean_placeholders:
        for slide in merged_prs.slides:
            total_cleaned_placeholders += remove_empty_placeholders(slide)

    # 두 번째 파일부터 병합 시작
    for file in uploaded_files[1:]:
        file_bytes = io.BytesIO(file.getvalue())
        prs = Presentation(file_bytes)

        for slide in prs.slides:
            if clean_placeholders:
                total_cleaned_placeholders += remove_empty_placeholders(slide)

            # 첫 번째 마스터 디자인을 보존하며 복사
            new_slide = copy_slide_content(merged_prs, slide, base_layouts_by_name)

            if clean_placeholders:
                total_cleaned_placeholders += remove_empty_placeholders(new_slide)

            total_slides += 1

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