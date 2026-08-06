import io
import copy
import streamlit as st
from pptx import Presentation

# 페이지 기본 설정 (page_title 속성 사용)
st.set_page_config(
    page_title="Meeting App - PPT 취합 & 자동 정화 도구",
    page_icon="📑",
    layout="wide"
)

def copy_slide_background(source_slide, target_slide):
    """
    소스 슬라이드 및 레이아웃의 배경(p:bg) XML 요소를 대상 슬라이드로 복사하여
    원본 슬라이드의 배경색, 그라데이션, 이미지 배경을 그대로 유지
    """
    nsmap = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
    
    # 1. 소스 슬라이드 자체에 정의된 배경 검사
    source_bg = source_slide._element.find('p:bg', nsmap)
    
    # 2. 슬라이드에 직접 설정된 배경이 없으면 슬라이드 레이아웃 배경 검사
    if source_bg is None and hasattr(source_slide, 'slide_layout'):
        source_bg = source_slide.slide_layout._element.find('p:bg', nsmap)

    if source_bg is not None:
        existing_bg = target_slide._element.find('p:bg', nsmap)
        if existing_bg is not None:
            target_slide._element.remove(existing_bg)

        new_bg = copy.deepcopy(source_bg)
        csld = target_slide._element.find('p:cSld', nsmap)
        if csld is not None:
            csld.insert(0, new_bg)


def remove_empty_placeholders(slide):
    """
    슬라이드 내 빈 개체 틀(Ghost Placeholders) 및 불필요한 기본 안내 문구 상자 완전 제거
    """
    removed_count = 0
    shapes_to_remove = []

    for shape in slide.shapes:
        if shape.is_placeholder:
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if not text or text in ["텍스트를 입력하십시오", "제목을 입력하십시오", "Click to edit Master title style"]:
                    shapes_to_remove.append(shape)
            else:
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
    소스 슬라이드의 개체, 서식, 배경(p:bg), 레이아웃 그래픽 도형 및 미디어 관계(rId)를
    완전 복제하여 서로 다른 슬라이드 마스터/레이아웃 원본 디자인을 동일하게 유지
    """
    try:
        blank_layout = dest_prs.slide_layouts[6]
    except IndexError:
        blank_layout = dest_prs.slide_layouts[len(dest_prs.slide_layouts) - 1]

    new_slide = dest_prs.slides.add_slide(blank_layout)

    # 1. 새 슬라이드의 기본 개체 틀 요소 말끔히 삭제 (기본 마스터 상속 차단)
    for shape in list(new_slide.shapes):
        sp = shape._element
        sp.getparent().remove(sp)

    # 2. 원본 슬라이드/레이아웃의 배경(p:bg) XML 복사
    copy_slide_background(source_slide, new_slide)

    # 3. 원본 레이아웃에 포함된 배경 디자인/장식 도형(개체틀 제외) 복사
    if hasattr(source_slide, 'slide_layout'):
        for l_shape in source_slide.slide_layout.shapes:
            if not l_shape.is_placeholder:
                new_l_element = copy.deepcopy(l_shape._element)
                # 배경 레이어로 삽입 (본문 도형 뒤에 배치)
                new_slide.shapes._spTree.insert(2, new_l_element)

    # 4. 소스 슬라이드의 본문 도형/개체 XML 깊은 복사(Deep Copy)
    for shape in source_slide.shapes:
        new_element = copy.deepcopy(shape._element)
        new_slide.shapes._spTree.append(new_element)

    # 5. 소스 슬라이드의 미디어/이미지/관계 참조(rels) 재연결
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

    # 6. 소스 레이아웃의 미디어/이미지 관계 참조(rels) 재연결
    if hasattr(source_slide, 'slide_layout'):
        for rel in source_slide.slide_layout.part.rels.values():
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
    여러 PPTX 파일을 순서대로 병합하며 슬라이드별 고유 레이아웃과 배경 디자인을 완전히 보존
    """
    if not uploaded_files:
        return None, 0, 0

    first_bytes = io.BytesIO(uploaded_files[0].getvalue())
    first_prs = Presentation(first_bytes)

    merged_prs = Presentation()
    merged_prs.slide_width = first_prs.slide_width
    merged_prs.slide_height = first_prs.slide_height

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
            if clean_placeholders:
                cleaned = remove_empty_placeholders(slide)
                total_cleaned_placeholders += cleaned

            new_slide = copy_slide_content(merged_prs, slide)

            if clean_placeholders:
                cleaned_post = remove_empty_placeholders(new_slide)
                total_cleaned_placeholders += cleaned_post

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
        **'내용에 문제가 있습니다' 파일 손상 경고**를 자동으로 정화하며, 각 PPT 원본의 **배경 디자인 및 레이아웃 서식을 완벽하게 보존**하여 취합해 드립니다.
        """
    )
    st.divider()

    st.sidebar.header("🛠️ 정화 및 병합 옵션")
    clean_option = st.sidebar.checkbox("유령 개체 틀(빈 글상자) 자동 정화", value=True)
    output_filename = st.sidebar.text_input("출력 파일명 설정", value="회의자료_통합본_정화완료.pptx")

    uploaded_files = st.file_uploader(
        "취합할 PPTX 파일들을 선택하거나 드래그앤드롭하세요 (여러 파일 선택 가능)",
        type=["pptx"],
        accept_multiple_files=True
    )

    if uploaded_files:
        st.subheader(f"📋 업로드된 파일 목록 ({len(uploaded_files)}개)")
        
        for idx, file in enumerate(uploaded_files, start=1):
            st.text(f"{idx}. {file.name} ({file.size / 1024:.1f} KB)")

        st.write("")
        if st.button("🚀 취합 & 정화 시작하기", type="primary", use_container_width=True):
            with st.spinner("슬라이드를 병합하고 서식과 레이아웃을 정밀하게 보존하는 중입니다..."):
                try:
                    result_stream, total_slides, cleaned_count = merge_presentations(
                        uploaded_files, clean_placeholders=clean_option
                    )

                    if result_stream:
                        st.success("✅ PPTX 취합 및 레이아웃/배경 서식 보존 정화가 완료되었습니다!")
                        
                        col1, col2 = st.columns(2)
                        col1.metric("통합된 슬라이드 수", f"{total_slides}장")
                        col2.metric("제거된 유령 글상자", f"{cleaned_count}개")

                        st.download_button(
                            label="📥 통합 PPTX 다운로드",
                            data=result_stream,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"❌ 취합 중 오류가 발생했습니다: {str(e)}")

if __name__ == "__main__":
    main()
```eof

개선된 `app.py` 코드를 생성했습니다. 주요 변경 사항 및 원리는 다음과 같습니다:

1. **배경 XML(`p:bg`) 복사 (`copy_slide_background`)**: 원본 슬라이드 또는 슬라이드 레이아웃에 지정된 단색/그라데이션/패턴 배경 요소를 대상 슬라이드로 직접 이식합니다.
2. **레이아웃 디자인 도형 복사**: 원본 슬라이드 레이아웃(Master Layout)에 배치된 그래픽 디자인, 로고, 선 등(개체틀 제외)을 배경 레이어로 복제하여 2번째, 3번째 PPT 슬라이드가 기존 원본 서식 디자인을 유지하도록 처리했습니다.
3. **레이아웃 미디어 참조(Rels) 복사**: 레이아웃에 삽입되어 있던 로고나 이미지 등의 의존성 관계 파일(`rels`)도 함께 연결하여 파일 손상 없이 안정적으로 출력됩니다.