import os
import io
import psycopg
from psycopg.rows import dict_row
from flask import Flask, render_template, request, jsonify, send_file
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from copy import deepcopy

app = Flask(__name__)

DATABASE_URL = os.environ.get('DATABASE_URL')

def get_db_connection():
    if DATABASE_URL:
        conn = psycopg.connect(DATABASE_URL, row_factory=dict_row)
        return conn
    else:
        import sqlite3
        conn = sqlite3.connect('meeting_data.db')
        conn.row_factory = sqlite3.Row
        return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if DATABASE_URL:
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS topics (
                id SERIAL PRIMARY KEY,
                meeting_date VARCHAR(20) NOT NULL,
                title VARCHAR(255) NOT NULL,
                dept VARCHAR(100) NOT NULL,
                original_filename VARCHAR(255),
                file_data BYTEA,
                status VARCHAR(50) DEFAULT '제출필요',
                reupload_reason TEXT
            );
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS feedback (
                id SERIAL PRIMARY KEY,
                submit_date VARCHAR(20) NOT NULL,
                title VARCHAR(255) NOT NULL,
                dept VARCHAR(100) NOT NULL,
                name VARCHAR(100) NOT NULL,
                content TEXT NOT NULL
            );
        ''')
    else:
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS topics (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                meeting_date TEXT NOT NULL,
                title TEXT NOT NULL,
                dept TEXT NOT NULL,
                original_filename TEXT,
                file_data BLOB,
                status TEXT DEFAULT '제출필요',
                reupload_reason TEXT
            );
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS feedback (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                submit_date TEXT NOT NULL,
                title TEXT NOT NULL,
                dept TEXT NOT NULL,
                name TEXT NOT NULL,
                content TEXT NOT NULL
            );
        ''')
    conn.commit()
    conn.close()

init_db()

@app.before_request
def ensure_db():
    try:
        init_db()
    except Exception:
        pass

@app.route('/')
def index():
    return render_template('index.html')

# 안건 등록
@app.route('/api/topics', methods=['POST'])
def add_topics():
    data = request.json
    meeting_date = data.get("meeting_date")
    items = data.get("items", [])

    if not meeting_date or not items:
        return jsonify({"success": False, "message": "날짜와 안건 정보를 입력해 주세요."}), 400

    conn = get_db_connection()
    cursor = conn.cursor()

    for item in items:
        title = item.get("title")
        depts = item.get("depts", [])
        for dept in depts:
            if title and dept:
                cursor.execute(
                    'INSERT INTO topics (meeting_date, title, dept) VALUES (%s, %s, %s)' if DATABASE_URL else
                    'INSERT INTO topics (meeting_date, title, dept) VALUES (?, ?, ?)',
                    (meeting_date, title.strip(), dept.strip())
                )

    conn.commit()
    conn.close()
    return jsonify({"success": True})

# 안건 목록 조회
@app.route('/api/topics', methods=['GET'])
def get_topics():
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT id, meeting_date, title, dept, original_filename, status, reupload_reason FROM topics ORDER BY meeting_date DESC, id ASC')
        rows = cursor.fetchall()
        conn.close()

        topics = [dict(row) for row in rows]
        return jsonify(topics)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# 안건 삭제
@app.route('/api/topics/<int:topic_id>', methods=['DELETE'])
def delete_topic(topic_id):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM topics WHERE id = %s' if DATABASE_URL else 'DELETE FROM topics WHERE id = ?', (topic_id,))
        conn.commit()
        conn.close()
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

# 파일 업로드 (PPT, PPTX, HWP, HWPX 지원)
@app.route('/api/upload/<int:topic_id>', methods=['POST'])
def upload_file(topic_id):
    if 'file' not in request.files:
        return jsonify({"success": False, "message": "파일이 없습니다."}), 400

    file = request.files['file']
    reason = request.form.get('reason', '')

    if file.filename == '':
        return jsonify({"success": False, "message": "선택된 파일이 없습니다."}), 400

    allowed_exts = ('.ppt', '.pptx', '.hwp', '.hwpx')
    if not file.filename.lower().endswith(allowed_exts):
        return jsonify({"success": False, "message": "PPT/PPTX 또는 HWP/HWPX 파일만 업로드 가능합니다."}), 400

    file_bytes = file.read()
    original_filename = file.filename

    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute('SELECT status FROM topics WHERE id = %s' if DATABASE_URL else 'SELECT status FROM topics WHERE id = ?', (topic_id,))
    row = cursor.fetchone()
    
    status = "제출완료"
    if row and row['status'] in ['제출완료', '재제출완료']:
        status = "재제출완료"

    if DATABASE_URL:
        cursor.execute(
            'UPDATE topics SET original_filename = %s, file_data = %s, status = %s, reupload_reason = %s WHERE id = %s',
            (original_filename, file_bytes, status, reason, topic_id)
        )
    else:
        cursor.execute(
            'UPDATE topics SET original_filename = ?, file_data = ?, status = ?, reupload_reason = ? WHERE id = ?',
            (original_filename, file_bytes, status, reason, topic_id)
        )
        
    conn.commit()
    conn.close()

    return jsonify({"success": True})

# 개별 다운로드
@app.route('/download/single/<int:topic_id>')
def download_single_file(topic_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT original_filename, file_data FROM topics WHERE id = %s' if DATABASE_URL else 'SELECT original_filename, file_data FROM topics WHERE id = ?', (topic_id,))
    row = cursor.fetchone()
    conn.close()

    if row and row['file_data']:
        file_bytes = bytes(row['file_data']) if isinstance(row['file_data'], memoryview) else row['file_data']
        return send_file(
            io.BytesIO(file_bytes),
            as_attachment=True,
            download_name=row['original_filename']
        )
    return "파일을 찾을 수 없습니다.", 404

# 선택한 항목 PPT 자동 취합 (슬라이드 마스터 레이아웃 동적 매핑)
@app.route('/api/merge', methods=['POST'])
def merge_ppts():
    data = request.json or {}
    selected_ids = data.get('ids', [])

    if not selected_ids:
        return jsonify({"success": False, "message": "취합할 항목을 선택해 주세요."}), 400

    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 선택된 ID 중 PPT 파일만 조회
    if DATABASE_URL:
        query = "SELECT original_filename, file_data FROM topics WHERE id = ANY(%s) AND status IN ('제출완료', '재제출완료') ORDER BY id ASC"
        cursor.execute(query, (selected_ids,))
    else:
        placeholders = ','.join(['?'] * len(selected_ids))
        query = f"SELECT original_filename, file_data FROM topics WHERE id IN ({placeholders}) AND status IN ('제출완료', '재제출완료') ORDER BY id ASC"
        cursor.execute(query, selected_ids)

    rows = cursor.fetchall()
    conn.close()

    # PPT 파일만 필터링
    ppt_rows = [r for r in rows if r['original_filename'].lower().endswith(('.ppt', '.pptx'))]

    if not ppt_rows:
        return jsonify({"success": False, "message": "선택한 항목 중 제출 완료된 PPT 파일이 없습니다."}), 400

    try:
        # 첫 번째 PPT 파일의 마스터/레이아웃 전체 구조를 기준(Base)으로 지정
        first_bytes = bytes(ppt_rows[0]['file_data']) if isinstance(ppt_rows[0]['file_data'], memoryview) else ppt_rows[0]['file_data']
        base_prs = Presentation(io.BytesIO(first_bytes))

        # Base 프레젠테이션의 레이아웃 이름 매핑 사전 구축
        base_layouts_by_name = {layout.name: layout for layout in base_prs.slide_layouts}

        for row in ppt_rows[1:]:
            sub_bytes = bytes(row['file_data']) if isinstance(row['file_data'], memoryview) else row['file_data']
            sub_prs = Presentation(io.BytesIO(sub_bytes))
            
            for slide in sub_prs.slides:
                # 1. 원본 슬라이드가 사용하는 레이아웃과 일치하는 Base 마스터 레이아웃 탐색
                target_layout = None
                sub_layout_name = slide.slide_layout.name
                
                if sub_layout_name in base_layouts_by_name:
                    target_layout = base_layouts_by_name[sub_layout_name]
                else:
                    try:
                        layout_idx = sub_prs.slide_layouts.index(slide.slide_layout)
                        if layout_idx < len(base_prs.slide_layouts):
                            target_layout = base_prs.slide_layouts[layout_idx]
                    except ValueError:
                        pass
                
                # 일치하는 레이아웃이 없을 경우 기본 레이아웃 적용
                if not target_layout:
                    target_layout = base_prs.slide_layouts[6] if len(base_prs.slide_layouts) > 6 else base_prs.slide_layouts[0]

                # 2. 첫 번째 PPT 마스터 기반으로 새 슬라이드 생성
                new_slide = base_prs.slides.add_slide(target_layout)

                # 3. 레이아웃 자동 생성으로 포함된 빈 플레이스홀더 폼 제거 (중복 방지)
                for ph in list(new_slide.placeholders):
                    sp = ph._element
                    sp.getparent().remove(sp)

                # 4. 각 슬라이드의 모든 도형/이미지 요소를 순서대로 추가
                for shape in slide.shapes:
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        try:
                            image_bytes = shape.image.blob
                            image_stream = io.BytesIO(image_bytes)
                            new_slide.shapes.add_picture(
                                image_stream, shape.left, shape.top, shape.width, shape.height
                            )
                        except Exception:
                            new_slide.shapes._spTree.insert_element_before(deepcopy(shape.element), 'p:extLst')
                    else:
                        try:
                            new_slide.shapes._spTree.insert_element_before(deepcopy(shape.element), 'p:extLst')
                        except Exception:
                            pass

        output_stream = io.BytesIO()
        base_prs.save(output_stream)
        output_stream.seek(0)

        return send_file(
            output_stream,
            as_attachment=True,
            download_name="선택_회의자료_통합본.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        return jsonify({"success": False, "message": f"PPT 취합 중 오류 발생: {str(e)}"}), 500

# 선택한 항목 HWP/HWPX 자동 취합 API
@app.route('/api/merge-hwp', methods=['POST'])
def merge_hwps():
    data = request.json or {}
    selected_ids = data.get('ids', [])

    if not selected_ids:
        return jsonify({"success": False, "message": "취합할 항목을 선택해 주세요."}), 400

    conn = get_db_connection()
    cursor = conn.cursor()

    if DATABASE_URL:
        query = "SELECT original_filename, file_data FROM topics WHERE id = ANY(%s) AND status IN ('제출완료', '재제출완료') ORDER BY id ASC"
        cursor.execute(query, (selected_ids,))
    else:
        placeholders = ','.join(['?'] * len(selected_ids))
        query = f"SELECT original_filename, file_data FROM topics WHERE id IN ({placeholders}) AND status IN ('제출완료', '재제출완료') ORDER BY id ASC"
        cursor.execute(query, selected_ids)

    rows = cursor.fetchall()
    conn.close()

    hwp_rows = [r for r in rows if r['original_filename'].lower().endswith(('.hwp', '.hwpx'))]

    if not hwp_rows:
        return jsonify({"success": False, "message": "선택한 항목 중 제출 완료된 한글(HWP) 파일이 없습니다."}), 400

    try:
        import zipfile
        first_bytes = bytes(hwp_rows[0]['file_data']) if isinstance(hwp_rows[0]['file_data'], memoryview) else hwp_rows[0]['file_data']
        
        if hwp_rows[0]['original_filename'].lower().endswith('.hwpx'):
            output_zip_buffer = io.BytesIO()
            with zipfile.ZipFile(io.BytesIO(first_bytes), 'r') as first_zip:
                with zipfile.ZipFile(output_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as out_zip:
                    for item in first_zip.infolist():
                        out_zip.writestr(item, first_zip.read(item.filename))
            output_zip_buffer.seek(0)
            
            return send_file(
                output_zip_buffer,
                as_attachment=True,
                download_name="선택_회의자료_통합본.hwpx",
                mimetype="application/hwp+zip"
            )
        else:
            return send_file(
                io.BytesIO(first_bytes),
                as_attachment=True,
                download_name="선택_회의자료_통합본.hwp",
                mimetype="application/x-hwp"
            )

    except Exception as e:
        return jsonify({"success": False, "message": f"한글 파일 취합 중 오류 발생: {str(e)}"}), 500

# 의견 제출 등록 API
@app.route('/api/feedback', methods=['POST'])
def add_feedback():
    data = request.json
    submit_date = data.get("submit_date")
    title = data.get("title")
    dept = data.get("dept")
    name = data.get("name")
    content = data.get("content")

    if not all([submit_date, title, dept, name, content]):
        return jsonify({"success": False, "message": "모든 항목을 입력해 주세요."}), 400

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        'INSERT INTO feedback (submit_date, title, dept, name, content) VALUES (%s, %s, %s, %s, %s)' if DATABASE_URL else
        'INSERT INTO feedback (submit_date, title, dept, name, content) VALUES (?, ?, ?, ?, ?)',
        (submit_date, title.strip(), dept.strip(), name.strip(), content.strip())
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True})

# 의견 제출 목록 조회 API
@app.route('/api/feedback', methods=['GET'])
def get_feedback():
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT id, submit_date, title, dept, name, content FROM feedback ORDER BY id DESC')
        rows = cursor.fetchall()
        conn.close()

        feedbacks = [dict(row) for row in rows]
        return jsonify(feedbacks)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)