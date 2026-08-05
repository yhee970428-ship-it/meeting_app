import os
import io
import zipfile
import psycopg
from psycopg.rows import dict_row
from flask import Flask, render_template, request, jsonify, send_file
from pptx import Presentation

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
        # 회의 안건 테이블
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
        # 의견 제출 게시판 테이블
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS feedback (
                id SERIAL PRIMARY KEY,
                submit_date VARCHAR(20) NOT NULL,
                dept VARCHAR(100) NOT NULL,
                name VARCHAR(100) NOT NULL,
                title VARCHAR(255) NOT NULL,
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
                dept TEXT NOT NULL,
                name TEXT NOT NULL,
                title TEXT NOT NULL,
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

# 1. 안건 등록 API
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

# 2. 안건 목록 조회 API
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

# 3. 안건 삭제 API
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

# 4. 파일 업로드 API (PPT, HWP, HWPX)
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

# 5. 개별 파일 다운로드 API
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

# 6. PPT 선택 및 전체 통합 취합 API (레이아웃/서식 유지 개선)
@app.route('/api/merge', methods=['POST'])
def merge_ppts():
    data = request.json or {}
    selected_ids = data.get('ids', [])

    conn = get_db_connection()
    cursor = conn.cursor()

    if selected_ids:
        placeholders = ', '.join(['%s' if DATABASE_URL else '?'] * len(selected_ids))
        query = f"SELECT original_filename, file_data FROM topics WHERE id IN ({placeholders}) AND status IN ('제출완료', '재제출완료') AND (LOWER(original_filename) LIKE '%%.ppt' OR LOWER(original_filename) LIKE '%%.pptx') ORDER BY id ASC"
        cursor.execute(query, selected_ids)
    else:
        query = "SELECT original_filename, file_data FROM topics WHERE status IN ('제출완료', '재제출완료') AND (LOWER(original_filename) LIKE '%.ppt' OR LOWER(original_filename) LIKE '%.pptx') ORDER BY id ASC"
        cursor.execute(query)

    rows = cursor.fetchall()
    conn.close()

    if not rows:
        return jsonify({"success": False, "message": "취합할 제출 완료된 PPT 파일이 없습니다."}), 400

    try:
        # 첫 번째 파일 기반으로Presentation 객체 생성 (기본 마스터 레이아웃 보유)
        first_bytes = bytes(rows[0]['file_data']) if isinstance(rows[0]['file_data'], memoryview) else rows[0]['file_data']
        base_prs = Presentation(io.BytesIO(first_bytes))

        # 두 번째 파일부터 서식 유지하며 슬라이드 개체 복사
        for row in rows[1:]:
            sub_bytes = bytes(row['file_data']) if isinstance(row['file_data'], memoryview) else row['file_data']
            sub_prs = Presentation(io.BytesIO(sub_bytes))
            
            for slide in sub_prs.slides:
                # 빈 슬라이드 레이아웃 추가 (index 6: Blank Layout)
                blank_layout = base_prs.slide_layouts[6] if len(base_prs.slide_layouts) > 6 else base_prs.slide_layouts[0]
                new_slide = base_prs.slides.add_slide(blank_layout)

                # 원본 슬라이드의 요소들을 XML 레벨에서 복사해 레이아웃 변형 최소화
                for shape in slide.shapes:
                    try:
                        el = shape.element
                        new_el = el.clone()
                        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
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

# 7. 한글(HWP/HWPX) 선택 및 전체 취합 API
@app.route('/api/merge-hwp', methods=['POST'])
def merge_hwps():
    data = request.json or {}
    selected_ids = data.get('ids', [])

    conn = get_db_connection()
    cursor = conn.cursor()

    if selected_ids:
        placeholders = ', '.join(['%s' if DATABASE_URL else '?'] * len(selected_ids))
        query = f"SELECT original_filename, file_data FROM topics WHERE id IN ({placeholders}) AND status IN ('제출완료', '재제출완료') AND (LOWER(original_filename) LIKE '%%.hwp' OR LOWER(original_filename) LIKE '%%.hwpx') ORDER BY id ASC"
        cursor.execute(query, selected_ids)
    else:
        query = "SELECT original_filename, file_data FROM topics WHERE status IN ('제출완료', '재제출완료') AND (LOWER(original_filename) LIKE '%.hwp' OR LOWER(original_filename) LIKE '%.hwpx') ORDER BY id ASC"
        cursor.execute(query)

    rows = cursor.fetchall()
    conn.close()

    if not rows:
        return jsonify({"success": False, "message": "취합할 제출 완료된 한글(HWP) 파일이 없습니다."}), 400

    try:
        first_bytes = bytes(rows[0]['file_data']) if isinstance(rows[0]['file_data'], memoryview) else rows[0]['file_data']
        first_filename = rows[0]['original_filename'].lower()
        
        if first_filename.endswith('.hwpx'):
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

# 8. 의견 제출 게시판 API
@app.route('/api/feedback', methods=['GET'])
def get_feedback():
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT id, submit_date, dept, name, title, content FROM feedback ORDER BY id DESC')
        rows = cursor.fetchall()
        conn.close()

        feedbacks = [dict(row) for row in rows]
        return jsonify(feedbacks)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/feedback', methods=['POST'])
def add_feedback():
    data = request.json or {}
    submit_date = data.get("submit_date")
    dept = data.get("dept")
    name = data.get("name")
    title = data.get("title")
    content = data.get("content")

    if not (submit_date and dept and name and title and content):
        return jsonify({"success": False, "message": "모든 항목을 입력해 주세요."}), 400

    conn = get_db_connection()
    cursor = conn.cursor()

    if DATABASE_URL:
        cursor.execute(
            'INSERT INTO feedback (submit_date, dept, name, title, content) VALUES (%s, %s, %s, %s, %s)',
            (submit_date, dept.strip(), name.strip(), title.strip(), content.strip())
        )
    else:
        cursor.execute(
            'INSERT INTO feedback (submit_date, dept, name, title, content) VALUES (?, ?, ?, ?, ?)',
            (submit_date, dept.strip(), name.strip(), title.strip(), content.strip())
        )

    conn.commit()
    conn.close()

    return jsonify({"success": True})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)