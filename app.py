import os
import sqlite3
import io
from flask import Flask, render_template, request, jsonify, send_file
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

app = Flask(__name__)

DB_FILE = 'meeting_data.db'

# DB 초기화 및 열 자동 추가(마이그레이션)
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    # 테이블 생성
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
        )
    ''')
    
    # 기존 DB에 file_data 컬럼이 없는 경우 대비한 자동 추가
    cursor.execute("PRAGMA table_info(topics)")
    columns = [col[1] for col in cursor.fetchall()]
    if 'file_data' not in columns:
        cursor.execute("ALTER TABLE topics ADD COLUMN file_data BLOB")
    if 'reupload_reason' not in columns:
        cursor.execute("ALTER TABLE topics ADD COLUMN reupload_reason TEXT")
        
    conn.commit()
    conn.close()

init_db()

def get_db_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

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

# 파일 업로드 (DB 영구 저장)
@app.route('/api/upload/<int:topic_id>', methods=['POST'])
def upload_file(topic_id):
    if 'file' not in request.files:
        return jsonify({"success": False, "message": "파일이 없습니다."}), 400

    file = request.files['file']
    reason = request.form.get('reason', '')

    if file.filename == '':
        return jsonify({"success": False, "message": "선택된 파일이 없습니다."}), 400

    if not file.filename.lower().endswith(('.ppt', '.pptx')):
        return jsonify({"success": False, "message": "PPT/PPTX 파일만 업로드 가능합니다."}), 400

    file_bytes = file.read()
    original_filename = file.filename

    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute('SELECT status FROM topics WHERE id = ?', (topic_id,))
    row = cursor.fetchone()
    
    status = "제출완료"
    if row and row['status'] in ['제출완료', '재제출완료']:
        status = "재제출완료"

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
    cursor.execute('SELECT original_filename, file_data FROM topics WHERE id = ?', (topic_id,))
    row = cursor.fetchone()
    conn.close()

    if row and row['file_data']:
        return send_file(
            io.BytesIO(row['file_data']),
            as_attachment=True,
            download_name=row['original_filename']
        )
    return "파일을 찾을 수 없습니다.", 404

# PPT 자동 병합
@app.route('/api/merge', methods=['POST'])
def merge_ppts():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT file_data FROM topics WHERE status IN ("제출완료", "재제출완료") AND file_data IS NOT NULL ORDER BY id ASC')
    rows = cursor.fetchall()
    conn.close()

    if not rows:
        return jsonify({"success": False, "message": "취합할 제출 완료된 PPT 파일이 없습니다."}), 400

    try:
        first_bytes = io.BytesIO(rows[0]['file_data'])
        base_prs = Presentation(first_bytes)

        for row in rows[1:]:
            sub_bytes = io.BytesIO(row['file_data'])
            sub_prs = Presentation(sub_bytes)
            
            for slide in sub_prs.slides:
                blank_layout = base_prs.slide_layouts[6]
                new_slide = base_prs.slides.add_slide(blank_layout)

                for shape in slide.shapes:
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        try:
                            image_bytes = shape.image.blob
                            image_stream = io.BytesIO(image_bytes)
                            new_slide.shapes.add_picture(
                                image_stream, shape.left, shape.top, shape.width, shape.height
                            )
                        except Exception:
                            from copy import deepcopy
                            new_slide.shapes._spTree.insert_element_before(deepcopy(shape.element), 'p:extLst')
                    else:
                        try:
                            from copy import deepcopy
                            new_slide.shapes._spTree.insert_element_before(deepcopy(shape.element), 'p:extLst')
                        except Exception:
                            pass

        output_stream = io.BytesIO()
        base_prs.save(output_stream)
        output_stream.seek(0)

        return send_file(
            output_stream,
            as_attachment=True,
            download_name="최종_회의자료_통합본.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        return jsonify({"success": False, "message": f"취합 중 오류 발생: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)