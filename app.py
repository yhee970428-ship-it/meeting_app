import os
import sqlite3
import io
from flask import Flask, render_template, request, jsonify, send_file
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
DB_FILE = 'meeting_data.db'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# 데이터베이스 초기화 (새로고침 시 데이터 유지)
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS topics (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            meeting_date TEXT NOT NULL,
            title TEXT NOT NULL,
            dept TEXT NOT NULL,
            file_path TEXT,
            original_filename TEXT,
            status TEXT DEFAULT '제출필요'
        )
    ''')
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

# 1. 안건 등록 (날짜 1개 + 안건 여러 개 + 부서 여러 개 다중 추가 지원)
@app.route('/api/topics', methods=['POST'])
def add_topics():
    data = request.json
    meeting_date = data.get("meeting_date")
    items = data.get("items", []) # [{title: "안건1", depts: ["기획팀", "재무팀"]}, ...]

    if not meeting_date or not items:
        return jsonify({"success": False, "message": "날짜와 안건 정보를 모두 입력해 주세요."}), 400

    conn = get_db_connection()
    cursor = conn.cursor()

    for item in items:
        title = item.get("title")
        depts = item.get("depts", []) # 배열 형태
        for dept in depts:
            if title and dept:
                cursor.execute(
                    'INSERT INTO topics (meeting_date, title, dept) VALUES (?, ?, ?)',
                    (meeting_date, title.strip(), dept.strip())
                )

    conn.commit()
    conn.close()
    return jsonify({"success": True})

# 2. 안건 목록 및 대시보드 조회
@app.route('/api/topics', methods=['GET'])
def get_topics():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM topics ORDER BY meeting_date DESC, id ASC')
    rows = cursor.fetchall()
    conn.close()

    topics = [dict(row) for row in rows]
    return jsonify(topics)

# 3. 파일 업로드
@app.route('/api/upload/<int:topic_id>', methods=['POST'])
def upload_file(topic_id):
    if 'file' not in request.files:
        return jsonify({"success": False, "message": "파일이 없습니다."}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"success": False, "message": "선택된 파일이 없습니다."}), 400

    if not file.filename.lower().endswith(('.ppt', '.pptx')):
        return jsonify({"success": False, "message": "PPT/PPTX 파일만 업로드 가능합니다."}), 400

    original_filename = file.filename
    ext = os.path.splitext(original_filename)[1]
    filename = f"topic_{topic_id}{ext}"
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    # DB 업데이트
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        'UPDATE topics SET file_path = ?, original_filename = ?, status = ? WHERE id = ?',
        (filepath, original_filename, "제출완료", topic_id)
    )
    conn.commit()
    conn.close()

    return jsonify({"success": True})

# 4. 개별 업로드 파일 다운로드
@app.route('/download/single/<int:topic_id>')
def download_single_file(topic_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT file_path, original_filename FROM topics WHERE id = ?', (topic_id,))
    row = cursor.fetchone()
    conn.close()

    if row and row['file_path'] and os.path.exists(row['file_path']):
        return send_file(row['file_path'], as_attachment=True, download_name=row['original_filename'])
    return "파일을 찾을 수 없습니다.", 404

# 5. PPT 자동 병합 및 다운로드
@app.route('/api/merge', methods=['POST'])
def merge_ppts():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT file_path FROM topics WHERE status = "제출완료" AND file_path IS NOT NULL ORDER BY id ASC')
    rows = cursor.fetchall()
    conn.close()

    valid_files = [row['file_path'] for row in rows if os.path.exists(row['file_path'])]

    if not valid_files:
        return jsonify({"success": False, "message": "취합할 제출 완료된 PPT 파일이 없습니다."}), 400

    try:
        base_prs = Presentation(valid_files[0])

        for filepath in valid_files[1:]:
            sub_prs = Presentation(filepath)
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

        output_path = os.path.join(OUTPUT_FOLDER, 'merged_meeting_materials.pptx')
        base_prs.save(output_path)
        return jsonify({"success": True, "download_url": "/download/merged"})

    except Exception as e:
        return jsonify({"success": False, "message": f"취합 중 오류 발생: {str(e)}"}), 500

@app.route('/download/merged')
def download_merged_file():
    path = os.path.join(OUTPUT_FOLDER, 'merged_meeting_materials.pptx')
    return send_file(path, as_attachment=True, download_name="최종_회의자료_통합본.pptx")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)