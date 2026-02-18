from flask import Flask, render_template, request, send_file, jsonify
import io
import json
from docx_generator import generate_exam_docx

app = Flask(__name__)

SUBJECTS = [
    {"name": "Mathematics", "code": "041"},
    {"name": "Science", "code": "086"},
    {"name": "English", "code": "101"},
    {"name": "Hindi", "code": "002"},
    {"name": "Social Science", "code": "087"},
    {"name": "Physics", "code": "042"},
    {"name": "Chemistry", "code": "043"},
    {"name": "Biology", "code": "044"},
    {"name": "History", "code": "027"},
    {"name": "Geography", "code": "029"},
    {"name": "Political Science", "code": "028"},
    {"name": "Economics", "code": "030"},
    {"name": "Computer Science", "code": "083"},
    {"name": "Information Technology", "code": "802"},
    {"name": "Sanskrit", "code": "122"},
    {"name": "EVS", "code": "006"},
    {"name": "Accountancy", "code": "055"},
    {"name": "Business Studies", "code": "054"},
    {"name": "Physical Education", "code": "048"},
    {"name": "Fine Arts", "code": "049"},
    {"name": "Music", "code": "031"},
    {"name": "Home Science", "code": "064"},
    {"name": "Psychology", "code": "037"},
    {"name": "Sociology", "code": "039"},
    {"name": "English Core", "code": "301"},
    {"name": "English Elective", "code": "001"},
    {"name": "Hindi Core", "code": "302"},
    {"name": "Hindi Elective", "code": "002"},
    {"name": "Mathematics Standard", "code": "041"},
    {"name": "Mathematics Basic", "code": "241"},
]

@app.route('/')
def index():
    return render_template('index.html', subjects=SUBJECTS)

@app.route('/api/subjects')
def get_subjects():
    return jsonify(SUBJECTS)

@app.route('/api/paper/generate', methods=['POST'])
def generate_paper():
    try:
        data = request.get_json()
        docx_buffer = generate_exam_docx(data)
        
        metadata = data.get('metadata', {})
        school = metadata.get('schoolName', 'ExamPaper').replace(' ', '_')
        subject = metadata.get('subject', 'Subject').replace(' ', '_')
        exam_type = metadata.get('examType', 'Exam').replace(' ', '_')
        filename = f"{school}_{subject}_{exam_type}.docx"
        
        return send_file(
            io.BytesIO(docx_buffer),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
