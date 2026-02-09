from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import os
import tempfile
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import PyPDF2
from docx import Document as DocxDocument
import io

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB 제한
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'

# 업로드 및 출력 폴더 생성
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf', 'docx', 'txt', 'pptx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def parse_document(file_path, file_type):
    """문서 파일을 파싱하여 텍스트 추출"""
    text_content = []
    
    if file_type == 'pdf':
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page_num, page in enumerate(pdf_reader.pages):
                    text = page.extract_text()
                    if text.strip():
                        text_content.append({
                            'page': page_num + 1,
                            'content': text.strip()
                        })
        except Exception as e:
            return {'error': f'PDF 파싱 오류: {str(e)}'}
    
    elif file_type == 'docx':
        try:
            doc = DocxDocument(file_path)
            paragraphs = []
            for para in doc.paragraphs:
                if para.text.strip():
                    paragraphs.append(para.text.strip())
            if paragraphs:
                text_content.append({
                    'page': 1,
                    'content': '\n'.join(paragraphs)
                })
        except Exception as e:
            return {'error': f'DOCX 파싱 오류: {str(e)}'}
    
    elif file_type == 'txt':
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
                if content.strip():
                    text_content.append({
                        'page': 1,
                        'content': content.strip()
                    })
        except Exception as e:
            return {'error': f'TXT 파싱 오류: {str(e)}'}
    
    return {'content': text_content}

def create_pptx_from_template(template_path, document_content, max_pages=40):
    """템플릿 PPTX를 로드하고 문서 내용을 적용하여 새로운 PPTX 생성"""
    try:
        # 템플릿 로드
        prs = Presentation(template_path)
        
        # 모든 내용을 하나의 텍스트로 합치기
        all_content = []
        for item in document_content:
            all_content.append(item['content'])
        
        combined_content = '\n\n'.join(all_content)
        
        # 내용을 슬라이드로 분할
        # 선택된 페이지 수만큼만 생성
        max_chars_per_slide = len(combined_content) // max_pages if max_pages > 0 else 1000
        
        # 문장 단위로 분할
        sentences = combined_content.split('\n')
        content_chunks = []
        current_chunk = []
        current_length = 0
        
        for sentence in sentences:
            sentence = sentence.strip()
            if not sentence:
                continue
                
            if current_length + len(sentence) > max_chars_per_slide and current_chunk:
                content_chunks.append('\n'.join(current_chunk))
                current_chunk = [sentence]
                current_length = len(sentence)
                
                # 최대 페이지 수에 도달하면 중단
                if len(content_chunks) >= max_pages:
                    break
            else:
                current_chunk.append(sentence)
                current_length += len(sentence) + 1
        
        # 마지막 청크 추가
        if current_chunk and len(content_chunks) < max_pages:
            content_chunks.append('\n'.join(current_chunk))
        
        # 최대 페이지 수로 제한
        content_chunks = content_chunks[:max_pages]
        
        # 각 청크를 슬라이드로 추가
        for idx, chunk in enumerate(content_chunks):
            # 빈 슬라이드 레이아웃 선택 (또는 첫 번째 레이아웃)
            slide_layout = prs.slide_layouts[1]  # Title and Content 레이아웃
            slide = prs.slides.add_slide(slide_layout)
            
            # 제목 설정
            title_shape = slide.shapes.title
            if title_shape:
                title_shape.text = f"슬라이드 {idx + 1}"
            
            # 내용 설정
            content_shape = slide.placeholders[1] if len(slide.placeholders) > 1 else None
            if content_shape:
                text_frame = content_shape.text_frame
                text_frame.text = chunk
                
                # 텍스트 포맷팅
                for paragraph in text_frame.paragraphs:
                    paragraph.font.size = Pt(14)
                    paragraph.alignment = PP_ALIGN.LEFT
                    paragraph.space_after = Pt(12)
        
        # 임시 파일에 저장
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'generated_presentation.pptx')
        prs.save(output_path)
        
        return output_path
    
    except Exception as e:
        raise Exception(f'PPTX 생성 오류: {str(e)}')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        if 'document' not in request.files or 'template' not in request.files:
            return jsonify({'error': '문서 파일과 템플릿 파일을 모두 업로드해주세요.'}), 400
        
        document_file = request.files['document']
        template_file = request.files['template']
        
        if document_file.filename == '' or template_file.filename == '':
            return jsonify({'error': '파일을 선택해주세요.'}), 400
        
        if not allowed_file(document_file.filename) or not allowed_file(template_file.filename):
            return jsonify({'error': '지원하지 않는 파일 형식입니다.'}), 400
        
        # 파일 저장
        doc_filename = secure_filename(document_file.filename)
        template_filename = secure_filename(template_file.filename)
        
        doc_path = os.path.join(app.config['UPLOAD_FOLDER'], doc_filename)
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], template_filename)
        
        document_file.save(doc_path)
        template_file.save(template_path)
        
        # 문서 타입 확인
        doc_ext = doc_filename.rsplit('.', 1)[1].lower()
        
        # 문서 파싱
        parsed_content = parse_document(doc_path, doc_ext)
        
        if 'error' in parsed_content:
            return jsonify({'error': parsed_content['error']}), 400
        
        if not parsed_content['content']:
            return jsonify({'error': '문서에서 내용을 추출할 수 없습니다.'}), 400
        
        # 페이지 수 가져오기
        page_count = request.form.get('pageCount', '40')
        try:
            page_count = int(page_count)
            if page_count < 10 or page_count > 40:
                return jsonify({'error': '페이지 수는 10~40 사이여야 합니다.'}), 400
        except ValueError:
            return jsonify({'error': '유효하지 않은 페이지 수입니다.'}), 400
        
        # PPTX 생성
        output_path = create_pptx_from_template(template_path, parsed_content['content'], page_count)
        
        return jsonify({
            'success': True,
            'message': 'PPTX 파일이 성공적으로 생성되었습니다.',
            'filename': 'generated_presentation.pptx'
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            return jsonify({'error': '파일을 찾을 수 없습니다.'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
