import os
import re
import platform
import subprocess
import tempfile
import time
import requests
from datetime import datetime
from werkzeug.utils import secure_filename
from flask import Blueprint, request, jsonify, send_file
from utils.decorators import require_auth
from utils.db_manager import (
    db,
    ProcessingJob,
    ProcessingJobHistory,
    JobType,
    JobStatus,
)
from utils.file_storage import FileStorageManager
from document_processor import DocumentProcessor
from config import TEMPLATE_DOCX, ALLOWED_EXTENSIONS, MAX_FILE_SIZE, STORAGE_DIR
from io import BytesIO

IS_WINDOWS = platform.system() == "Windows"

if IS_WINDOWS:
    import pythoncom
    from docx2pdf import convert

# Helper Functions
def clean_html(text):
    """Remove HTML tags from a string."""
    if not text: return ""
    return re.sub(r'<[^>]+>', '', text)

def normalize_filename(s):
    """Sanitize a string for use as a filename."""
    if not s: return "document"
    # Remove special chars, replace spaces with underscores
    s = re.sub(r'[^\w\s-]', '', s).strip()
    return re.sub(r'[-\s]+', '_', s)
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def _read_best_available_file(history, file_kind='output'):
    if file_kind == 'upload':
        preferred_path = history.UploadFileServerPath
        preferred_data = history.UploadFileData
        preferred_name = history.UploadFileName
        path_attr = 'UploadFileServerPath'
        data_attr = 'UploadFileData'
        name_attr = 'UploadFileName'
    else:
        preferred_path = history.OutputFileServerPath
        preferred_data = history.OutputFileData
        preferred_name = history.OutputFileName
        path_attr = 'OutputFileServerPath'
        data_attr = 'OutputFileData'
        name_attr = 'OutputFileName'

    try:
        return file_storage_manager.retrieve_file(
            preferred_path, preferred_data, preferred_name
        ).getvalue(), preferred_name
    except FileNotFoundError:
        pass

    if preferred_name:
        candidates = ProcessingJobHistory.query.filter(
            getattr(ProcessingJobHistory, name_attr) == preferred_name
        ).order_by(ProcessingJobHistory.CreatedDate.desc()).all()

        for candidate in candidates:
            candidate_path = getattr(candidate, path_attr)
            candidate_data = getattr(candidate, data_attr)
            try:
                return file_storage_manager.retrieve_file(
                    candidate_path, candidate_data, preferred_name
                ).getvalue(), preferred_name
            except FileNotFoundError:
                continue

    return None, preferred_name


def _is_allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def _coerce_bool(value, default=False):
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {'true', '1', 'yes', 'on'}
    return bool(value)

def get_first_sentence(text):
    """Fallback if Groq title fails: extract first sentence or first 50 chars."""
    if not text: return "New Chat"
    clean = clean_html(text).strip()
    if not clean: return "New Chat"
    
    # Try to find first sentence boundary
    match = re.split(r'[.!?\n]', clean)
    first = match[0].strip() if match else clean[:50]
    return first[:60] if first else "New Chat"


def convert_docx_to_pdf(file_bytes):
    system_os = platform.system()
    print(f"[PDF] Running on OS: {system_os}")
    libreoffice_path = os.getenv("LIBREOFFICE_PATH", "/usr/bin/libreoffice")

    def windows_method():
        pythoncom.CoInitialize()
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                docx_path = os.path.join(tmpdir, "file.docx")
                pdf_path = os.path.join(tmpdir, "file.pdf")

                with open(docx_path, 'wb') as f:
                    f.write(file_bytes)

                convert(docx_path, pdf_path)

                with open(pdf_path, 'rb') as f:
                    return f.read()
        finally:
            pythoncom.CoUninitialize()

    def convert_to_pdf_linux(input_path, output_path):
        subprocess.run([
            libreoffice_path,
            "--headless",
            "--convert-to", "pdf",
            input_path,
            "--outdir", os.path.dirname(output_path)
        ], check=True)

    def linux_method():
        with tempfile.TemporaryDirectory() as tmpdir:
            docx_path = os.path.join(tmpdir, "file.docx")
            pdf_path = os.path.join(tmpdir, "file.pdf")

            with open(docx_path, 'wb') as f:
                f.write(file_bytes)

            convert_to_pdf_linux(docx_path, pdf_path)

            with open(pdf_path, 'rb') as f:
                return f.read()

    # =========================
    # WINDOWS
    # =========================
    if system_os == "Windows":
        return windows_method()

    else:
        return linux_method()

process_bp = Blueprint('process', __name__)
file_storage_manager = FileStorageManager(STORAGE_DIR)


# ==============================
# SMART TITLE GENERATOR
# ==============================
def generate_title(text):
    """
    Generate a smart summary title using Groq API.
    Fallback to first sentence if API fails.
    """
    try:
        api_key = os.getenv("GROQ_API")
        if not api_key:
            logger.warning("GROQ_API key not found, skipping smart title")
            raise ValueError("Missing API Key")
        clean_prompt = clean_html(text)[:1000]
        if not clean_prompt:
            return "New Chat"

        response = requests.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={
                "model": "llama-3.3-70b-versatile",
                "messages": [
                    {"role": "system", "content": "Generate a catchy 3-5 word title for the following text. Respond ONLY with the plain text title, no punctuation or quotes."},
                    {"role": "user", "content": clean_prompt}
                ],
                "max_tokens": 20
            },
            timeout=5
        )

        if response.status_code == 200:
            title = response.json().get('choices', [{}])[0].get('message', {}).get('content', '').strip()
            if title:
                return normalize_filename(title).replace("_", " ")[:60]

    except Exception as e:
        logger.error(f"Groq Title API failed: {e}")

    return get_first_sentence(text)


# ==============================
# UNIQUE FILENAME GENERATOR
# ==============================
def generate_unique_filename(base_name, current_user):
    clean_name = normalize_filename(base_name)
    # Compact timestamp: No gaps, includes seconds
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    filename_base = f"CODES_{clean_name}_{timestamp}"
    filename = f"{filename_base}.docx"

    counter = 1
    while ProcessingJobHistory.query.join(ProcessingJob).filter(
        ProcessingJob.UserId == current_user.Id,
        ProcessingJobHistory.OutputFileName == filename
    ).first():
        filename = f"{filename_base}_{counter}.docx"
        counter += 1

    return filename


# ==============================
# PROCESS TEXT
# ==============================
@process_bp.route('/process-text', methods=['POST'])
@require_auth
def process_text(current_user):
    try:
        start_time = time.time()
        data = request.get_json()
        text_input = data.get('text', '')

        user_font = data.get('fontFamily', 'Calibri')
        user_size = int(data.get('fontSize', 11))
        include_cover = _coerce_bool(data.get('includeCover', False))
        include_toc = _coerce_bool(data.get('includeTOC', False))

        processor = DocumentProcessor(
            template_path=TEMPLATE_DOCX,
            font_family=user_font,
            font_size=user_size,
            include_cover=include_cover,
            include_toc=include_toc
        )

        doc_obj = processor.html_to_docx(text_input)

        if include_toc:
            with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_doc:
                temp_docx_path = tmp_doc.name
            try:
                doc_obj.save(temp_docx_path)
                processor.refresh_saved_docx(temp_docx_path)
                with open(temp_docx_path, 'rb') as f:
                    final_file_data = f.read()
            finally:
                if os.path.exists(temp_docx_path):
                    os.remove(temp_docx_path)
        else:
            output_buffer = BytesIO()
            doc_obj.save(output_buffer)
            output_buffer.seek(0)
            final_file_data = output_buffer.read()

        job_id = data.get('jobId', None)

        # Job handling
        if job_id:
            job = ProcessingJob.query.filter_by(Id=job_id, UserId=current_user.Id).first()
            if not job:
                return jsonify({'success': False, 'error': 'Conversation not found'}), 404
        else:
            job_name = generate_title(text_input)
            job = ProcessingJob(JobName=job_name, UserId=current_user.Id)
            db.session.add(job)
            db.session.flush()

        # Generate an output file name based on the actual text input, not the chat's title
        if job_id:
            file_base_name = get_first_sentence(text_input)
        else:
            file_base_name = job_name
            
        filename = generate_unique_filename(file_base_name, current_user)

        processing_time = round(time.time() - start_time, 2)
        stored_output = file_storage_manager.save_file(
            final_file_data, filename, current_user.Id, job.Id
        )

        history = ProcessingJobHistory(
            ProcessJobId=job.Id,
            JobType=JobType.TEXT,
            ProcessingTime=processing_time,
            UploadFileData=text_input.encode('utf-8'),
            Status=JobStatus.SUCCESS,
            OutputFileData=stored_output['db_bytes'],
            OutputFileServerPath=stored_output['server_path'],
            OutputFileName=filename,
            FontFamily=user_font,
            FontSize=user_size,
            IncludeCover=include_cover,
            IncludeTOC=include_toc,
            CreatedBy=current_user.Id,
            ModifiedBy=current_user.Id
        )

        db.session.add(history)
        db.session.commit()
        processing_count = history.ProcessingCount

        return jsonify({
            'success': True, 
            'historyId': history.Id, 
            'jobId': job.Id,
            'jobName': job.JobName,
            'processingTime': processing_time,
            'time': datetime.now().strftime('%d %b %Y, %I:%M %p'),
            'processingCount': processing_count
        })

    except Exception as e:
        db.session.rollback()
        logger.error(f"Error in process_text: {e}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)}), 500


# ==============================
# PROCESS FILE
# ==============================
@process_bp.route('/process-file', methods=['POST'])
@require_auth
def process_file(current_user):
    try:
        start_time = time.time()
        if request.content_length and request.content_length > MAX_FILE_SIZE:
            return jsonify({'success': False, 'error': 'File exceeds maximum allowed size'}), 400

        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file uploaded'})

        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'})
        if not _is_allowed_file(file.filename):
            return jsonify({'success': False, 'error': 'Unsupported file type'}), 400

        user_font = request.form.get('fontFamily', 'Calibri')
        user_size = int(request.form.get('fontSize', 11))
        include_cover = _coerce_bool(request.form.get('includeCover'))
        include_toc = _coerce_bool(request.form.get('includeTOC'))
        job_id = request.form.get('jobId', None)

        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = os.path.join(tmpdir, secure_filename(file.filename))
            output_path = os.path.join(tmpdir, "output.docx")

            file.save(input_path)

            processor = DocumentProcessor(
                template_path=TEMPLATE_DOCX,
                font_family=user_font,
                font_size=user_size,
                include_cover=include_cover,
                include_toc=include_toc
            )

            result = processor.universal_extract(input_path, output_path)

            if not result.get('success', False):
                return jsonify({'success': False, 'error': result.get('error')}), 500

            with open(input_path, 'rb') as f:
                original_file_data = f.read()

            with open(output_path, 'rb') as f:
                final_file_data = f.read()

        # Job handling
        if job_id and job_id != 'null':
            job = ProcessingJob.query.filter_by(Id=job_id, UserId=current_user.Id).first()
            if not job:
                return jsonify({'success': False, 'error': 'Conversation not found'}), 404
        else:
            base_name = os.path.splitext(file.filename)[0]
            job = ProcessingJob(JobName=base_name, UserId=current_user.Id)
            db.session.add(job)
            db.session.flush()

        base_name = os.path.splitext(file.filename)[0]
        filename = generate_unique_filename(base_name, current_user)
        processing_time = round(time.time() - start_time, 2)

        stored_upload = file_storage_manager.save_file(
            original_file_data, secure_filename(file.filename), current_user.Id, job.Id
        )
        stored_output = file_storage_manager.save_file(
            final_file_data, filename, current_user.Id, job.Id
        )

        history = ProcessingJobHistory(
            ProcessJobId=job.Id,
            JobType=JobType.FILE,
            UploadFileName=secure_filename(file.filename),
            UploadFileData=stored_upload['db_bytes'],
            UploadFileServerPath=stored_upload['server_path'],
            OutputFileName=filename,
            OutputFileData=stored_output['db_bytes'],
            OutputFileServerPath=stored_output['server_path'],
            ProcessingTime=processing_time,
            FontFamily=user_font,
            FontSize=user_size,
            IncludeCover=include_cover,
            IncludeTOC=include_toc,
            Status=JobStatus.SUCCESS,
            CreatedBy=current_user.Id,
            ModifiedBy=current_user.Id
        )

        db.session.add(history)
        db.session.commit()

        processing_count = history.ProcessingCount

        return jsonify({
            'success': True, 
            'historyId': history.Id, 
            'jobId': job.Id,
            'jobName': job.JobName,
            'processingTime': processing_time,
            'time': datetime.now().strftime('%d %b %Y, %I:%M %p'),
            'processingCount': processing_count
        })

    except Exception as e:
        db.session.rollback()
        logger.error(f"Error in process_file: {e}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)}), 500


# ==============================
# UPDATE TITLE (UNCHANGED)
# ==============================
@process_bp.route('/update-title/<int:job_id>', methods=['POST'])
@require_auth
def update_title(current_user, job_id):
    try:
        data = request.json
        new_title = data.get('title', '').strip()
        if not new_title:
            return jsonify({'success': False, 'error': 'Title is required'}), 400

        job = ProcessingJob.query.filter_by(Id=job_id, UserId=current_user.Id).first()
        if not job:
            return jsonify({'success': False, 'error': 'Job not found'}), 404

        job.JobName = new_title
        db.session.commit()
        return jsonify({'success': True})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


# ==============================
# DELETE
# ==============================
@process_bp.route('/delete/<int:job_id>', methods=['DELETE'])
@require_auth
def delete_conversation(current_user, job_id):
    try:
        job = ProcessingJob.query.filter_by(Id=job_id, UserId=current_user.Id).first()
        if not job:
            return jsonify({'success': False, 'error': 'Job not found'}), 404

        ProcessingJobHistory.query.filter_by(ProcessJobId=job.Id).delete()
        db.session.delete(job)
        db.session.commit()

        return jsonify({'success': True})

    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'error': str(e)}), 500


# ==============================
# SIDEBAR
# ==============================
@process_bp.route('/conversations', methods=['GET'])
@require_auth
def get_conversations(current_user):
    try:
        jobs = ProcessingJob.query.filter_by(UserId=current_user.Id)\
            .order_by(ProcessingJob.CreatedDate.desc()).all()

        return jsonify([{
            "id": job.Id,
            "title": job.JobName,
            "lastUpdate": job.CreatedDate.isoformat() if job.CreatedDate else None,
            "messageCount": ProcessingJobHistory.query.filter_by(ProcessJobId=job.Id).count(),
            "IsFavorite": job.IsFavorite
        } for job in jobs]), 200
    except Exception as e:
        logger.error(f"Error in get_conversations: {e}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)}), 500

@process_bp.route('/favorite/<int:job_id>', methods=['POST'])
@require_auth
def toggle_favorite(current_user, job_id):
    job = ProcessingJob.query.filter_by(
        Id=job_id, 
        UserId=current_user.Id
    ).first_or_404()
    
    job.IsFavorite = not job.IsFavorite   # ✅ FIXED
    
    db.session.commit()
    
    return jsonify({
        'success': True, 
        'isFavorite': job.IsFavorite
    })

# ==============================
# EDIT & RE-PROCESS TEXT
# ==============================
@process_bp.route('/edit-text/<int:history_id>', methods=['POST'])
@require_auth
def edit_text(current_user, history_id):
    try:
        start_time = time.time()
        data = request.json
        
        # 1. Fetch the existing history record first so we have defaults
        original_history = ProcessingJobHistory.query.get_or_404(history_id)
        job = ProcessingJob.query.get(original_history.ProcessJobId)

        if job.UserId != current_user.Id:
            return jsonify({'success': False, 'error': 'Unauthorized'}), 403

        # 2. Now extract new settings if provided, else keep originals
        new_text = data.get('text', '')
        font_family = data.get('fontFamily', original_history.FontFamily)
        font_size = int(data.get('fontSize', original_history.FontSize))
        include_cover = _coerce_bool(data.get('includeCover'), original_history.IncludeCover)
        include_toc = _coerce_bool(data.get('includeTOC'), original_history.IncludeTOC)

        # 3. Re-run the DocumentProcessor with new settings
        processor = DocumentProcessor(
            template_path=TEMPLATE_DOCX,
            font_family=font_family,
            font_size=font_size,
            include_cover=include_cover,
            include_toc=include_toc
        )

        doc_obj = processor.html_to_docx(new_text)

        if include_toc:
            with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_doc:
                temp_docx_path = tmp_doc.name
            try:
                doc_obj.save(temp_docx_path)
                processor.refresh_saved_docx(temp_docx_path)
                with open(temp_docx_path, 'rb') as f:
                    final_file_data = f.read()
            finally:
                if os.path.exists(temp_docx_path):
                    os.remove(temp_docx_path)
        else:
            output_buffer = BytesIO()
            doc_obj.save(output_buffer)
            output_buffer.seek(0)
            final_file_data = output_buffer.read()

        # 3. Create a NEW history entry (to keep the "chat" thread intact)
        new_count = original_history.ProcessingCount + 1
        
        # 🟢 INCREMENT ORIGINAL History Item's Count as well
        original_history.ProcessingCount = new_count
        original_history.ModifiedBy = current_user.Id
        db.session.add(original_history)
        
        # Generate a descriptive filename based on the edited text
        file_base_name = get_first_sentence(new_text)
        new_filename = generate_unique_filename(file_base_name, current_user)
        processing_time = round(time.time() - start_time, 2)

        stored_output = file_storage_manager.save_file(
            final_file_data, new_filename, current_user.Id, job.Id
        )

        new_history = ProcessingJobHistory(
            ProcessJobId=job.Id,
            JobType=JobType.TEXT,
            ProcessingTime=processing_time,
            UploadFileData=new_text.encode('utf-8'),
            Status=JobStatus.SUCCESS,
            OutputFileData=stored_output['db_bytes'],
            OutputFileServerPath=stored_output['server_path'],
            OutputFileName=new_filename,
            FontFamily=font_family,
            FontSize=font_size,
            IncludeCover=include_cover,
            IncludeTOC=include_toc,
            ProcessingCount=new_count,
            CreatedBy=current_user.Id,
            ModifiedBy=current_user.Id
        )

        db.session.add(new_history)
        db.session.commit()

        processing_time = round(time.time() - start_time, 2)

        return jsonify({
            'success': True, 
            'message': 'Text updated and re-processed',
            'historyId': new_history.Id,
            'processingTime': processing_time,
            'time': datetime.now().strftime('%d %b %Y, %I:%M %p'),
            'processingCount': new_count
        })

    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'error': str(e)}), 500

# ==============================
# LOAD CHAT
# ==============================
@process_bp.route('/conversation/<int:job_id>', methods=['GET'])
@require_auth
def get_conversation(current_user, job_id):
    """Load full conversation history"""
    try:
        job = ProcessingJob.query.filter_by(Id=job_id, UserId=current_user.Id).first()
        if not job:
            return jsonify({'error': 'Unauthorized or not found'}), 404
        
        histories = ProcessingJobHistory.query.filter_by(ProcessJobId=job_id)\
            .order_by(ProcessingJobHistory.CreatedDate.asc()).all()
        
        # ✅ Include file data for rendering
        return jsonify([h.to_dict(include_file_data=True) for h in histories]), 200
        
    except Exception as e:
        logger.error(f"Error in get_conversation: {e}", exc_info=True)
        return jsonify({'error': str(e)}), 500

# ==============================
# DOWNLOAD
# ==============================
@process_bp.route('/download/<int:history_id>', methods=['GET'])
@require_auth
def download_file(current_user, history_id):
    try:
        history = ProcessingJobHistory.query.get_or_404(history_id)
        job = ProcessingJob.query.get(history.ProcessJobId)
        if job.UserId != current_user.Id:
            return jsonify({'success': False, 'error': 'Unauthorized'}), 403

        try:
            file_buffer = file_storage_manager.retrieve_file(
                history.OutputFileServerPath,
                history.OutputFileData,
                history.OutputFileName
            )
        except FileNotFoundError as e:
            return jsonify({'success': False, 'error': str(e)}), 404

        requested_format = request.args.get('format', 'docx').lower()

        if requested_format == 'pdf':
            try:
                pdf_data = convert_docx_to_pdf(history.OutputFileData)

                return send_file(
                    BytesIO(pdf_data),
                    mimetype='application/pdf',
                    as_attachment=True,
                    download_name=history.OutputFileName.replace('.docx', '.pdf')
                )

            except Exception as e:
                return jsonify({'success': False, 'error': str(e)}), 500

        else:
            return send_file(
                file_buffer,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                as_attachment=True,
                download_name=history.OutputFileName
            )

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': f'Download failed: {str(e)}'}), 500
