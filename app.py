from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import tempfile
import zipfile
from pptx import Presentation
from PIL import Image
import io
import uuid
from datetime import datetime, timedelta
import threading
import time

app = Flask(__name__)
CORS(app, origins=["https://keen-creponne-d8adb1.netlify.app"])

# Configure maximum file size (100MB)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB in bytes

# Configure upload folder
UPLOAD_FOLDER = 'temp_files'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Clean up old files every hour
def cleanup_old_files():
    while True:
        try:
            cutoff_time = datetime.now() - timedelta(hours=1)
            for filename in os.listdir(UPLOAD_FOLDER):
                filepath = os.path.join(UPLOAD_FOLDER, filename)
                file_time = datetime.fromtimestamp(os.path.getctime(filepath))
                if file_time < cutoff_time:
                    os.remove(filepath)
        except Exception as e:
            print(f"Cleanup error: {e}")
        time.sleep(3600)  # Wait 1 hour

# Start cleanup thread
cleanup_thread = threading.Thread(target=cleanup_old_files, daemon=True)
cleanup_thread.start()

def compress_image(image_data, quality=60):
    """Compress image data"""
    try:
        image = Image.open(io.BytesIO(image_data))
        
        # Convert to RGB if necessary
        if image.mode in ('RGBA', 'LA', 'P'):
            rgb_image = Image.new('RGB', image.size, (255, 255, 255))
            if image.mode == 'P':
                image = image.convert('RGBA')
            rgb_image.paste(image, mask=image.split()[-1] if image.mode == 'RGBA' else None)
            image = rgb_image
        
        # Compress the image
        output = io.BytesIO()
        image.save(output, format='JPEG', quality=quality, optimize=True)
        return output.getvalue()
    except Exception as e:
        print(f"Image compression error: {e}")
        return image_data

def compress_ppt(input_path, output_path, quality=70, target_size=40 * 1024 * 1024):
    """Compress PPTX by optimizing images and re-zipping contents"""

    import shutil

    temp_dir = tempfile.mkdtemp()
    try:
        # 1. Extract pptx zip
        with zipfile.ZipFile(input_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        media_folder = os.path.join(temp_dir, 'ppt', 'media')
        if os.path.exists(media_folder):
            for fname in os.listdir(media_folder):
                path = os.path.join(media_folder, fname)
                if not fname.lower().endswith(('.jpg', '.jpeg', '.png')):
                    continue
                try:
                    img = Image.open(path)

                    # convert transparency to white background
                    if img.mode in ('RGBA', 'LA', 'P'):
                        bg = Image.new('RGB', img.size, (255, 255, 255))
                        if img.mode == 'P':
                            img = img.convert('RGBA')
                        bg.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
                        img = bg

                    # compress to JPEG
                    buf = io.BytesIO()
                    img.save(buf, format='JPEG', quality=quality, optimize=True)

                    with open(path, 'wb') as f:
                        f.write(buf.getvalue())
                except Exception as e:
                    print(f"[Image Skipped] {fname}: {e}")
                    continue

        # 2. Rezip the folder into a new .pptx
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for foldername, _, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    archive_name = os.path.relpath(file_path, temp_dir)
                    zipf.write(file_path, arcname=archive_name)

        original_size = os.path.getsize(input_path)
        compressed_size = os.path.getsize(output_path)

        # 3. Retry with lower quality if still too big
        if target_size and compressed_size > target_size and quality > 30:
            print(f"Retrying compression: current size {compressed_size//1e6:.2f}MB, reducing quality")
            return compress_ppt(input_path, output_path, quality=quality - 10, target_size=target_size)

        return {
            'success': True,
            'original_size': original_size,
            'compressed_size': compressed_size,
            'compression_ratio': round((1 - compressed_size / original_size) * 100, 2)
        }

    except Exception as e:
        return {'success': False, 'error': str(e)}
    finally:
        shutil.rmtree(temp_dir)


@app.route('/')
def index():
    return jsonify({
        'message': 'PPT Compressor API',
        'endpoints': {
            'upload': '/upload - POST',
            'download': '/download/<file_id> - GET'
        }
    })

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not file.filename.lower().endswith(('.ppt', '.pptx')):
        return jsonify({'error': 'Please upload a PPT or PPTX file'}), 400
    
    # Check file size (100MB limit)
    if request.content_length and request.content_length > 100 * 1024 * 1024:
        return jsonify({'error': 'File size must be less than 100MB'}), 400
    
    try:
        # Generate unique filename
        file_id = str(uuid.uuid4())
        original_filename = file.filename
        temp_input = os.path.join(UPLOAD_FOLDER, f"{file_id}_input.pptx")
        temp_output = os.path.join(UPLOAD_FOLDER, f"{file_id}_output.pptx")
        
        # Set longer timeout for large files
        app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 300  # 5 minutes
        
        # Save uploaded file
        file.save(temp_input)
        
        # Compress the file
        # compress with target under 40MB
        result = compress_ppt(temp_input, temp_output, quality=70, target_size=40 * 1024 * 1024)

        
        if result['success']:
            return jsonify({
                'success': True,
                'file_id': file_id,
                'original_filename': original_filename,
                'original_size': result['original_size'],
                'compressed_size': result['compressed_size'],
                'compression_ratio': result['compression_ratio'],
                'download_url': f'/download/{file_id}'
            })
        else:
            return jsonify({
                'success': False,
                'error': result['error']
            }), 500
            
    except Exception as e:
        return jsonify({'error': f'Processing failed: {str(e)}'}), 500

@app.route('/download/<file_id>')
def download_file(file_id):
    try:
        output_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_output.pptx")
        
        if not os.path.exists(output_path):
            return jsonify({'error': 'File not found or expired'}), 404
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name=f'compressed_{file_id}.pptx',
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        
    except Exception as e:
        return jsonify({'error': f'Download failed: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)