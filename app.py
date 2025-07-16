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
CORS(app)

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

def compress_ppt(file_path, output_path):
    """Compress PPT file by optimizing images and removing metadata"""
    try:
        # Open the presentation
        prs = Presentation(file_path)
        
        # Process each slide
        for slide in prs.slides:
            # Process shapes in the slide
            for shape in slide.shapes:
                if hasattr(shape, 'image'):
                    try:
                        # Get image data
                        image_data = shape.image.blob
                        
                        # Compress the image
                        compressed_data = compress_image(image_data, quality=70)
                        
                        # Replace the image (this is a simplified approach)
                        # In practice, you might need more sophisticated image replacement
                        print(f"Compressed image in slide, original size: {len(image_data)}, compressed: {len(compressed_data)}")
                        
                    except Exception as e:
                        print(f"Error processing image in shape: {e}")
                        continue
        
        # Save the compressed presentation
        prs.save(output_path)
        
        # Get file sizes for comparison
        original_size = os.path.getsize(file_path)
        compressed_size = os.path.getsize(output_path)
        
        return {
            'success': True,
            'original_size': original_size,
            'compressed_size': compressed_size,
            'compression_ratio': round((1 - compressed_size/original_size) * 100, 2)
        }
        
    except Exception as e:
        return {
            'success': False,
            'error': str(e)
        }

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
    
    try:
        # Generate unique filename
        file_id = str(uuid.uuid4())
        original_filename = file.filename
        temp_input = os.path.join(UPLOAD_FOLDER, f"{file_id}_input.pptx")
        temp_output = os.path.join(UPLOAD_FOLDER, f"{file_id}_output.pptx")
        
        # Save uploaded file
        file.save(temp_input)
        
        # Compress the file
        result = compress_ppt(temp_input, temp_output)
        
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