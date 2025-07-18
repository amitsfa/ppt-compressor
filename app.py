from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import tempfile
import zipfile
from PIL import Image
import io
import uuid
from datetime import datetime, timedelta
import threading
import time
import shutil
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app, origins=["https://keen-creponne-d8adb1.netlify.app"])

# Configure maximum file size (100MB)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB in bytes

# Configure upload folder
UPLOAD_FOLDER = 'temp_files'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

class AdvancedPPTCompressor:
    """Enhanced PPT compressor with better compression ratios"""
    
    def __init__(self, target_size_mb: int = 40, min_quality: int = 20, max_quality: int = 85):
        self.target_size_mb = target_size_mb
        self.min_quality = min_quality
        self.max_quality = max_quality
        self.supported_formats = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.gif', '.webp'}
        
    def compress_pptx(self, input_path: str, output_path: str) -> Dict:
        """Advanced PPTX compression with adaptive quality"""
        try:
            original_size = os.path.getsize(input_path)
            logger.info(f"Starting compression: {original_size / (1024*1024):.2f} MB")
            
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)
                
                # Extract PPTX contents
                self._extract_pptx(input_path, temp_path)
                
                # Find media files
                media_files = self._find_media_files(temp_path)
                logger.info(f"Found {len(media_files)} media files")
                
                # Remove unnecessary metadata
                self._clean_metadata(temp_path)
                
                # Compress with adaptive quality
                result = self._compress_with_adaptive_quality(temp_path, media_files, output_path)
                
                # Add original size to result
                result['original_size'] = original_size
                result['original_size_mb'] = original_size / (1024*1024)
                
                return result
                
        except Exception as e:
            logger.error(f"Compression failed: {e}")
            return {'success': False, 'error': str(e)}
    
    def _extract_pptx(self, pptx_path: str, extract_path: Path):
        """Extract PPTX file contents"""
        with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
            zip_ref.extractall(extract_path)
    
    def _find_media_files(self, extract_path: Path) -> List[Path]:
        """Find all media files in extracted PPTX"""
        media_files = []
        
        # Check all possible media locations
        media_dirs = [
            extract_path / 'ppt' / 'media',
            extract_path / 'word' / 'media',
            extract_path / 'xl' / 'media'
        ]
        
        for media_dir in media_dirs:
            if media_dir.exists():
                for file_path in media_dir.rglob('*'):
                    if file_path.suffix.lower() in self.supported_formats:
                        media_files.append(file_path)
        
        return media_files
    
    def _clean_metadata(self, extract_path: Path):
        """Remove metadata to reduce file size"""
        removable_items = [
            'docProps/app.xml',
            'docProps/core.xml', 
            'docProps/custom.xml',
            'customXml'
        ]
        
        for item in removable_items:
            item_path = extract_path / item
            try:
                if item_path.exists():
                    if item_path.is_file():
                        item_path.unlink()
                    else:
                        shutil.rmtree(item_path)
                    logger.info(f"Removed metadata: {item}")
            except Exception as e:
                logger.warning(f"Could not remove {item}: {e}")
    
    def _compress_image(self, image_path: Path, quality: int) -> bool:
        """Compress individual image with smart format conversion"""
        try:
            with Image.open(image_path) as img:
                original_format = img.format
                
                # Handle different image modes
                if img.mode in ('RGBA', 'LA'):
                    # Create white background for transparency
                    background = Image.new('RGB', img.size, (255, 255, 255))
                    if img.mode == 'RGBA':
                        background.paste(img, mask=img.split()[-1])
                    else:
                        background.paste(img, mask=img.split()[-1])
                    img = background
                elif img.mode == 'P':
                    # Convert palette mode
                    img = img.convert('RGBA')
                    background = Image.new('RGB', img.size, (255, 255, 255))
                    background.paste(img, mask=img.split()[-1])
                    img = background
                elif img.mode not in ('RGB', 'L'):
                    img = img.convert('RGB')
                
                # Determine optimal format and compression
                if image_path.suffix.lower() == '.png' and img.mode == 'RGB':
                    # Convert PNG to JPEG if no transparency
                    jpeg_path = image_path.with_suffix('.jpg')
                    img.save(jpeg_path, 'JPEG', quality=quality, optimize=True, progressive=True)
                    image_path.unlink()  # Remove original PNG
                    return True
                elif image_path.suffix.lower() in ['.jpg', '.jpeg']:
                    # Optimize JPEG
                    img.save(image_path, 'JPEG', quality=quality, optimize=True, progressive=True)
                    return True
                else:
                    # Convert other formats to JPEG
                    jpeg_path = image_path.with_suffix('.jpg')
                    img.save(jpeg_path, 'JPEG', quality=quality, optimize=True, progressive=True)
                    image_path.unlink()  # Remove original
                    return True
                    
        except Exception as e:
            logger.error(f"Failed to compress {image_path}: {e}")
            return False
    
    def _create_compressed_pptx(self, extract_path: Path, output_path: str) -> int:
        """Create compressed PPTX from extracted contents"""
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as zipf:
            for file_path in extract_path.rglob('*'):
                if file_path.is_file():
                    rel_path = file_path.relative_to(extract_path)
                    zipf.write(file_path, rel_path)
        
        return os.path.getsize(output_path)
    
    def _compress_with_adaptive_quality(self, extract_path: Path, media_files: List[Path], output_path: str) -> Dict:
        """Compress with adaptive quality control"""
        target_size = self.target_size_mb * 1024 * 1024
        
        # Store original media data for restoration
        original_media_data = {}
        for media_file in media_files:
            original_media_data[media_file] = media_file.read_bytes()
        
        # Try different quality levels
        best_result = None
        
        for quality in range(self.max_quality, self.min_quality - 1, -10):
            logger.info(f"Trying quality: {quality}")
            
            # Restore original media files
            for media_file, original_data in original_media_data.items():
                media_file.write_bytes(original_data)
            
            # Compress all media files
            successful_compressions = 0
            for media_file in media_files:
                if self._compress_image(media_file, quality):
                    successful_compressions += 1
            
            # Create test PPTX
            test_output = output_path + '.tmp'
            compressed_size = self._create_compressed_pptx(extract_path, test_output)
            
            # Calculate compression ratio
            original_total = sum(len(data) for data in original_media_data.values())
            compression_ratio = round((1 - compressed_size / max(original_total, 1)) * 100, 2)
            
            result = {
                'success': True,
                'compressed_size': compressed_size,
                'compressed_size_mb': compressed_size / (1024*1024),
                'compression_ratio': compression_ratio,
                'quality_used': quality,
                'images_processed': successful_compressions,
                'target_achieved': compressed_size <= target_size
            }
            
            # If target achieved, use this result
            if result['target_achieved']:
                shutil.move(test_output, output_path)
                logger.info(f"Target achieved with quality {quality}")
                return result
            
            # Keep track of best result so far
            if best_result is None or compressed_size < best_result['compressed_size']:
                best_result = result.copy()
                if os.path.exists(output_path):
                    os.remove(output_path)
                shutil.move(test_output, output_path)
            else:
                os.remove(test_output)
        
        # Return best result achieved
        logger.warning(f"Could not reach target size, best compression: {best_result['compressed_size_mb']:.2f}MB")
        return best_result or {'success': False, 'error': 'Compression failed'}

# Global compressor instance
compressor = AdvancedPPTCompressor()

# Clean up old files every hour
def cleanup_old_files():
    while True:
        try:
            cutoff_time = datetime.now() - timedelta(hours=1)
            for filename in os.listdir(UPLOAD_FOLDER):
                filepath = os.path.join(UPLOAD_FOLDER, filename)
                if os.path.isfile(filepath):
                    file_time = datetime.fromtimestamp(os.path.getctime(filepath))
                    if file_time < cutoff_time:
                        os.remove(filepath)
                        logger.info(f"Cleaned up old file: {filename}")
        except Exception as e:
            logger.error(f"Cleanup error: {e}")
        time.sleep(3600)  # Wait 1 hour

# Start cleanup thread
cleanup_thread = threading.Thread(target=cleanup_old_files, daemon=True)
cleanup_thread.start()

@app.route('/')
def index():
    return jsonify({
        'message': 'Enhanced PPT Compressor API',
        'version': '2.0',
        'features': [
            'Advanced image compression with adaptive quality',
            'Smart format conversion (PNG → JPEG when possible)',
            'Metadata removal for size reduction',
            'Progressive JPEG optimization',
            'Automatic quality adjustment to meet target size'
        ],
        'endpoints': {
            'upload': '/upload - POST',
            'download': '/download/<file_id> - GET',
            'status': '/status - GET'
        }
    })

@app.route('/status')
def status():
    """API status endpoint"""
    temp_files = len([f for f in os.listdir(UPLOAD_FOLDER) if os.path.isfile(os.path.join(UPLOAD_FOLDER, f))])
    return jsonify({
        'status': 'running',
        'temp_files': temp_files,
        'max_file_size_mb': app.config['MAX_CONTENT_LENGTH'] / (1024*1024),
        'target_size_mb': compressor.target_size_mb
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
    
    # Check file size
    if request.content_length and request.content_length > app.config['MAX_CONTENT_LENGTH']:
        return jsonify({'error': 'File size must be less than 100MB'}), 400
    
    try:
        # Generate unique filename
        file_id = str(uuid.uuid4())
        original_filename = file.filename
        temp_input = os.path.join(UPLOAD_FOLDER, f"{file_id}_input.pptx")
        temp_output = os.path.join(UPLOAD_FOLDER, f"{file_id}_output.pptx")
        
        # Save uploaded file
        file.save(temp_input)
        
        # Get target size from request (default 40MB)
        target_size = request.form.get('target_size', 40, type=int)
        compressor.target_size_mb = target_size
        
        logger.info(f"Processing file: {original_filename}, Target: {target_size}MB")
        
        # Compress the file using advanced compressor
        result = compressor.compress_pptx(temp_input, temp_output)
        
        # Clean up input file
        try:
            os.remove(temp_input)
        except:
            pass
        
        if result['success']:
            # Calculate actual compression ratio based on original file
            actual_compression = round((1 - result['compressed_size'] / result['original_size']) * 100, 2)
            
            return jsonify({
                'success': True,
                'file_id': file_id,
                'original_filename': original_filename,
                'original_size': result['original_size'],
                'original_size_mb': result['original_size_mb'],
                'compressed_size': result['compressed_size'],
                'compressed_size_mb': result['compressed_size_mb'],
                'compression_ratio': actual_compression,
                'quality_used': result.get('quality_used', 'N/A'),
                'images_processed': result.get('images_processed', 0),
                'target_achieved': result.get('target_achieved', False),
                'download_url': f'/download/{file_id}'
            })
        else:
            return jsonify({
                'success': False,
                'error': result.get('error', 'Unknown compression error')
            }), 500
            
    except Exception as e:
        logger.error(f"Upload processing failed: {e}")
        return jsonify({'error': f'Processing failed: {str(e)}'}), 500

@app.route('/download/<file_id>')
def download_file(file_id):
    try:
        output_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_output.pptx")
        
        if not os.path.exists(output_path):
            return jsonify({'error': 'File not found or expired'}), 404
        
        # Get file size for logging
        file_size = os.path.getsize(output_path)
        logger.info(f"Downloading file: {file_id}, Size: {file_size / (1024*1024):.2f}MB")
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name=f'compressed_{file_id}.pptx',
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        
    except Exception as e:
        logger.error(f"Download failed: {e}")
        return jsonify({'error': f'Download failed: {str(e)}'}), 500

@app.errorhandler(413)
def file_too_large(e):
    return jsonify({'error': 'File size exceeds 100MB limit'}), 413

@app.errorhandler(Exception)
def handle_exception(e):
    logger.error(f"Unhandled exception: {e}")
    return jsonify({'error': 'Internal server error'}), 500

if __name__ == '__main__':
    logger.info("Starting Enhanced PPT Compressor API...")
    app.run(debug=True, host='0.0.0.0', port=5000)