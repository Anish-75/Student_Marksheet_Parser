"""
Flask API for Student Marksheet Parser
"""
from flask import Flask, request, send_file, jsonify
from werkzeug.utils import secure_filename
from parser import parse_marksheet, create_individual_report, validate_marksheet_structure
import os
from zipfile import ZipFile
from datetime import datetime
import traceback

app = Flask(__name__)

# Configuration
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size


def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def cleanup_old_files():
    """Clean up old uploaded and output files"""
    for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            try:
                if os.path.isfile(file_path):
                    # Delete files older than 1 hour
                    if os.path.getmtime(file_path) < datetime.now().timestamp() - 3600:
                        os.remove(file_path)
            except Exception as e:
                print(f"Error deleting {file_path}: {e}")


@app.route('/', methods=['GET'])
def home():
    """Home endpoint with API information"""
    return jsonify({
        "message": "Student Marksheet Parser API",
        "version": "1.0",
        "endpoints": {
            "health": "/health [GET]",
            "upload": "/upload [POST]",
            "download": "/download/<filename> [GET]"
        }
    })


@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({
        "status": "healthy", 
        "message": "API is running",
        "timestamp": datetime.now().isoformat()
    })


@app.route('/upload', methods=['POST'])
def upload_marksheet():
    """
    Upload marksheet and get individual student reports
    
    Expected: multipart/form-data with 'file' field containing Excel file
    
    Returns: JSON with processing results and download URL
    """
    # Clean up old files first
    cleanup_old_files()
    
    # Check if file is in request
    if 'file' not in request.files:
        return jsonify({"error": "No file provided in request"}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({"error": "No file selected"}), 400
    
    if not allowed_file(file.filename):
        return jsonify({
            "error": "Invalid file type. Only .xlsx and .xls files are allowed"
        }), 400
    
    try:
        # Save uploaded file
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        input_filename = f"{timestamp}_{filename}"
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
        file.save(input_path)
        
        # Validate file structure
        try:
            validate_marksheet_structure(input_path)
        except ValueError as ve:
            return jsonify({"error": f"Validation failed: {str(ve)}"}), 400
        
        # Parse the marksheet
        students = parse_marksheet(input_path)
        
        if not students:
            return jsonify({"error": "No student data found in file"}), 400
        
        # Create individual reports
        output_files = []
        student_list = []
        
        for student in students:
            # Create safe filename from PR number
            safe_pr = student['pr_number'].replace('/', '_').replace('\\', '_')
            output_filename = f"Grade_Report_{safe_pr}.xlsx"
            output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
            
            # Generate individual report
            create_individual_report(student, output_path)
            output_files.append(output_filename)
            
            # Add to student list for response
            student_list.append({
                "pr_number": student['pr_number'],
                "name": student['name'],
                "seat_number": student['seat_number'],
                "overall_grade": student['overall_letter_grade']
            })
        
        # Create a ZIP file with all reports
        zip_filename = f'All_Grade_Reports_{timestamp}.zip'
        zip_path = os.path.join(app.config['OUTPUT_FOLDER'], zip_filename)
        
        with ZipFile(zip_path, 'w') as zipf:
            for output_file in output_files:
                file_path = os.path.join(app.config['OUTPUT_FOLDER'], output_file)
                zipf.write(file_path, output_file)
        
        return jsonify({
            "success": True,
            "message": f"Successfully processed {len(students)} students",
            "students_count": len(students),
            "students": student_list,
            "download_url": f"/download/{zip_filename}",
            "timestamp": timestamp
        }), 200
        
    except Exception as e:
        error_trace = traceback.format_exc()
        print(f"Error processing file: {error_trace}")
        return jsonify({
            "error": "Failed to process marksheet",
            "details": str(e)
        }), 500


@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    """
    Download the generated reports ZIP file
    
    Args:
        filename: Name of the file to download
    """
    file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    
    if not os.path.exists(file_path):
        return jsonify({"error": "File not found"}), 404
    
    return send_file(
        file_path, 
        as_attachment=True,
        download_name=filename
    )


@app.route('/test', methods=['GET'])
def test_parsing():
    """
    Test endpoint to check if sample file exists and can be parsed
    """
    sample_file = 'sample_marksheet.xlsx'
    sample_path = os.path.join(UPLOAD_FOLDER, sample_file)
    
    if not os.path.exists(sample_path):
        return jsonify({
            "error": "Sample file not found",
            "message": "Please upload a sample_marksheet.xlsx file to the uploads folder"
        }), 404
    
    try:
        students = parse_marksheet(sample_path)
        return jsonify({
            "success": True,
            "students_found": len(students),
            "first_student": students[0] if students else None
        })
    except Exception as e:
        return jsonify({
            "error": "Failed to parse sample file",
            "details": str(e)
        }), 500


if __name__ == '__main__':
    print("=" * 50)
    print("Student Marksheet Parser API")
    print("=" * 50)
    print("Server starting on http://localhost:5000")
    print("\nAvailable endpoints:")
    print("  - GET  /health           - Health check")
    print("  - POST /upload           - Upload marksheet")
    print("  - GET  /download/<file>  - Download results")
    print("  - GET  /test             - Test parsing")
    print("=" * 50)
    app.run(debug=True, port=5000, host='0.0.0.0')