"""
Test script for parsing functionality
Run this to test the parser with a sample Excel file
"""
from parser import parse_marksheet, create_individual_report, validate_marksheet_structure
import os
import json


def test_parsing():
    """Test the parsing functionality"""
    
    print("=" * 60)
    print("STUDENT MARKSHEET PARSER - TEST SCRIPT")
    print("=" * 60)
    
    # Check if sample file exists
    sample_files = [
        'sample_marksheet.xlsx',
        'uploads/sample_marksheet.xlsx',
        '../sample_marksheet.xlsx'
    ]
    
    input_file = None
    for sample_file in sample_files:
        if os.path.exists(sample_file):
            input_file = sample_file
            break
    
    if not input_file:
        print("\n❌ ERROR: Sample file not found!")
        print("\nPlease create a sample Excel file with the following structure:")
        print("  - Row 4: Course codes")
        print("  - Row 5: Course names")
        print("  - Row 12+: Student data")
        print("\nExpected columns:")
        print("  - Column C: PR Number")
        print("  - Column E: Student Name")
        print("  - Column G: Seat Number")
        print("  - See config.py for course column mappings")
        print("\nSave it as 'sample_marksheet.xlsx' in the current directory.")
        return
    
    print(f"\n✓ Found sample file: {input_file}")
    
    # Validate the file structure
    print("\n" + "-" * 60)
    print("STEP 1: Validating file structure...")
    print("-" * 60)
    
    try:
        validate_marksheet_structure(input_file)
        print("✓ File structure is valid")
    except Exception as e:
        print(f"❌ Validation failed: {str(e)}")
        return
    
    # Parse the marksheet
    print("\n" + "-" * 60)
    print("STEP 2: Parsing marksheet...")
    print("-" * 60)
    
    try:
        students = parse_marksheet(input_file)
        print(f"✓ Successfully parsed marksheet")
        print(f"✓ Found {len(students)} student(s)")
    except Exception as e:
        print(f"❌ Parsing failed: {str(e)}")
        import traceback
        print(traceback.format_exc())
        return
    
    if not students:
        print("❌ No student data found in the file")
        return
    
    # Display first student data
    print("\n" + "-" * 60)
    print("STEP 3: Sample student data")
    print("-" * 60)
    
    first_student = students[0]
    print(f"\nPR Number: {first_student['pr_number']}")
    print(f"Name: {first_student['name']}")
    print(f"Seat Number: {first_student['seat_number']}")
    print(f"\nCourses enrolled: {len(first_student['courses'])}")
    
    print("\nCourse Details:")
    for course in first_student['courses']:
        print(f"  - {course['course_code']}: {course['course_name']}")
        print(f"    Credits: {course['credits_assigned']} | Grade: {course['letter_grade']} | GP: {course['grade_point']}")
    
    print(f"\nTotal Credits Assigned: {first_student['total_credits_assigned']}")
    print(f"Total Credits Earned: {first_student['total_credits_earned']}")
    print(f"Overall Grade: {first_student['overall_letter_grade']}")
    
    # Create individual reports
    print("\n" + "-" * 60)
    print("STEP 4: Generating individual reports...")
    print("-" * 60)
    
    output_folder = 'test_outputs'
    os.makedirs(output_folder, exist_ok=True)
    
    generated_files = []
    for i, student in enumerate(students, 1):
        safe_pr = student['pr_number'].replace('/', '_').replace('\\', '_')
        output_filename = f"Grade_Report_{safe_pr}.xlsx"
        output_path = os.path.join(output_folder, output_filename)
        
        try:
            create_individual_report(student, output_path)
            generated_files.append(output_filename)
            print(f"  ✓ Generated: {output_filename}")
        except Exception as e:
            print(f"  ❌ Failed to generate report for {student['pr_number']}: {str(e)}")
    
    # Summary
    print("\n" + "=" * 60)
    print("TEST SUMMARY")
    print("=" * 60)
    print(f"✓ Students processed: {len(students)}")
    print(f"✓ Reports generated: {len(generated_files)}")
    print(f"✓ Output folder: {os.path.abspath(output_folder)}")
    
    # Save parsed data as JSON for inspection
    json_output = os.path.join(output_folder, 'parsed_data.json')
    with open(json_output, 'w', encoding='utf-8') as f:
        json.dump(students, f, indent=2, ensure_ascii=False)
    print(f"✓ JSON data saved: {json_output}")
    
    print("\n" + "=" * 60)
    print("✓ TEST COMPLETED SUCCESSFULLY!")
    print("=" * 60)
    print("\nNext steps:")
    print("1. Check the generated Excel files in the 'test_outputs' folder")
    print("2. Verify the data is correct")
    print("3. Run the API: python app.py")
    print("4. Test the API with curl or Postman")


if __name__ == "__main__":
    test_parsing()