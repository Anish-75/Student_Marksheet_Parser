"""
Sample Marksheet Template Generator
Creates a sample Excel file showing the expected structure
"""
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from config import COURSE_CODES, COURSE_COLUMNS, STUDENT_COLUMNS, SUMMARY_COLUMNS


def create_sample_template():
    """Generate a sample marksheet template"""
    
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Marksheet"
    
    # Header styling
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
    
    # Row 1-3: Header space (you can add university name, semester info, etc.)
    sheet['A1'] = 'UNIVERSITY NAME'
    sheet['A1'].font = Font(size=14, bold=True)
    sheet['A2'] = 'SEMESTER MARKSHEET - SAMPLE TEMPLATE'
    sheet['A2'].font = Font(size=12, bold=True)
    
    # Row 4: Course codes (index 3)
    row = 4
    for course_code in COURSE_CODES:
        cols = COURSE_COLUMNS[course_code]
        # Place course code at Credits Assigned column
        col_idx = cols['credits_assigned']
        sheet.cell(row=row, column=col_idx + 1, value=course_code)
        sheet.cell(row=row, column=col_idx + 1).font = header_font
        sheet.cell(row=row, column=col_idx + 1).fill = header_fill
        sheet.cell(row=row, column=col_idx + 1).alignment = Alignment(horizontal='center')
    
    # Row 5: Course names (index 4)
    row = 5
    course_names = {
        'CMP-100': 'Computer Programming',
        'CMP-101': 'Data Structures',
        'MCV-111': 'Calculus I',
        'MCV-112': 'Calculus II',
        'SHM-132': 'English Communication',
        'SHM-133': 'Professional Ethics',
        'AEC-153': 'Environmental Studies',
        'VAC-158': 'Physical Education',
        'VAC-159': 'NSS/NCC',
        'SEC-143': 'Skill Enhancement'
    }
    
    for course_code in COURSE_CODES:
        cols = COURSE_COLUMNS[course_code]
        col_idx = cols['credits_assigned']
        course_name = course_names.get(course_code, course_code)
        sheet.cell(row=row, column=col_idx + 1, value=course_name)
        sheet.cell(row=row, column=col_idx + 1).font = Font(bold=True)
    
    # Row 6-10: Sub-headers for each course
    row = 6
    sheet.cell(row=row, column=STUDENT_COLUMNS['pr_number'] + 1, value='PR Number')
    sheet.cell(row=row, column=STUDENT_COLUMNS['name'] + 1, value='Student Name')
    sheet.cell(row=row, column=STUDENT_COLUMNS['seat_number'] + 1, value='Seat Number')
    
    for course_code in COURSE_CODES:
        cols = COURSE_COLUMNS[course_code]
        sheet.cell(row=row, column=cols['credits_assigned'] + 1, value='Cr.Asgn')
        sheet.cell(row=row, column=cols['credits_earned'] + 1, value='Cr.Earn')
        sheet.cell(row=row, column=cols['grade_point'] + 1, value='GP')
        sheet.cell(row=row, column=cols['letter_grade'] + 1, value='Grade')
    
    sheet.cell(row=row, column=SUMMARY_COLUMNS['total_credits_assigned'] + 1, value='Tot.Cr.Asgn')
    sheet.cell(row=row, column=SUMMARY_COLUMNS['total_credits_earned'] + 1, value='Tot.Cr.Earn')
    sheet.cell(row=row, column=SUMMARY_COLUMNS['overall_letter_grade'] + 1, value='Overall Grade')
    
    # Make row 6 bold
    for cell in sheet[row]:
        if cell.value:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
    
    # Row 12+: Sample student data
    students_sample = [
        {
            'pr_number': '2023001',
            'name': 'John Doe',
            'seat_number': 'S001',
            'grades': {'CMP-100': 4, 'CMP-101': 4, 'MCV-111': 3, 'MCV-112': 3, 
                      'SHM-132': 4, 'SHM-133': 3, 'AEC-153': 3, 'VAC-158': 4, 
                      'VAC-159': 4, 'SEC-143': 3}
        },
        {
            'pr_number': '2023002',
            'name': 'Jane Smith',
            'seat_number': 'S002',
            'grades': {'CMP-100': 4, 'CMP-101': 3, 'MCV-111': 4, 'MCV-112': 3, 
                      'SHM-132': 4, 'SHM-133': 4, 'AEC-153': 3, 'VAC-158': 4, 
                      'VAC-159': 3, 'SEC-143': 4}
        },
        {
            'pr_number': '2023003',
            'name': 'Mike Johnson',
            'seat_number': 'S003',
            'grades': {'CMP-100': 3, 'CMP-101': 3, 'MCV-111': 3, 'MCV-112': 3, 
                      'SHM-132': 3, 'SHM-133': 3, 'AEC-153': 3, 'VAC-158': 3, 
                      'VAC-159': 3, 'SEC-143': 3}
        }
    ]
    
    grade_to_letter = {4: 'A', 3: 'B', 2: 'C', 1: 'D', 0: 'F'}
    
    row = 12
    for student in students_sample:
        # Student basic info
        sheet.cell(row=row, column=STUDENT_COLUMNS['pr_number'] + 1, value=student['pr_number'])
        sheet.cell(row=row, column=STUDENT_COLUMNS['name'] + 1, value=student['name'])
        sheet.cell(row=row, column=STUDENT_COLUMNS['seat_number'] + 1, value=student['seat_number'])
        
        # Course grades
        total_credits_assigned = 0
        total_credits_earned = 0
        total_grade_points = 0
        
        for course_code in COURSE_CODES:
            cols = COURSE_COLUMNS[course_code]
            grade_point = student['grades'].get(course_code, 0)
            credits = 4 if course_code in ['CMP-100', 'CMP-101', 'MCV-111', 'MCV-112'] else 2
            
            sheet.cell(row=row, column=cols['credits_assigned'] + 1, value=credits)
            sheet.cell(row=row, column=cols['credits_earned'] + 1, value=credits if grade_point > 0 else 0)
            sheet.cell(row=row, column=cols['grade_point'] + 1, value=grade_point)
            sheet.cell(row=row, column=cols['letter_grade'] + 1, value=grade_to_letter[grade_point])
            
            total_credits_assigned += credits
            total_credits_earned += credits if grade_point > 0 else 0
            total_grade_points += grade_point * credits
        
        # Summary
        sheet.cell(row=row, column=SUMMARY_COLUMNS['total_credits_assigned'] + 1, value=total_credits_assigned)
        sheet.cell(row=row, column=SUMMARY_COLUMNS['total_credits_earned'] + 1, value=total_credits_earned)
        
        # Calculate overall grade
        avg_gp = total_grade_points / total_credits_assigned if total_credits_assigned > 0 else 0
        overall_grade = 'A' if avg_gp >= 3.5 else 'B' if avg_gp >= 2.5 else 'C' if avg_gp >= 1.5 else 'D'
        sheet.cell(row=row, column=SUMMARY_COLUMNS['overall_letter_grade'] + 1, value=overall_grade)
        
        row += 1
    
    # Adjust column widths
    sheet.column_dimensions['A'].width = 5
    sheet.column_dimensions['C'].width = 12  # PR Number
    sheet.column_dimensions['E'].width = 25  # Name
    sheet.column_dimensions['G'].width = 12  # Seat Number
    
    # Save the template
    filename = 'sample_marksheet_template.xlsx'
    wb.save(filename)
    print(f"✓ Sample template created: {filename}")
    print("\nThis template shows the expected structure:")
    print("  - Row 4: Course codes")
    print("  - Row 5: Course names")
    print("  - Row 6: Column headers")
    print("  - Row 12+: Student data (3 sample students included)")
    print("\nYou can use this as a reference or fill it with your actual data.")


if __name__ == "__main__":
    create_sample_template()
    print("\n" + "="*60)
    print("Creating sample OUTPUT format for reference...")
    print("="*60)
    
    # Create sample output
    wb_out = Workbook()
    sheet_out = wb_out.active
    sheet_out.title = "Grade Report"
    
    info_font = Font(name='Arial', size=11, bold=True)
    normal_font = Font(name='Arial', size=11)
    header_font = Font(name='Arial', size=11, bold=True)
    
    from openpyxl.styles import Border, Side
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Student Information
    sheet_out['A7'] = 'NAME OF CANDIDATE:'
    sheet_out['A7'].font = info_font
    sheet_out['B7'] = 'SAMPLE STUDENT NAME'
    
    sheet_out['A8'] = 'SEAT NUMBER:'
    sheet_out['A8'].font = info_font
    sheet_out['B8'] = 'S001'
    
    sheet_out['A9'] = 'P.R. NUMBER:'
    sheet_out['A9'].font = info_font
    sheet_out['B9'] = '2023001'
    
    sheet_out['A10'] = 'DEPARTMENT:'
    sheet_out['A10'].font = info_font
    sheet_out['B10'] = 'ALL DEPARTMENT'
    
    sheet_out['A11'] = 'COLLEGE:'
    sheet_out['A11'].font = info_font
    sheet_out['B11'] = 'SHREE RAYESHWAR INSTITUTE OF ENGINEERING AND INFORMATION TECHNOLOGY'
    
    sheet_out['A12'] = 'EXAMINATION:'
    sheet_out['A12'].font = info_font
    sheet_out['B12'] = 'NOVEMBER/DECEMBER 2024'
    
    sheet_out['A13'] = 'SCHEME:'
    sheet_out['A13'].font = info_font
    sheet_out['B13'] = 'RC 2024-25'
    
    sheet_out['A14'] = 'SEMESTER I'
    sheet_out['A14'].font = Font(name='Arial', size=11, bold=True)
    
    # Table Headers
    headers = {
        'A16': 'Course Code',
        'B16': 'Nomenclature',
        'G16': 'Credits Assigned',
        'H16': 'Credits Earned',
        'I16': 'Grade Point',
        'J16': 'Letter Grade'
    }
    
    for cell, value in headers.items():
        sheet_out[cell] = value
        sheet_out[cell].font = header_font
        sheet_out[cell].border = border
        sheet_out[cell].alignment = Alignment(horizontal='center', vertical='center')
    
    # Sample data
    row_out = 17
    sample_courses = [
        ['CMP-100', 'Computer Programming', 4, 4, 4, 'A'],
        ['CMP-101', 'Data Structures', 4, 4, 3, 'B'],
        ['MCV-111', 'Calculus I', 3, 3, 4, 'A']
    ]
    
    for course in sample_courses:
        sheet_out.cell(row=row_out, column=1, value=course[0]).border = border
        sheet_out.cell(row=row_out, column=2, value=course[1]).border = border
        sheet_out.cell(row=row_out, column=7, value=course[2]).border = border
        sheet_out.cell(row=row_out, column=8, value=course[3]).border = border
        sheet_out.cell(row=row_out, column=9, value=course[4]).border = border
        sheet_out.cell(row=row_out, column=10, value=course[5]).border = border
        row_out += 1
    
    # Total row
    row_out += 1
    sheet_out.cell(row=row_out, column=1, value='TOTAL').font = Font(bold=True)
    sheet_out.cell(row=row_out, column=1).border = border
    sheet_out.cell(row=row_out, column=7, value=11).font = Font(bold=True)
    sheet_out.cell(row=row_out, column=7).border = border
    sheet_out.cell(row=row_out, column=8, value=11).font = Font(bold=True)
    sheet_out.cell(row=row_out, column=8).border = border
    sheet_out.cell(row=row_out, column=10, value='A').font = Font(bold=True)
    sheet_out.cell(row=row_out, column=10).border = border
    
    # Column widths
    sheet_out.column_dimensions['A'].width = 15
    sheet_out.column_dimensions['B'].width = 40
    sheet_out.column_dimensions['G'].width = 16
    sheet_out.column_dimensions['H'].width = 16
    sheet_out.column_dimensions['I'].width = 14
    sheet_out.column_dimensions['J'].width = 14
    
    wb_out.save('sample_output_format.xlsx')
    print("✓ Sample OUTPUT format created: sample_output_format.xlsx")
    print("\nThis shows how the individual student reports will look.")