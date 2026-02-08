"""
Configuration file for marksheet parsing
Contains all column mappings and structure definitions
"""

# Helper function to convert Excel column to index
def col_to_index(col):
    """Convert Excel column letter(s) to 0-based index"""
    index = 0
    for char in col:
        index = index * 26 + (ord(char.upper()) - ord('A') + 1)
    return index - 1

# Row indices (0-based for pandas)
COURSE_CODE_ROW = 3  # Row 4 in Excel
COURSE_NAME_ROW = 4  # Row 5 in Excel
STUDENT_DATA_START_ROW = 11  # Row 12 in Excel

# Student basic info columns
STUDENT_COLUMNS = {
    'pr_number': 2,      # Column C
    'name': 4,           # Column E
    'seat_number': 6     # Column G
}

# Course grade columns mapping
# Format: 'course_code': {'credits_assigned': index, 'credits_earned': index, 'grade_point': index, 'letter_grade': index}
COURSE_COLUMNS = {
    'CMP-100': {
        'credits_assigned': col_to_index('O'),   # 14
        'credits_earned': col_to_index('P'),     # 15
        'grade_point': col_to_index('Q'),        # 16
        'letter_grade': col_to_index('S')        # 18
    },
    'CMP-101': {
        'credits_assigned': col_to_index('AE'),  # 30
        'credits_earned': col_to_index('AF'),    # 31
        'grade_point': col_to_index('AG'),       # 32
        'letter_grade': col_to_index('AI')       # 34
    },
    'MCV-111': {
        'credits_assigned': col_to_index('AQ'),  # 42
        'credits_earned': col_to_index('AR'),    # 43
        'grade_point': col_to_index('AS'),       # 44
        'letter_grade': col_to_index('AU')       # 46
    },
    'MCV-112': {
        'credits_assigned': col_to_index('BG'),  # 58
        'credits_earned': col_to_index('BH'),    # 59
        'grade_point': col_to_index('BI'),       # 60
        'letter_grade': col_to_index('BK')       # 62
    },
    'SHM-132': {
        'credits_assigned': col_to_index('BS'),  # 70
        'credits_earned': col_to_index('BT'),    # 71
        'grade_point': col_to_index('BU'),       # 72
        'letter_grade': col_to_index('BW')       # 74
    },
    'SHM-133': {
        'credits_assigned': col_to_index('CI'),  # 86
        'credits_earned': col_to_index('CJ'),    # 87
        'grade_point': col_to_index('CK'),       # 88
        'letter_grade': col_to_index('CM')       # 90
    },
    'AEC-153': {
        'credits_assigned': col_to_index('CU'),  # 98
        'credits_earned': col_to_index('CV'),    # 99
        'grade_point': col_to_index('CW'),       # 100
        'letter_grade': col_to_index('CY')       # 102
    },
    'VAC-158': {
        'credits_assigned': col_to_index('DH'),  # 111
        'credits_earned': col_to_index('DI'),    # 112
        'grade_point': col_to_index('DJ'),       # 113
        'letter_grade': col_to_index('DL')       # 115
    },
    'VAC-159': {
        'credits_assigned': col_to_index('DX'),  # 127
        'credits_earned': col_to_index('DY'),    # 128
        'grade_point': col_to_index('DZ'),       # 129
        'letter_grade': col_to_index('EB')       # 131
    },
    'SEC-143': {
        'credits_assigned': col_to_index('EM'),  # 142
        'credits_earned': col_to_index('EN'),    # 143
        'grade_point': col_to_index('EO'),       # 144
        'letter_grade': col_to_index('EQ')       # 146 (assuming pattern continues)
    }
}

# Summary columns
SUMMARY_COLUMNS = {
    'total_credits_assigned': col_to_index('ER'),  # 143
    'total_credits_earned': col_to_index('ES'),    # 144 (assuming typo in original, should be ES not ER)
    'overall_letter_grade': col_to_index('EX')     # 149
}

# List of all course codes in order
COURSE_CODES = [
    'CMP-100', 'CMP-101', 'MCV-111', 'MCV-112', 
    'SHM-132', 'SHM-133', 'AEC-153', 'VAC-158', 
    'VAC-159', 'SEC-143'
]