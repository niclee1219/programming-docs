"""
AGM Generator — Configuration
Ordinals, sheet name candidates, keyword search terms, default cell fallbacks.
Local file-path constants removed — persistence is now handled by SQLAlchemy.
"""

ORDINALS = {
    1: "First", 2: "Second", 3: "Third", 4: "Fourth", 5: "Fifth",
    6: "Sixth", 7: "Seventh", 8: "Eighth", 9: "Ninth", 10: "Tenth",
    11: "Eleventh", 12: "Twelfth", 13: "Thirteenth", 14: "Fourteenth",
    15: "Fifteenth", 16: "Sixteenth", 17: "Seventeenth", 18: "Eighteenth",
    19: "Nineteenth", 20: "Twentieth", 21: "Twenty-First", 22: "Twenty-Second",
    23: "Twenty-Third", 24: "Twenty-Fourth", 25: "Twenty-Fifth",
    26: "Twenty-Sixth", 27: "Twenty-Seventh", 28: "Twenty-Eighth",
    29: "Twenty-Ninth", 30: "Thirtieth",
}

# Sheet names tried in order when locating the export/data tab
SHEET_CANDIDATES = [
    "Export", "EXPORT", "export",
    "Data", "DATA", "data",
    "Info", "INFO", "Company", "COMPANY",
    "Summary", "SUMMARY", "Details", "DETAILS",
]

# Keywords searched in the first 80 rows of the sheet.
# The value is expected in the adjacent right cell or below.
FIELD_KEYWORDS = {
    "company_name":       ["Company Name", "Name of Company", "Company:"],
    "reg_no":             ["Reg No", "Registration No", "UEN", "Reg. No", "Company Reg"],
    "address":            ["Registered Address", "Office Address", "Address"],
    "financial_year_end": ["Financial Year End", "FYE", "Year End", "Year Ended", "FY End"],
    "agm_number":         ["AGM No", "AGM Number", "AGM #", "Meeting No", "Annual General Meeting No"],
    "agm_date":           ["AGM Date", "Meeting Date", "Date of AGM", "Date of Meeting"],
    "directors_start":    ["Directors", "Board of Directors", "Director Names", "Name of Directors"],
}

# Fallback cell positions tried if keyword search finds nothing
DEFAULT_CELLS = {
    "company_name":       ["B2", "B1", "C2", "A2"],
    "reg_no":             ["B3", "B4", "C3", "D3"],
    "address":            ["B5", "B6", "C5", "D5"],
    "financial_year_end": ["B7", "B8", "C7"],
    "agm_number":         ["B9", "B10", "C9"],
    "agm_date":           ["B11", "B12", "C11"],
}
