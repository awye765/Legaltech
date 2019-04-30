import openpyxl
import re

# Returns Workbook File
workbook = openpyxl.load_workbook('Arb_Test_Data_01.xlsx')

# Returns Worksheet
worksheet = workbook.get_sheet_by_name('AW Test')

# Returns cell with Kira extracted arbitration clause
arb_clause = worksheet['G2'].value

# Dummy list of arbitration arb_institution_list
arb_institution_list = ["DIS"]

arb_institution_match = []

# H2 - Arbitration Clause (Institution)

def searchWholeWord(w):
    return re.compile(r'\b({0})\b'.format(w), flags=re.IGNORECASE).search

def extractWholeWord(arb_institution_list):
    for cat in arb_institution_list:
        return_value = searchWholeWord(cat)(arb_clause)
        if return_value:
            arb_institution_match.append(cat)
            print(arb_institution_match)
        else:
            print("Could not extract arbitration institution")

extractWholeWord(arb_institution_list)
# I2 - Arbitration Clause (Seat / Place / Venue) - CITY

# J2 - Arbitration Clause (Seat / Place / Venue) - COUNTRY

# K2 - S/P/V

# L2 - Arbitation Clause (Governing Law)

# M2 - Arbitration Clause (Governing Law - COUNTRY)

# Match Institution
