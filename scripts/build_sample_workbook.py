"""
Generates an advanced Excel workbook containing macroeconomic data and formula templates.
Requires: pip install openpyxl
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# Sheet 1: Raw Data
ws_data = wb.active
ws_data.title = "Raw_WorldBank_Data"

headers = ["Country_Code", "Country_Name", "Region", "Income_Group", "Year", "GDP_USD_Trillions", "Exchange_Rate_USD", "Inflation_Rate_Pct"]
ws_data.append(headers)

sample_rows = [
    ["IND", "India", "South Asia", "Lower middle income", 2010, 1.67, 45.7, 11.9],
    ["IND", "India", "South Asia", "Lower middle income", 2015, 2.10, 64.1, 4.9],
    ["IND", "India", "South Asia", "Lower middle income", 2020, 2.67, 74.1, 6.2],
    ["USA", "United States", "North America", "High income", 2010, 15.00, 1.0, 1.6],
    ["USA", "United States", "North America", "High income", 2015, 18.24, 1.0, 0.1],
    ["USA", "United States", "North America", "High income", 2020, 21.06, 1.0, 1.2],
    ["DEU", "Germany", "Europe & Central Asia", "High income", 2010, 3.42, 0.75, 1.1],
    ["DEU", "Germany", "Europe & Central Asia", "High income", 2015, 3.38, 0.90, 0.5],
    ["DEU", "Germany", "Europe & Central Asia", "High income", 2020, 3.89, 0.88, 0.5],
]

for row in sample_rows:
    ws_data.append(row)

# Sheet 2: Executive Dashboard
ws_dash = wb.create_sheet(title="Interactive_Dashboard")

# Title Header
ws_dash.merge_cells("B2:G2")
title_cell = ws_dash["B2"]
title_cell.value = "GLOBAL MACROECONOMIC INDICATORS DASHBOARD"
title_cell.font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
title_cell.fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
title_cell.alignment = Alignment(horizontal="center", vertical="center")

# Dynamic Parameter Selection Card
ws_dash["B4"] = "Select Country Code:"
ws_dash["B4"].font = Font(bold=True)
ws_dash["C4"] = "IND"  # Dropdown cell
ws_dash["C4"].fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

ws_dash["B5"] = "Select Year:"
ws_dash["B5"].font = Font(bold=True)
ws_dash["C5"] = 2020  # Dropdown cell
ws_dash["C5"].fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

# KPI Output Cards using Excel Lookup Formulas
ws_dash["B7"] = "Country Name:"
ws_dash["C7"] = '=XLOOKUP(C4, Raw_WorldBank_Data!A:A, Raw_WorldBank_Data!B:B, "Not Found")'

ws_dash["B8"] = "Region:"
ws_dash["C8"] = '=XLOOKUP(C4, Raw_WorldBank_Data!A:A, Raw_WorldBank_Data!C:C, "Not Found")'

ws_dash["B9"] = "Income Group:"
ws_dash["C9"] = '=XLOOKUP(C4, Raw_WorldBank_Data!A:A, Raw_WorldBank_Data!D:D, "Not Found")'

ws_dash["B11"] = "GDP (USD Trillions):"
ws_dash["C11"] = '=SUMIFS(Raw_WorldBank_Data!F:F, Raw_WorldBank_Data!A:A, C4, Raw_WorldBank_Data!E:E, C5)'
ws_dash["C11"].number_format = '$#,##0.00'

ws_dash["B12"] = "Exchange Rate (to USD):"
ws_dash["C12"] = '=SUMIFS(Raw_WorldBank_Data!G:G, Raw_WorldBank_Data!A:A, C4, Raw_WorldBank_Data!E:E, C5)'
ws_dash["C12"].number_format = '#,##0.00'

# Auto-adjust column widths
for ws in [ws_data, ws_dash]:
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

wb.save("dashboard/Global_Macroeconomic_Dashboard.xlsx")
print("Generated dashboard/Global_Macroeconomic_Dashboard.xlsx successfully!")
