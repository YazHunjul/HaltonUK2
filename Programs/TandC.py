from Programs import genInfo as GI, canopy as canopy
import openpyxl
from openpyxl.styles import Font, Border, Side, PatternFill
import streamlit as st

def TandC():
    GI.titleAndLogo("Testing And Commissioning")
    genInfo = GI.getGenInfo()
    hoods = canopy.numCanopies()
    print(hoods)
    for k,v in hoods.items():
        print(v.model)
    
    if st.button("Save to Excel"):
        saveToExcel(genInfo, hoods)
        
def saveToExcel(genInfo, hoods):
    wb = openpyxl.load_workbook('T&C_Templates.xlsx')
    ws = wb.active
    
    #General Information
    row=4
    genFont(ws, 'A', row, f"CLIENT: {genInfo['client']} ")
    row+=2
    genFont(ws, 'A', row, f"PROJECT NAME: {genInfo['project_name']} ")
    row+=2
    genFont(ws, 'A', row, f"PROJECT NUMBER: {genInfo['project_number']} ")
    row+=2
    genFont(ws, 'A', row, f"DATE OF VISIT: {genInfo['date_of_visit']} ")
    row+=2
    genFont(ws, 'A', row, f"ENGINEER(s): {genInfo['engineers']} ")
    row+=2
    #Canopy Air Readings
    titleFont(ws, 'A', row, "KITCHEN CANOPY AIR READINGS")
    row+=2
    
    fill_style = PatternFill(start_color="9ac9f4", end_color="9ac9f4", fill_type="solid")
# Define a border (thin border for all sides)
    thin_border = Border(
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

    # Apply the fill and border dynamically to a range of cells


    for k,v in hoods.items():
        for col in range(1, 10):  # Column A to B (1 to 2)
            cell = ws.cell(row=row, column=col)
            cell.fill = fill_style  # Apply fill
            cell.border = thin_border  # Apply border
        ws[f'A{row}'].border = Border(left=Side(style='thin'),top=Side(style='thin'),
        bottom=Side(style='thin'))
        ws[f'j{row}'].border = Border(left=Side(style='thin'))
        ws.row_dimensions[row].height = 20
        genFont(ws, 'A', row, "EXTRACT AIR DATA")
        row+=1
        genFont(ws, 'A', row,v.drawingNum)
        fillBorder(ws, 'A', row)
        row+=1
    

    # You can also add more rows if necessary, e.g., canopy details, etc.
    
    # Define the path for the new Excel file
    updated_workbook_path = f'T&C.xlsx'
    
    # Save the workbook to the specified path
    wb.save(updated_workbook_path)
    
    # Provide the download link
    with open(updated_workbook_path, "rb") as file:
        st.download_button(label=f"Download Report", data=file, file_name=f"Testing And Commissioning.xlsx")
        
def genFont(ws, letter, row, word):
    ws[f'{letter}{row}'] = word
    ws[f'{letter}{row}'].font = Font(size=11, bold=True)
    
def titleFont(ws, letter, row, word):
    ws[f'{letter}{row}'] = word
    ws[f'{letter}{row}'].font = Font(size=14, bold=True)
    
def fillBorder(ws, letter, row):
    thin_border = Border(
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
    for col in range(1, 10):  # Column A to B (1 to 2)
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border  # Apply border