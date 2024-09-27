from Programs import genInfo as GI, canopy as canopy
import openpyxl
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
import streamlit as st
import math

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
    ws.row_breaks.append(openpyxl.worksheet.pagebreak.Break(id=row))
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
        #Capture Jet Hoods
        if v.model in ['KVF', 'KVI', 'KCH-F', 'KCH-I', 'KSR-S', 'KSR-F', 'KSR-M']:
            colorFill(ws, row)
            ws[f'A{row}'].border = Border(top=Side(style='thin'),
            bottom=Side(style='thin'))
            ws[f'j{row}'].border = Border(left=Side(style='thin'))
            ws.row_dimensions[row].height = 20
            ws.merge_cells(f'A{row}:I{row}')
            genFont(ws, 'A', row, "EXTRACT AIR DATA")
            makeCenter(ws, 'A', row)
            row+=1
            genFont(ws, 'A', row,f'Drawing Number: \t{v.drawingNum}')
            fillBorder(ws, 'A', row)
            iBorder(ws, 'i', row)
            row+=1
            genFont(ws, 'A', row, f'Location: \t{v.location}')
            fillBorder(ws, 'A', row)
            iBorder(ws, 'i', row)
            row+=1
            genFont(ws, 'A', row, f'Model: \t{v.model}')
            fillBorder(ws, 'A', row)
            iBorder(ws, 'i', row)
            row+=1
            genFont(ws, 'A', row, f'Quantity of Canopy Sections: \t{v.quantityOfSections}')
            fillBorder(ws, 'A', row)
            iBorder(ws, 'i', row)
            row+=1
            genFont(ws, 'A', row, f'Calculation: \tQV = Kf x √Pa')
            fillBorder(ws, 'A', row)
            iBorder(ws, 'i', row)
            row+=1
            colorFill(ws, row)
            ws.merge_cells(f'A{row}:I{row}')
            genFont(ws, 'A', row, "EXTRACT AIR READINGS")
            makeCenter(ws, 'A', row)
            row+=1
            #Module Number
            #ws.merge_cells(f'A{row}:B{row}')
            extendRow(ws, row)
            genFont(ws, 'A', row, 'Module #')
            makeCenter(ws, 'A', row)
            sectionBorder(ws, row, 1,2)
            #TAB Point Reading
            ws.merge_cells(f'B{row}:C{row}')
            extendRow(ws, row)
            genFont(ws, 'B', row, 'T.A.B Point Reading (Pa)')
            makeCenter(ws, 'B', row)
            sectionBorder(ws, row, 2,4)
            #K-Factor (m3/h)
            ws.merge_cells(f'D{row}:E{row}')
            extendRow(ws, row)
            genFont(ws, 'D', row, 'K-Factor (m3 /h)')
            makeCenter(ws, 'D', row)
            sectionBorder(ws, row, 4,6)
            #Achieved Flowrate
            #ws.merge_cells(f'F{row}:G{row}')
            extendRow(ws, row)
            genFont(ws, 'F', row, 'Flowrate (m3 /h)')
            makeCenter(ws, 'F', row)
            sectionBorder(ws, row, 6,7)
            #Design Flowrate
            #ws.merge_cells(f'H{row}:I{row}')
            extendRow(ws, row)
            genFont(ws, 'G', row, 'Flowrate (m3 /s)')
            makeCenter(ws, 'G', row)
            sectionBorder(ws, row, 7,8)
            #Percentage
            ws.merge_cells(f'H{row}:I{row}')
            extendRow(ws, row)
            genFont(ws, 'H', row, 'Percentage')
            makeCenter(ws, 'H', row)
            sectionBorder(ws, row, 8,10)
            row+=1
            # Now Fill the Info (v.sections)
            print(v.sections)
            totalFlowRate = 0
            percentage = 0
            for section, info in v.sections.items():
                print(info)
                print(f'{section} : {info}')
                #Mod Number
                genFont(ws, 'A', row, section)
                makeCenter(ws, 'A', row)
                sectionBorder(ws, row, 1,2)
                #Tab Reading
                ws.merge_cells(f'B{row}:C{row}')
                genFont(ws, 'B', row, f' {info['tab_reading']} pa')
                makeCenter(ws, 'B', row)
                sectionBorder(ws, row, 2, 4)
                #K-Factor
                ws.merge_cells(f'D{row}:E{row}')
                genFont(ws, 'D', row, f' {info['k_factor']}')
                makeCenter(ws, 'D', row)
                sectionBorder(ws, row, 4, 6)
                #AchievedFlow
                #ws.merge_cells(f'F{row}:G{row}')
                genFont(ws, 'F', row, f' {round(info['achieved'],2)}')
                makeCenter(ws, 'F', row)
                sectionBorder(ws, row, 6, 7)
                #Design Flow
                #ws.merge_cells(f'H{row}:I{row}')
                genFont(ws, 'G', row, f' {info['designFlow']}')
                makeCenter(ws, 'G', row)
                sectionBorder(ws, row, 7,8)
                totalFlowRate += info['designFlow']
                #Percentage
                ws.merge_cells(f'H{row}:I{row}')
                genFont(ws, 'H', row, f' {round(info['achieved']/info['designFlow'],0)}%')
                makeCenter(ws, 'H', row)
                sectionBorder(ws, row, 8,10)
                percentage += round(info['achieved']/info['designFlow'],0)

                row+=1
                #Total Flow Rate
            colorFill2(ws, row)
            row+=1
            ws.merge_cells(f'A{row}:D{row}')    
            genFont(ws, 'A', row, f'Total Flowrate                                {totalFlowRate} m3/s')
            sectionBorder(ws, row, 1, 5)
            row+=1
            ws.merge_cells(f'A{row}:D{row}')    
            genFont(ws, 'A', row, f'Total Percentage                                {percentage}%')
            sectionBorder(ws, row, 1, 5)
            if v.model in ['KVF', 'KCH-F']:
                #Supply Air Readings
                row +=2
                colorFill(ws, row)
                ws.merge_cells(f'A{row}:I{row}')
                genFont(ws, 'A', row, "SUPPLY AIR DATA")
                makeCenter(ws, 'A', row)
                row+=1
                genFont(ws, 'A', row,f'Drawing Number: \t{v.drawingNum}')
                fillBorder(ws, 'A', row)
                iBorder(ws, 'i', row)
                row+=1
                genFont(ws, 'A', row, f'Location: \t{v.location}')
                fillBorder(ws, 'A', row)
                iBorder(ws, 'i', row)
                row+=1
                genFont(ws, 'A', row, f'Model: \t{v.model}')
                fillBorder(ws, 'A', row)
                iBorder(ws, 'i', row)
                row+=1
                genFont(ws, 'A', row, f'Quantity of Canopy Sections: \t{v.quantityOfSections}')
                fillBorder(ws, 'A', row)
                iBorder(ws, 'i', row)
                row+=1
                genFont(ws, 'A', row, f'Calculation: \tQV = Kf x √Pa')
                fillBorder(ws, 'A', row)
                iBorder(ws, 'i', row)
                row+=1
                colorFill(ws, row)
                ws.merge_cells(f'A{row}:I{row}')
                genFont(ws, 'A', row, "SUPPLY AIR READINGS")
                makeCenter(ws, 'A', row)
                row+=1
                extendRow(ws, row)
                genFont(ws, 'A', row, 'Module #')
                makeCenter(ws, 'A', row)
                sectionBorder(ws, row, 1,2)
                #TAB Point Reading
                ws.merge_cells(f'B{row}:C{row}')
                extendRow(ws, row)
                genFont(ws, 'B', row, 'T.A.B Point Reading (Pa)')
                makeCenter(ws, 'B', row)
                sectionBorder(ws, row, 2,4)
                #K-Factor (m3/h)
                ws.merge_cells(f'D{row}:E{row}')
                extendRow(ws, row)
                genFont(ws, 'D', row, 'K-Factor (m3 /h)')
                makeCenter(ws, 'D', row)
                sectionBorder(ws, row, 4,6)
                #Achieved Flowrate
                #ws.merge_cells(f'F{row}:G{row}')
                extendRow(ws, row)
                genFont(ws, 'F', row, 'Flowrate (m3 /h)')
                makeCenter(ws, 'F', row)
                sectionBorder(ws, row, 6,7)
                #Design Flowrate
                extendRow(ws, row)
                genFont(ws, 'G', row, 'Flowrate (m3 /s)')
                makeCenter(ws, 'G', row)
                sectionBorder(ws, row, 7,8)
                #Percentage
                ws.merge_cells(f'H{row}:I{row}')
                extendRow(ws, row)
                genFont(ws, 'H', row, 'Percentage')
                makeCenter(ws, 'H', row)
                sectionBorder(ws, row, 8,10)
                row+=1
                totalFlowRateSup = 0
                totPercentage = 0
                for section2, info2 in v.sections.items():
                    #Mod Number
                    print(v.sections.items())
                    genFont(ws, 'A', row, section2)
                    makeCenter(ws, 'A', row)
                    sectionBorder(ws, row, 1,2)
                    #Tab Reading
                    ws.merge_cells(f'B{row}:C{row}')
                    genFont(ws, 'B', row, f' {info2['supplyTab']} pa')
                    makeCenter(ws, 'B', row)
                    sectionBorder(ws, row, 2, 4)
                    #K-Factor
                    ws.merge_cells(f'D{row}:E{row}')
                    genFont(ws, 'D', row, f' {info2['supplyKFactor']}')
                    makeCenter(ws, 'D', row)
                    sectionBorder(ws, row, 4, 6)
                    #AchievedFlow
                    #ws.merge_cells(f'F{row}:G{row}')
                    genFont(ws, 'F', row, f' {round(info2['achievedSupply'],2)}')
                    makeCenter(ws, 'F', row)
                    sectionBorder(ws, row, 6, 7)
                    #Design Flow
                    
                    genFont(ws, 'G', row, f' {info2['supplyDesign']}')
                    makeCenter(ws, 'G', row)
                    sectionBorder(ws, row, 7, 8)
                    totalFlowRateSup += info2['supplyDesign']
                    #Percentage
                    ws.merge_cells(f'H{row}:I{row}')
                    genFont(ws, 'H', row, f' {round(info2['achievedSupply']/info2['supplyDesign'], 0)}%')
                    makeCenter(ws, 'H', row)
                    sectionBorder(ws, row, 8, 10)
                    totPercentage += round(info2['achievedSupply']/info2['supplyDesign'], 0)
                    row+=1
                colorFill2(ws, row)
                row+=1
                ws.merge_cells(f'A{row}:D{row}')    
                genFont(ws, 'A', row, f'Total Flowrate                                {totalFlowRateSup} m3/s')
                sectionBorder(ws, row, 1, 5)
                row+=1
                ws.merge_cells(f'A{row}:D{row}')    
                genFont(ws, 'A', row, f'Total Percentage                                {totPercentage}%')
                sectionBorder(ws, row, 1, 5)
                row+=2
                #Start of Result summary (Extract Air)
                
                
        elif v.model in ['KSR-S', 'KSR-F', 'KSR-M']:
            print('lol')
            
            
        ws.row_breaks.append(openpyxl.worksheet.pagebreak.Break(id=row))  # Break after Row 20 (before row 21)
            #End of Capture Jet Hoods

        row+=2
        
    

    # You can also add more rows if necessary, e.g., canopy details, etc.
    
    # Define the path for the new Excel file
    updated_workbook_path = f'T&C.xlsx'
    
    # Save the workbook to the specified path
    wb.save(updated_workbook_path)
    
    # Provide the download link
    with open(updated_workbook_path, "rb") as file:
        st.download_button(label=f"Download Report", data=file, file_name=f"Testing And Commissioning.xlsx")
      
def iBorder(ws, letter, row):
    ws[f'i{row}'].border = Border(right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    ws[f'A{row}'].border = Border(left=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
def genFont(ws, letter, row, word):
    ws[f'{letter}{row}'] = word
    ws[f'{letter}{row}'].font = Font(size=11, bold=True)
def extendRow(ws, row):
    ws.row_dimensions[row].height = 30
def makeCenter(ws, letter, row):
    ws[f'{letter}{row}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
def sectionBorder(ws,row, first, second):
    thin_border = Border(
    top=Side(style='thin'),
    bottom=Side(style='thin'),
    left=Side(style='thin'),
    right=Side(style='thin')
    )
    for col in range(first, second):  # Columns A to D
        cell = ws.cell(row=row, column=col)
        cell.border = thin_border
        
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
            
def colorFill(ws, row):
    fill_style = PatternFill(start_color="9ac9f4", end_color="9ac9f4", fill_type="solid")
# Define a border (thin border for all sides)
    thin_border = Border(
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
    for col in range(1, 10):  # Column A to B (1 to 2)
                cell = ws.cell(row=row, column=col)
                cell.fill = fill_style  # Apply fill
                cell.border = thin_border  # Apply border
def colorFill2(ws, row):
    fill_style = PatternFill(start_color="9ac9f4", end_color="9ac9f4", fill_type="solid")
# Define a border (thin border for all sides)
    thin_border = Border(
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
    for col in range(1, 5):  # Column A to B (1 to 2)
                cell = ws.cell(row=row, column=col)
                cell.fill = fill_style  # Apply fill
                cell.border = thin_border  # Apply border