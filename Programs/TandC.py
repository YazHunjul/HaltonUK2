import io

from Programs import genInfo as GI, canopy as canopy
import openpyxl
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.drawing.image import Image
import streamlit as st
from PIL import Image as PILImage
from openpyxl.worksheet.header_footer import HeaderFooter
import math
from openpyxl.styles import Font

def TandC():
    GI.titleAndLogo("Testing And Commissioning")
    
    # Only load URL parameters once per session
    if 'url_params_loaded' not in st.session_state:
        try:
            GI.load_from_query_params()
            # Mark that we've loaded the URL parameters
            st.session_state['url_params_loaded'] = True
        except Exception as e:
            st.error(f"Error loading form data: {e}")
    
    genInfo = GI.getGenInfo()
    
    # Create sidebar for navigation
    with st.sidebar:
        st.header("Navigation")
        st.markdown("#### Jump to canopy:")
        
        # Create anchor for top of page
        st.markdown('<a name="top"></a>', unsafe_allow_html=True)
        
        # Add link to top of page
        st.markdown('[General Information](#top)', unsafe_allow_html=True)
    
    # Continue with the main form elements
    hoods, edge_box_details = canopy.numCanopies()  # Unpack returned values
    
    # Update sidebar with canopy links after we have the hood data
    with st.sidebar:
        for key, hood in hoods.items():
            # Extract the number from the key since hood.id might not exist
            # Key format is '{selection} {num}'
            hood_num = key.split()[-1]  # Get the last part which should be the number
            
            # Create unique anchor ID for each canopy
            anchor_id = f"canopy_{hood_num}"
            
            # Format the link to show model (drawing number) - location
            drawing_info = hood.drawingNum if hood.drawingNum else f"#{int(hood_num)+1}"
            st.markdown(f'[{hood.model} ({drawing_info}) - {hood.location}](#{anchor_id})', unsafe_allow_html=True)
        
        # Add links to other sections
        st.markdown("#### Other Sections:")
        st.markdown('[Comments](#comments)', unsafe_allow_html=True)
        st.markdown('[Signature](#signature)', unsafe_allow_html=True)
    
    # Get comments BEFORE signature
    st.markdown('<a name="comments"></a>', unsafe_allow_html=True)
    st.markdown("### Comments")
    comments = GI.get_comments()
    
    # Store the signature data to prevent it from disappearing on re-renders
    if 'signature_data' not in st.session_state:
        st.session_state.signature_data = None
    
    # Get the signature input AFTER comments
    st.markdown('<a name="signature"></a>', unsafe_allow_html=True)
    st.markdown("### Signature")
    sign = GI.get_sign()
    if sign.image_data is not None:
        st.session_state.signature_data = sign.image_data
    elif st.session_state.signature_data is not None:
        sign.image_data = st.session_state.signature_data
    
    # Add sharing options at the bottom
    st.markdown("---")
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("Save to Excel"):
            saveToExcel(genInfo, hoods, comments, sign, edge_box_details)
    with col2:
        with st.expander("📤 Share Your Form Data", expanded=False):
            st.write("Generate a link you can share with others. When they open the link, all your form data will be pre-filled.")
            GI.display_shareable_link()
        
def saveToExcel(genInfo, hoods, comments, sign, edge_box_details):
    wb = openpyxl.load_workbook('T&C_Templates.xlsx')
    ws = wb.active
    
    #General Information
    row=4
    genFont(ws, 'A', row, f"CLIENT: {genInfo['client'].title()} ")
    row+=2
    genFont(ws, 'A', row, f"PROJECT NAME: {genInfo['project_name'].title()} ")
    row+=2
    genFont(ws, 'A', row, f"PROJECT NUMBER: {genInfo['project_number'].title()} ")
    row+=2
    genFont(ws, 'A', row, f"DATE OF VISIT: {genInfo['date_of_visit']} ")
    row+=2
    genFont(ws, 'A', row, f"ENGINEER(s): {genInfo['engineers'].title()} ")
    row+=4
    # Apply the fill and border dynamically to a range of cells
    if canopy.features and len(canopy.features) > 0:
        genFont(ws, 'A', row, "Your Kitchen Ventilation System has the following serviceable technology:")
        ws[f'A{row}'].font = Font(bold=True, size=15)
        row += 3

        # Display UV Image if Selected
        if "CJ" in canopy.features and canopy.features["CJ"] == True:
            original_image = PILImage.open("features/C-RAY.png")
            resized_image = original_image.resize((600, 150))  # Resize the image
            resized_image.save("resized_image.png")
            img = Image("resized_image.png")
            ws.add_image(img, f'A{row}')
            row += 8  # Adjust the row position after the image

        # Display M.A.R.V.E.L. Image if Selected
        if "marvel" in canopy.features and canopy.features["marvel"] == True:
            original_image = PILImage.open("features/Marvel.png")
            resized_image = original_image.resize((620, 130))  # Resize the image
            resized_image.save("resized_Marvel.png")
            img = Image("resized_Marvel.png")
            ws.add_image(img, f'A{row}')
            row += 5  # Adjust the row position after the image
    row+=10
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


    for k,v in hoods.items():
        #Capture Jet Hoods
        if v.model in ['KVF', 'KVI', 'KCH-F', 'KCH-I', 'KSR-S', 'KSR-F', 'KSR-M', 'UVF', 'UVI', 'USR-S', 'USR-F', 'USR-M', 'UWF', 'CMW-FMOD']:
            colorFill(ws, row)
            ws[f'A{row}'].border = Border(top=Side(style='thin'),
            bottom=Side(style='thin'))
            #ws[f'j{row}'].border = Border(left=Side(style='thin'))
            ws.row_dimensions[row].height = 20
            ws.merge_cells(f'A{row}:I{row}')
            colorFill(ws, row)  # Apply the fill and border to the entire range
            genFont(ws, 'A', row, "EXTRACT AIR DATA", "FFFFFF")
            makeCenter(ws, 'A', row)
            row+=1
            genFont(ws, 'A', row,f'Drawing Number: {v.drawingNum}')
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
            genFont(ws, 'A', row, f'Total Design Airflow: \t{v.total_design_flow_ms} m³/s')
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
            # genFont(ws, 'A', row, f'Design Airflow: \t{v.totalDesignFlow} M³/s')
            # fillBorder(ws, 'A', row)
            # iBorder(ws, 'i', row)
            row+=1
            colorFill(ws, row)
            ws.merge_cells(f'A{row}:I{row}')
            genFont(ws, 'A', row, "EXTRACT AIR READINGS", "FFFFFF")
            makeCenter(ws, 'A', row)
            row+=1
            #Module Number
            #ws.merge_cells(f'A{row}:B{row}')
            extendRow(ws, row)
            genFont(ws, 'A', row, 'Section #')
            makeCenter(ws, 'A', row)
            sectionBorder(ws, row, 1,2)

            ws.merge_cells(f'B{row}:C{row}')
            genFont(ws, 'B', row, 'T.A.B Point Reading (Pa)')
            makeCenter(ws, 'B', row)
            sectionBorder(ws, row, 2,4)

            ws.merge_cells(f'D{row}:E{row}')
            genFont(ws, 'D', row, 'K-Factor (m³/h)')
            makeCenter(ws, 'D', row)
            sectionBorder(ws, row, 4,6)
            
            #Achieved Flowrate
            extendRow(ws, row)
            genFont(ws, 'F', row, 'Flowrate (m³/h)')
            makeCenter(ws, 'F', row)
            sectionBorder(ws, row, 6,7)
            #Design Flowrate
            #ws.merge_cells(f'H{row}:I{row}')
            extendRow(ws, row)
            genFont(ws, 'G', row, 'Flowrate (m³/s)')
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
            totalFlowRate = 0
            percentage = 0
            totalAchieved = 0
            for section, info in v.sections.items():
                #Mod Number
                genFont(ws, 'A', row, section)
                makeCenter(ws, 'A', row)
                sectionBorder(ws, row, 1,2)
                #Tab Reading
                ws.merge_cells(f'B{row}:C{row}')
                genFont(ws, 'B', row, f' {round(info['tab_reading'], 0)} Pa',)
                makeCenter(ws, 'B', row)
                sectionBorder(ws, row, 2, 4)
                #K-Factor
                ws.merge_cells(f'D{row}:E{row}')
                genFont(ws, 'D', row, f' {info['k_factor']}')
                makeCenter(ws, 'D', row)
                sectionBorder(ws, row, 4, 6)
                #AchievedFlow
                #ws.merge_cells(f'F{row}:G{row}')
                genFont(ws, 'F', row, f' {round(info['achieved'],0)}')
                makeCenter(ws, 'F', row)
                sectionBorder(ws, row, 6, 7)
                totalAchieved+=round(info['achieved'],2)
                #Design Flow
                #ws.merge_cells(f'H{row}:I{row}')
                genFont(ws, 'G', row, f' {round((info["k_factor"] * math.sqrt(info["tab_reading"])) / 3600, 2)}')
                makeCenter(ws, 'G', row)
                sectionBorder(ws, row, 7,8)
                totalFlowRate += (info["k_factor"] * math.sqrt(info["tab_reading"])) / 3600
                #Percentage
                ws.merge_cells(f'H{row}:I{row}')
                if v.total_design_flow_ms:
                    actual_flow = round((info["k_factor"] * math.sqrt(info["tab_reading"])) / 3600, 2)
                    genFont(ws, 'H', row, f' {round((actual_flow/v.total_design_flow_ms) * 100, 0)}%')
                makeCenter(ws, 'H', row)
                sectionBorder(ws, row, 8,10)
                if v.total_design_flow_ms:
                    percentage += round((actual_flow/v.total_design_flow_ms) * 100, 0)

                row+=1
                #Total Flow Rate
            colorFill2(ws, row)
            row+=1
            ws.merge_cells(f'A{row}:D{row}')    
            genFont(ws, 'A', row, f'Total Flowrate                                {round(totalFlowRate, 2)} m³/s')
            sectionBorder(ws, row, 1, 5)
            row+=1
            ws.merge_cells(f'A{row}:D{row}')
            if v.total_design_flow_ms:
                genFont(ws, 'A', row, f'Percentage of Design                                {round((totalFlowRate/v.total_design_flow_ms * 100),1)}%')
            sectionBorder(ws, row, 1, 5)
            row+=1
            if v.model in ['KVF', 'KCH-F', 'UVF', 'USR-F', 'KSR-F', 'KWF', 'UWF', 'CMW-FMOD']:
                #Supply Air Readings
                row +=2
                colorFill(ws, row)
                ws.merge_cells(f'A{row}:I{row}')
                genFont(ws, 'A', row, "SUPPLY AIR DATA", "FFFFFF")
                makeCenter(ws, 'A', row)
                row+=1
                genFont(ws, 'A', row,f"Drawing Number: {v.drawingNum}")
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
                genFont(ws, 'A', row, f'Total Supply Design Airflow: \t{v.total_supply_design_flow_ms} m³/s')
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
                # genFont(ws, 'G', row, f' {round((info2["supplyKFactor"] * math.sqrt(info2["supplyTab"])) / 3600, 2)} M³/s')
                # row+=1
                colorFill(ws, row)
                ws.merge_cells(f'A{row}:I{row}')
                genFont(ws, 'A', row, "SUPPLY AIR READINGS", "FFFFFF")
                makeCenter(ws, 'A', row)
                row+=1
                extendRow(ws, row)
                genFont(ws, 'A', row, 'Section #')
                makeCenter(ws, 'A', row)
                sectionBorder(ws, row, 1,2)
                
                ws.merge_cells(f'B{row}:C{row}')
                genFont(ws, 'B', row, 'T.A.B Point Reading (Pa)')
                makeCenter(ws, 'B', row)
                sectionBorder(ws, row, 2,4)
                #K-Factor (m3/h)
                ws.merge_cells(f'D{row}:E{row}')
                extendRow(ws, row)
                genFont(ws, 'D', row, 'K-Factor (m³/h)')
                makeCenter(ws, 'D', row)
                sectionBorder(ws, row, 4,6)
                #Achieved Flowrate
                #ws.merge_cells(f'F{row}:G{row}')
                extendRow(ws, row)
                genFont(ws, 'F', row, 'Flowrate (m³/h)')
                makeCenter(ws, 'F', row)
                sectionBorder(ws, row, 6,7)
                #Design Flowrate
                extendRow(ws, row)
                genFont(ws, 'G', row, 'Flowrate (m³/s)')
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
                    genFont(ws, 'A', row, section2)
                    makeCenter(ws, 'A', row)
                    sectionBorder(ws, row, 1,2)
                    #Tab Reading
                    ws.merge_cells(f'B{row}:C{row}')
                    genFont(ws, 'B', row, f' {round(info2['supplyTab'], 0)} Pa')
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
                    
                    genFont(ws, 'G', row, f' {round((info2["supplyKFactor"] * math.sqrt(info2["supplyTab"])) / 3600, 2)}')
                    makeCenter(ws, 'G', row)
                    sectionBorder(ws, row, 7, 8)
                    totalFlowRateSup += round((info2["supplyKFactor"] * math.sqrt(info2["supplyTab"])) / 3600, 2)
                    #Percentage for Supply
                    ws.merge_cells(f'H{row}:I{row}')
                    if v.total_supply_design_flow_ms:
                        actual_flow_supply_ms = round((info2["supplyKFactor"] * math.sqrt(info2["supplyTab"])) / 3600, 2)
                        genFont(ws, 'H', row, f' {round((actual_flow_supply_ms/v.total_supply_design_flow_ms * 100), 0)}%')
                    makeCenter(ws, 'H', row)
                    sectionBorder(ws, row, 8,10)
                    if v.total_supply_design_flow_ms:
                        totPercentage += round((actual_flow_supply_ms/v.total_supply_design_flow_ms * 100), 0)
                    row+=1
                colorFill2(ws, row)
                row+=1
                ws.merge_cells(f'A{row}:D{row}')    
                genFont(ws, 'A', row, f'Total Flowrate                                {totalFlowRateSup} m³/s')
                sectionBorder(ws, row, 1, 5)
                row+=1
                ws.merge_cells(f'A{row}:D{row}')    
                genFont(ws, 'A', row, f'Percentage of Design                                {totPercentage}%')
                sectionBorder(ws, row, 1, 5)
                row+=2
                #Start of Result summary (Extract Air)
        #CheckList

        if "UV Capture Jet" in v.checklist:
            row += 2
            # Add color-filled title for UV section
            colorFill(ws, row)
            ws.merge_cells(f'A{row}:I{row}')
            genFont(ws, 'A', row, f"UV CAPTURE RAY SYSTEM FOR {v.model}", "FFFFFF")
            makeCenter(ws, 'A', row)
            row += 1

            # Add checklist items
            for check, value in v.checklist["UV Capture Jet"].items():
                # Add description in A:G
                ws.merge_cells(f'A{row}:G{row}')
                genFont(ws, 'A', row, check)
                
                # Add value in H:I
                ws.merge_cells(f'H{row}:I{row}')
                if value is True:
                    ws[f'H{row}'] = "✓"
                elif value is False:
                    ws[f'H{row}'] = ""
                else:
                    ws[f'H{row}'] = value
                makeCenter(ws, 'H', row)

                # Apply borders to H and I
                border = Border(
                    top=Side(style='thin'),
                    bottom=Side(style='thin'),
                    left=Side(style='thin'),
                    right=Side(style='thin')
                )
                ws[f'H{row}'].border = border
                ws[f'I{row}'].border = border

                # Apply borders to the entire row
                sectionBorder(ws, row, 1, 9)
                row += 1

        # Add M.A.R.V.E.L. details if available
        if "M.A.R.V.E.L. System" in v.checklist:
            row += 2
            # Add color-filled title for Marvel section
            colorFill(ws, row)
            ws.merge_cells(f'A{row}:I{row}')
            genFont(ws, 'A', row, f"M.A.R.V.E.L. SYSTEM FOR {v.model}", "FFFFFF")
            makeCenter(ws, 'A', row)
            row += 1

            # Add checklist items
            for check, value in v.checklist["M.A.R.V.E.L. System"].items():
                # Add description in A:G
                ws.merge_cells(f'A{row}:G{row}')
                genFont(ws, 'A', row, check)
                
                # Add value in H:I
                ws.merge_cells(f'H{row}:I{row}')
                if value is True:
                    ws[f'H{row}'] = "✓"
                elif value is False:
                    ws[f'H{row}'] = ""
                else:
                    ws[f'H{row}'] = value
                makeCenter(ws, 'H', row)

                # Apply borders to H and I
                border = Border(
                    top=Side(style='thin'),
                    bottom=Side(style='thin'),
                    left=Side(style='thin'),
                    right=Side(style='thin')
                )
                ws[f'H{row}'].border = border
                ws[f'I{row}'].border = border

                # Apply borders to the entire row
                sectionBorder(ws, row, 1, 9)
                row += 1

        elif v.model == 'CXW':
            colorFill(ws, row)
            ws[f'A{row}'].border = Border(top=Side(style='thin'),
            bottom=Side(style='thin'))
            ws.row_dimensions[row].height = 20
            ws.merge_cells(f'A{row}:I{row}')
            colorFill(ws, row)  # Apply the fill and border to the entire range
            genFont(ws, 'A', row, "EXTRACT AIR DATA", "FFFFFF")
            makeCenter(ws, 'A', row)
            row+=1
            
            genFont(ws, 'A', row,f'Drawing Number: {v.drawingNum}')
            fillBorder(ws, 'A', row)
            iBorder(ws, 'i', row)
            row+=1
            
            genFont(ws, 'A', row, f'Location: \t{v.location}')
            fillBorder(ws, 'A', row)
            iBorder(ws, 'i', row)
            row+=1
            
            genFont(ws, 'A', row, f'Model: \tCXW')
            fillBorder(ws, 'A', row)
            iBorder(ws, 'i', row)
            row+=1
            
            genFont(ws, 'A', row, f'Design Flowrate: \t{v.total_design_flow_ms} m³/s')
            fillBorder(ws, 'A', row)
            iBorder(ws, 'i', row)
            row+=1
            
            genFont(ws, 'A', row, f'Quantity of Grills: \t{v.quantityOfSections}')
            fillBorder(ws, 'A', row)
            iBorder(ws, 'i', row)
            row+=1
            
            genFont(ws, 'A', row, f'Grill Size (mm): \tx')
            fillBorder(ws, 'A', row)
            iBorder(ws, 'i', row)
            row+=1
            
            genFont(ws, 'A', row, f'Calculation: \tQ = Av')
            fillBorder(ws, 'A', row)
            iBorder(ws, 'i', row)
            row+=1
            
            colorFill(ws, row)
            ws.merge_cells(f'A{row}:I{row}')
            genFont(ws, 'A', row, "EXTRACT AIR READINGS", "FFFFFF")
            makeCenter(ws, 'A', row)
            row+=1
            
            # Column headers
            ws.merge_cells(f'A{row}:B{row}')
            genFont(ws, 'A', row, 'Anemometer Reading\n(Average m/s)')
            makeCenter(ws, 'A', row)
            
            ws.merge_cells(f'C{row}:D{row}')
            genFont(ws, 'C', row, 'Free Area\n(m²)')
            makeCenter(ws, 'C', row)
            
            ws.merge_cells(f'E{row}:F{row}')
            genFont(ws, 'E', row, 'Flowrate\n(m³/h)')
            makeCenter(ws, 'E', row)
            
            ws.merge_cells(f'G{row}:I{row}')
            genFont(ws, 'G', row, 'Flowrate\n(m³/s)')
            makeCenter(ws, 'G', row)
            
            # Apply borders to all header cells
            for col in range(1, 10):
                ws.cell(row=row, column=col).border = Border(
                    top=Side(style='thin'),
                    bottom=Side(style='thin'),
                    left=Side(style='thin'),
                    right=Side(style='thin')
                )
            row+=1
            
            # Data rows
            total_flowrate_ms = 0
            for section, data in v.sections.items():
                ws.merge_cells(f'A{row}:B{row}')
                genFont(ws, 'A', row, f'{round(data["anemometer_reading"], 1)}')
                makeCenter(ws, 'A', row)
                
                ws.merge_cells(f'C{row}:D{row}')
                genFont(ws, 'C', row, f'{round(data["free_area"], 1)}')
                makeCenter(ws, 'C', row)
                
                ws.merge_cells(f'E{row}:F{row}')
                genFont(ws, 'E', row, f'{round(data["flowrate_m3h"], 1)}')
                makeCenter(ws, 'E', row)
                
                ws.merge_cells(f'G{row}:I{row}')
                genFont(ws, 'G', row, f'{round(data["flowrate_m3s"], 2)}')
                makeCenter(ws, 'G', row)
                
                # Apply borders to all data cells
                for col in range(1, 10):
                    ws.cell(row=row, column=col).border = Border(
                        top=Side(style='thin'),
                        bottom=Side(style='thin'),
                        left=Side(style='thin'),
                        right=Side(style='thin')
                    )
                
                total_flowrate_ms += data["flowrate_m3s"]
                row+=1
            # Add total row with color fill
            colorFill(ws, row)
            ws.merge_cells(f'A{row}:D{row}')
           

            # Apply borders to all cells in the total row
            for col in range(1, 1):
                ws.cell(row=row, column=col).border = Border(
                    top=Side(style='thin'),
                    bottom=Side(style='thin'),
                    left=Side(style='thin'),
                    right=Side(style='thin')
                )
            # Total Flowrate row
            row+=1
            ws.merge_cells(f'A{row}:D{row}')    
            genFont(ws, 'A', row, f'Total Flowrate                                {round(total_flowrate_ms, 2)} m³/s')
            sectionBorder(ws, row, 1, 5)
            
            # Percentage of Design row
            row+=1
            ws.merge_cells(f'A{row}:D{row}')    
            if v.total_design_flow_ms:
                percentage = round((total_flowrate_ms/v.total_design_flow_ms * 100), 1)
                genFont(ws, 'A', row, f'Percentage of Design                                {percentage}%')
            sectionBorder(ws, row, 1, 5)
            row+=1

        elif v.model in ['KSR-S', 'KSR-F', 'KSR-M']:
            print('lol')
            
        row+=3
        #Results
       # Results Summary - Extract Air
         # Results Summary for Extract Air
    extract_summary = []
    supply_summary = []
    total_extract_design = 0
    total_extract_actual = 0
    total_supply_design = 0
    total_supply_actual = 0

    # Iterate over each canopy to process both Extract and Supply Air
    for canopy_num, hood in hoods.items():
        drawing_number = f"{hood.drawingNum}"  # Format the drawing number

        # Initialize per-drawing totals
        extract_design_total = 0
        extract_actual_total = 0
        supply_design_total = 0
        supply_actual_total = 0

        # Process Extract Air sections
        if hood.model in ['KVF', 'KVI', 'KCH-F', 'KCH-I', 'KSR-S', 'KSR-F', 'KSR-M', 'UVF', 'UVI', 'USR-S', 'USR-F', 'USR-M', 'CMW-FMOD']:
            for section_data in hood.sections.values():
                extract_design_total = hood.total_design_flow_ms  # Use the total extract design flow
                extract_actual_total += round((section_data["k_factor"] * math.sqrt(section_data["tab_reading"])) / 3600, 2)

            # Calculate percentage using total values
            if extract_design_total > 0:
                extract_percentage = round((extract_actual_total / extract_design_total) * 100, 1)

        # Process Supply Air sections if applicable
        if hood.model in ['KVF', 'KCH-F', 'UVF', 'USR-F', 'KSR-F', 'KWF', 'UWF', 'CMW-FMOD']:
            for section_data in hood.sections.values():
                supply_design_total = hood.total_supply_design_flow_ms  # Use the total design flow for whole canopy
                supply_actual_total += round((section_data["supplyKFactor"] * math.sqrt(section_data["supplyTab"])) / 3600, 2)

            # Calculate percentage using total values
            if supply_design_total > 0:
                supply_percentage = round((supply_actual_total / supply_design_total) * 100, 1)

        # Calculate percentage and store results for each drawing
        if extract_design_total > 0:
            extract_summary.append({
                "drawing_number": v.drawingNum,  # Use the actual drawing number instead of "TOTAL"
                "design_flow_rate_total": extract_design_total,
                "actual_flow_rate_total": round(extract_actual_total, 2),
                "percentage": extract_percentage
            })
            total_extract_design += extract_design_total
            total_extract_actual += extract_actual_total

        if supply_design_total > 0:
            supply_summary.append({
                "drawing_number": drawing_number,
                "design_flow_rate_total": supply_design_total,
                "actual_flow_rate_total": round(supply_actual_total, 2),
                "percentage": supply_percentage
            })
            total_supply_design += supply_design_total
            total_supply_actual += supply_actual_total

    # Display Results Summary for Extract Air
    row += 2
    genFont(ws, 'A', row, "RESULTS SUMMARY - EXTRACT AIR")
    row += 1
    colorFill(ws, row)

    # Header for Extract Air Summary Table
    ws.merge_cells('A{0}:B{0}'.format(row))
    genFont(ws, 'A', row, "Drawing Number", "FFFFFF")
    ws.merge_cells('C{0}:D{0}'.format(row))
    genFont(ws, 'C', row, "Design Flow Rate (m³/s)", "FFFFFF")
    ws.merge_cells('E{0}:F{0}'.format(row))
    genFont(ws, 'E', row, "Actual Flowrate (m³/s)", "FFFFFF")
    ws.merge_cells('G{0}:I{0}'.format(row))
    genFont(ws, 'G', row, "Percentage of Design", "FFFFFF")
    sectionBorder(ws, row, 1, 10)
    colorFill(ws, row)
    
    row+=1

    # Populate Extract Air Summary Table
    for result in extract_summary:
        ws.merge_cells('A{0}:B{0}'.format(row))
        genFont(ws, 'A', row, result["drawing_number"])
        makeCenter(ws, 'A', row)
        
        ws.merge_cells('C{0}:D{0}'.format(row))
        genFont(ws, 'C', row, f"{result['design_flow_rate_total']} ")
        makeCenter(ws, 'C', row)
        
        ws.merge_cells('E{0}:F{0}'.format(row))
        genFont(ws, 'E', row, f"{round(result['actual_flow_rate_total'], 2)}")
        makeCenter(ws, 'E', row)
        
        ws.merge_cells('G{0}:I{0}'.format(row))
        genFont(ws, 'G', row, f"{result['percentage']}%")
        makeCenter(ws, 'G', row)
        
        sectionBorder(ws, row, 1, 10)
        row += 1

    # Totals row for Extract Air Summary
    ws.merge_cells('A{0}:B{0}'.format(row))
    genFont(ws, 'A', row, "TOTAL")
    makeCenter(ws, 'A', row)
    
    ws.merge_cells('C{0}:D{0}'.format(row))
    genFont(ws, 'C', row, f"{total_extract_design}")
    makeCenter(ws, 'C', row)
    
    ws.merge_cells('E{0}:F{0}'.format(row))
    genFont(ws, 'E', row, f"{total_extract_actual}")
    makeCenter(ws, 'E', row)
    
    ws.merge_cells('G{0}:I{0}'.format(row))
    overall_extract_percentage = round((total_extract_actual / total_extract_design) * 100, 1) if total_extract_design > 0 else 0
    genFont(ws, 'G', row, f"{round(overall_extract_percentage, 1)}%")
    makeCenter(ws, 'G', row)
    
    sectionBorder(ws, row, 1, 10)
    row += 3

    # Display Results Summary for Supply Air
    genFont(ws, 'A', row, "RESULTS SUMMARY - SUPPLY AIR")
    row += 1
    colorFill(ws, row)

    # Header for Supply Air Summary Table
    ws.merge_cells('A{0}:B{0}'.format(row))
    genFont(ws, 'A', row, "Drawing Number", "FFFFFF")
    
    ws.merge_cells('C{0}:D{0}'.format(row))
    genFont(ws, 'C', row, "Design Flow Rate (m³/s)", "FFFFFF")
    
    ws.merge_cells('E{0}:F{0}'.format(row))
    genFont(ws, 'E', row, "Actual Flowrate (m³/s)", "FFFFFF")
    
    ws.merge_cells('G{0}:I{0}'.format(row))
    genFont(ws, 'G', row, "Percentage of Design", "FFFFFF")
    
    sectionBorder(ws, row, 1, 10)
    row += 1

    # Populate Supply Air Summary Table
    for result in supply_summary:
        ws.merge_cells('A{0}:B{0}'.format(row))
        genFont(ws, 'A', row, result["drawing_number"])
        makeCenter(ws, 'A', row)
        
        ws.merge_cells('C{0}:D{0}'.format(row))
        genFont(ws, 'C', row, f"{result['design_flow_rate_total']}")
        makeCenter(ws, 'C', row)
        
        ws.merge_cells('E{0}:F{0}'.format(row))
        genFont(ws, 'E', row, f"{round(result['actual_flow_rate_total'], 2)}")
        makeCenter(ws, 'E', row)
        
        ws.merge_cells('G{0}:I{0}'.format(row))
        genFont(ws, 'G', row, f"{result['percentage']}%")
        makeCenter(ws, 'G', row)
        
        sectionBorder(ws, row, 1, 10)
        row += 1

    # Totals row
    ws.merge_cells('A{0}:B{0}'.format(row))
    genFont(ws, 'A', row, "TOTAL")
    makeCenter(ws, 'A', row)
    
    ws.merge_cells('C{0}:D{0}'.format(row))
    genFont(ws, 'C', row, f"{total_supply_design}")
    makeCenter(ws, 'C', row)
    
    ws.merge_cells('E{0}:F{0}'.format(row))
    genFont(ws, 'E', row, f"{round(total_supply_actual, 2)}")
    makeCenter(ws, 'E', row)
    
    ws.merge_cells('G{0}:I{0}'.format(row))
    overall_supply_percentage = round((total_supply_actual / total_supply_design) * 100, 1) if total_supply_design > 0 else 0
    genFont(ws, 'G', row, f"{overall_supply_percentage}%")
    makeCenter(ws, 'G', row)
    
    sectionBorder(ws, row, 1, 10)
    
    if edge_box_details:
    # Add header row with blue fill
        row += 5
        genFont(ws, 'A', row, "FINAL SYSTEM CHECKS")
        row += 1
        colorFill(ws, row)  # Blue background
        ws.merge_cells(f'A{row}:G{row}')
        ws.merge_cells(f'H{row}:I{row}')
        genFont(ws, 'A', row, "EDGE BOX", "FFFFFF")
        makeCenter(ws, 'A', row)

        # Add Edge Box parameters
        row += 1
        for key, value in edge_box_details.items():
            # Merge columns A:G for the parameter name
            ws.merge_cells(f'A{row}:G{row}')
            genFont(ws, 'A', row, key)

            # Merge columns H:I for the parameter value
            ws.merge_cells(f'H{row}:I{row}')
            if value is True:
                ws[f'H{row}'] = "✓"
            elif value is False:
                ws[f'H{row}'] = ""
            else:
                ws[f'H{row}'] = value  # For non-boolean values

            # Center-align the value
            makeCenter(ws, 'H', row)

            # Add borders for the row
            sectionBorder(ws, row, 1, 10)
            row += 1
                
                
    #Additional Notes
    row+=5
    colorFill(ws, row)
    ws.merge_cells('A{0}:I{0}'.format(row))
    genFont(ws, 'A', row, f'ADDITIONAL NOTES', "FFFFFF")
    row+=1
    for comment in comments:
        ws[f'A{row}'] = comment
        extendRow(ws, row)
        ws.merge_cells(f'A{row}:I{row}')
        makeCenter(ws, 'A', row)
        ws[f'A{row}'].border = Border(left=Side(style='thin'))
        ws[f'I{row}'].border = Border(right=Side(style='thin'))
        row+=1
    ws[f'A{row}'].border = Border(top=Side(style='thin'))
    ws[f'B{row}'].border = Border(top=Side(style='thin'))
    ws[f'C{row}'].border = Border(top=Side(style='thin'))
    ws[f'D{row}'].border = Border(top=Side(style='thin'))
    ws[f'E{row}'].border = Border(top=Side(style='thin'))
    ws[f'F{row}'].border = Border(top=Side(style='thin'))
    ws[f'G{row}'].border = Border(top=Side(style='thin'))
    ws[f'H{row}'].border = Border(top=Side(style='thin'))
    ws[f'I{row}'].border = Border(top=Side(style='thin'))
    genFont(ws, 'A', row, "Date")
    genFont(ws, 'C', row, f"{genInfo['date_of_visit']}")
    ws.merge_cells(f'A{row}:B{row}')
    ws.merge_cells(f'C{row}:I{row}')
    sectionBorder(ws, row, 1, 3)
    sectionBorder(ws, row, 3, 10)
    row+=1
    genFont(ws, 'A', row, "Print")
    genFont(ws, 'C', row, f"{genInfo['engineers']}")
    ws.merge_cells(f'A{row}:B{row}')
    ws.merge_cells(f'C{row}:I{row}')
    sectionBorder(ws, row, 1, 3)
    sectionBorder(ws, row, 3, 10)
    row += 1
    if sign.image_data is not None:
        img = PILImage.fromarray(sign.image_data.astype('uint8'), 'RGBA')

        # Step 1: Resize the image (change dimensions as needed, e.g., 150x75 pixels)
        resized_img = img.resize((200, 75))  # Resize to desired dimensions

        # Step 2: Save the resized image in-memory and to a file
        img_io = io.BytesIO()
        resized_img.save(img_io, format='PNG')
        img_io.seek(0)

        with open("resized_signature_image.png", "wb") as f:
            f.write(img_io.read())

        # Step 3: Load the resized image into openpyxl
        signature_img = Image("resized_signature_image.png")

        # Step 4: Insert the signature into the Excel file
        ws.merge_cells(f'A{row}:I{row}')
        genFont(ws, 'A', row, "Sign")
        # Apply only left and right borders to the merged range
        ws[f'A{row}'].border = Border(left=Side(style='thin'))
        ws[f'I{row}'].border = Border(right=Side(style='thin'))
        row += 1  # Move to the row under "Sign"
        
        start_row = row
        # Create 6 vertical rows for signature
        for i in range(6):
            ws.merge_cells(f'A{start_row + i}:I{start_row + i}')
            # Add left and right borders
            ws[f'A{start_row + i}'].border = Border(left=Side(style='thin'))
            ws[f'I{start_row + i}'].border = Border(right=Side(style='thin'))
        
        # Add bottom, left, and right borders to the last row
        ws[f'A{start_row + 5}'].border = Border(left=Side(style='thin'), bottom=Side(style='thin'))
        ws[f'I{start_row + 5}'].border = Border(right=Side(style='thin'), bottom=Side(style='thin'))
        for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:
            ws[f'{col}{start_row + 5}'].border = Border(bottom=Side(style='thin'))
        
        # Add the signature image with adjusted size and position
        ws.add_image(signature_img, f"D{start_row + 1}")  # Moved to column D and down one row
        
        row += 6

            
        ws.row_breaks.append(openpyxl.worksheet.pagebreak.Break(id=row))  # Break after Row 20 (before row 21)
            #End of Capture Jet Hoods

        row+=2
        
    

    # You can also add more rows if necessary, e.g., canopy details, etc.
    #applyMulishFont(ws)
    # Define the path for the new Excel file
    updated_workbook_path = f'T&C.xlsx'
    
    # Save the workbook to the specified path
    wb.save(updated_workbook_path)
    
    # Provide the download link
    with open(updated_workbook_path, "rb") as file:
        st.download_button(
            label="Download Report", 
            data=file, 
            file_name=f"{genInfo['project_number']} {genInfo['project_name']} - Canopy Commissioning Report.xlsx"
        )
      
def iBorder(ws, letter, row):
    ws[f'i{row}'].border = Border(right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    ws[f'A{row}'].border = Border(left=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
def genFont(ws, col, row, text, color="000000"):  # Added color parameter with default black
    ws[f'{col}{row}'] = text
    ws[f'{col}{row}'].font = Font(name='Calibri', size=11, color=color)
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
    fill_style = PatternFill(start_color="2499D5", end_color="2499D5", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Define the start and end columns for the merged range
    start_col = 1  # Column A
    end_col = 9    # Column I

    # Loop through each column in the merged range
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = fill_style
        cell.border = thin_border

    # Ensure the left and right borders for the first and last columns
    ws.cell(row=row, column=start_col).border = Border(
        left=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    ws.cell(row=row, column=end_col).border = Border(
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

def colorFill2(ws, row):
    ws.merge_cells(f'A{row}:D{row}')
    for col in ['A', 'B', 'C', 'D']:
        ws[f'{col}{row}'].fill = PatternFill(start_color="2499D5", end_color="2499D5", fill_type="solid")
    
    # Add borders only on the left of A and right of D
    ws[f'A{row}'].border = Border(left=Side(style='thin'))
    ws[f'D{row}'].border = Border(right=Side(style='thin'))

def applyMulishFont(ws):
    """
    Apply Mulish font to all cells in the worksheet.
    """
    mulish_font = Font(name='Arial', size=11)  # Define the Mulish font

    for row in ws.iter_rows():
        for cell in row:
            if cell.value:  # Apply font only if the cell contains content
                cell.font = mulish_font

            print("Current row:", row)
            print("Merged cells before clearing:", [str(r) for r in ws.merged_cells.ranges])
            
            # Clear any existing merges in this row
            for merge_range in ws.merged_cells.ranges.copy():
                if row in range(merge_range.min_row, merge_range.max_row + 1):
                    print(f"Unmerging: {str(merge_range)}")
                    ws.unmerge_cells(str(merge_range))
            
            print("Attempting to merge G to I:", f"G{row}:I{row}")
            ws.merge_cells(f'G{row}:I{row}')
            print("Merged cells after operation:", [str(r) for r in ws.merged_cells.ranges])
            
            genFont(ws, 'G', row, 'Percentage of Design', "FFFFFF")
            makeCenter(ws, 'G', row)