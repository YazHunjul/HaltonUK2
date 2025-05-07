import streamlit as st
import math

cj_hoods = ['KVF', 'KVI', 'KCH-F', 'KCH-I']
cj_lp = ['KSR-S', 'KSR-F', 'KSR-M']
cr_hoods = ['UVF', 'UVI']
cr_lp_hoods = ['USR-S', 'USR-F', 'USR-M']
ww_hoods = ['KWF', 'KWI', 'UWF', 'UWI', 'CMW-FMOD', 'CMW-IMOD']
new_hoods = ['XYZ1', 'XYZ2', 'XYZ3']  # Add the new hoods here

cj_hoods_k_factor = {
    1: 71.7,
    2: 143.5,
    3: 215.2,
    4: 287,
    5: 358.8,
    6: 430.5
}

cj_lp_k_factor = {
    1: 68.6,
    2: 128.6,
    3: 184.2,
    4: 245.6,
    5: 307,
    6: 368.4
}

cj_hoods_supply_kfactor = {
    1000: 121.7,
    1500: 182.6,
    2000: 243.4,
    2500: 304.2,
    3000: 365.1
}

cr_hoods_k_factor = {
    1: 71.8,
    2: 143.6,
    3: 215.4,
    4: 287.2,
    5: 359,
    6: 430.8
}

cr_lp_k_factor = {
     1: 68.6,
    2: 128.6,
    3: 184.2,
    4: 245.6,
    5: 307,
    6: 368.4
}

ww_hoods_kfactor = {
    1: 65.5,
    2: 131,
    3: 196.5,
    4: 262,
    5: 327.5,
    6: 393
}

new_hoods_kfactor = {
    1: 80.0,  # Replace with actual values
    2: 160.0,
    3: 240.0,
    4: 320.0,
    5: 400.0,
    6: 480.0
}

# Add CMW k-factor mapping after other k-factor dictionaries
cmw_kfactor = {
    1000: 115,
    1500: 172.5,
    2000: 230,
    2500: 287.5,
    3000: 345
}

features = {}

class cjHoods():
    def __init__(self, drawingNum, location, model, idNum, quantityOfSections, total_design_flow_ms, total_supply_design_flow_ms) -> None:
        self.drawingNum = drawingNum
        self.location = location
        self.model = model
        self.quantityOfSections = quantityOfSections
        self.total_design_flow_ms = total_design_flow_ms
        self.total_supply_design_flow_ms = total_supply_design_flow_ms
        self.ksaQuantity = 0
        self.sections = {}
        self.k_factor = 0
        self.tab_Reading = 0
        self.plenumLength = 0
        self.design_supply = 0
        self.tab_Reading_supply = 0
        self.checklist = {}
        self.id = idNum

    def getCJSections(self, selection, num):
        # Remove the CSS styling from here since it's now at the program start
        if selection in ["CMW-FMOD", "CMW-IMOD"]:
            return
            
        
        # Check if MARVEL is enabled for this hood
        marvel_enabled = 'M.A.R.V.E.L. System' in self.checklist
        
        for i in range(self.quantityOfSections):
            st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)

            st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>{selection} Section {i+1} Extract Info ({self.location})<h4>", unsafe_allow_html=True)
            
            # Create different column layouts based on MARVEL status
            if marvel_enabled:
                columns = st.columns(5)  # 5 columns for KSA, TAB, Min, Idle, Design
            else:
                columns = st.columns(2)  # 2 columns for standard layout
                
            st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)

            # Column 1: KSA's
            with columns[0]:
                # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>Section {i+1} KSA's<h4>", unsafe_allow_html=True)
                ksa_key = f'ksaQuantity{selection}{num}{i}'
                if ksa_key in st.session_state and not isinstance(st.session_state[ksa_key], int):
                    try:
                        st.session_state[ksa_key] = int(float(st.session_state[ksa_key]))
                    except:
                        st.session_state[ksa_key] = 1
                self.ksaQuantity = st.number_input(f"Section {i+1} KSA's", key=ksa_key, min_value=1, max_value=6)
            
            # Column 2: TAB Point Reading
            with columns[1]:
                # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} T.A.B Reading<h4>", unsafe_allow_html=True)
                tab_key = f'canopyLocs{selection}{num}{i}'
                if tab_key in st.session_state and not isinstance(st.session_state[tab_key], (float, int)):
                    try:
                        st.session_state[tab_key] = float(st.session_state[tab_key])
                    except:
                        st.session_state[tab_key] = 0.0
                self.tab_Reading = st.number_input(f'Section {i+1} T.A.B Reading', key=tab_key, min_value=0.0)
            
            # If MARVEL is enabled, add Min, Idle, Design columns for Extract
            marvel_data = {}
            if marvel_enabled:
                # Column 3: Extract Min %
                with columns[2]:
                    # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} Min (%)<h4>", unsafe_allow_html=True)
                    min_pct_key = f'extract_min_pct_{selection}_{i}_{num}'
                    try:
                        current_value = float(st.session_state.get(min_pct_key, 0.0))
                    except (ValueError, TypeError):
                        current_value = 0.0
                    st.session_state[min_pct_key] = current_value
                    min_pct = st.number_input(f'Section {i+1} Min (%)', min_value=0.0, max_value=100.0, value=current_value, step=0.1, key=min_pct_key,)
                    marvel_data['extract_min_pct'] = float(min_pct)
                
                # Column 4: Extract Idle %
                with columns[3]:
                    # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} Idle (%)<h4>", unsafe_allow_html=True)
                    idle_pct_key = f'extract_idle_pct_{selection}_{i}_{num}'
                    try:
                        current_value = float(st.session_state.get(idle_pct_key, 0.0))
                    except (ValueError, TypeError):
                        current_value = 0.0
                    st.session_state[idle_pct_key] = current_value
                    idle_pct = st.number_input(f'Section {i+1} Idle (%)', min_value=0.0, max_value=100.0, value=current_value, step=0.1, key=idle_pct_key,)
                    marvel_data['extract_idle_pct'] = float(idle_pct)
                
                # Column 5: Extract Design Flow
                with columns[4]:
                    # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} Design<h4>", unsafe_allow_html=True)
                    design_flow_per_section = 0.0
                    design_flow_key = f'extract_design_flow_{selection}_{i}_{num}'
                    if design_flow_key in st.session_state:
                        try:
                            st.session_state[design_flow_key] = float(st.session_state[design_flow_key])
                        except (ValueError, TypeError):
                            st.session_state[design_flow_key] = 0.0
                    else:
                        st.session_state[design_flow_key] = 0.0
                    design_flow = st.number_input(f'Section {i+1} Design', min_value=0.0, format='%.3f', value=float(st.session_state[design_flow_key]), step=0.001, key=design_flow_key)
                    marvel_data['extract_design_flow'] = float(design_flow)

            # Assign K-factor based on hood type
            if selection in cj_hoods:
                self.k_factor = cj_hoods_k_factor[self.ksaQuantity]
            elif selection in cj_lp:
                self.k_factor = cj_lp_k_factor[self.ksaQuantity]
            elif selection in cr_hoods:
                self.k_factor = cr_hoods_k_factor[self.ksaQuantity]
            elif selection in cr_lp_hoods:
                self.k_factor = cr_lp_k_factor[self.ksaQuantity]
            elif selection in ww_hoods:
                self.k_factor = ww_hoods_kfactor[self.ksaQuantity]
            elif selection in new_hoods:
                self.k_factor = new_hoods_kfactor[self.ksaQuantity]

            # Ensure tab_Reading is a float for calculation
            tab_reading_float = float(self.tab_Reading)
            achievedFlowRate = self.k_factor * math.sqrt(tab_reading_float)

            # Handle supply air logic if applicable
            if selection in ['KVF', 'KCH-F', 'UVF', 'USR-F', 'KSR-F', 'KWF', 'UWF'] and selection not in ["CMW-FMOD", "CMW-IMOD"]:
                st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>{selection} Section {i+1} Supply Info ({self.location})<h4>", unsafe_allow_html=True)
                
                if marvel_enabled:
                    supply_cols = st.columns(5)  # 5 columns for supply settings when MARVEL is enabled
                else:
                    supply_cols = st.columns(2)  # 2 columns for standard supply settings

                # Column 1: Plenum Length
                with supply_cols[0]:
                    # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>Section {i+1} Plenum Length<h4>", unsafe_allow_html=True)
                    plenum_key = f'plenumLength{selection}{num}{i}'
                    if plenum_key in st.session_state and not isinstance(st.session_state[plenum_key], int):
                        try:
                            val = int(float(st.session_state[plenum_key]))
                            if val in [1000, 1500, 2000, 2500, 3000]:
                                st.session_state[plenum_key] = val
                            else:
                                st.session_state[plenum_key] = 1000
                        except:
                            st.session_state[plenum_key] = 1000
                    self.plenumLength = st.selectbox(f'Section {i+1} Plenum Length', [1000, 1500, 2000, 2500, 3000], key=plenum_key)
                
                supplyKFactor = cj_hoods_supply_kfactor[self.plenumLength]
                
                # Column 2: Supply TAB Reading
                with supply_cols[1]:
                    # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} T.A.B Point Reading<h4>", unsafe_allow_html=True)
                    supply_tab_key = f'sup_tab{selection}{num}{i}'
                    if supply_tab_key in st.session_state and not isinstance(st.session_state[supply_tab_key], (float, int)):
                        try:
                            st.session_state[supply_tab_key] = float(st.session_state[supply_tab_key])
                        except:
                            st.session_state[supply_tab_key] = 0.0
                    self.tab_Reading_supply = st.number_input(f'Section {i+1} T.A.B Point Reading', key=supply_tab_key, min_value=0.0)

                # If MARVEL is enabled, add Min, Idle, Design columns for Supply
                supply_marvel_data = {}
                if marvel_enabled:
                    # Column 3: Supply Min %
                    with supply_cols[2]:
                        # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} Supply Min (%)<h4>", unsafe_allow_html=True)
                        min_pct_key = f'supply_min_pct_{selection}_{i}_{num}'
                        try:
                            current_value = float(st.session_state.get(min_pct_key, 0.0))
                        except (ValueError, TypeError):
                            current_value = 0.0
                        st.session_state[min_pct_key] = current_value
                        min_pct = st.number_input(f'Section {i+1} Supply Min (%)', min_value=0.0, max_value=100.0, value=current_value, step=0.1, key=min_pct_key)
                        supply_marvel_data['supply_min_pct'] = float(min_pct)
                    
                    # Column 4: Supply Idle %
                    with supply_cols[3]:
                        # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} Supply Idle (%)<h4>", unsafe_allow_html=True)
                        idle_pct_key = f'supply_idle_pct_{selection}_{i}_{num}'
                        try:
                            current_value = float(st.session_state.get(idle_pct_key, 0.0))
                        except (ValueError, TypeError):
                            current_value = 0.0
                        st.session_state[idle_pct_key] = current_value
                        idle_pct = st.number_input(f'Section {i+1} Supply Idle (%)', min_value=0.0, max_value=100.0, value=current_value, step=0.1, key=idle_pct_key)
                        supply_marvel_data['supply_idle_pct'] = float(idle_pct)
                    
                    # Column 5: Supply Design Flow
                    with supply_cols[4]:
                        # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} Supply Design<h4>", unsafe_allow_html=True)
                        supply_design_flow_per_section = 0.0
                        design_flow_key = f'supply_design_flow_{selection}_{i}_{num}'
                        if design_flow_key in st.session_state:
                            try:
                                st.session_state[design_flow_key] = float(st.session_state[design_flow_key])
                            except (ValueError, TypeError):
                                st.session_state[design_flow_key] = 0.0
                        else:
                            st.session_state[design_flow_key] = 0.0
                        design_flow = st.number_input(f'Section {i+1} Supply Design', min_value=0.0, format='%.3f', value=float(st.session_state[design_flow_key]), step=0.001, key=design_flow_key)
                        supply_marvel_data['supply_design_flow'] = float(design_flow)

                # Ensure supply tab reading is a float for calculation
                supply_tab_float = float(self.tab_Reading_supply)
                supplyAchievedFlowRate = supplyKFactor * math.sqrt(supply_tab_float)
                
                # Create section data
                section_data = {
                    "ksaQuantity": self.ksaQuantity,
                    "k_factor": self.k_factor,
                    'tab_reading': self.tab_Reading,
                    'achieved': round(achievedFlowRate, 2),
                    'supplyKFactor': supplyKFactor,
                    'supplyTab': self.tab_Reading_supply,
                    'supplyDesign': round(self.design_supply, 2),
                    'achievedSupply': round(supplyAchievedFlowRate, 2)
                }
                
                # Add MARVEL data if present
                if marvel_enabled:
                    section_data.update(marvel_data)  # Add extract MARVEL data
                    section_data.update(supply_marvel_data)  # Add supply MARVEL data
                    
                self.sections[f'{i+1}'] = section_data
            else:
                # Create section data
                section_data = {
                    "ksaQuantity": self.ksaQuantity,
                    "k_factor": self.k_factor,
                    'tab_reading': self.tab_Reading,
                    'achieved': round(achievedFlowRate, 2)
                }
                
                # Add MARVEL data if present
                if marvel_enabled and marvel_data:
                    section_data.update(marvel_data)
                    
                self.sections[f'{i+1}'] = section_data
            

def numCanopies():
    # Add custom CSS for input styling
    st.markdown("""
    <style>
    /* Style for all empty inputs */
    input:not([value]),
    input[value=""],
    input[value="0"],
    input[value="0.0"],
    input[value="0.00"],
    input[value="0.000"],
    textarea:empty,
    select:has(option[value=""]:checked),
    select:has(option[value="0"]:checked),
    div[data-baseweb] input:not([value]),
    div[data-baseweb] input[value=""],
    div[data-baseweb] input[value="0"],
    div[data-baseweb] input[value="0.0"],
    div[data-baseweb] input[value="0.00"],
    div[data-baseweb] input[value="0.000"],
    .stTextInput input:placeholder-shown,
    .stNumberInput input:placeholder-shown,
    .stTextInput input[value=""],
    .stNumberInput input[value="0"],
    .stNumberInput input[value="0.0"],
    .stNumberInput input[value="0.00"],
    .stNumberInput input[value="0.000"],
    .stSelectbox select:has(option[value=""]:checked),
    .stSelectbox select:has(option[value="0"]:checked) {
        background-color: #ffcccb !important;
    }

    /* Style for empty textareas in Streamlit */
    .stTextArea [data-baseweb="textarea"]:empty,
    .stTextArea textarea:placeholder-shown {
        background-color: #ffcccb !important;
    }

    /* Style for empty multiselect */
    .stMultiSelect [data-baseweb="select"]:not(:has(span[role="option"])) {
        background-color: #ffcccb !important;
    }
    </style>
    """, unsafe_allow_html=True)

    hoods = {}
    # st.markdown("<h1 style='text-align: center;margin-top: 30px;margin-bottom:-100px;'>Enter Number Of Canopies<h1>", unsafe_allow_html=True)
    
    # Ensure num_canopies is an integer
    if 'num_canopies' in st.session_state and not isinstance(st.session_state['num_canopies'], int):
        try:
            # Try to convert to int
            st.session_state['num_canopies'] = int(float(st.session_state['num_canopies']))
        except:
            # Default to 0 if conversion fails
            st.session_state['num_canopies'] = 0
    
    # Use consistent key for number of canopies to help with state restoration
    numInput = st.number_input('Enter Number Of Canopies', min_value=0, key="num_canopies")
    
    for num in range(numInput):
        # Add an anchor for this canopy
        # st.markdown(f'<a name="canopy_{num}"></a>', unsafe_allow_html=True)
        # Include canopy number in the heading
        # st.markdown(f"<h3 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy #{num+1}</h3>", unsafe_allow_html=True)
        selection = st.selectbox(f'Canopy {num+1}', cj_hoods + cj_lp + cr_hoods + cr_lp_hoods + ww_hoods + new_hoods + ['CXW'], key=f'selection{num}')
        if selection == 'CXW':
            hood = createCXWHood(selection, num)
        else:
            hood = createCJHood(selection, num)
            hood.getCJSections(selection, num)
        hoods[f'{selection} {num}'] = hood
        st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)

    edge_box_checked = st.checkbox("Final System Checks: Edge Box", key="edge_box_checked")
    edge_box_details = {}
    if edge_box_checked:
        edge_box_details = edgeBoxForm()

    return hoods, edge_box_details


def edgeBoxForm():
    """
    Displays the Edge Box details form and returns the entered data.
    """
    st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Edge Box Details<h3>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    with col1:
        edge_installed = st.checkbox("Edge Installed", key="edge_installed")
        edge_id = st.text_input("Edge ID", key="edge_id", value="")
        edge_4g_status = st.selectbox("Edge 4G Status", ["", "Online", "Offline"], key="edge_4g_status")
    with col2:
        modbus_operation = st.checkbox("Modbus Operation", key="modbus_operation")
        lan_connection = st.selectbox("LAN Connection Available", ["", "Yes", "No", "N/A"], key="lan_connection")
    st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)
    # Return the collected Edge Box data as a dictionary
    return {
        "Edge Installed": edge_installed,
        "Edge ID": edge_id,
        "Edge 4G Status": edge_4g_status,
        "LAN Connection Available": lan_connection,
        "Modbus Operation": modbus_operation,
    }


def checklist(hood, location, drawingNum):
    """
    Checklist for UV Capture Jet or M.A.R.V.E.L. System.
    - "With UV" is only shown for hoods starting with 'UV'.
    - "With M.A.R.V.E.L." is shown for all hoods.
    """
    system_checks = {}

    # UV Capture Jet System Checklist (only for hoods starting with 'UV')
    if hood.startswith("UV"):
        with_uv = st.checkbox(f"{(hood.split(' '))[0]} ({location}) With UV Capture Jet?", key=f"uv_{hood}_{location}_{drawingNum}")
        if with_uv:
            features["CJ"] = True
            st.write("UV Capture Ray System Checklist")
            col1, col2 = st.columns(2)
            with col1:
                # Ensure quantity is an integer
                quantity_key = f'quantity_uv_{hood}_{location}'
                if quantity_key in st.session_state and not isinstance(st.session_state[quantity_key], int):
                    try:
                        st.session_state[quantity_key] = int(float(st.session_state[quantity_key]))
                    except:
                        st.session_state[quantity_key] = 0
                quantity = st.number_input("Quantity of Slaves per System", min_value=0, key=quantity_key)
                
                # Ensure pressure is a number
                pressure_key = f'pressure_uv_{hood}_{location}'
                if pressure_key in st.session_state and not isinstance(st.session_state[pressure_key], (int, float)):
                    try:
                        st.session_state[pressure_key] = float(st.session_state[pressure_key])
                    except:
                        st.session_state[pressure_key] = 0
                pressure = st.number_input("UV Pressure Setpoint (Pa)", min_value=0, key=pressure_key)
                
                # Ensure CJ reading is a number
                cj_reading_key = f'cjreading_uv_{hood}_{location}'
                if cj_reading_key in st.session_state and not isinstance(st.session_state[cj_reading_key], (int, float)):
                    try:
                        st.session_state[cj_reading_key] = float(st.session_state[cj_reading_key])
                    except:
                        st.session_state[cj_reading_key] = 0
                cj_reading = st.number_input("Capture Jet Average Pressure Reading (Pa)", min_value=0, key=cj_reading_key)
            with col2:
                airflow = st.checkbox("Airflow Proved on Each Controller", value=True, key=f'airflow_uv_{hood}_{location}')
                safety = st.checkbox('All UV Safety Switches Tested', value=True, key=f'safety_uv_{hood}_{location}')
                comms = st.checkbox("Communication To All Canopy controllers", value=True, key=f'comms_uv_{hood}_{location}')
                filters = st.checkbox("All Filter Safety Switches Tested", value=True, key=f'filters_uv_{hood}_{location}')
                ops = st.checkbox("UV System Tested and Fully Operational", value=True, key=f'ops_uv_{hood}_{location}')
                cj = st.checkbox("Capture Jet Operational", value=True, key=f'cj_uv_{hood}_{location}')
            system_checks['UV Capture Jet'] = {
                "Quantity of Slaves per System": quantity,
                "UV Pressure Setpoint (Pa)": pressure,
                "Capture Jet Average Pressure Reading (Pa)": cj_reading,
                "Airflow Proved on Each Controller": airflow,
                "Communication To All Canopy Controllers": comms,
                "UV System Tested and Fully Operational": ops,
                "All UV Safety Switches Tested": safety,
                "All Filter Safety Switches Tested": filters,
                # "UV Capture Jet Operational": cj,
            }

    # M.A.R.V.E.L. System Checklist (available for all hoods)
    with_marvel = st.checkbox(f"{(hood.split(' '))[0]} ({location}) With M.A.R.V.E.L. System?", key=f"marvel_{hood}_{location}_{drawingNum}")
    
    if with_marvel:
        features["marvel"] = True
        system_checks['M.A.R.V.E.L. System'] = {}

    return system_checks


        

def createCJHood(selection, num):
    """
    Create a hood object and retrieve system-specific checks (UV or M.A.R.V.E.L.).
    """
    st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>{selection} Info<h3>", unsafe_allow_html=True)
    
    # For CMW-FMOD and CMW-IMOD, use a special layout with more columns
    if selection in ["CMW-FMOD", "CMW-IMOD"]:
        col1, col2, col3, col4 = st.columns(4)
        
        # Keep track of used drawing numbers
        if 'used_drawing_numbers' not in st.session_state:
            st.session_state.used_drawing_numbers = set()

        with col1:
            # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Drawing Number<h4>", unsafe_allow_html=True)
            drawingNum = st.text_input('Drawing Number', key=f'drawNum{num}')
            
            # After we have the drawing number, update the header
            if drawingNum:
                # Replace the header with one that includes the drawing number
                drawing_header = f"{selection} ({drawingNum}) Info"
                st.markdown(f"""
                <script>
                document.querySelector('h3:contains("{selection} Info")').innerText = '{drawing_header}';
                </script>
                """, unsafe_allow_html=True)
        
        with col2:
            # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy Location<h4>", unsafe_allow_html=True)
            canopyLocation = st.text_input(f'Canopy Location', key=f'canopyLoc{num}')
        
        with col3:
            # Design Flow
            design_flow_key = f'design_flow_{num}'
            if design_flow_key in st.session_state and not isinstance(st.session_state[design_flow_key], (float, int)):
                try:
                    st.session_state[design_flow_key] = float(st.session_state[design_flow_key])
                except:
                    st.session_state[design_flow_key] = 0.0
            
            # Add the label back
            # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Extract Design Airflow <h4>", unsafe_allow_html=True)
            
            # Instead of setting the session state value and then reading it,
            # just use a direct value
            total_design_flow_ms = st.number_input(f'Design Airflow', 
                                                key=design_flow_key, 
                                                min_value=0.0,
                                                step=0.001,
                                                value=0.0)
            
            # Show supply design flow for supply-capable hoods
            if selection == 'CMW-FMOD':
                with col3:
                    # Supply flow
                    supply_flow_key = f'total_supply_design_flow_ms{num}'
                    if supply_flow_key in st.session_state and not isinstance(st.session_state[supply_flow_key], (float, int)):
                        try:
                            st.session_state[supply_flow_key] = float(st.session_state[supply_flow_key])
                        except:
                            st.session_state[supply_flow_key] = 0.0
                        
                    # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Supply Design Airflow <h4>", unsafe_allow_html=True)
                    
                    # Instead of setting the session state value and then reading it,
                    # just use a direct value
                    total_supply_design_flow_ms = st.number_input(f'Supply Design Airflow', 
                                                                key=supply_flow_key, 
                                                                
                                                                min_value=0.0,
                                                                format="%.3f",
                                                                value=0.0)
            else:
                total_supply_design_flow_ms = 0.0
        
        with col4:
            # Ensure quantity is an integer
            quantity_key = f'quantity{num}'
            if quantity_key in st.session_state and not isinstance(st.session_state[quantity_key], int):
                try:
                    st.session_state[quantity_key] = int(float(st.session_state[quantity_key]))
                except:
                    st.session_state[quantity_key] = 0
                    
            # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy Sections<h4>", unsafe_allow_html=True)
            quantity = st.number_input(f'Canopy Sections', key=quantity_key, min_value=0)
        
        # Second row with slot dimensions
        # st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)
        slot_col1, slot_col2 = st.columns(2)
        
        with slot_col1:
            # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Length of Slot (mm)<h4>", unsafe_allow_html=True)
            # Add validation to ensure slot_length is a number
            slot_length_key = f"slot_length_{num}"
            if slot_length_key in st.session_state and not isinstance(st.session_state[slot_length_key], (int, float)):
                try:
                    st.session_state[slot_length_key] = int(float(st.session_state[slot_length_key]))
                    if st.session_state[slot_length_key] not in [1000, 1500, 2000, 2500, 3000]:
                        st.session_state[slot_length_key] = 1000
                except (ValueError, TypeError):
                    st.session_state[slot_length_key] = 1000  # Default value
            
            slot_length = st.number_input(f'Length of Slot (mm)', min_value=1000, max_value=3000, value=1000, step=500, key=slot_length_key)
        
        with slot_col2:
            # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Width of Slot (mm)<h4>", unsafe_allow_html=True)
            # Add validation to ensure slot_width is a number
            slot_width_key = f"slot_width_{num}"
            if slot_width_key in st.session_state and not isinstance(st.session_state[slot_width_key], (int, float)):
                try:
                    st.session_state[slot_width_key] = int(float(st.session_state[slot_width_key]))
                    if st.session_state[slot_width_key] < 1 or st.session_state[slot_width_key] > 200:
                        st.session_state[slot_width_key] = 85
                except (ValueError, TypeError):
                    st.session_state[slot_width_key] = 85  # Default value
                    
            slot_width = st.number_input(f'Width of Slot (mm)', min_value=1, max_value=200, value=85, step=1, key=slot_width_key)
        marvel_enabled = st.checkbox(f"{selection} ({canopyLocation}) With M.A.R.V.E.L. System?", key=f"marvel_{selection}_{num}")
        # Create the hood object
        hood = cjHoods(drawingNum, canopyLocation, selection, num, quantity, total_design_flow_ms, total_supply_design_flow_ms)
        
        # Store the slot dimensions in the sections dictionary
        hood.sections["slot_length"] = slot_length
        hood.sections["slot_width"] = slot_width
        
        # Set k-factor based on slot length for CMW hoods
        if selection in ["CMW-FMOD", "CMW-IMOD"]:
            # Find the closest slot length in the k-factor mapping
            valid_lengths = sorted(cmw_kfactor.keys())
            closest_length = min(valid_lengths, key=lambda x: abs(x - slot_length))
            k_factor = cmw_kfactor[closest_length]
            
            # Add M.A.R.V.E.L. checkbox before Extract Air Readings
            # st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)
            
            
            # Add section inputs for TAB readings if sections > 0
            if quantity > 0:
                st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)
                st.markdown("### Extract Air Readings")
                
                for i in range(quantity):
                    tab_reading_key = f'tab_reading_{selection}_{num}_{i}'
                    if tab_reading_key in st.session_state and not isinstance(st.session_state[tab_reading_key], (int, float)):
                        try:
                            st.session_state[tab_reading_key] = float(st.session_state[tab_reading_key])
                            if st.session_state[tab_reading_key] < 0.0:
                                st.session_state[tab_reading_key] = 0.0
                        except (ValueError, TypeError):
                            st.session_state[tab_reading_key] = 0.0
                    
                    tab_reading = st.number_input(f'Section {i+1} TAB Reading', 
                                                key=tab_reading_key,
                                                min_value=0.0,
                                                step=0.001,
                                                format="%.3f")
            
            # Add M.A.R.V.E.L. to checklist if enabled
            if marvel_enabled:
                hood.checklist['M.A.R.V.E.L. System'] = {}
        
        # Apply checklist
        # hood.checklist = checklist(selection, canopyLocation, drawingNum)
        
        return hood
    
    # Regular hood layout (non-CMW models)
    else:
        col1, col2, col3, col4 = st.columns(4)
        
        # Keep track of used drawing numbers
        if 'used_drawing_numbers' not in st.session_state:
            st.session_state.used_drawing_numbers = set()

        with col1:
            # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Drawing Number<h4>", unsafe_allow_html=True)
            drawingNum = st.text_input('Drawing Number', key=f'drawNum{num}')
            
            # After we have the drawing number, update the header
            if drawingNum:
                # Replace the header with one that includes the drawing number
                drawing_header = f"{selection} ({drawingNum}) Info"
                st.markdown(f"""
                <script>
                document.querySelector('h3:contains("{selection} Info")').innerText = '{drawing_header}';
                </script>
                """, unsafe_allow_html=True)
        with col2:
            # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy Location<h4>", unsafe_allow_html=True)
            canopyLocation = st.text_input('Canopy Location', key=f'canopyLoc{num}')
        with col3:
            # Design Flow
            design_flow_key = f'design_flow_{num}'
            if design_flow_key in st.session_state and not isinstance(st.session_state[design_flow_key], (float, int)):
                try:
                    st.session_state[design_flow_key] = float(st.session_state[design_flow_key])
                except:
                    st.session_state[design_flow_key] = 0.0
            
            # Add the label back
            # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Extract Design Airflow <h4>", unsafe_allow_html=True)
            
            # Instead of setting the session state value and then reading it,
            # just use a direct value
            total_design_flow_ms = st.number_input(f'Design Airflow', 
                                                key=design_flow_key, 
                                                min_value=0.0,
                                                step=0.001,
                                                value=0.0)
            
            # Show supply design flow for supply-capable hoods
            if selection in ['KVF', 'KCH-F', 'UVF', 'USR-F', 'KSR-F', 'KWF', 'UWF'] or (selection == 'CMW-FMOD'):
                with col3:
                    # Supply flow
                    supply_flow_key = f'total_supply_design_flow_ms{num}'
                    if supply_flow_key in st.session_state and not isinstance(st.session_state[supply_flow_key], (float, int)):
                        try:
                            st.session_state[supply_flow_key] = float(st.session_state[supply_flow_key])
                        except:
                            st.session_state[supply_flow_key] = 0.0
                        
                    # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Supply Design Airflow<h4>", unsafe_allow_html=True)
                    
                    # Instead of setting the session state value and then reading it,
                    # just use a direct value
                    total_supply_design_flow_ms = st.number_input('Supply Design Airflow', 
                                                                key=supply_flow_key, 
                                                                
                                                                min_value=0.0,
                                                                format="%.3f",
                                                                value=0.0)
            else:
                total_supply_design_flow_ms = 0.0
        with col4:
            # Ensure quantity is an integer
            quantity_key = f'quantity{num}'
            if quantity_key in st.session_state and not isinstance(st.session_state[quantity_key], int):
                try:
                    st.session_state[quantity_key] = int(float(st.session_state[quantity_key]))
                except:
                    st.session_state[quantity_key] = 0
                    
            # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy Sections<h4>", unsafe_allow_html=True)
            quantity = st.number_input('Number of Sections', key=quantity_key, min_value=0)

        hood = cjHoods(drawingNum, canopyLocation, selection, num, quantity, total_design_flow_ms, total_supply_design_flow_ms)

        # Retrieve system-specific checks
        hood.checklist = checklist(selection, canopyLocation, drawingNum)

        return hood

# Add CXW to hood types
def createCXWHood(selection, num):
    """
    Create a CXW hood object with anemometer readings instead of k-factors
    """
    st.markdown(f"<h3 style='text-align: center;margin-top: 30px;margin-bottom:-50px;'>{selection} Info<h3>", unsafe_allow_html=True)
    
    # Create a 2-column layout for basic info inputs
    col1, col2 = st.columns(2)
    
    with col1:
        # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Drawing Number<h4>", unsafe_allow_html=True)
        drawingNum = st.text_input(f'Drawing Number', key=f'drawing_num_{num}')
    
    with col2:
        # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy Location<h4>", unsafe_allow_html=True)
        canopyLocation = st.text_input(f'Canopy Location', key=f'location_{num}')
    
    # Second row with 3 columns for remaining inputs
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Ensure design flow is a float
        design_flow_key = f'design_flow_{num}'
        if design_flow_key in st.session_state and not isinstance(st.session_state[design_flow_key], (float, int)):
            try:
                st.session_state[design_flow_key] = float(st.session_state[design_flow_key])
            except:
                st.session_state[design_flow_key] = 0.0
                
        # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Design Flowrate<h4>", unsafe_allow_html=True)
        total_design_flow_ms = st.number_input(f'Design Flowrate', key=design_flow_key, min_value=0.0, format="%.3f", value=0.0)
    
    with col2:
        # Ensure quantity is an integer
        quantity_key = f'quantity_{num}'
        if quantity_key in st.session_state and not isinstance(st.session_state[quantity_key], int):
            try:
                st.session_state[quantity_key] = int(float(st.session_state[quantity_key]))
            except:
                st.session_state[quantity_key] = 1
                
        # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Quantity of Grills<h4>", unsafe_allow_html=True)
        quantity = st.number_input(f'Quantity of Grills', key=quantity_key, min_value=1)
        
    with col3:
        # Get grill size with default value for entire canopy (just once)
        grill_size_key = f'grill_size_{num}'
        default_size = "350x300"  # Default size
        
        # Set default in session state only if not already present
        if grill_size_key not in st.session_state:
            st.session_state[grill_size_key] = default_size
            
        # st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Grill Size (mm)<h4>", unsafe_allow_html=True)
        grill_size = st.text_input(f'Grill Size (mm)', key=grill_size_key)
    
    # Calculate free area based on grill size (hidden from UI)
    free_area = 0
    if grill_size:
        try:
            dimensions = [float(x) for x in grill_size.split('x')]
            # Calculate free area and format to exactly 3 decimal places
            free_area = (dimensions[0]/1000) * (dimensions[1]/1000)  # Convert mm to m²
            free_area = float(f"{free_area:.3f}")  # Format to 3 decimal places
        except:
            st.error(f"Invalid grill size format. Please use format like '350x300'")
    
    # Display header with drawing number if available
    marvel_enabled = st.checkbox(f"{selection} ({canopyLocation}) With M.A.R.V.E.L. System?", key=f"marvel_{selection}_{num}")
    if drawingNum:
        st.markdown(f"<h3 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>{selection} ({drawingNum}) - {canopyLocation}<h3>", unsafe_allow_html=True)
    
    # Add M.A.R.V.E.L. checkbox before Extract Air Readings

    
    # Create sections based on quantity
    sections = {}
    total_flowrate_m3s = 0
    
    # if quantity > 0:
        # st.markdown("### EXTRACT AIR READINGS")
        # st.markdown("<h4>Anemometer Reading (Average m/s)</h4>", unsafe_allow_html=True)
    
    # Display each grill's reading in a separate, clearly labeled section
    for i in range(quantity):
        st.markdown(f"#### Grill {i+1}")
        
        # Create columns for inputs - more columns if M.A.R.V.E.L. is enabled
        if marvel_enabled:
            cols = st.columns(4)  # 4 columns for anemometer, min%, idle%, design
        else:
            cols = st.columns(1)  # Just 1 column for anemometer reading
        
        with cols[0]:
            # Ensure anemometer reading is a float
            anem_key = f'anem_{num}_{i}'
            if anem_key in st.session_state and not isinstance(st.session_state[anem_key], (float, int)):
                try:
                    st.session_state[anem_key] = float(st.session_state[anem_key])
                except:
                    st.session_state[anem_key] = 0.0
            
            anemometer_reading = st.number_input(f'Anemometer Reading', key=anem_key, min_value=0.0, step=0.01, format="%g")
        
        # Add M.A.R.V.E.L. specific inputs if enabled
        marvel_data = {}
        if marvel_enabled:
            with cols[1]:
                min_pct_key = f'min_pct_{selection}_{i}_{num}'
                if min_pct_key in st.session_state and not isinstance(st.session_state[min_pct_key], (int, float)):
                    try:
                        st.session_state[min_pct_key] = float(st.session_state[min_pct_key])
                        if st.session_state[min_pct_key] < 0.0 or st.session_state[min_pct_key] > 100.0:
                            st.session_state[min_pct_key] = 30.0
                    except (ValueError, TypeError):
                        st.session_state[min_pct_key] = 30.0
                
                min_pct = st.number_input(f'Min (%)', 
                                        min_value=0.0, 
                                        max_value=100.0, 
                                        value=30.0, 
                                        step=0.1, 
                                        key=min_pct_key,
                                        format="%.1f")
                marvel_data['min_pct'] = min_pct
            
            with cols[2]:
                idle_pct_key = f'idle_pct_{selection}_{i}_{num}'
                if idle_pct_key in st.session_state and not isinstance(st.session_state[idle_pct_key], (int, float)):
                    try:
                        st.session_state[idle_pct_key] = float(st.session_state[idle_pct_key])
                        if st.session_state[idle_pct_key] < 0.0 or st.session_state[idle_pct_key] > 100.0:
                            st.session_state[idle_pct_key] = 70.0
                    except (ValueError, TypeError):
                        st.session_state[idle_pct_key] = 70.0
                                    
                idle_pct = st.number_input(f'Idle (%)', 
                                         min_value=0.0, 
                                         max_value=100.0, 
                                         value=70.0, 
                                         step=0.1, 
                                         key=idle_pct_key,
                                         format="%.1f")
                marvel_data['idle_pct'] = idle_pct
            
            with cols[3]:
                design_flow_key = f'design_flow_{selection}_{i}_{num}'
                if design_flow_key in st.session_state and not isinstance(st.session_state[design_flow_key], (int, float)):
                    try:
                        st.session_state[design_flow_key] = float(st.session_state[design_flow_key])
                        if st.session_state[design_flow_key] < 0.0:
                            st.session_state[design_flow_key] = 0.0
                    except (ValueError, TypeError):
                        st.session_state[design_flow_key] = 0.0
                                    
                design_flow = st.number_input(f'Design (m³/s)', 
                                            min_value=0.0, 
                                            value=0.0, 
                                            step=0.001, 
                                            key=design_flow_key,
                                            format="%.3f")
                marvel_data['design_flow'] = design_flow
        
        # Calculate flowrates with precise formatting
        flowrate_m3h = int(free_area * anemometer_reading * 3600)  # m³/h, no decimal places
        
        # Calculate m³/s directly (not from m³/h) and format to exactly 3 decimal places
        flowrate_m3s = free_area * anemometer_reading  # Direct calculation m³/s
        flowrate_m3s = float(f"{flowrate_m3s:.3f}")  # Format to exactly 3 decimal places
        
        total_flowrate_m3s += flowrate_m3s
        
        # Store section data
        section_data = {
            "grill_size": grill_size,
            "free_area": free_area,
            "anemometer_reading": anemometer_reading,
            "flowrate_m3h": flowrate_m3h,
            "flowrate_m3s": flowrate_m3s
        }
        
        # Add M.A.R.V.E.L. data if enabled
        if marvel_enabled:
            section_data.update(marvel_data)
        
        sections[f"{i+1}"] = section_data
        
        st.markdown("---")
    
    # Format total flowrate to exactly 3 decimal places
    total_flowrate_m3s = float(f"{total_flowrate_m3s:.3f}")
    
    # Calculate percentage of design
    percentage_of_design = 0
    if total_design_flow_ms > 0:
        percentage_of_design = (total_flowrate_m3s / total_design_flow_ms) * 100
        percentage_of_design = float(f"{percentage_of_design:.1f}")  # Format to 1 decimal place
    
    # Display the summary information
    # if quantity > 0:
    #     st.markdown("### Summary Information")
    #     st.markdown(f"**Total Flow Rate: {total_flowrate_m3s:.3f} m³/s**")
    #     st.markdown(f"**Percentage of Design: {percentage_of_design}%**")
    #     st.markdown(f"**Calculation Method: Qv = A x v**")
    #     st.markdown("*Where: Qv = Flow rate, A = Free area, v = Air velocity*")
    
    # Create the hood object
    hood = CXWHood(drawingNum, canopyLocation, selection, num, quantity, total_design_flow_ms, sections)
    hood.total_flowrate_m3s = total_flowrate_m3s
    hood.percentage_of_design = percentage_of_design
    
    # Add M.A.R.V.E.L. to checklist if enabled
    if marvel_enabled:
        hood.checklist['M.A.R.V.E.L. System'] = {}
    
    return hood

class CXWHood:
    def __init__(self, drawingNum, location, model, num, quantity, total_design_flow_ms, sections):
        self.drawingNum = drawingNum
        self.location = location
        self.model = model
        self.num = num
        self.quantityOfSections = quantity
        self.total_design_flow_ms = total_design_flow_ms
        self.sections = sections
        self.checklist = {}  # Add this for consistency with other hood types
        self.total_flowrate_m3s = 0  # Will be set by createCXWHood
        self.percentage_of_design = 0  # Will be set by createCXWHood

    def getCJSections(self, selection, num):
        """
        This method exists for compatibility with the hood interface,
        but for CXW hoods, the sections are already created during initialization.
        """
        # Check if MARVEL is enabled for this hood
        marvel_enabled = 'M.A.R.V.E.L. System' in self.checklist
        
        if marvel_enabled and self.quantityOfSections > 0:
            # Calculate default design flow per section
            design_flow_per_section = self.total_design_flow_ms / self.quantityOfSections if self.quantityOfSections > 0 else 0
            design_flow_per_section = round(design_flow_per_section, 3)
            
            # Update each section with M.A.R.V.E.L. data
            for i in range(self.quantityOfSections):
                section_key = str(i + 1)
                if section_key in self.sections:
                    # Add M.A.R.V.E.L. data with default values if not present
                    if 'min_pct' not in self.sections[section_key]:
                        self.sections[section_key]['min_pct'] = 30.0
                    if 'idle_pct' not in self.sections[section_key]:
                        self.sections[section_key]['idle_pct'] = 70.0
                    if 'design_flow' not in self.sections[section_key]:
                        self.sections[section_key]['design_flow'] = 0.0
        
    # Helper method to format flowrate to exactly 3 decimal places
    def format_flowrate(self, value):
        """Format a value to exactly 3 decimal places as a float"""
        return float(f"{value:.3f}")
        