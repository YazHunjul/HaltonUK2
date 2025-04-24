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
        st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)

        for i in range(self.quantityOfSections):
            drawing_info = f"({self.drawingNum})" if self.drawingNum else f"#{num+1}"
            st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>{selection} {drawing_info} Section {i+1} Extract Info ({self.location})<h4>", unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)

            if selection in ['KVF', 'KCH-F', 'UVF', 'USR-F', 'KSR-F']:
                st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>{selection} {drawing_info} Section {i+1} Supply Info ({self.location})<h4>", unsafe_allow_html=True)
                col3, col4 = st.columns(2)

            with col1:
                st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>Section {i+1} KSA's<h4>", unsafe_allow_html=True)
                # Ensure KSA quantity is an integer
                ksa_key = f'ksaQuantity{selection}{num}{i}'
                if ksa_key in st.session_state and not isinstance(st.session_state[ksa_key], int):
                    try:
                        st.session_state[ksa_key] = int(float(st.session_state[ksa_key]))
                    except:
                        st.session_state[ksa_key] = 1
                self.ksaQuantity = st.number_input('.', key=ksa_key, label_visibility='hidden', min_value=1, max_value=6)
            
            with col2:
                st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} T.A.B Point Reading<h4>", unsafe_allow_html=True)
                # Ensure tab reading is a float
                tab_key = f'canopyLocs{selection}{num}{i}'
                if tab_key in st.session_state and not isinstance(st.session_state[tab_key], (float, int)):
                    try:
                        st.session_state[tab_key] = float(st.session_state[tab_key])
                    except:
                        st.session_state[tab_key] = 0.0
                self.tab_Reading = st.number_input('.', key=tab_key, label_visibility='hidden', min_value=0.0)

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
            if selection in ['KVF', 'KCH-F', 'UVF', 'USR-F', 'KSR-F', 'KWF', 'UWF', 'CMW-FMOD']:
                col3, col4 = st.columns(2)
                with col3:
                    st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>Section {i+1} Plenum Length<h4>", unsafe_allow_html=True)
                    # Ensure plenum length selection has a valid initial value
                    plenum_key = f'plenumLength{selection}{num}{i}'
                    if plenum_key in st.session_state and not isinstance(st.session_state[plenum_key], int):
                        try:
                            # Try to convert to integer if it's a valid value in the list
                            val = int(float(st.session_state[plenum_key]))
                            if val in [1000, 1500, 2000, 2500, 3000]:
                                st.session_state[plenum_key] = val
                            else:
                                st.session_state[plenum_key] = 1000  # Default
                        except:
                            st.session_state[plenum_key] = 1000  # Default if conversion fails
                    
                    self.plenumLength = st.selectbox('.', [1000, 1500, 2000, 2500, 3000], key=plenum_key, label_visibility='hidden')
                
                supplyKFactor = cj_hoods_supply_kfactor[self.plenumLength]
                
                with col4:
                    st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} T.A.B Point Reading<h4>", unsafe_allow_html=True)
                    # Ensure supply tab reading is a float
                    supply_tab_key = f'sup_tab{selection}{num}{i}'
                    if supply_tab_key in st.session_state and not isinstance(st.session_state[supply_tab_key], (float, int)):
                        try:
                            st.session_state[supply_tab_key] = float(st.session_state[supply_tab_key])
                        except:
                            st.session_state[supply_tab_key] = 0.0
                            
                    self.tab_Reading_supply = st.number_input('.', key=supply_tab_key, label_visibility='hidden', min_value=0.0)
                
                # Ensure supply tab reading is a float for calculation
                supply_tab_float = float(self.tab_Reading_supply)
                supplyAchievedFlowRate = supplyKFactor * math.sqrt(supply_tab_float)
                
                self.sections[f'{i+1}'] = {
                    "ksaQuantity": self.ksaQuantity,
                    "k_factor": self.k_factor,
                    'tab_reading': self.tab_Reading,
                    'achieved': round(achievedFlowRate, 2),
                    'supplyKFactor': supplyKFactor,
                    'supplyTab': self.tab_Reading_supply,
                    'supplyDesign': round(self.design_supply, 2),
                    'achievedSupply': round(supplyAchievedFlowRate, 2)
                }
            else:
                self.sections[f'{i+1}'] = {
                    "ksaQuantity": self.ksaQuantity,
                    "k_factor": self.k_factor,
                    'tab_reading': self.tab_Reading,
                    'achieved': round(achievedFlowRate, 2)
                }


def numCanopies():
    hoods = {}
    st.markdown("<h1 style='text-align: center;margin-top: 30px;margin-bottom:-100px;'>Enter Number Of Canopies<h1>", unsafe_allow_html=True)
    
    # Ensure num_canopies is an integer
    if 'num_canopies' in st.session_state and not isinstance(st.session_state['num_canopies'], int):
        try:
            # Try to convert to int
            st.session_state['num_canopies'] = int(float(st.session_state['num_canopies']))
        except:
            # Default to 0 if conversion fails
            st.session_state['num_canopies'] = 0
    
    # Use consistent key for number of canopies to help with state restoration
    numInput = st.number_input('.', label_visibility='hidden', min_value=0, key="num_canopies")
    
    for num in range(numInput):
        # Add an anchor for this canopy
        st.markdown(f'<a name="canopy_{num}"></a>', unsafe_allow_html=True)
        # Include canopy number in the heading
        st.markdown(f"<h3 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy #{num+1}</h3>", unsafe_allow_html=True)
        selection = st.selectbox('.', cj_hoods + cj_lp + cr_hoods + cr_lp_hoods + ww_hoods + new_hoods + ['CXW'], key=f'selection{num}', label_visibility='hidden')
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
        st.write("M.A.R.V.E.L. System Checklist")
        col1, col2 = st.columns(2)
        with col1:
            # Ensure min_airflow is a number
            min_airflow_key = f'minAirflow_marvel_{hood}_{location}'
            if min_airflow_key in st.session_state and not isinstance(st.session_state[min_airflow_key], (int, float)):
                try:
                    st.session_state[min_airflow_key] = float(st.session_state[min_airflow_key])
                except:
                    st.session_state[min_airflow_key] = 0
            min_airflow = st.number_input("Canopy 'Minimum' Air Flow Set Point % of Design", min_value=0, key=min_airflow_key)
            
            # Ensure idle_airflow is a number
            idle_airflow_key = f'idleAirflow_marvel_{hood}_{location}'
            if idle_airflow_key in st.session_state and not isinstance(st.session_state[idle_airflow_key], (int, float)):
                try:
                    st.session_state[idle_airflow_key] = float(st.session_state[idle_airflow_key])
                except:
                    st.session_state[idle_airflow_key] = 0
            idle_airflow = st.number_input("Canopy 'Idle' Air Flow Set % Of Design", min_value=0, key=idle_airflow_key)
            
            # Ensure cook_airflow is a number
            cook_airflow_key = f'cookAirflow_marvel_{hood}_{location}'
            if cook_airflow_key in st.session_state and not isinstance(st.session_state[cook_airflow_key], (int, float)):
                try:
                    st.session_state[cook_airflow_key] = float(st.session_state[cook_airflow_key])
                except:
                    st.session_state[cook_airflow_key] = 0
            cook_airflow = st.number_input("Canopy 'Cook' Air Flow Set % Of Design", min_value=0, key=cook_airflow_key)
            
            # Ensure cook_time is a number
            cook_time_key = f'cookTime_marvel_{hood}_{location}'
            if cook_time_key in st.session_state and not isinstance(st.session_state[cook_time_key], (int, float)):
                try:
                    st.session_state[cook_time_key] = int(float(st.session_state[cook_time_key]))
                except:
                    st.session_state[cook_time_key] = 0
            cook_time = st.number_input("Cook Mode Run Time (Seconds)", min_value=0, key=cook_time_key)
            
            # Ensure override_temp is a number
            override_temp_key = f'overrideTemp_marvel_{hood}_{location}'
            if override_temp_key in st.session_state and not isinstance(st.session_state[override_temp_key], (int, float)):
                try:
                    st.session_state[override_temp_key] = float(st.session_state[override_temp_key])
                except:
                    st.session_state[override_temp_key] = 0
            override_temp = st.number_input("M.A.R.V.E.L. System Override Temp Setpoint (°C)", min_value=0, key=override_temp_key)
        with col2:
            marvel_min = st.checkbox("M.A.R.V.E.L. System Tested in 'Minimum' Mode", value=True, key=f'marvelMin_marvel_{hood}_{location}')
            marvel_idle = st.checkbox("M.A.R.V.E.L. System Tested in 'Idle' Mode", value=True, key=f'marvelIdle_marvel_{hood}_{location}')
            marvel_cook = st.checkbox("M.A.R.V.E.L. System Tested in 'Cook' Mode", value=True, key=f'marvelCook_marvel_{hood}_{location}')
            lir_2_test = st.checkbox("Testing of Canopy LIR-2 Sensors Successful", value=True, key=f'lir2_marvel_{hood}_{location}')
            hmi_override = st.checkbox("Override Test on Halton HMI Successful", value=True, key=f'hmiOverride_marvel_{hood}_{location}')
            auto_balance = st.checkbox("Auto-Balance of M.A.R.V.E.L. System Successful", value=True, key=f'autoBalance_marvel_{hood}_{location}')
        system_checks['M.A.R.V.E.L. System'] = {
            "Canopy 'Minimum' Air Flow Set Point % of Design": min_airflow,
            "Canopy 'Idle' Air Flow Set % Of Design": idle_airflow,
            "Canopy 'Cook' Air Flow Set % Of Design": cook_airflow,
            "M.A.R.V.E.L. System Tested in 'Minimum' Mode": marvel_min,
            "M.A.R.V.E.L. System Tested in 'Idle' Mode": marvel_idle,
            "M.A.R.V.E.L. System Tested in 'Cook' Mode": marvel_cook,
            "Cook Mode Run Time (Seconds)": cook_time,
            "Testing of Canopy LIR-2 Sensors Successful": lir_2_test,
            "M.A.R.V.E.L. System Override Temp Setpoint (°C)": override_temp,
            "Override Test on Halton HMI Successful": hmi_override,
            "Auto-Balance of M.A.R.V.E.L. System Successful": auto_balance,
        }

    return system_checks


        

def createCJHood(selection, num):
    """
    Create a hood object and retrieve system-specific checks (UV or M.A.R.V.E.L.).
    """
    st.markdown(f"<h3 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>{selection} Info<h3>", unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns(4)
    
    # Keep track of used drawing numbers
    if 'used_drawing_numbers' not in st.session_state:
        st.session_state.used_drawing_numbers = set()

    with col1:
        st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Drawing Number<h4>", unsafe_allow_html=True)
        drawingNum = st.text_input('.', key=f'drawNum{num}', label_visibility='hidden')
        
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
        st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy Location<h4>", unsafe_allow_html=True)
        canopyLocation = st.text_input('.', key=f'canopyLoc{num}', label_visibility='hidden')
    with col3:
        # Ensure the number input receives a float
        flow_key = f'total_design_flow_ms{num}'
        if flow_key in st.session_state and not isinstance(st.session_state[flow_key], (float, int)):
            try:
                # Try to convert to float if it's a string
                st.session_state[flow_key] = float(st.session_state[flow_key])
            except:
                # If conversion fails, reset to 0.0
                st.session_state[flow_key] = 0.0
                
        st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Extract Design Airflow (m³/s)<h4>", unsafe_allow_html=True)
        total_design_flow_ms = st.number_input('.', key=flow_key, label_visibility='hidden', min_value=0.0)
        
        # Only show supply design flow for supply-capable hoods
        if selection in ['KVF', 'KCH-F', 'UVF', 'USR-F', 'KSR-F', 'KWF', 'UWF', 'CMW-FMOD']:
            # Ensure the supply flow number input receives a float
            supply_flow_key = f'total_supply_design_flow_ms{num}'
            if supply_flow_key in st.session_state and not isinstance(st.session_state[supply_flow_key], (float, int)):
                try:
                    st.session_state[supply_flow_key] = float(st.session_state[supply_flow_key])
                except:
                    st.session_state[supply_flow_key] = 0.0
                    
            st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Supply Design Airflow (m³/s)<h4>", unsafe_allow_html=True)
            total_supply_design_flow_ms = st.number_input('.', key=supply_flow_key, label_visibility='hidden', min_value=0.0)
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
                
        st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy Sections<h4>", unsafe_allow_html=True)
        quantity = st.number_input('.', key=quantity_key, label_visibility='hidden', min_value=0)

    hood = cjHoods(drawingNum, canopyLocation, selection, num, quantity, total_design_flow_ms, total_supply_design_flow_ms)

    # Retrieve system-specific checks
    hood.checklist = checklist(selection, canopyLocation, drawingNum)

    return hood

# Add CXW to hood types
def createCXWHood(selection, num):
    """
    Create a CXW hood object with anemometer readings instead of k-factors
    """
    st.markdown(f"<h3 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>{selection} #{num+1} Info<h3>", unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        drawingNum = st.text_input('Drawing Number', key=f'drawing_num_{num}')
    with col2:
        canopyLocation = st.text_input('Location', key=f'location_{num}')
    with col3:
        # Ensure design flow is a float
        design_flow_key = f'design_flow_{num}'
        if design_flow_key in st.session_state and not isinstance(st.session_state[design_flow_key], (float, int)):
            try:
                st.session_state[design_flow_key] = float(st.session_state[design_flow_key])
            except:
                st.session_state[design_flow_key] = 0.0
                
        total_design_flow_ms = st.number_input('Design Flowrate (m³/s)', key=design_flow_key, min_value=0.0)
    with col4:
        # Ensure quantity is an integer
        quantity_key = f'quantity_{num}'
        if quantity_key in st.session_state and not isinstance(st.session_state[quantity_key], int):
            try:
                st.session_state[quantity_key] = int(float(st.session_state[quantity_key]))
            except:
                st.session_state[quantity_key] = 1
                
        quantity = st.number_input('Quantity of Grills', key=quantity_key, min_value=1)
    
    # Create sections based on quantity
    sections = {}
    for i in range(quantity):
        st.markdown(f"##### Section {i+1}")
        col1, col2 = st.columns(2)
        
        with col1:
            grill_size = st.text_input(f'Grill Size (mm)', key=f'grill_size_{num}_{i}')
            # Calculate free area based on grill size (assuming square)
            if grill_size:
                try:
                    dimensions = [float(x) for x in grill_size.split('x')]
                    free_area = (dimensions[0]/1000) * (dimensions[1]/1000)  # Convert mm to m
                except:
                    free_area = 0
            else:
                free_area = 0
                
        with col2:
            # Ensure anemometer reading is a float
            anem_key = f'anem_{num}_{i}'
            if anem_key in st.session_state and not isinstance(st.session_state[anem_key], (float, int)):
                try:
                    st.session_state[anem_key] = float(st.session_state[anem_key])
                except:
                    st.session_state[anem_key] = 0.0
                    
            anemometer_reading = st.number_input(f'Anemometer Reading (m/s)', key=anem_key, min_value=0.0)
            
        # Calculate flowrate
        flowrate_m3h = free_area * anemometer_reading * 3600  # Convert to m³/h
        flowrate_m3s = flowrate_m3h / 3600  # Convert to m³/s
        
        sections[f"Section {i+1}"] = {
            "grill_size": grill_size,
            "free_area": free_area,
            "anemometer_reading": anemometer_reading,
            "flowrate_m3h": flowrate_m3h,
            "flowrate_m3s": flowrate_m3s
        }
    
    hood = CXWHood(drawingNum, canopyLocation, selection, num, quantity, total_design_flow_ms, sections)
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

    def getCJSections(self, selection, num):
        """
        This method exists for compatibility with the hood interface,
        but for CXW hoods, the sections are already created during initialization
        """
        pass  # All section data is already collected in __init__
        