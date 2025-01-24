import streamlit as st
import math

cj_hoods = ['KVF', 'KVI', 'KCH-F', 'KCH-I']
cj_lp = ['KSR-S', 'KSR-F', 'KSR-M']
cr_hoods = ['UVF', 'UVI']
cr_lp_hoods = ['USR-S', 'USR-F', 'USR-M']
ww_hoods = ['KWF', 'KWI', 'UWF', 'UWI', 'CMW-FMOD', 'CMW-IMOD']
new_hoods = ['XYZ1', 'XYZ2', 'XYZ3']  # Add the new hoods here

cj_hoods_k_factor = {
    1: 71.8,
    2: 143.6,
    3: 215.4,
    4: 287.2,
    5: 359,
    6: 430.8
}

cj_lp_k_factor = {
    1: 67.7,
    2: 135.4,
    3: 203.1,
    4: 270.8,
    5: 338.5,
    6: 406.2
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
    1: 67.7,
    2: 135.4,
    3: 203.1,
    4: 270.8,
    5: 338.5,
    6: 406.2
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
            st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>{self.model} Section {i+1} Extract Info ({self.location})<h4>", unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)

            if selection in ['KVF', 'KCH-F', 'UVF', 'USR-F', 'KSR-F']:
                st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>{self.model} Section {i+1} Supply Info ({self.location})<h4>", unsafe_allow_html=True)
                col3, col4 = st.columns(2)

            with col1:
                st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>Section {i+1} KSA's<h4>", unsafe_allow_html=True)
                self.ksaQuantity = st.number_input('.', key=f'ksaQuantity{selection}{num}{i}', label_visibility='hidden', min_value=1, max_value=6)
            with col2:
                st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} T.A.B Point Reading<h4>", unsafe_allow_html=True)
                self.tab_Reading = st.number_input('.', key=f'canopyLocs{selection}{num}{i}', label_visibility='hidden', min_value=0.0)
            # with col1:
            #     st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Section {i+1} Flowrate (m^3/h)<h4>", unsafe_allow_html=True)
            #     self.design_flow = st.number_input('.', key=f'designflow{selection}{num}{i}', label_visibility='hidden', min_value=0.0)

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

            achievedFlowRate = self.k_factor * math.sqrt(self.tab_Reading)

            # Handle supply air logic if applicable
            if selection in ['KVF', 'KCH-F', 'UVF', 'USR-F', 'KSR-F', 'KWF', 'UWF', 'CMW-FMOD']:
                col3, col4 = st.columns(2)
                with col3:
                    st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>Section {i+1} Plenum Length<h4>", unsafe_allow_html=True)
                    self.plenumLength = st.selectbox('.', [1000, 1500, 2000, 2500, 3000], key=f'plenumLength{selection}{num}{i}', label_visibility='hidden')
                supplyKFactor = cj_hoods_supply_kfactor[self.plenumLength]
                with col4:
                    st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} T.A.B Point Reading<h4>", unsafe_allow_html=True)
                    self.tab_Reading_supply = st.number_input('.', key=f'sup_tab{selection}{num}{i}', label_visibility='hidden', min_value=0.0)
                # with col3:
                #     st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} Design Flow <h4>", unsafe_allow_html=True)
                #     self.design_supply = st.number_input('.', key=f'design_sup{selection}{num}{i}', label_visibility='hidden', min_value=0.0)
                supplyAchievedFlowRate = supplyKFactor * math.sqrt(self.tab_Reading_supply)
                self.sections[f'{i+1}'] = {
                    "ksaQuantity": self.ksaQuantity,
                    "k_factor": self.k_factor,
                    'tab_reading': self.tab_Reading,
                    # 'designFlow': self.design_flow,
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
                    'designFlow': self.design_flow,
                    'achieved': round(achievedFlowRate, 2)
                }


def numCanopies():
    hoods = {}
    st.markdown("<h1 style='text-align: center;margin-top: 30px;margin-bottom:-100px;'>Enter Number Of Canopies<h1>", unsafe_allow_html=True)
    numInput = st.number_input('.', label_visibility='hidden', min_value=0)
    for num in range(numInput):
        st.markdown(f"<h3 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy {num+1}<h3>", unsafe_allow_html=True)
        selection = st.selectbox('.', cj_hoods + cj_lp + cr_hoods + cr_lp_hoods + ww_hoods + new_hoods, key=f'selection{num}', label_visibility='hidden')
        hood = createCJHood(selection, num)
        hood.getCJSections(selection, num)
        hoods[f'{selection} {num}'] = hood
        st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)

    edge_box_checked = st.checkbox("Final System Checks: Edge Box")
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
        edge_installed = st.checkbox("Edge Installed")
        edge_id = st.text_input("Edge ID", key="edge_id", value="")
        edge_4g_status = st.selectbox("Edge 4G Status", ["", "Online", "Offline"], key="edge_4g_status")
    with col2:
        modbus_operation = st.checkbox("Modbus Operation")
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


def checklist(hood, location):
    """
    Checklist for UV Capture Jet or M.A.R.V.E.L. System.
    - "With UV" is only shown for hoods starting with 'UV'.
    - "With M.A.R.V.E.L." is shown for all hoods.
    """
    system_checks = {}

    # UV Capture Jet System Checklist (only for hoods starting with 'UV')
    if hood.startswith("UV"):
        with_uv = st.checkbox(f"{(hood.split(' '))[0]} ({location}) With UV Capture Jet?", key=f"uv_{hood}_{location}")
        if with_uv:
            features["CJ"] = "yes"  # Add UV system to features
            st.write("UV Capture Ray System Checklist")
            col1, col2 = st.columns(2)
            with col1:
                quantity = st.number_input("Quantity of Slaves per System", min_value=0, key=f'quantity_uv_{hood}_{location}')
                pressure = st.number_input("UV Pressure Setpoint (Pa)", min_value=0, key=f'pressure_uv_{hood}_{location}')
                cj_reading = st.number_input("Capture Jet Average Pressure Reading (Pa)", min_value=0, key=f'cjreading_uv_{hood}_{location}')
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
    with_marvel = st.checkbox(f"{(hood.split(' '))[0]} ({location}) With M.A.R.V.E.L. System?", key=f"marvel_{hood}_{location}")
    if with_marvel:
        features["marvel"] = "yes"  # Add M.A.R.V.E.L. system to features
        st.write("M.A.R.V.E.L. System Checklist")
        col1, col2 = st.columns(2)
        with col1:
            min_airflow = st.number_input("Canopy 'Minimum' Air Flow Set Point % of Design", min_value=0, key=f'minAirflow_marvel_{hood}_{location}')
            idle_airflow = st.number_input("Canopy 'Idle' Air Flow Set % Of Design", min_value=0, key=f'idleAirflow_marvel_{hood}_{location}')
            cook_airflow = st.number_input("Canopy 'Cook' Air Flow Set % Of Design", min_value=0, key=f'cookAirflow_marvel_{hood}_{location}')
            cook_time = st.number_input("Cook Mode Run Time (Seconds)", min_value=0, key=f'cookTime_marvel_{hood}_{location}')
            override_temp = st.number_input("M.A.R.V.E.L. System Override Temp Setpoint (°C)", min_value=0, key=f'overrideTemp_marvel_{hood}_{location}')
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
    with col1:
        st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Drawing Number<h4>", unsafe_allow_html=True)
        drawingNum = st.text_input('.', key=f'drawNum{num}', label_visibility='hidden')
    with col2:
        st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy Location<h4>", unsafe_allow_html=True)
        canopyLocation = st.text_input('.', key=f'canopyLoc{num}', label_visibility='hidden')
    with col3:
        st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Extract Design Airflow (m³/s)<h4>", unsafe_allow_html=True)
        total_design_flow_ms = st.number_input('.', key=f'total_design_flow_ms{num}', label_visibility='hidden', min_value=0.0)
        
        # Only show supply design flow for supply-capable hoods
        if selection in ['KVF', 'KCH-F', 'UVF', 'USR-F', 'KSR-F', 'KWF', 'UWF', 'CMW-FMOD']:
            st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Supply Design Airflow m³/s)<h4>", unsafe_allow_html=True)
            total_supply_design_flow_ms = st.number_input('.', key=f'total_supply_design_flow_ms{num}', label_visibility='hidden', min_value=0.0)
        else:
            total_supply_design_flow_ms = 0.0
    with col4:
        st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy Sections<h4>", unsafe_allow_html=True)
        quantity = st.number_input('.', key=f'quantity{num}', label_visibility='hidden', min_value=0)

    hood = cjHoods(drawingNum, canopyLocation, selection, num, quantity, total_design_flow_ms, total_supply_design_flow_ms)

    # Retrieve system-specific checks
    hood.checklist = checklist(selection, canopyLocation)

    return hood
        