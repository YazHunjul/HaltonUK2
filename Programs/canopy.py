import streamlit as st
import math
cj_hoods = ['KVF', 'KVI', 'KCH-F', 'KCH-I']
cj_lp = ['KSR-S', 'KSR-F', 'KSR-M']
cr_hoods = ['UVF', 'UVI']
cr_lp_hoods = ['USR-S', 'USR-F', 'USR-M']

cj_hoods_k_factor = {
    1 : 71.8,
    2: 143.6,
    3 : 215.4,
    4 : 287.2,
    5 : 359,
    6 : 430.8
}

cj_lp_k_factor = {
    1 : 67.7,
    2: 135.4,
    3 : 203.1,
    4 : 270.8,
    5 : 338.5,
    6 : 406.2
}

cj_hoods_supply_kfactor = {
    1000 : 121.7,
    1500 : 182.6,
    2000 : 243.4,
    2500 : 304.2,
    3000 : 365.1
}

cr_hoods_k_factor = {
    1: 71.8,
    2: 143.6,
    3 : 215.4,
    4 : 287.2,
    5 : 359,
    6: 430.8
    }

cr_lp_k_factor = {
    1: 67.7,
    2: 135.4,
    3 : 203.1,
    4 : 270.8,
    5 : 338.5,
    6: 406.2
    }

ww_hoods = ['KWF', 'KWI', "UWF", "UWI", "CMW-FMOD", "CMW-IMOD"]

ww_hoods_kfactor = {
    1 : 65.5,
    2 : 131,
    3 : 196.5,
    4 : 262,
    5 : 327.5,
    6 : 393
}

features = {

}

class cjHoods():
    def __init__(self, drawingNum, location, model, idNum,  quantityOfSections) -> None:
        self.drawingNum = drawingNum
        self.location = location
        self.model = model
        self.quantityOfSections = quantityOfSections
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
        if self.model.startswith('UV'):
            self.checklist = checklist(f'{selection} {self.id}', self.location)
            st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)
        for i in range(self.quantityOfSections):
            st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>{self.model} Section {i+1} Extract Info ({self.location})<h4>", unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)
            if selection in ['KVF', 'KCH-F', 'UVF', 'USR-F','KSR-F']:
                st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>{self.model} Section {i+1} Supply Info ({self.location})<h4>", unsafe_allow_html=True)
                col3, col4 = st.columns(2)

            
            with col1:
                st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>Section {i+1} KSA's<h4>", unsafe_allow_html=True)
                self.ksaQuantity = st.number_input('.', key=f'ksaQuantity{selection}{num}{i}', label_visibility='hidden', min_value=1, max_value=6)
            with col2:
                st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} T.A.B Point Reading<h4>", unsafe_allow_html=True)
                self.tab_Reading = st.number_input('.', key=f'canopyLocs{selection}{num}{i}', label_visibility='hidden', min_value=0.0)
            with col1:
                st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Section {i+1} Design Airflow<h4>", unsafe_allow_html=True)
                self.design_flow = st.number_input('.', key=f'designflow{selection}{num}{i}', label_visibility='hidden', min_value=0.0)
            #CJ HOODS KFACTOR
            if selection in cj_hoods:   
                self.k_factor = cj_hoods_k_factor[self.ksaQuantity]
            #CJ LP HOODS KFACTOR
            elif selection in cj_lp:
                self.k_factor = cj_lp_k_factor[self.ksaQuantity]
            #CR HOODS
            elif selection in cr_hoods:
                self.k_factor = cr_hoods_k_factor[self.ksaQuantity]
            elif selection in cr_lp_hoods:
                self.k_factor = cr_lp_k_factor[self.ksaQuantity]
            elif selection in ww_hoods:
                self.k_factor = ww_hoods_kfactor[self.ksaQuantity]
                
            achievedFlowRate = self.k_factor * math.sqrt(self.tab_Reading)
            #IF FRESH AIR
            if selection in ['KVF', 'KCH-F', 'UVF', 'KCH-F', 'USR-F', 'KSR-F', 'KWF', 'UWF', 'CMW-FMOD']:
                col3, col4 = st.columns(2)
                with col3:
                    st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>Section {i+1} Plenum Length<h4>", unsafe_allow_html=True)
                    self.plenumLength = st.selectbox('.', [1000, 1500, 2000, 2500, 3000],key=f'plenumLength{selection}{num}{i}', label_visibility='hidden')
                supplyKFactor = cj_hoods_supply_kfactor[self.plenumLength]
                with col4:
                    st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} T.A.B Point Reading<h4>", unsafe_allow_html=True)
                    self.tab_Reading_supply = st.number_input('.', key=f'sup_tab{selection}{num}{i}', label_visibility='hidden', min_value=0.0)
                with col3:
                    st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} Design Flow <h4>", unsafe_allow_html=True)
                    self.design_supply = st.number_input('.', key=f'design_sup{selection}{num}{i}', label_visibility='hidden', min_value=0.0)
                supplyAchievedFlowRate = cj_hoods_supply_kfactor[self.plenumLength] * math.sqrt(self.tab_Reading_supply)
                self.sections[f'{i+1}'] = {"ksaQuantity" : self.ksaQuantity, "k_factor" : self.k_factor, 'tab_reading' : self.tab_Reading, 'designFlow' : self.design_flow, 'achieved' : round(achievedFlowRate,2), 'supplyKFactor' : supplyKFactor, 'supplyTab' : self.tab_Reading_supply, 'supplyDesign' : round(self.design_supply,2), 'achievedSupply' : round(supplyAchievedFlowRate,2)}
            else:
                self.sections[f'{i+1}'] = {"ksaQuantity" : self.ksaQuantity, "k_factor" : self.k_factor, 'tab_reading' : self.tab_Reading, 'designFlow' : self.design_flow, 'achieved' : round(achievedFlowRate,2)}
                

def numCanopies():
    hoods = {}
    st.markdown("<h1 style='text-align: center;margin-top: 30px;margin-bottom:-100px;'>Enter Number Of Canopies<h1>", unsafe_allow_html=True)
    numInput = st.number_input('.', label_visibility='hidden', min_value=0)
    for num in range(numInput):
        st.markdown(f"<h3 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy {num+1}<h3>", unsafe_allow_html=True)
        selection = st.selectbox('.',['','KVF', "KVI", "KCH-F", "KCH-I", "KSR-S", "KSR-F", "KSR-M", "UVI", "UVF", "USR-S", "USR-F", "USR-M", 'KWF', 'KWI', "UWF", "UWI", "CMW-FMOD", "CMW-IMOD"],key=f'selection{num}', label_visibility='hidden')
        #CJ HOODS
        hood = createCJHood(selection,num)
        hood.getCJSections(selection,num)

        hoods[f'{selection} {num}'] = hood
    
            
        st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)
    return hoods


def checklist(hood, location):
    if hood.startswith("UV"):
        with_cj = st.checkbox(f"{(hood.split(' '))[0]} ({location}) With Capture Jet?")
        if with_cj:
            features['CJ'] = "yes"

            st.write("UV Capture Ray System")
            col1, col2,col3, col4 = st.columns(4)
            with col1:
                quantity = st.number_input("Quantity of Slaves per System", min_value=0, key=f'quantity{hood}')
            with col2:
                pressure = st.number_input("UV Pressure Setpoint (Pa)",min_value=0, key=f'pressure{hood}')
            with col3:
                airflow = st.checkbox("Airflow Proved on Each Controller", value=True, key=f'airflow{hood}')
            with col4:
                safety = st.checkbox('All UV Safety Switches Tested', value=True, key=f'safety{hood}')
            with col1:
                comms = st.checkbox("Communication To All Canopy controllers", value=True, key=f'comms{hood}')
            with col2:
                filters = st.checkbox("All Filter Safety Switches Tested", value=True, key=f'filters{hood}')
            with col3:
                ops = st.checkbox("UV System Tested and Fully Operational", value=True, key=f'ops{hood}')
            with col4:
                cj = st.checkbox("Capture Jet Operational", value=True, key=f'cj{hood}')
            with col1:
                cj_reading = st.number_input("Capture Jet average pressure Reading", min_value=0, key=f'cjreading{hood}')
            return {
                "Quantity of Slaves per System": quantity,
                "UV Pressure Setpoint (Pa)": pressure,
                "Airflow Proved on Each Controller": airflow,
                "Communication To All Canopy Controllers": comms,
                "UV System Tested and Fully Operational": ops,
                "All UV Safety Switches Tested": safety,
                "All Filter Safety Switches Tested": filters,
                "UV Capture Jet Operational": cj,
                "Capture Jet average pressure Reading (PA)": cj_reading

            }

    return {}
        
def createCJHood(selection, num):
    st.markdown(f"<h3 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>{selection} Info<h3>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Drawing Number<h4>", unsafe_allow_html=True)
        drawingNum = st.text_input('.', key=f'drawNum{num}', label_visibility='hidden')
    with col2:
        st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy Location<h4>", unsafe_allow_html=True)
        canopyLocation = st.text_input('.', key=f'canopyLoc{num}', label_visibility='hidden')
    with col3:
        st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy Sections<h4>", unsafe_allow_html=True)
        quantity = st.number_input('.', key=f'quantity{num}', label_visibility='hidden', min_value=0)
    return cjHoods(drawingNum, canopyLocation, selection, num, quantity)
        