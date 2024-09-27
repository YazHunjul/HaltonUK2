import streamlit as st
import math
cj_hoods = ['KVF', 'KVI', 'KCH-F', 'KCH-I']
cj_lp = ['KSR-S', 'KSR-F', 'KSR-M']

cj_hoods_k_factor =  {
    1 : 71.8,
    2: 143.6,
    3 : 215.4,
    4 : 287.2,
    5 : 359,
    6 : 430.8
}

cj_lp_k_factor =  {
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

class cjHoods():
    def __init__(self, drawingNum, location, model, quantityOfSections) -> None:
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
    
    def getCJSections(self, selection, num):
        st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)
        for i in range(self.quantityOfSections):
            st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>{self.model} Section {i+1} Extract Info ({self.location})<h4>", unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)
            if selection in ['KVF', 'KCH-F']:
                st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>{self.model} Section {i+1} Supply Info ({self.location})<h4>", unsafe_allow_html=True)
                col3, col4 = st.columns(2)
                st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)
            
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
                
            achievedFlowRate = self.k_factor * math.sqrt(self.tab_Reading)
            #IF FRESH AIR
            if selection in ['KVF', 'KCH-F']:
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
                print(math.sqrt(self.tab_Reading))
                self.sections[f'{i+1}'] = {"ksaQuantity" : self.ksaQuantity, "k_factor" : self.k_factor, 'tab_reading' : self.tab_Reading, 'designFlow' : self.design_flow, 'achieved' : round(achievedFlowRate,2), 'supplyKFactor' : cj_hoods_supply_kfactor[self.plenumLength], 'supplyTab' : self.tab_Reading_supply, 'supplyDesign' : round(self.design_supply,2), 'achievedSupply' : round(supplyAchievedFlowRate,2)}
            else:
                self.sections[f'{i+1}'] = {"ksaQuantity" : self.ksaQuantity, "k_factor" : self.k_factor, 'tab_reading' : self.tab_Reading, 'designFlow' : self.design_flow, 'achieved' : round(achievedFlowRate,2)}

def numCanopies():
    hoods = {}
    st.markdown("<h1 style='text-align: center;margin-top: 30px;margin-bottom:-100px;'>Enter Number Of Canopies<h1>", unsafe_allow_html=True)
    numInput = st.number_input('.', label_visibility='hidden', min_value=0)
    for num in range(numInput):
        st.markdown(f"<h3 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy {num+1}<h3>", unsafe_allow_html=True)
        selection = st.selectbox('.',['','KVF', "KVI", "KCH-F", "KCH-I", "KSR-S", "KSR-F", "KSR-M", "UVI", "UVF"],key=f'selection{num}', label_visibility='hidden')
        #CJ HOODS
        hood = createCJHood(selection,num)
        hood.getCJSections(selection,num)
        hoods[f'{selection} {num}'] = hood
            
        st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)
    return hoods

        
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
    return cjHoods(drawingNum, canopyLocation, selection, quantity)
        