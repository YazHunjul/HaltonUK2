import streamlit as st

cj_hoods = ['KVF', 'KVI', 'KCH-F', 'KCH-I']

cj_hoods_k_factor =  {
    1 : 71.8,
    2: 143.6,
    3 : 215.4,
    4 : 287.2,
    5 : 359,
    6 : 430.8
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
    
    def getSections(self, selection, num):
        st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)
        for i in range(self.quantityOfSections):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-62px;'>Section {i+1} KSA's<h4>", unsafe_allow_html=True)
                self.ksaQuantity = st.number_input('.', key=f'ksaQuantity{selection}{num}{i}', label_visibility='hidden', min_value=1, max_value=6)
            with col2:
                st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-60px;'>Section {i+1} T.A.B Point Reading<h4>", unsafe_allow_html=True)
                self.tab_Reading = st.number_input('.', key=f'canopyLocs{selection}{num}{i}', label_visibility='hidden', min_value=0.0)
            with col1:
                st.markdown(f"<h4 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Section {i+1} Design Airflow<h4>", unsafe_allow_html=True)
                self.tab_Reading = st.number_input('.', key=f'designflow{selection}{num}{i}', label_visibility='hidden', min_value=0.0)
            self.k_factor = cj_hoods_k_factor[self.ksaQuantity]
            
            self.sections[f'section {i+1}'] = {"ksaQuantity" : self.ksaQuantity, "k_factor" : self.k_factor, 'tab_reading' : self.tab_Reading}
            st.markdown("<hr style='border:1px solid white;'>", unsafe_allow_html=True)
    
            
    
def numCanopies():
    hoods = {}
    st.markdown("<h1 style='text-align: center;margin-top: 30px;margin-bottom:-100px;'>Enter Number Of Canopies<h1>", unsafe_allow_html=True)
    numInput = st.number_input('.', label_visibility='hidden', min_value=0)
    for num in range(numInput):
        st.markdown(f"<h3 style='text-align: center;margin-top: 30px;margin-bottom:-70px;'>Canopy {num+1}<h3>", unsafe_allow_html=True)
        selection = st.selectbox('.',['','KVF', "KVI", "KCH-F", "KCH-I"],key=f'selection{num}', label_visibility='hidden')
        if selection in cj_hoods:
            hood = createCJHood(selection,num)
            hood.getSections(selection,num)
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
        