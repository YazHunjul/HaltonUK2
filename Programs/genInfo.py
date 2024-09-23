import streamlit as st
import base64


def titleAndLogo(title):
    st.markdown(f"<h1 style='text-align: center;'>{title}<h1>", unsafe_allow_html=True)
    return


def getGenInfo():
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("<h2 style='text-align:center; margin-bottom:-40px;'>Client</h2>", unsafe_allow_html=True)
        client = st.text_input('.', key='client', label_visibility='hidden')
    with col2:
        st.markdown("<h2 style='text-align:center;margin-bottom:-40px;'>Project Name</h2>", unsafe_allow_html=True)
        projectName = st.text_input('.', key='projectName', label_visibility='hidden')
    with col1:
        st.markdown("<h2 style='text-align:center;margin-bottom:-40px;'>Project Number</h2>", unsafe_allow_html=True)
        projectNumber= st.text_input('.', key='projectNumber', label_visibility='hidden')
    with col2:
        st.markdown("<h2 style='text-align:center;margin-bottom:-40px;'>Date Of Visit</h2>", unsafe_allow_html=True)
        DateOfVisit = st.date_input('.', key='date', label_visibility='hidden')
    st.markdown("<h2 style='text-align:center;margin-bottom:-40px;'>Engineer(s)</h2>", unsafe_allow_html=True)
    engineers = st.text_input('.', key='engineers', label_visibility='hidden')
    
    return {
        'client': client,
        'project_name': projectName,
        'project_number': projectNumber,
        'date_of_visit': DateOfVisit,
        'engineers': engineers
    }
    
    