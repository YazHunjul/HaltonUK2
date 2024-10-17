import streamlit as st
import base64
from streamlit_drawable_canvas import st_canvas

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

def get_comments():
    comments = st.text_area('Comments (Enter Each in new space)', key='comments')
    comments_split = comments.split('\n')
    return comments_split

def get_sign():
    st.write("Signed By Engineer(s)")

    canvas_result = st_canvas(
        stroke_width=5,
        stroke_color="#000000",
        background_color="#FFFFFF",
        height=150,
        width=700,
        drawing_mode="freedraw",
        key="canvas_test"
    )
    return canvas_result
    