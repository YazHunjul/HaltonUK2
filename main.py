import streamlit as st

def main():
    st.sidebar.title("Navigation")
    selection = st.sidebar.radio("Go to", ['Home',"Testing And Commissioning", "Service Reports"])
    
    if selection == "Testing And Commissioning":
        from Programs import TandC
        TandC.TandC()
        return
    
    
main()