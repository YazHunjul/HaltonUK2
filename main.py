import streamlit as st

def main():
    st.sidebar.title("Navigation")
    selection = st.sidebar.radio("Go to", ["Testing And Commissioning"])
    
    if selection == "Testing And Commissioning":
        from Programs import TandC
        TandC.TandC()
        return
    
    
main()