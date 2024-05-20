import streamlit as st

property = ""

def user(new_value):
    global property
    property = new_value
    st.session_state.my_global_variable = property
    # return file

def get_user():
    return st.session_state.my_global_variable