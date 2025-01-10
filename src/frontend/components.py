import streamlit as st
from pathlib import Path
from typing import Callable, Optional

def setup_page():
    """Configure the Streamlit page with custom styling."""
    st.set_page_config(
        page_title="Excel Template Inserter",
        page_icon="📊",
        layout="wide"
    )
    
    # Custom CSS for modern UI
    st.markdown("""
        <style>
        .stButton > button {
            background-color: #3b82f6;
            color: white;
            border-radius: 6px;
            padding: 0.5rem 1rem;
            border: none;
            box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
        }
        .stButton > button:hover {
            background-color: #2563eb;
        }
        .upload-box {
            border: 2px dashed #e5e7eb;
            border-radius: 8px;
            padding: 20px;
            text-align: center;
        }
        </style>
    """, unsafe_allow_html=True)

def action_button(label: str, callback: Callable, key: str):
    """Create a styled action button."""
    if st.button(label, key=key):
        callback()

def success_message(message: str):
    """Display a success message."""
    st.success(message)

def error_message(message: str):
    """Display an error message."""
    st.error(message)

def info_card(title: str, content: str):
    """Display an info card with title and content."""
    st.markdown(f"""
        <div style='
            padding: 1rem;
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            margin-bottom: 1rem;
        '>
            <h3 style='margin-top: 0;'>{title}</h3>
            <p style='margin-bottom: 0;'>{content}</p>
        </div>
    """, unsafe_allow_html=True) 