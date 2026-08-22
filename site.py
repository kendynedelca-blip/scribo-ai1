import base64
import hashlib
import hmac
import sqlite3
import json
import asyncio
import re
import secrets
from datetime import datetime
from io import BytesIO


import streamlit as st
import streamlit.components.v1 as components
from groq import Groq


st.markdown(
    """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    [data-testid="stStatusWidget"] {visibility: hidden;}
    .viewerBadge_container__1QS13 {display: none !important;}
    </style>
    """,
    unsafe_allow_html=True
)






# ============================================================
# DEPENDENȚE OPȚIONALE
# ============================================================

