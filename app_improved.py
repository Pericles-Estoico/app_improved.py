import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from io import BytesIO
import os
import pickle
import datetime

# Configurações
st.set_page_config(page_title="Pure & Posh Baby - Sistema de Relatórios", page_icon="👑", layout="wide")

# Header com logo - Responsivo
# CSS mais agressivo para centralização
st.markdown("""
<style>
.centered-logo {
    display: flex;
    justify-content: center;
    align-items: center;
    width: 100%;
    margin: 0 auto;
    text-align: center;
}
.centered-title {
    text-align: center;
    width: 100%;
    margin: 0 auto;
}
/* Forçar centralização em todos os elementos de imagem */
div[data-testid="stImage"] {
    display: flex !important;
    justify-content: center !important;
    align-items: center !important;
    width: 100% !important;
}
div[data-testid="stImage"] > div {
    display: flex !important;
    justify-content: center !important;
    align-items: center !important;
    width: 100% !important;
}
</style>
""", unsafe_allow_html=True)

# HTML direto para centralização garantida
st.markdown("""
<div class="centered-logo">
    <img src="data:image/png;base64,{}" width="200" style="display: block; margin: 0 auto;">
</div>
