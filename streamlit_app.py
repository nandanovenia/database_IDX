import streamlit as st
from streamlit_option_menu import option_menu
from bs4 import BeautifulSoup
import os, io
import pandas as pd
from xlsxwriter import Workbook
import altair as alt

st.title('HALO')
