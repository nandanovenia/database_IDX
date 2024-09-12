import streamlit as st
from streamlit_option_menu import option_menu
from bs4 import BeautifulSoup
import os, io
import pandas as pd
from xlsxwriter import Workbook
import altair as alt

folder = "D:/Project A3/Database Connection/IDX_data - extracted"

def change_date_format(data):
    for entry in data:
        if '2023-12-31' in entry:
            entry['31 December 2023'] = entry.pop('2023-12-31')
        if '2022-12-31' in entry:
            entry['31 December 2022'] = entry.pop('2022-12-31')
    return data

# Function to convert text to sentence case
def sentence_case(text):
    return text.capitalize()

def tabel_lengkap_BS(folder_efek,emiten):
    year = ['2020','2021','2022','2023']
    data_account_list = []
    for i in range(len(year)):
        filename =  next((path for path in [
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/1210000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/2210000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/3210000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/1220000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/2220000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/8220000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/4220000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/5220000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/6220000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/7220000.html"
        ] if os.path.exists(path)), None)

        if filename is None:
            continue
        with open(filename, "r", encoding="utf-8") as HTMLFileToBeOpened:
            contents = HTMLFileToBeOpened.read()
            soup = BeautifulSoup(contents, 'html.parser')

            # DATE
            date_headers = soup.find_all('td', class_="colHeader01")
            if len(date_headers) >= 2:
                year_current = date_headers[0].text.strip()
                year_prior = date_headers[1].text.strip()
            else:
                continue

            def process_accounts():
                if (year[i] == '2020') or (year[i] == '2021'):
                    accounts = soup.find_all(
                        'td',
                        class_="rowHeaderID01"
                    )
                elif year[i] == '2022':
                    accounts = soup.find_all(
                        'td',
                        class_="rowHeaderLeft"
                    )
                elif year[i] == '2023':
                    accounts = soup.find_all(
                        'td',
                        class_="rowHeaderLeft"
                    )
                
                for item in accounts:
                    values_items = item.find_next_siblings(class_="valueCell")
                    if len(values_items) >= 2:
                        data_account = {
                            "Account": item.get_text(strip=True),
                            year_current: values_items[0].get_text(strip=True),
                            year_prior: values_items[1].get_text(strip=True)
                        }
                        data_account_list.append(data_account)
            
            # Process different levels of accounts
            process_accounts()

    df_date_formatted=change_date_format(data_account_list)

    for entry in data_account_list:
        entry['Account'] = sentence_case(entry['Account'])

    combined_data = {}
    for entry in df_date_formatted:
        account = entry['Account']
        if account not in combined_data:
            combined_data[account] = {}
            
        for year, value in entry.items():
            if year != 'Account':
                combined_data[account][year] = value

    df = pd.DataFrame.from_dict(combined_data, orient='index').reset_index()
    sorted_columns = sorted(df.columns[1:], key=lambda x: int(x.split()[-1]))

    # Reorder the DataFrame columns
    df = df[['index'] + sorted_columns]
    df = df.rename(columns={'index': 'Account'})
    df.columns = [col.split()[-1] if 'December' in col else col for col in df.columns]
    df = df.set_index(['Account'])
    return df

def tabel_lengkap_LR(folder_efek,emiten):
    data_account_list = []
    year = ['2020','2021','2022','2023']
    for i in range(len(year)):
        filename =  next((path for path in [
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/1311000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/2311000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/3311000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/5311000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/1321000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/2321000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/3321000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/5321000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/1312000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/2312000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/3312000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/4312000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/6312000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/7312000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/8312000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/1322000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/2322000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/3322000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/4322000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/6322000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/7322000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/8322000.html"
        ] if os.path.exists(path)), None)

        if filename is None:
            continue
        with open(filename, "r", encoding="utf-8") as HTMLFileToBeOpened:
            contents = HTMLFileToBeOpened.read()
            soup = BeautifulSoup(contents, 'html.parser')

            # DATE
        date_headers = soup.find_all('td', class_="colHeader01")
        if len(date_headers) >= 2:
            year_current = date_headers[0].text.strip()
            year_prior = date_headers[1].text.strip()
            print(year_current)
        else:
            print("Date headers not found.")
            continue

        # Function to process accounts
        def process_accounts():
            if (year[i] == '2020') or (year[i] == '2021'):
                accounts = soup.find_all(
                    'td',
                    class_="rowHeaderID01"
                )
            elif year[i] == '2022':
                accounts = soup.find_all(
                    'td',
                    class_="rowHeaderLeft",
                    style=f"width:30%;"
                )
            elif year[i] == '2023':
                accounts = soup.find_all(
                    'td',
                    class_="rowHeaderLeft"
                )
            
            for item in accounts:
                values_items = item.find_next_siblings(class_="valueCell")
                if len(values_items) >= 2:
                    data_account = {
                        "Account": item.get_text(strip=True),
                        year_current: values_items[0].get_text(strip=True),
                        year_prior: values_items[1].get_text(strip=True)
                    }
                    data_account_list.append(data_account)
        process_accounts()    

    df_date_formatted=change_date_format(data_account_list)
    
    for entry in data_account_list:
        entry['Account'] = sentence_case(entry['Account'])

    combined_data = {}
    for entry in df_date_formatted:
        account = entry['Account']
        if account not in combined_data:
            combined_data[account] = {}
            
        for year, value in entry.items():
            if year != 'Account':
                combined_data[account][year] = value

    df = pd.DataFrame.from_dict(combined_data, orient='index').reset_index()
    sorted_columns = sorted(df.columns[1:], key=lambda x: int(x.split()[-1]))

    # Reorder the DataFrame columns
    df = df[['index'] + sorted_columns]
    df.columns = [col.split()[-1] if 'December' in col else col for col in df.columns]
    df = df.rename(columns={'index': 'Account'})
    df = df.set_index(['Account'])
    return df

def tabel_lengkap_cashflow(folder_efek,emiten):
    data_account_list = []
    year = ['2020','2021','2022','2023']

    for i in range(len(year)):
        filename =  next((path for path in [
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/1510000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/2510000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/3510000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/4510000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/5510000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/6510000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/7510000.html",
            f"{folder_efek}/{emiten}/{emiten}{year[i]}/8510000.html"
        ] if os.path.exists(path)), None)

        if filename is None:
            continue
        with open(filename, "r", encoding="utf-8") as HTMLFileToBeOpened:
            contents = HTMLFileToBeOpened.read()
            soup = BeautifulSoup(contents, 'html.parser')

        # DATE
        date_headers = soup.find_all('td', class_="colHeader01")
        if len(date_headers) >= 2:
            year_current = date_headers[0].text.strip()
            year_prior = date_headers[1].text.strip()
            print(year_current)
        else:
            print("Date headers not found.")
            continue

        def process_accounts():
            if (year[i] == '2020') or (year[i] == '2021'):
                accounts = soup.find_all(
                    'td',
                    class_="rowHeaderID01"
                )
            elif year[i] == '2022':
                accounts = soup.find_all(
                    'td',
                    class_="rowHeaderLeft"
                )
            elif year[i] == '2023':
                accounts = soup.find_all(
                    'td',
                    class_="rowHeaderLeft")
            
            for item in accounts:
                values_items = item.find_next_siblings(class_="valueCell")
                if len(values_items) >= 2:
                    data_account = {
                        "Account": item.get_text(strip=True),
                        year_current: values_items[0].get_text(strip=True),
                        year_prior: values_items[1].get_text(strip=True)
                    }
                    data_account_list.append(data_account)
        
        # Process different levels of accounts
        process_accounts()    
    df_date_formatted=change_date_format(data_account_list)

    for entry in data_account_list:
        entry['Account'] = sentence_case(entry['Account'])

    combined_data = {}
    for entry in df_date_formatted:
        account = entry['Account']
        if account not in combined_data:
            combined_data[account] = {}
            
        for year, value in entry.items():
            if year != 'Account':
                combined_data[account][year] = value

    df = pd.DataFrame.from_dict(combined_data, orient='index').reset_index()
    sorted_columns = sorted(df.columns[1:], key=lambda x: int(x.split()[-1]))

    # Reorder the DataFrame columns
    df = df[['index'] + sorted_columns]
    df.columns = [col.split()[-1] if 'December' in col else col for col in df.columns]
    df = df.rename(columns={'index': 'Account'})
    df = df.set_index(['Account'])
    return df

#-------------------- STREAMLIT -------------------#
#Efek selectbox
efek=st.selectbox('Jenis Efek', ['Saham','Obligasi'],key='Pilih jenis efek')

#Ticker Selectbox
if (efek=='Saham'):
    folder_efek=f"{folder}/Saham"
    sub_folders = [name for name in os.listdir(folder_efek) if os.path.isdir(os.path.join(folder_efek, name))]
    emiten=st.selectbox('Ticker',sub_folders, key='Pilih ticker')

elif (efek=='Obligasi'):
    folder_efek=f"{folder}/Obligasi"
    sub_folders = [name for name in os.listdir(folder_efek) if os.path.isdir(os.path.join(folder_efek, name))]
    emiten=st.selectbox('Ticker',sub_folders, key='Pilih ticker')

tab1, tab2 = st.tabs(["Data", "Financial Highlights"])

tab1.title("Laporan Keuangan")
tab2.title("Rasio Keuangan")

with tab1:
    jenis_lapkeu=st.selectbox('Jenis Laporan Keuangan',['Balance Sheet','Laporan Laba/Rugi','Laporan Arus Kas'], key='Pilih jenis Laporan Keuangan')
    # Load data based on selected report type
    if jenis_lapkeu == "Balance Sheet":
        df = tabel_lengkap_BS(folder_efek, emiten)
    elif jenis_lapkeu == 'Laporan Laba/Rugi':
        df = tabel_lengkap_LR(folder_efek, emiten)
    elif jenis_lapkeu == 'Laporan Arus Kas':
        df = tabel_lengkap_cashflow(folder_efek, emiten)

    #Data Lengkap
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer)
    output.seek(0)
    
    #Data Sederhana
    df_sederhana=df.rename_axis('Account').reset_index()
    df_sederhana=df_sederhana[df_sederhana.Account.str.startswith(('Jumlah', 'Penjualan', 'Beban'))].reset_index(drop=True)
    df_sederhana = df_sederhana.set_index(['Account'])

    output_2 = io.BytesIO()
    with pd.ExcelWriter(output_2, engine='xlsxwriter') as writer:
        df_sederhana.to_excel(writer)
    output_2.seek(0)

    #DISPLAY
    st.subheader('Data Lengkap')
    st.download_button(label="Download as Excel",
                data=output,
                file_name=f'{emiten}_{jenis_lapkeu}_lengkap.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    
    st.dataframe(df, width=1500)

    st.subheader('Data Sederhana')
    st.download_button(label="Download as Excel",
                data=output_2,
                file_name=f'{emiten}_{jenis_lapkeu}_sederhana.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    
    st.dataframe(df_sederhana, width=1500)

with tab2:

    def clean_number(value):
        if isinstance(value, str):  # Check if the value is a string
            value = value.replace(',', '')  # Remove commas
            if '(' in value and ')' in value:  # Convert parentheses to negative sign
                value = value.replace('(', '-').replace(')', '')
            return pd.to_numeric(value, errors='coerce')  # Convert to numeric
        return value

    # Transform Balance Sheet untuk perhitungan Rasio Keuangan
    BS_df = tabel_lengkap_BS(folder_efek, emiten)
    BS_statement = BS_df.T
    BS_statement = BS_statement.applymap(clean_number)
    
    # Transform laporan Laba/Rugi untuk perhitungan Rasio Keuangan
    LR_df = tabel_lengkap_LR(folder_efek, emiten)
    LR_statement = LR_df.T
    LR_statement = LR_statement.applymap(clean_number)

    #Perhitungan rasio keuangan
    ## Rasio Likuiditas
    BS_statement['Current Ratio'] = BS_statement['Jumlah aset lancar'] / BS_statement['Jumlah liabilitas jangka pendek']
    ## Rasio Profitabilitas
    LR_statement['Gross Profit Margin'] = LR_statement['Jumlah laba bruto'] / LR_statement['Penjualan dan pendapatan usaha'] * 100
    LR_statement['Net Profit Margin'] = LR_statement['Jumlah laba (rugi)'] / LR_statement['Penjualan dan pendapatan usaha'] * 100
    LR_statement['Return on Asset'] = LR_statement['Jumlah laba (rugi)'] / BS_statement['Jumlah aset'] * 100
    LR_statement['Return on Equity'] = LR_statement['Jumlah laba (rugi)'] / BS_statement['Jumlah ekuitas'] * 100   
    ## Rasio Solvabilitas
    BS_statement['Debt to Asset Ratio'] = BS_statement['Jumlah liabilitas'] / BS_statement['Jumlah aset'] * 100
    BS_statement['Debt to Equity Ratio'] = BS_statement['Jumlah liabilitas'] / BS_statement['Jumlah ekuitas'] * 100
    BS_statement['Long Term Debt Ratio'] = BS_statement['Jumlah liabilitas jangka panjang'] / BS_statement['Jumlah aset'] * 100

    # Display Rasio Keuangan
    st.header("Rasio Likuiditas")
    st.dataframe(BS_statement["Current Ratio"])
    #st.line_chart(BS_statement['Current Ratio'])
    c = alt.Chart(BS_statement.reset_index()).mark_line().encode(x='index',y='Current Ratio').configure_axisX(
        labelAngle=0).properties(width=800,height=300)
    st.altair_chart(c)
    
    st.header("Rasio Profitabilitas")
    LR_statement=LR_statement[['Gross Profit Margin','Net Profit Margin','Return on Asset','Return on Equity']]
    st.dataframe(LR_statement)
    
    # Create a multi-line Altair chart
    df_long = LR_statement.reset_index().melt('index', var_name='Rasio Profitabilitas', value_name='Value')
    c = alt.Chart(df_long).mark_line().encode(
        x=alt.X('index:O', title='Year', axis=alt.Axis(labelAngle=0)),
        y=alt.Y('Value:Q', title='Percentage'),
        color='Rasio Profitabilitas:N'  # Color lines by Metric
    ).properties(width=800,height=300)
    st.altair_chart(c)

    st.header("Rasio Solvabilitas")
    BS_statement=BS_statement[['Debt to Asset Ratio','Debt to Equity Ratio','Long Term Debt Ratio']]
    st.dataframe(BS_statement)
    
    # Create a multi-line Altair chart
    df_long = BS_statement.reset_index().melt('index', var_name='Rasio Solvabilitas', value_name='Value')
    c = alt.Chart(df_long).mark_line().encode(
        x=alt.X('index:O', title='Year', axis=alt.Axis(labelAngle=0)),
        y=alt.Y('Value:Q', title='Percentage'),
        color='Rasio Solvabilitas:N'  # Color lines by Metric
    ).properties(width=800,height=300)
    st.altair_chart(c)
