#Getting Phone Numbers

import os
import tempfile
import numpy as np
import pandas as pd, phonenumbers, re
import streamlit as st


@st.dialog("Error!!")
def raiseError(text):
    st.text(text)

def country_code(number):
    try:
        return phonenumbers.parse("+"+number ).country_code
    except:
        return pd.NA

def national_number(number):    
    try:
        return phonenumbers.parse("+"+number ).national_number
    except:
        return pd.NA

def getNumber(number):
    
    try:
        mobile = re.findall(r'\d+', str(number))[0]
        if len(mobile)> 7:
            return mobile

    except:
        return pd.NA

def getNewName(NewColName, df):
    newcol = 1
    while NewColName in df.columns:  #Checks whether the newly generated column name already exists or not
        NewColName = NewColName+"_"+str(newcol)
        newcol = newcol+1
    
    return NewColName
      
def save_upload(uploaded_file):
    tmp_dir = tempfile.mkdtemp()
    tmp_path = os.path.join(tmp_dir, uploaded_file.name)
    with open(tmp_path, "wb") as f:
        f.write(uploaded_file.read())
    return tmp_path

def get_country_code(df, select_columns, fileextn, select_sheets = 0, LSQFormat = False):

    if fileextn != "manual":
        if fileextn == "xlsx":
            df = pd.read_excel(filePath, sheet_name=select_sheets)
        else:
            df= pd.read_csv(filePath,sep=",", encoding_errors="ignore", low_memory=False)

    ColLocation = df.columns.get_loc(select_columns)
    df["Test"] = df[select_columns].map(lambda x:  getNumber(x) )
    
    country_code_column = getNewName("Country_Code", df)
    phone_number_column = getNewName("Phone_Number", df)

    df.insert(loc=ColLocation, column=country_code_column, value=df["Test"].map(lambda x: country_code(str(x))), allow_duplicates=True )
    df.insert(loc=ColLocation+1, column=phone_number_column, value=df["Test"].map(lambda x: national_number(str(x))), allow_duplicates=True )
    
    totalLength =  df[phone_number_column].notna().sum() 

    if totalLength>0:
        if LSQFormat:  
            LSQ_Phone = getNewName("Phone_LSQ", df)
            df.insert(loc=ColLocation+2, column=LSQ_Phone,  
                    value=np.where(df[phone_number_column].notna(), 
                                                "+" + df[country_code_column].astype(str) + "-" + df[phone_number_column].astype(str),
                                                pd.NA), allow_duplicates=True )
            
        df.drop(columns=["Test"], inplace=True)

        df.rename(columns={i: " " for i in df.columns if i.find("Unnamed") != -1}, errors="ignore", inplace=True)
        #df.drop(columns=[i for i in df.columns if i.find("Unnamed") != -1], inplace=True, errors="ignore")
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx" if (fileextn in ["xlsx", "manual"]) else ".csv")
        tmp.close()
        if (fileextn in ["xlsx", "manual"]):
            with pd.ExcelWriter(tmp.name, engine="xlsxwriter") as file:
                df.to_excel(file, index=False)
        else:
            df.to_csv(tmp.name,sep=",", index=False)

        return tmp.name
    
    else:
        df.drop(columns=[country_code_column, phone_number_column], inplace=True)
    
        raiseError("Incorrect column selected. Please select the phone number column.")

if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0  

def clearUploads():
    if "tempData" in st.session_state:
        st.session_state["tempData"] = False
        st.rerun()

def reset_text():
    st.session_state.manualText = ""

st.set_page_config(page_title=" Country Code Formatter", page_icon="🌎", layout="wide")
st.header("🌎 Country Code Formatter")

text = st.text_area("Enter Phone Number(s) manually.",   key="manualText", on_change=clearUploads)

st.markdown("<h3 style='text-align: center;'>OR</h3>", unsafe_allow_html=True)
fileLoc = st.file_uploader("Upload File", type=["csv", "xlsx"], on_change=reset_text, key=f"uploader_{st.session_state.uploader_key}")

getinLSQFormat = st.checkbox("LeadSquared Format.")

if fileLoc:
    filePath = save_upload(fileLoc)
    fileName,  fileextn = os.path.basename(filePath).split(".") 
    select_sheets = 0
    if fileextn == "xlsx":
        with pd.ExcelFile(filePath, engine="openpyxl") as file:
            sheets = file.sheet_names
        if len(sheets) > 0:
            select_sheets = st.selectbox(label = "Multiple Sheets detected. Select a sheet to process", options=sheets)
            if select_sheets:
                df = pd.read_excel(filePath,nrows=100 , sheet_name=select_sheets)
        else:
            df = pd.read_excel(filePath,nrows=100 , sheet_name=sheets[0])
    else:
        df= pd.read_csv(filePath,sep=",", nrows = 100, encoding_errors="replace")

    eligibleColumns = df.select_dtypes(include=["float", "int"]).columns
    select_columns = st.selectbox(label = "Select a column to process", options=eligibleColumns)

    btn = st.button(label="Generate Data", type="primary")
    
    if btn:
        with st.spinner("Processing..",show_time=True):
            
            tmp_path = get_country_code(df, select_columns, fileextn, select_sheets, getinLSQFormat)

            if tmp_path:
                if (fileextn == "csv") and getinLSQFormat:
                    new_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                    new_tmp.close()
                    new_df = pd.read_csv(tmp_path, sep=",", low_memory=False).to_excel(new_tmp.name, index=False)
                    tmp_path = new_tmp.name
                    fileextn = "xlsx"

                outputFileName = fileName+"."+fileextn
                with open(tmp_path, "rb") as f:
                    
                    st.download_button("Download Files", f,  file_name=outputFileName)

                os.unlink(tmp_path) 
            #st.download_button(label="Download File", data=df, file_name=filePath.name) 
elif len(text) > 0:
    
    df = pd.DataFrame(data={"Phone": text.split("\n")})
    df["Phone"] = df["Phone"].apply(lambda x: pd.NA if x== "" else x)
    df.dropna(how="all", inplace=True)
    st.dataframe(df, key="tempData", hide_index=True)
    
    btn = st.button(label="Generate Data", type="primary")
    if btn:
        with st.spinner("Processing..",show_time=True):
                
                tmp_path = get_country_code(df, "Phone", "manual", None, getinLSQFormat)
                output_df = pd.read_excel(tmp_path)

                st.dataframe(output_df.style.set_properties(**{'text-align': 'right'}) , hide_index=True)  

    

    
