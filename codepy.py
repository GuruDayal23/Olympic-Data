
import numpy as np
import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridUpdateMode
from st_aggrid.grid_options_builder import GridOptionsBuilder

def main():
     try:
         st.set_page_config(page_title="Data Validator",page_icon=" ",layout="wide")
         st.sidebar.file_uploader("Please upload your criteria file")

         @st.cache
         def rd():

             df = pd.read_excel(r"C:\Users\DELL\OneDrive\Desktop\Test_dummy_data.xlsx",sheet_name="Sheet1")
             return df
         df = rd()

         ch = st.multiselect("Select base columns for transformation",df.columns)


         meltcol = st.selectbox("select the to be transformed", df.columns)

         df_new = pd.DataFrame()

         df_new = df.copy(deep=True)
         df_new['New'] = ""

         def get_data():
             for i in ch:
                df_new['New'] = df_new['New'] + "___" + df_new[str(i)]
             df_new1 = df_new[['New',meltcol]]
             df_new1 =df_new1.astype(str)
             df3= df_new1[meltcol].str.split(',', expand=True)
             df4 = pd.concat([df_new1,df3], axis=1, join='inner')
             del df4[meltcol]
             df5 = pd.melt(df4, id_vars="New", value_vars=np.arange(df3.shape[1]),var_name='metric', value_name=meltcol)
             del df5['metric']
             df5 = df5[~df5[meltcol].isnull()]
             ct = 1
             for i in ch:
                 df5[i] = df5['New'].apply(lambda x: str(x).replace("___",",").split(',')[ct])
                 ct+=1
             del df5['New']
             df5 = df5[ch+[meltcol]]
             # st.write(df5.shape)
             gd = GridOptionsBuilder.from_dataframe(df5)
             gd.configure_pagination(enabled=True)
             gd.configure_side_bar()
             gd.configure_default_column(editable=True, enableRowGroup=True, groupable=True, value=True, aggFunc="sum")
             gridoptions = gd.build()

             AgGrid(df5, gridOptions=gridoptions,
                     allow_unsafe_jscode=True,
                     height=500, enable_enterprise_modules=True,
                     theme='streamlit')
             if st.button("Send Email"):
                 to_mail = st.text_input("Please enter Email ID")
             # send_emails(to_mail)
         if st.checkbox("Show report"):
             get_data()

             # TODO: add sending of email
         def send_emails(x):
             import win32com.client as win32
             outlook = win32.Dispatch('outlook.application')
             mail = outlook.CreateItem(0)
             mail.To = x
             mail.Subject = 'Test Streamlit Report'
             mail.Body = 'Message body'
             mail.HTMLBody = '<h2>HTML Message body</h2>' # this field is optional

             # To attach a file to the email (optional):
             attachment = "Path to the attachment"
             mail.Attachments.Add(attachment)

             mail.Send()


         rd()




     except Exception as e:
        print(e)
        
#         st.write("Error: Can not select transform column in both the fields")


if __name__ == '__main__':
    main()
