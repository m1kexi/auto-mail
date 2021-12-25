#!/usr/bin/env python3

import pandas as pd
import numpy as np
from pandas.core.indexes import base
import streamlit as st
import os

from mail_tc import *

base_path = os.path.join(os.path.expanduser('~'),'code/Repositories/auto-mail')

df = pd.read_excel(os.path.join(base_path,'mail_list.xlsx'))
df = df.dropna(subset=['西法老师'])
names = df['西法老师'].drop_duplicates().to_list()
mail_mapping = df.set_index('西法老师').to_dict()['邮箱']

files = [i for i in os.listdir(os.path.join(base_path,'files')) if (not i.startswith(('~','.'))) and i.endswith('.xlsx')]
files_name = [i.split('.')[0] for i in files]

files_not_in_names = [i for i in files_name if i not in names]
names_not_in_files = [i for i in names if i not in files_name]

# title
st.title(':fire:Auto-mail:dash:')

# modules
module = st.sidebar.radio('Modules',('Split Files','Send Mails'), index=1)

if module == 'Split Files':

    # sidebar
    st.sidebar.header('Existing Files')
    e_files = [file for file in os.listdir(os.path.join(base_path,'files'))]
    st.sidebar.table(e_files)

    if st.sidebar.button('remove old files'):
        for file in os.listdir(os.path.join(base_path,'files')):
            os.remove(os.path.join(base_path,'files',file))

    # main
    uploaded_file = st.file_uploader("Choose a file")
    if uploaded_file is not None:
        xl = pd.ExcelFile(uploaded_file)
        selected_sheet = st.selectbox('Choose sheet',xl.sheet_names)
        df = xl.parse(selected_sheet)
        
        columns = df.columns.to_list()
        s_columns = st.multiselect('Choose columns',columns,columns[11:15])
        col1,col2,col3 = st.columns(3)
        with col1:
            name = st.selectbox('name column',columns,0)
        with col2:
            role = st.selectbox('role column',columns,1)
        with col3:
            role_list = df[role].drop_duplicates().to_list()
            s_role = st.selectbox('special role',role_list,role_list.index('全职系数'))
        df = df.dropna(subset=[name])
        df = df[~df[name].str.contains('汇总')]
        role_check = df.groupby(name)[role].nunique()
        check_list = role_check[role_check!=1]
        role_map = df[[name,role]].drop_duplicates().set_index(name).to_dict()[role]

        if len(check_list) != 0:
            st.warning('One person with more than 1 role: ' + ', '.join(check_list))
        else:
            pass

        st.header('Confirm and Split')
        if st.button('Split'):
            for pname in df[name]:
                _df = df[df[name] == pname]
                if role_map[pname] == s_role:
                    pass
                else:
                    _df = _df[[column for column in _df.columns if column not in s_columns]]
                _df.to_excel(os.path.join(base_path,f'files/{pname}.xlsx'),index=False)

elif module == 'Send Mails':

    # side bar
    st.sidebar.subheader('Files not in the mail list')
    st.sidebar.warning(', '.join([f'{i}.xlsx' for i in files_not_in_names]))

    st.sidebar.subheader('Names in the mail list has no correspondent file')
    st.sidebar.warning(', '.join(names_not_in_files))

    mail_subject = st.sidebar.text_input('Please input the mail subject.')
    mail_body = st.sidebar.text_input('Please input the mail body')

    st.sidebar.subheader('Confirm and send')
    button_send = st.sidebar.button('Send')
    my_bar = st.sidebar.progress(0)
    if button_send:
        # df['sent flag'] = 'x'
        flag = []
        ttl = len(names)
        count = 0
        for name in names:
            if name in files_name:
                address = mail_mapping[name]
                # error_msg = send_mail(address, mail_subject, mail_body,os.path.join(base_path,f'files/{name}.xlsx'), f'{name}.xlsx')
                error_msg = 'success'
                if error_msg == 'success':
                    flag.append('sent')
                else:
                    flag.append('send failed')
            else:
                flag.append('not correspondent file')
            count += 1
            my_bar.progress(count/ttl)
        df['sent flag'] = flag
    else:
        pass

    # mail list table
    st.table(df)