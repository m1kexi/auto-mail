#!/usr/bin/env python3

import pandas as pd
import numpy as np
import streamlit as st
import os

from mail_tc import *


df = pd.read_excel('mail_list.xlsx')
df = df.dropna(subset=['西法老师'])
names = df['西法老师'].drop_duplicates().to_list()
mail_mapping = df.set_index('西法老师').to_dict()['邮箱']

files = [i for i in os.listdir('files') if (not i.startswith(('~','.'))) and i.endswith('.xlsx')]
files_name = [i.split('.')[0] for i in files]

files_not_in_names = [i for i in files_name if i not in names]
names_not_in_files = [i for i in names if i not in files_name]

# title
st.title(':fire:Auto-mail:dash:')

# side bar
st.sidebar.subheader('Files not in the mail list')
st.sidebar.warning(', '.join([f'{i}.xlsx' for i in files_not_in_names]))

st.sidebar.subheader('Names in the mail list has no correspondent file')
st.sidebar.warning(', '.join(names_not_in_files))

mail_subject = st.sidebar.text_input('Please input the mail subject.')
mail_body = st.sidebar.text_input('Please input the mail body')

st.sidebar.subheader('Confirm and send')
if st.sidebar.button('Send'):
    # df['sent flag'] = 'x'
    flag = []
    for name in names:
        if name in files_name:
            address = mail_mapping[name]
            error_msg = send_mail(address, mail_subject, mail_body,f'files/{name}.xlsx', f'{name}.xlsx')
            if error_msg == 'success':
                flag.append('sent')
            else:
                flag.append('send failed')
        else:
            flag.append('not correspondent file')
    df['sent flag'] = flag
else:
    pass

# mail list table
st.table(df)