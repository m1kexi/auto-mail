#!/usr/bin/env python3

import pandas as pd
import numpy as np
import streamlit as st
import os


df = pd.read_excel('mail_list.xlsx')
df = df.dropna(subset=['西法老师'])
names = df['西法老师'].drop_duplicates().to_list()

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

st.sidebar.subheader('Confirm and send')
if st.sidebar.button('Send'):
    df['sent flag'] = 'x'
else:
    pass

# mail list table
st.table(df)