#!/usr/bin/python3
# encoding: utf-8
"""
@version: python3.6
@author: ‘Song‘
@software: HMO
@file: 
@time: 9:49
"""
import os

user_choice = input('Please choose "Upgrade or Initilize":')

if user_choice == 'U' or 'u'：
    os.system('pip install --upgrade python-docx')
    os.system('pip install --upgrade openpyxl')
else:
    os.system('pip install python-docx')
    os.system('pip install openpyxl')