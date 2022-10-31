#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from enum import unique
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog as fd
import pandas as pd
import os

def choose_file():
    file_path_list = fd.askopenfilenames(title='Choose files', initialdir=os.getcwd())
    file_path_list = '|'.join(file_path_list)
    input_entry.insert('0', file_path_list)
    # return list(file_path_list)

def choose_dir():
    dir_path = fd.askdirectory(title='Choose directory', initialdir=os.getcwd())
    output_entry.insert('0', dir_path)
    # return dir_path

def load_file(file_path_list):
    data_list = []
    for file_path in file_path_list:
        # using file name to be grouping tags
        file_name = os.path.basename(file_path).split('.')[:-1]
        file_name = '.'.join(file_name)
        # read data and add column of grouping tag
        data = pd.read_excel(file_path)
        data['FILE_NAME'] = file_name
        data_list.append(data)
    return data_list

def grouping_list(data_list):
    # rbind the list of data
    data_combine = pd.concat(data_list, axis=0)
    data_combine = data_combine.reset_index(drop=True)
    # specific column that be groupby index
    column_list = list(data_combine.columns)
    column_list = [x  for x in column_list if x != 'FILE_NAME']
    # aggregate grouping tag
    data_combine = data_combine.groupby(column_list).agg('_'.join)
    data_combine = data_combine.reset_index()
    return data_combine

def save_file(df, output_dir):
    new_file_list = df['FILE_NAME'].unique()
    for new_file in new_file_list:
        df_temp = df[df['FILE_NAME']==new_file]
        df_temp = df_temp.drop('FILE_NAME', axis=1)
        number_of_rows = len(df_temp)
        df_temp.to_excel(f'{output_dir}/{new_file}_rows{number_of_rows}.xlsx', index=False)

def main():
    file_path_list = input_entry.get()
    file_path_list = file_path_list.split("|")
    output_dir = output_entry.get()
    df = load_file(file_path_list)
    df = grouping_list(df)
    save_file(df, output_dir)
    tk.messagebox.showinfo(title = 'Done!', message = "OK")



# file_path_list = choose_file()
# # file_path_list = [
# #     'C:/Project/List_processing/Input/2012.xlsx',
# #     'C:/Project/List_processing/Input/2032.xlsx',
# #     'C:/Project/List_processing/Input/2040.xlsx'
# #     ]
# output_dir = choose_dir()
# # output_dir = 'C:/Project/List_processing/Output'

# df = load_file(file_path_list)
# df = grouping_list(df)
# save_file(df, output_dir)



# GUI
# Basic window
root_window = tk.Tk()
root_window.geometry("650x130")
root_window.resizable(False, False)
# root_window.resizable(True, True)
root_window.title("Group lists")

# Some function
input_label = tk.Label(root_window, text='List files:')
input_label.grid(row=0, column=0)
input_entry = tk.Entry(root_window, width=50)
input_entry.grid(row=0, column=1)

output_label = tk.Label(root_window, text='Output path:')
output_label.grid(row=1, column=0)
output_entry = tk.Entry(root_window, width=50)
output_entry.grid(row=1, column=1)

choose_file_button = tk.Button(root_window, text='Choose lists', command=choose_file)
choose_file_button.grid(row=0, column=3)

choose_dir_button = tk.Button(root_window, text='Choose directory', command=choose_dir)
choose_dir_button.grid(row=1, column=3)

    
start_button = tk.Button(root_window, text="Group it", command=main)
start_button.grid(row=2, column=0, columnspan=2)
close_button = tk.Button(root_window, text="Close the window", command=root_window.quit)
close_button.grid(row=3, column=0, columnspan=2)

# Run it
root_window.mainloop() 


