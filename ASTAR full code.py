import os

import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter import messagebox

import matplotlib.pyplot as plt

import csv

import numpy as np
import pandas as pd

import re

adj_field = []
adj_moment = []
data_dict = {}
data = pd.DataFrame()
compile_data = pd.DataFrame()
ms_total = pd.DataFrame(columns=['Ms Saturation', 'Thickness'])

#Creating tkinter UI
root = tk.Tk()
root.title('ASTAR AGM & Yield Quality Checking')
label = tk.Label(root, text = 'Select Mode', font=('Arial', 15))
label.pack(padx = 5, pady = 5)
root.resizable(False, False)
root.geometry('250x150')

#directory for AGM
def select_folder(ori):
    foldernames = fd.askdirectory(initialdir = '/Users/zhuowei/Desktop/A*Star', title = f'{ori}files')
    return foldernames

def select_file(ori):
    file_path = fd.askopenfilename()
    return file_path


#AGM pack
def agm():
    global compile_data
    global ms_total
    global saturation_value
    #selecting folder directory
    folder = select_folder('Select Data Folder')
    folder_save = select_folder('Select Save Directory')
    folder_str = os.listdir(folder)
    folder_str.sort()

    #read files in folder
    for filename in folder_str:
        f = os.path.join(folder, filename)
        if filename.endswith('.xlsx'):
            df = pd.read_excel(f)
        
            for index, row in df.iterrows():
                key = row['Sample']
                value = row['Volume']
                data_dict[key] = value
            
        #Read data files from lines 88 to second last
        else:
            with open (f, 'r') as file:
                lines = file.readlines()
                data = csv.DictReader(lines[88:-1], fieldnames=("Field", "Moment", "Adjusted Field", "Adjusted Moment"))
                lookup_key = str(filename)[7]
                vol = data_dict.get(lookup_key)
            
                for row in data:
                    adj_field.append(float(row["Adjusted Field"]))
                    if vol is not None:
                        adj_moment.append(float(row["Adjusted Moment"])/vol)
                    else:
                        adj_moment.append(0)
                    df_data = pd.DataFrame({'Adj_field': adj_field, 'Ms': adj_moment})

                for i in range(1):
                    compile_data['Adj_field'] = adj_field

                pattern = r'([A-Za-z])(\d+\.\d+)'
                match = re.search(pattern, filename)
                if match:
                    letter = match.group(1)
                    digits = match.group(2)

                compile_data = pd.concat([compile_data, df_data['Ms']], axis=1)
                compile_data = compile_data.rename(columns= {'Ms': f'{letter} ({digits}nm)'})
                #compile_data = compile_data.rename(columns = {'Ms': f'{str(filename)[7]}' + ' (' + f'{str(filename)[8:11]}' + 'nm)'})
                data_path = f'{folder_save}/' + str(filename[0:6]) + '.csv'

                if folder_save:
                    file_path  = f'{folder_save}/' + str(filename) + '.csv'
                    df_data.to_csv(file_path, index=False)
                    print('CSV files saved sucessfully')
                else:
                    print('No folder path selected')
                
                data_saturation = csv.reader(lines[61:63])
                saturation_value = None

                for row in data_saturation:
                    if row[0].strip().startswith('Saturation'):
                        saturation_value = row[0].split()[1]
                    
                    if not ms_total['Thickness'].str.contains(digits).any(): 
                        ms_total = pd.concat([ms_total, pd.DataFrame({'Ms Saturation': [float(saturation_value)/vol], 'Thickness': f'{digits}nm'})], ignore_index=True)
            
            adj_moment.clear()
            adj_field.clear()

    compile_data = pd.concat([compile_data, ms_total],axis = 1)
    compile_data.to_csv(data_path, index = False)    

def agm_no_excel():
    global saturation_value
    compile_data = pd.DataFrame()
    ms_total = pd.DataFrame(columns=['Ms Saturation', 'Sample'])
    folder = select_folder('Select Data Folder')
    folder_save = select_folder('Select Save Directory')
    folder_str = os.listdir(folder)
    folder_str.sort()

    for filename in folder_str:
        f = os.path.join(folder, filename)
        with open (f, 'r') as file:
            lines = file.readlines()
            data = csv.DictReader(lines[88:-1], fieldnames=("Field", "Moment", "Adjusted Field", "Adjusted Moment"))
            for row in data:
                adj_field.append(float(row["Adjusted Field"]))
                adj_moment.append(float(row["Adjusted Moment"]))
                df_data = pd.DataFrame({'Adj_field': adj_field, 'Ms': adj_moment})

            file_path  = f'{folder_save}/' + str(filename) + '.csv'
            df_data.to_csv(file_path, index=False)
 
            for i in range(1):
                compile_data['Adj_field'] = adj_field

            compile_data = pd.concat([compile_data, df_data['Ms']], axis=1)
            compile_data = compile_data.rename(columns = {'Ms': f'{str(filename)[7]}' + ' (' + f'{str(filename)[8:11]})'})
            data_path = f'{folder_save}/' + str(filename[0:6]) + '.csv'
            
            data_saturation = csv.reader(lines[61:63])
            saturation_value = None

            for row in data_saturation:
                if row[0].strip().startswith('Saturation'):
                    saturation_value = row[0].split()[1]

                sample_name = filename[8:11]
                if not (ms_total['Sample'] == sample_name).any():
                    ms_total = pd.concat([ms_total, pd.DataFrame({'Ms Saturation': [saturation_value], 'Sample': sample_name})], ignore_index=True)
    
        adj_moment.clear()
        adj_field.clear()

    compile_data = pd.concat([compile_data, ms_total],axis = 1)
    compile_data.to_csv(data_path, index = False) 


def msh_graph():
    x_data = []
    y_data = []
    f = select_file('Select File for data to plot')
    fig = plt.figure
    with open(f, 'r') as file:
        reader = csv.reader(file)
        header = next(reader)
        saturation_index = header.index('Ms Saturation')
        
        for row in reader:
            x_data.append(float(row[0]))

        for column in range(1, saturation_index):
            y_values = []
            file.seek(0)
            next(reader)
            for row in reader:
                y_values.append(float(row[column]))
            y_data.append(y_values)
    
    for i in range(len(y_data)):
        plt.plot(x_data, y_data[i], label=header[i+1])

    plt.xlabel('H (Oe)')
    plt.ylabel('Ms (emu/cc)')
    plt.legend()
    plt.show()
    

def yqc():
    data = pd.DataFrame()
    current = []
    voltage = []
    resistance = []
    sensitivity = None
    
    folder = select_folder('Select Data Folder')
    folder_save = select_folder('Select Save Directory')
    folder_str = os.listdir(folder)
    folder_str.sort()

    for filename in folder_str:
        f = os.path.join(folder, filename)

        if f.endswith('.txt'):
            with open(f, 'r') as file:
                reader = csv.DictReader(file, delimiter= '\t')
                
                sensitivity_pattern = r'PA10-(\d+)'
                match_pattern = re.search(sensitivity_pattern, filename)
                if match_pattern:
                    sensitivity = match_pattern.group(1)

                for row in reader:
                    current.append(float(row['I (A)']))
                    value_x = float(row['Y (V)'])
                    value = (0.004/value_x*10**float(sensitivity))*10**-3
                    resistance.append(float(value))
                    data_y = pd.DataFrame({'Resistance (Ohm)': resistance})
                for i in range(1):
                    data['H (A)'] = current

            data = pd.concat([data, data_y['Resistance (Ohm)']], axis = 1)

            pattern = r'([A-Z])(\d+)D(\d+)'
            match = re.search(pattern, filename)
            if match:
                letter = match.group(1)
                digits = match.group(2) + 'D' + match.group(3)
                data = data.rename(columns={'Resistance (Ohm)': f'{letter}{digits}'})

            current.clear()
            resistance.clear()
            print(f'{letter}{digits} successfully converted!')

    match = re.search(r'R(\d+)C(\d+)', filename)
    if match:
        row = match.group(1)
        column = match.group(2)
        data_path = f'{folder_save}/FP2756-R{row}C{column}.csv'

    data.to_csv(data_path, index = False)

#AGM interface
def agm_interface():
    win = Toplevel()
    win.title("AGM Data")
    win.geometry('250x150')
    message = "Select Folder"
    Label(win, text=message).pack()
    Button(win, text='AGM', command=agm).pack()
    Button(win, text='Moment vs H-Field', command = agm_no_excel).pack()
    Button(win, text = 'Ms.H Graph', command = msh_graph).pack()
    Button(win, text = 'Quit', command = win.quit).pack()

#Yield Interface
def yield_interface():
    win = Toplevel()
    win.title("Yield Quality Checking Data")
    win.geometry('250x150')
    message = "Select Folder"
    Label(win, text=message).pack()
    Button(win, text='Yield Quality Check', command=yqc).pack()
    Button(win, text = 'Quit', command = win.quit).pack()
  
#Quit 
def quit_application():
    root.destroy()
    
#AGM button
agm_button = ttk.Button(root, text='AGM Data', command = agm_interface)
agm_button.pack(expand=True)

#Yield-Quality-Check Button
yield_button = ttk.Button(root, text = 'Yield Quality Checking Data', command = yield_interface)
yield_button.pack(expand=True)

#Quit Button
quit_button = tk.Button(root, text="Quit", command=quit_application)
quit_button.pack(expand=False)

root.mainloop()