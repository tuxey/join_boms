# -*- coding: utf-8 -*-
"""
Created on Fri Jul 30 18:47:16 2021
@author: Pablo
"""

import os
import glob
import itertools
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

class MainApplication(tk.Frame):
  def __init__(self, parent):
    tk.Frame.__init__(self, parent)
    self.parent = parent

    # the rest of the GUI
    
    self.frm_getfolder = tk.Frame(master=parent, width=2000, height=50)
    self.frm_getfolder.pack()

    btn_choosedir = tk.Button(
      text='Choose the folder with CSV BOMs',
      master=self.frm_getfolder,
      command=self.get_workfolder
    )
    btn_choosedir.pack()

    self.lbl_folder = tk.Label(text='The path to the chosen folder will appear here', master=self.frm_getfolder)
    self.lbl_folder.pack()

    self.frm_readfiles = tk.Frame(height=5)
    self.frm_readfiles.pack()
    
    self.lbl_files = tk.Label(master=self.frm_readfiles)
    self.lbl_files.pack()

    # main variables
    
    self.headers = ['Designator', 'Description', 'Comment', 'Footprint', 'Quantity']
    self.file_header = "BOM_unified_buying-"
    self.allowed_to_continue = False

  # 1. get file directory and set as working/current directory
  def get_workfolder(self):
    self.filepath = filedialog.askdirectory()
    print(self.filepath)
    self.lbl_folder['text'] = self.filepath
    os.chdir(self.filepath)

    if self.filepath == None:
      messagebox.showinfo(message='Please, choose a folder', title='Choose a folder')
    else:
      # continue the script
      self.find_csv_files()
    
    if self.allowed_to_continue:
      # Show 'Join BOMs' button

      self.frm_btnjoin = tk.Frame(master=self.parent, width=50, height=50)
      self.frm_btnjoin.pack()

      btn_join_boms = tk.Button(
        text='Join BOMs',
        master=self.frm_btnjoin,
        command=self.join_boms
      )
      btn_join_boms.pack()

  # 2. Find all csv files in the folder, check for errors in filenames
  def find_csv_files(self):
    extension = "csv"
    self.all_filenames = [i for i in glob.glob(f"*.{extension}")] # glob pattern matching
    self.n_files = len(self.all_filenames)
    self.allowed_to_continue = False

    if self.n_files >0:
      # 3. Get project names from filenames
      self.project_names = [name.replace(self.file_header, "").replace(f".{extension}", "") for name in self.all_filenames]
      self.lbl_files['text'] = [str(i + '\n') for i in self.all_filenames]
      
      flag_header_error = False
      for name in self.all_filenames:
        if (name.find(self.file_header) == -1):
          print('File header not found.')
          messagebox.showwarning(
            message='Error. CSV files must start with {}\nFile: {}'.format(self.file_header, name)
          )
          flag_header_error = True
      
      if not flag_header_error:
        self.read_BOM_files()


    else: # 0 csv files
      self.lbl_files['text'] = 'No CSV files found in selected folder.'
      messagebox.showwarning(
        message='No CSV files found in selected folder.\nSelect a valid folder.'
      )

  # 3. Read BOM files and check for errors in headers
  def read_BOM_files(self):
    self.boms = []
    header_mismatch = False
    error_msg = ''
    for i in range(self.n_files):
      # read csv
      bom_csv = pd.read_csv(self.all_filenames[i], encoding_errors='replace') # replace micro symbol with '�'

      # append to the list
      self.boms.append(bom_csv)

      # check if all CSVs have correct headers
      curr_headers = self.boms[i].columns.values.tolist()

      if (not(curr_headers == self.headers)):
        error_msg += 'Header mismatch in file:\n\t' + self.all_filenames[i] + '\n'
        error_msg += 'Wrong headers:\n\t'+ ', '.join(curr_headers) + '\n'
        
        header_mismatch = True
        
    if header_mismatch:
      error_msg += '\nThe correct headers are:' + '\t' + ', '.join(self.headers) + '\n'
      messagebox.showerror(message=error_msg, title = 'Header error')
    
    else:
      self.allowed_to_continue = True

  # 4. Do all the joining stuff
  def join_boms(self):
    self.make_project_qty_columns()

    combined_csv = pd.concat(self.boms)

    out_csv = self.group_components(combined_csv)

    self.remove_decimals_and_mu(out_csv)
    self.designator_column_only_letters(out_csv, self.headers[0])

    self.sort_by_columns(out_csv)
    self.rename_column_to(out_csv, self.headers[0], 'Type')

    self.export_to_csv(out_csv)

  # 4.1. Rename 'Quantity' column to Project Name
  def make_project_qty_columns(self):
    for i in range(len(self.project_names)):
      self.boms[i].rename(
        columns=({ self.headers[4]: self.project_names[i]}),
        inplace=True,
      )

  # 4.2. Group same components (only if all Description, Comment and Footprint match)
  def group_components(self, in_csv):
    aggregation_functions = {
      self.headers[0]: 'first',
      self.headers[1]: 'first', 
      self.headers[2]: 'first', 
      self.headers[3]: 'first'
    }
    for project in self.project_names:
      aggregation_functions[project] = 'first'
    
    return in_csv.groupby([
      in_csv[self.headers[1]],
      in_csv[self.headers[2]],
      in_csv[self.headers[3]]
    ]).aggregate(aggregation_functions)

  # 4.3. Make quantities columns without decimals
  def remove_decimals_and_mu(self, in_csv):
    for project in self.project_names:
      in_csv[project] = in_csv[project].apply(lambda x : str(x).replace('.0',''))
      in_csv[project] = in_csv[project].apply(lambda x : str(x).replace('nan','0'))

      in_csv[self.headers[2]]= in_csv[self.headers[2]].apply(lambda x : str(x).replace('�F','uF'))
      in_csv[self.headers[2]]= in_csv[self.headers[2]].apply(lambda x : str(x).replace('�H','uH'))

  # 4.4. From Designator column, get only the reference designator letter (like R, C, etc.)
  def designator_column_only_letters(self, in_csv, column):
    for item in in_csv[column]:
      curr_replacement = str("".join(itertools.takewhile(str.isalpha, item))) # remove all except the first letters before a number
      in_csv.replace(item, value = curr_replacement, inplace=True)

  # 4.5. Sort by columns
  def sort_by_columns(self, to_sort_csv):
    to_sort_csv.sort_values(by=['Designator'], axis=0,inplace=True)

  # 4.6. Rename 'Designator' column to 'Type'
  def rename_column_to(self, to_rename_csv, column_from, column_to):
    to_rename_csv.rename(columns={column_from: column_to}, inplace=True)

  # 4.7. Export to csv
  def export_to_csv(self, to_export_csv):
    Path((str(self.filepath) + "\\output")).mkdir(parents=True, exist_ok=True)
    to_export_csv.to_csv((str(self.filepath) + "\\output\\joined_BOM.csv"), index=False, encoding='utf-8-sig')

    success_message = f'Successfully merged {self.n_files} csv BOMs:\n'
    for file in self.all_filenames:
      success_message += f'- {file}\n'
    success_message += '\nFinal BOM is placed in ./output/joined_BOM.csv'
    messagebox.showinfo(message=success_message, title='Success')
    print(success_message)


  



if __name__ == '__main__':
  window = tk.Tk()
  window.minsize(height=20, width=200)
  mainapp = MainApplication(window).pack(side='top', fill='both', expand=True)
  window.mainloop()
