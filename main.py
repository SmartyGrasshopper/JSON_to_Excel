from genericpath import isdir
import tkinter as tk
from tkinter import scrolledtext
from tkinter import filedialog
from tkinter import messagebox
import Widgets as widgets
import os
import json
from json_excel_converter import Converter
from json_excel_converter.xlsx import Writer

class App(tk.Tk):
    ''' - Class representing the main application
        - Author(s): @author1  (See autors.txt file)
    '''
    def __init__(self):
        super().__init__()

        # variables for widgets and UI that need access throughout the app
        self.padx, self.pady = 5, 5
        self.buttonColor = 'Cyan3'
        self.accentButtonColor = 'Cyan2'
        self.JSONFilesDisplay:scrolledtext = None
        self.exportFileNameBox:tk.Entry = None
        self.exportDirectoryBox:tk.Entry = None

        # variables for data
        self.listOfJSONFiles = []
        self.ExportFileName = ''
        self.ExportFolder = ''
        self.lastPath = os.getcwd()

    def launch(self):
        self._createAppUI()
        self.mainloop()
    
    def _createAppUI(self):
        ''' function to build/create the UI of the application '''
        self.title('JSON to Excel')
        self.minsize(height = 500, width = 600)

        # creating the frame for handling JSON files
        JSON_Frame = tk.LabelFrame(
            master = self,
            text = 'Select JSON file(s)'
        )
        JSON_Frame.pack(side = 'top', fill = 'both', expand = True, padx = self.padx, pady = self.pady)

        self.JSONFilesDisplay = scrolledtext.ScrolledText(master = JSON_Frame, state = 'disabled')
        self.JSONFilesDisplay.pack(side = 'left', fill = 'both', expand = 'True', padx = self.padx, pady = self.pady)

        ButtonsFrame = tk.Frame(master = JSON_Frame)
        ButtonsFrame.pack(side = 'right', padx = self.padx, pady = self.pady, fill = 'y')

        widgets.customFlatButton(
            master = ButtonsFrame,
            text = 'Add JSON File', 
            backgroundColor = self.buttonColor,
            accentColor = self.accentButtonColor,
            command = self._AddNewJSONFile
        ).pack(pady = self.pady, fill = 'x')

        widgets.customFlatButton(
            master = ButtonsFrame,
            text = 'Clear Files List', 
            backgroundColor = self.buttonColor,
            accentColor = self.accentButtonColor,
            command = self._ClearJSONFiles
        ).pack(pady = self.pady, fill = 'x')

        
        # creating the frame for exporting the files
        ExportFrame = tk.LabelFrame(
            master = self,
            text = 'Export/Flatten the JSON file(s) to an Excel file'
        )
        ExportFrame.pack(side = 'bottom', fill = 'x', expand = False, padx = self.padx, pady = self.pady)
        '''
        tk.Label(
            master = ExportFrame,
            text = 'File Name (excel)'
        ).grid(row = 0, column = 0, padx = self.padx, pady = self.pady, sticky = tk.W)

        self.exportFileNameBox = tk.Entry(
            master = ExportFrame,
            width = 50
        )
        self.exportFileNameBox.grid(row = 0, column = 1, padx = self.padx, pady = self.pady, sticky = (tk.E, tk.W), columnspan = 2)
        '''
        tk.Label(
            master = ExportFrame,
            text = 'Folder/Directory'
        ).grid(row = 1, column = 0, padx = self.padx, pady = self.pady, sticky = tk.W)

        self.exportDirectoryBox = tk.Entry(
            master = ExportFrame,
            width = 50
        )
        self.exportDirectoryBox.grid(row = 1, column = 1, padx = self.padx, pady = self.pady, sticky = tk.E)

        widgets.customFlatButton(
            master = ExportFrame,
            text = 'Browse', 
            backgroundColor = self.buttonColor,
            accentColor = self.accentButtonColor,
            command = self._getExportDirectory
        ).grid(row = 1, column = 2, padx = self.padx, pady = self.pady)


        widgets.customFlatButton(
            master = ExportFrame,
            text = 'Export Unflattened', 
            backgroundColor = self.buttonColor,
            accentColor = self.accentButtonColor,
            command = lambda: self._flattenAndExport(flatten = False)
        ).grid(row = 2, column = 0, padx = self.padx, pady = self.pady, columnspan = 3, sticky = (tk.E, tk.W))

        widgets.customFlatButton(
            master = ExportFrame,
            text = 'Flatten & Export', 
            backgroundColor = self.buttonColor,
            accentColor = self.accentButtonColor,
            command = lambda: self._flattenAndExport(flatten = True)
        ).grid(row = 3, column = 0, padx = self.padx, pady = self.pady, columnspan = 3, sticky = (tk.E, tk.W))

    def _getExportDirectory(self):
        directory = filedialog.askdirectory(title = 'Select Export Directory', initialdir = self.lastPath)
        if(directory != None and directory != ''):
            self.exportDirectoryBox.delete(0, tk.END)
            self.exportDirectoryBox.insert(0, directory)

    def _flattenAndExport(self, flatten=True):
        ''' function to handle all the flattening and export work '''
        if(self.listOfJSONFiles != []):
            #JSON_DataFrames = pd.DataFrame()
            if(self.exportDirectoryBox.get().strip() != '' and isdir(self.exportDirectoryBox.get().strip())):
                try:
                    for filepath in self.listOfJSONFiles:
                        with open(filepath, 'r') as file:
                            #pd.DataFrame(json.loads(file.read())).to_csv('{}/{}.csv'.format(self.exportDirectoryBox.get().strip(), filepath.split('/')[-1].split('.')[0]), index = False)
                            Conv = Converter()
                            Conv.convert(
                                [self._flatten(json.load(file))] if flatten else [json.load(file)], 
                                Writer(
                                    file = '{}/{}.xlsx'.format(self.exportDirectoryBox.get().strip(), filepath.split('/')[-1].split('.')[0])
                                    )
                                )

                except:
                    messagebox.showerror(title = 'Error', message = 'Some error occred while reading the json files. Ensure they are not corrupt and export path is accessible tp write.')
                else:
                    messagebox.showinfo(title = 'Success', message = 'All files successfully exported in {}'.format(self.exportDirectoryBox.get().strip()))
            else:
                messagebox.showerror(message = 'Select a valid export folder to export the excel files.')
        else:
            messagebox.showinfo(message = 'No JSON files selected (to flatten).')

    def _AddNewJSONFile(self):
        filepaths = filedialog.askopenfilenames(
            title = 'Select json/JSON Files to Flatten',
            filetypes = (
                ('json', '*.json'),
                ('JSON', '*.JSON')
            ),
            initialdir = self.lastPath
        )

        if(filepaths != '' and filepaths != None and filepaths != ()):
            self.JSONFilesDisplay.configure(state = 'normal')
            for filepath in filepaths: self.JSONFilesDisplay.insert(tk.END, '{}\n'.format(filepath))
            self.JSONFilesDisplay.configure(state = 'disabled')
            self.listOfJSONFiles.extend(filepaths)

            # updating the last opened path variable
            self.lastPath = filepaths[-1].removesuffix('/{}'.format(filepath.split(r'/')[-1]))

    def _ClearJSONFiles(self):
        self.listOfJSONFiles = []
        self.JSONFilesDisplay.configure(state = 'normal')
        self.JSONFilesDisplay.delete(0.0, tk.END)
        self.JSONFilesDisplay.configure(state = 'disabled')

    def _flatten(self, json:dict) -> dict:
        ''' - function to flatten a json data (given in form of json dictionary)
            - returns json data (in for of dictionary)
        '''
        out = {}
        def flatten(x, name = ''):
            # source: geeksforgeeks.org
            if type(x) is dict:
                for a in x: flatten(x[a], name + a + '-')
            elif type(x) is list:
                i = 0
                for a in x:
                    flatten(a, name + str(i) + '-')
            else:
                out[name[:-1]] = x

        flatten(json)
        return out

if __name__ == '__main__':
    App().launch()
