import pandas as pd  
import os
from pathlib import Path
import tkinter as tk
from tkinter import messagebox as msg
from tkinter import filedialog as fld
from zipfile import BadZipFile
from PIL import Image, ImageTk
import pyarrow

INPUT_FILES = []
OUTPUT_FILES = {}
KEY_VAL = {}
SEL_KEY_VAL = {}
CONV_PARQ_FILES = []
DICT_PARQ_FILES = {}
SEL_OUTPUT_FILES = {}
SEL_INPUT_FILES = {}
STATUS = ()

root = tk.Tk()
root.state("zoomed")
root.title("Converter")

bcg_import = Image.open(os.path.join("Graphics/Background.png"))
bcg_picture = ImageTk.PhotoImage(bcg_import)

bg_canvas = tk.Canvas(root, highlightthickness=0)
bg_canvas.place(x=0, y=0, relwidth=1, relheight=1)

bg_canvas.create_image(0, 0, anchor="nw", image=bcg_picture)

def sel_input_files():
    input_files = []
    root.withdraw()
    files = fld.askopenfilenames(
        title="Select just Excel files",
        filetypes=[
            ("Excel Files", "*.xlsx"),
            ("Excel Binary Files", "*.xlsb"),
            ("Excel Files (97-2003)", "*.xls"),
            ("Excel CSV Files", "*.csv"),
        ]
    )
    input_files.clear()
    input_files.extend(files)

    for file in files:
        already_exists = False
        for key, key_val in KEY_VAL.items():
            if file in key_val:
                already_exists = True
                break

        if not already_exists:
            INPUT_FILES.append(file)
            print(f"Item: {file}, inserted")
        else:
            print(f"Item: {file}, is already included in the list")

    KEY_VAL.clear()
    instance.load_input_items()
    instance.load_destination()
    root.deiconify()
    root.state("zoomed")

def sel_destination(event=None):
    if not SEL_INPUT_FILES:
        print("NO VALUES")
        return

    dest_path = fld.askdirectory(title="Select destination path")
    if not dest_path:
        return

    for key in SEL_INPUT_FILES:
        if key in OUTPUT_FILES:
            OUTPUT_FILES[key] = [dest_path]

    instance.output_list.delete(0, tk.END)
    for key, value in OUTPUT_FILES.items():
        instance.output_list.insert(tk.END, value[0])

    SEL_INPUT_FILES.clear()

def convert_to_parquet():
    for key in SEL_INPUT_FILES:
        if key in OUTPUT_FILES:
            input_file = SEL_INPUT_FILES[key]
            output_file = OUTPUT_FILES[key]
            for key,item in OUTPUT_FILES.items():
                output_file = item[0]
            if output_file != "Destination NOT Selected" and Path(output_file).exists():
                extension = Path(input_file).suffix.lower()
                if extension in (".xlsx"):
                    try:
                        file_read = pd.read_excel(input_file, engine="openpyxl")
                        for col in file_read.columns:
                            if file_read[col].dtype == "object":
                                file_read[col] = file_read[col].astype(str)
                        output_folder = Path(output_file)
                        output_folder.mkdir(parents=True, exist_ok=True)
                        input_name = Path(input_file).stem
                        parquet_path = output_folder / f"{input_name}.parquet"
                        file_read.to_parquet(parquet_path, engine="pyarrow")
                        msg.showinfo("SYSTEM", f"Successfully converted:\n{parquet_path}")
                        CONV_PARQ_FILES.append(parquet_path)
                    except BadZipFile:
                        msg.showerror(
                            "Invalid Excel File",
                            "This file is not a valid Excel (.xlsx) file.\n"
                            "It may be corrupted or has the wrong extension."
                        )
                    except Exception as e:
                        msg.showerror("SYSTEM",str(e))

                elif extension in (".xlsb"):
                    try:
                        file_read = pd.read_excel(input_file, engine="pyxlsb")
                        for col in file_read.columns:
                            if file_read[col].dtype == "object":
                                file_read[col] = file_read[col].astype(str)
                        output_folder = Path(output_file)
                        output_folder.mkdir(parents=True, exist_ok=True)
                        input_name = Path(input_file).stem
                        parquet_path = output_folder / f"{input_name}.parquet"
                        file_read.to_parquet(parquet_path, engine="pyarrow")
                        msg.showinfo("SYSTEM", f"Successfully converted:\n{parquet_path}")
                        CONV_PARQ_FILES.append(parquet_path)
                    except BadZipFile:
                        msg.showerror(
                            "Invalid Excel File",
                            "This file is not a valid Binary-Excel (.pyxlsb) file.\n"
                            "It may be corrupted or has the wrong extension."
                        )
                    except Exception as e:
                        msg.showerror("SYSTEM",str(e))

                elif extension in (".xlrd"):
                    try:
                        file_read = pd.read_excel(input_file, engine="xlrd")
                        for col in file_read.columns:
                            if file_read[col].dtype == "object":
                                file_read[col] = file_read[col].astype(str)
                        output_folder = Path(output_file)
                        output_folder.mkdir(parents=True, exist_ok=True)
                        input_name = Path(input_file).stem
                        parquet_path = output_folder / f"{input_name}.parquet"
                        file_read.to_parquet(parquet_path, engine="pyarrow")
                        msg.showinfo("SYSTEM", f"Successfully converted:\n{parquet_path}")
                        CONV_PARQ_FILES.append(parquet_path)
                    except BadZipFile:
                        msg.showerror(
                            "Invalid Excel File",
                            "This file is not a valid Excel (.xls) file.\n"
                            "It may be corrupted or has the wrong extension."
                        )
                    except Exception as e:
                        msg.showerror("SYSTEM",str(e))

                elif extension in (".csv"):
                    try:
                        file_read = pd.read_excel(input_file, engine="pyarrow")
                        for col in file_read.columns:
                            if file_read[col].dtype == "object":
                                file_read[col] = file_read[col].astype(str)
                        output_folder = Path(output_file)
                        output_folder.mkdir(parents=True, exist_ok=True)
                        input_name = Path(input_file).stem
                        parquet_path = output_folder / f"{input_name}.parquet"
                        file_read.to_parquet(parquet_path, engine="pyarrow")
                        msg.showinfo("SYSTEM", f"Successfully converted:\n{parquet_path}")
                        CONV_PARQ_FILES.append(parquet_path)
                    except BadZipFile:
                        msg.showerror(
                            "Invalid Excel File",
                            "This file is not a valid Excel(comma separated value) (.csv) file.\n"
                            "It may be corrupted or has the wrong extension."
                        )
                    except Exception as e:
                        msg.showerror("SYSTEM",str(e))
            else:
                msg.showerror("SYSTEM","Destination NOT Selected")
                break
    instance.load_converted_files()

class Main_Window(tk.Frame):
    def __init__ (self,widgets):
        super().__init__(widgets)
        self.pack(fill="both",expand=True)
        self.frame_toolbar = tk.Frame(self)
        self.frame_toolbar.pack(fill="both",anchor="ne")
        self.show_widgets()
        
    def show_widgets(self):

        tk.Label(self.frame_toolbar, text="").pack(side="left", padx=50)
        tk.Label(self.frame_toolbar, text="Choose Files to convert",font=("Segoe UI", 14, "bold")).pack(side="left", padx=5)

# Button 1
        self.button_1_img = Image.open("Graphics/Button_Select_files.png")
        self.button_1_img = self.button_1_img.resize((200, 50), Image.LANCZOS)
        self.button_1_photo = ImageTk.PhotoImage(self.button_1_img)
        self.button_1 = tk.Button(
            self.frame_toolbar,
            image=self.button_1_photo,
            command=sel_input_files,borderwidth=0
        )
        self.button_1.pack(side="left", padx=5)
# Button 2
        self.button_2_img = Image.open("Graphics/Button_select_destination.png")
        self.button_2_img = self.button_2_img.resize((200, 50), Image.LANCZOS)
        self.button_2_photo = ImageTk.PhotoImage(self.button_2_img)
        self.button_2 = tk.Button(
            self.frame_toolbar,
            image=self.button_2_photo,
            command=sel_destination,borderwidth=0
        )
        self.button_2.pack(side="left", padx=5)
# Button 3
        self.button_3_img = Image.open("Graphics/Button_convert_parquet.png")
        self.button_3_img = self.button_3_img.resize((200, 50), Image.LANCZOS)
        self.button_3_photo = ImageTk.PhotoImage(self.button_3_img)
        self.button_3 = tk.Button(
            self.frame_toolbar,
            image=self.button_3_photo,
            command=convert_to_parquet,borderwidth=0
        )
        self.button_3.pack(side="left", padx=5)

# Button 4
        self.button_4_img = Image.open("Graphics/Button_Open_converted.png")
        self.button_4_img = self.button_4_img.resize((200, 50), Image.LANCZOS)
        self.button_4_photo = ImageTk.PhotoImage(self.button_4_img)
        self.button_4 = tk.Button(
            self.frame_toolbar,
            image=self.button_4_photo,
            command=self.open_converted_files,borderwidth=0
        )
        self.button_4.pack(side="left", padx=5)

        frame_input = tk.Frame(self)
        frame_input.pack(expand=True,fill=tk.BOTH)
        scroll_x = tk.Scrollbar(frame_input,orient=tk.HORIZONTAL)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        scroll_y = tk.Scrollbar(frame_input,orient=tk.VERTICAL)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        frame_input2 = tk.Frame(self)
        frame_input2.pack(expand=True,fill=tk.BOTH,side="bottom")
        tk.Label(frame_input2,text="Converted Files",font=("Segoe UI", 14, "bold")).pack(side="top",anchor="nw")
        self.conv_list = tk.Listbox(frame_input2,selectmode=tk.SINGLE,width=100,height=100)

        self.input_list = tk.Listbox(frame_input,selectmode=tk.MULTIPLE,width=100,height=35,xscrollcommand=scroll_x.set,yscrollcommand=scroll_y.set)
        self.input_list.pack(side="left", padx=10, pady=10, fill="both", expand=True)

        self.output_list = tk.Listbox(frame_input,width=100,height=35,xscrollcommand=scroll_x.set,yscrollcommand=scroll_y.set)
        self.output_list.pack(side="right", padx=10, pady=10, fill="both", expand=True)

        self.conv_list.pack(side="bottom", padx=10, pady=10, fill="both", expand=True)

        self.input_list.bind("<<ListboxSelect>>", self.sel_input_files)
        self.output_list.bind("<<Listboxselect>>",self,sel_destination)
        self.conv_list.bind("<<Listboxselect>>",self.open_converted_files)

    def load_input_items(self):
        start_indx = len(KEY_VAL)

        for index, file in enumerate(INPUT_FILES, start=start_indx):
            KEY_VAL[index] = [file]
        self.input_list.delete(0, tk.END)

        for key, value in KEY_VAL.items():
            self.input_list.insert(tk.END, value[0])
        print("INPUT:", KEY_VAL)
        self.load_existing_ouput_items()

    def load_destination(self):
        self.output_list.delete(0, tk.END)
        placeholder = "Destination NOT Selected"
        for index, file in enumerate(INPUT_FILES):
            destination = OUTPUT_FILES.get(index)
            if not destination:
                OUTPUT_FILES[index] = placeholder
                self.output_list.insert(tk.END, placeholder)
            else:
                self.output_list.insert(tk.END, destination)
        print("OUTPUT DESTINATION VALUES:", OUTPUT_FILES)

    def load_existing_ouput_items(self):

        self.output_list.delete(0, tk.END)
        for key,item in OUTPUT_FILES.items():
            self.output_list.insert(tk.END,item[0])
        print("OUTPUT DESTINATION VALUES:",OUTPUT_FILES)
        

    def sel_output_files(self,event):
        global SEL_OUTPUT_FILES
        SEL_OUTPUT_FILES=[self.output_list.get(i) for i in self.output_list.curselection()]
        print(SEL_OUTPUT_FILES)       
    
    def sel_input_files(self,event):
        global SEL_INPUT_FILES
        global SEL_KEY_VAL
        SEL_INPUT_FILES.clear()
        for i in self.input_list.curselection():
            SEL_INPUT_FILES[i] = self.input_list.get(i)
            print("ITEMS_SELECTED:",SEL_INPUT_FILES)

    def load_converted_files(self):
        for key,item in enumerate(CONV_PARQ_FILES):
            DICT_PARQ_FILES[key] = item
            print("adscascasc",DICT_PARQ_FILES)
        
        self.conv_list.delete(0,tk.END)
        for index,item in DICT_PARQ_FILES.items():
            self.conv_list.insert(tk.END, str(item))

    def open_converted_files(self):
        selected_files = [self.conv_list.get(i) for i in self.conv_list.curselection()]
        
        for file_path in selected_files:
            print("OPEN THIS FILE:", file_path)
            if Path(file_path).exists():
                os.startfile(file_path)
            else:
                print(f"File does not exist: {file_path}")

def data_folder_loops(sel_years,*sel_dept):
    source_path = Path(r'\\micro-p.com\mpnorth\users\Accounts\Accounts Shared\Commercial Finance - Ops\Operations\Operations\FY26\01_Resource Planning\00_SAP Data')
    if source_path.exists() == False:
        print("Var: source_path: I Do not exist")
    
    for dept in sel_dept:
        path = os.path.join(source_path,dept)

if __name__ == "__main__":
    instance = Main_Window(root)
    root.mainloop()
