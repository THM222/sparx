import tkinter as tk
import tkinter.messagebox as msgbox
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter import filedialog as fd
from modules import sparx

import os, platform, subprocess, sys

import pdb

WINDOW_WIDTH = 650
WINDOW_MIN_HEIGHT = 180
WINDOW_MAX_HEIGHT = 400


class Window(tk.Tk):
    def __init__(self):
        super().__init__()
        self.eval('tk::PlaceWindow . center')

        self.title = 'THM Sparx Data Processor'
        self.resizable(True, True)
        self.geometry(f'{WINDOW_WIDTH}x{WINDOW_MIN_HEIGHT}')
        self.advanced = tk.Frame(self)

        self.args = sparx.Parameters()

        self.open_file_text = tk.StringVar()
        self.save_dir_text = tk.StringVar()

        self.num_weeks_text = tk.IntVar()
        self.xp_top_n_text = tk.IntVar()
        self.il_top_n_text = tk.IntVar()
        self.il_min_time_mins_text = tk.IntVar()

        self.process_yg_chk = tk.BooleanVar()
        self.process_rg_chk = tk.BooleanVar()
        self.process_mc_chk = tk.BooleanVar()
        self.process_xp_chk = tk.BooleanVar()
        self.process_il_chk = tk.BooleanVar()

        self.set_defaults()

        open_label = ttk.Label(self, text='Sparx input data file:')
        open_label.grid(row=0, column=0, sticky='w', padx=(10, 10), pady=(10,0))
        
        self.open_file_entry = tk.Entry(self, width=50, textvar=self.open_file_text)
        self.open_file_entry.grid(row=1, column=0, sticky='nsew', padx=10, pady=0)

        open_button = ttk.Button(self, text='Choose File', command=self.select_file)
        open_button.grid(row=1, column=1, sticky='w', padx=0, pady=0)

        save_label = ttk.Label(self, text='Output directory:')
        save_label.grid(row=2, column=0, sticky='w', padx=(10, 10), pady=(10,0))

        self.save_dir_entry = tk.Entry(self, width=50, textvar=self.save_dir_text)
        self.save_dir_entry.grid(row=3, column=0, sticky='nsew', padx=10, pady=0)

        save_button = ttk.Button(self, text='Choose Directory', command=self.select_directory)
        save_button.grid(row=3, column=1, sticky='w', padx=0, pady=0)

        # Run
        run_button = ttk.Button(self, text='Run!', command=self.process_input)
        run_button.grid(row=5, column=0, sticky='w', padx=10, pady=10)

        self.advanced_button = ttk.Button(self, text='Advanced ►', command=self.advanced_on_click)
        self.advanced_button.grid(row=5, column=1, sticky='w', padx=(10, 10), pady=(10,0))
        self.advanced_visible = False
        
        self.create_advanced_frame()


    def create_advanced_frame(self): #▼ ►
        num_weeks_label = ttk.Label(self.advanced, text='Weeks to process:').grid(row=0, column=0, padx=10, pady=0, sticky="e")
        self.num_weeks_entry = tk.Entry(self.advanced, width=5, textvar=self.num_weeks_text).grid(row=0, column=1, padx=10, pady=0)
        
        xp_top_n_label = ttk.Label(self.advanced, text='XP Boost Top N:').grid(row=1, column=0, padx=10, pady=0, sticky="e")
        self.xp_top_n_entry = tk.Entry(self.advanced, width=5, textvar=self.xp_top_n_text).grid(row=1, column=1, padx=10, pady=0)
        
        il_top_n_label = ttk.Label(self.advanced, text='Independent Learning Top N:').grid(row=2, column=0, padx=10, pady=0, sticky="e")
        self.il_top_n_entry = tk.Entry(self.advanced, width=5, textvar=self.il_top_n_text).grid(row=2, column=1, padx=10, pady=0)
        
        il_min_time_mins_label = ttk.Label(self.advanced, text='Minimum minutes:').grid(row=2, column=2, padx=10, pady=0, sticky="e")
        self.il_min_time_mins_entry = tk.Entry(self.advanced, width=5, textvar=self.il_min_time_mins_text).grid(row=2, column=3, padx=10, pady=0)

        # Create a checkbox
        chk_yg = tk.Checkbutton(self.advanced, text="Process year group data", var=self.process_yg_chk).grid(row=4, column=0, columnspan=3, sticky="w")
        chk_mc = tk.Checkbutton(self.advanced, text="Process maths class data", var=self.process_mc_chk).grid(row=5, column=0, columnspan=3, sticky="w")
        chk_rg = tk.Checkbutton(self.advanced, text="Process registration group data", var=self.process_rg_chk).grid(row=6, column=0, columnspan=3, sticky="w")
        chk_il = tk.Checkbutton(self.advanced, text="Process independent learning", var=self.process_il_chk).grid(row=7, column=0, columnspan=3, sticky="w")
        chk_xp = tk.Checkbutton(self.advanced, text="Process XP boost", var=self.process_xp_chk).grid(row=8, column=0, columnspan=3, sticky="w")


    def advanced_on_click(self):
        if self.advanced_visible:
            self.advanced_visible = False
            self.advanced.grid_remove()
            self.geometry(f'{WINDOW_WIDTH}x{WINDOW_MIN_HEIGHT}')
            self.advanced_button.configure(text='Advanced ►')
        else:
            self.advanced_visible = True
            self.advanced.grid(row=7, column=0, columnspan=5, sticky='w', padx=10, pady=10)
            self.geometry(f'{WINDOW_WIDTH}x{WINDOW_MAX_HEIGHT}')
            self.advanced_button.configure(text='Advanced ▼')
     

    def set_defaults(self):
        self.num_weeks_text.set(2)
        self.xp_top_n_text.set(10)
        self.il_top_n_text.set(10)
        self.il_min_time_mins_text.set(20)
        self.process_yg_chk.set(True)
        self.process_rg_chk.set(True)
        self.process_mc_chk.set(True)
        self.process_xp_chk.set(True)
        self.process_il_chk.set(True)


    def select_file(self):
        filetypes = (
            ('xlsx files', '*.xlsx'),
            ('All files', '*.*')
        )
        f = fd.askopenfilename(title='Select file', filetypes=filetypes, initialdir="/")
        self.open_file_text.set(f)
        self.args.input_file = f


    def select_directory(self):
        d = fd.askdirectory(title='Select', mustexist=True, initialdir="/")
        self.save_dir_text.set(d)
        self.args.output_dir = d


    def process_input(self):
        self.update_args()
        print(f"Running with parameters: {self.args}")
        sparx.run(self.args)
        msgbox.showinfo('Complete!', f'Saved to {self.args.output_dir}', icon=msgbox.INFO)
        if platform.system() == "Windows":
           os.startfile(self.args.output_dir)
        else:
           opener = "open" if sys.platform == "darwin" else "xdg-open"
           subprocess.call([opener, self.args.output_dir])


    def update_args(self):
        self.args.num_weeks = self.num_weeks_text.get()
        self.args.input_file = self.open_file_entry.get()
        self.args.output_dir = self.save_dir_entry.get()
        self.args.xp_top_n = self.xp_top_n_text.get()
        self.args.il_top_n = self.il_top_n_text.get()
        self.args.il_min_time_mins = self.il_min_time_mins_text.get()

        self.args.process_year_group = self.process_yg_chk.get()
        self.args.process_maths_class = self.process_mc_chk.get()
        self.args.process_reg_group = self.process_rg_chk.get()
        self.args.process_xp_boost = self.process_xp_chk.get()
        self.args.process_independent_learning = self.process_il_chk.get()


if __name__ == "__main__":
    window = Window()
    window.mainloop()

