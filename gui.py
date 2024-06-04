import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog as fd
import pandas as pd
from topsis import Topsis
import numpy as np
import itertools


class Topsis_GUI:
    def __init__(self, root):
        self.root = root
        self.root.geometry('250x300')
        self.root.title('Rekomendasi Sekolah')

        self.titel = tk.Label(self.root, text='TOPSIS')
        self.titel.pack(pady=20)

        self.input_btn = tk.Button(root, text="Pilih File", command=self.open_file)
        self.input_btn.pack(pady=10, ipadx=3, ipady=3)
        
    def topsis_calculate(self): #function untuk memanggil topsis dari mmodul topsis 
        evaluation_matrix = self.get_array()
        weight = self.weights
        criteria = np.array([True, True, True, False, False])
        
        tp = Topsis(evaluation_matrix, weight, criteria)

        tp.calc() #kalkulasi topsis

        self.sklh = list(itertools.chain(*self.sekolah.to_numpy().tolist())) #merubah array 2d menjadi 1d

        self.normal = tp.normalized_decision
        self.normal = self.normal.tolist()
        for i in range(len(self.normal)):
            self.normal[i].insert(0, self.sklh[i]) #membuat berisi berisi normalisasi

        self.weighted = tp.weighted_normalized 
        self.weighted = self.weighted.tolist()
        for i in range(len(self.weighted)): 
            self.weighted[i].insert(0, self.sklh[i]) #membuat berisi berisi normalisasi terbobot

        self.best_alt = tp.best_alternatives.tolist()
        self.worst_alt = tp.worst_alternatives.tolist()

        print(self.best_alt, self.worst_alt)

        self.best_alt.insert(0, 'A+')
        self.worst_alt.insert(0, 'A-')

        self.weighted.append(self.best_alt)
        self.weighted.append(self.worst_alt) #memasukkan baris A+ dan A- ke tabel normalisasi terbobot

        self.dminus = tp.worst_distance
        self.dplus = tp.best_distance

        self.dplus = self.dplus.tolist()
        self.dminus = self.dminus.tolist()

        self.d = np.vstack((self.dplus, self.dminus))
        self.d = self.d.T.tolist()

        for i in range(len(self.d)):
            self.d[i].insert(0, self.sklh[i]) #membuat baris D+ dan D-
        
        self.worst = tp.worst_similarity.tolist()
        self.rank = self.rank_array_descending(self.worst)
        self.rank = self.rank.tolist()

        self.v = np.vstack((self.worst, self.rank))
        self.v = self.v.T.tolist() 
        
        for i in range(len(self.v)):
            self.v[i].insert(0, self.sklh[i]) #membuat row untuk v dan ranking

    def get_array(self): #mendapatkan array matrix dari excel
        self.array = self.df.to_numpy()
        self.array = np.delete(self.array, 0, axis=1)
        self.array = self.array.astype(float)
        return self.array
    
    def rank_array_descending(self, arr): #mengurutkan hasil dari v atau meranking
        sorted_indices = np.argsort(arr)
        sorted_indices_descending = sorted_indices[::-1]
        ranks = np.empty_like(sorted_indices_descending)
        ranks[sorted_indices_descending] = np.arange(1, len(arr) + 1)
        return ranks
    
    def open_file(self): #memilih file
        self.file = fd.askopenfilename(title='Pilih File', filetypes=[('Excel file', '.xls*')])
        if self.file:
            self.msg = messagebox.showinfo("File Status", 'File telah tepilih')
            self.window_bobot()
        else:
            self.msg = messagebox.showinfo("File Status", 'File belum tepilih')

    def read_file(self): #membaca file excel
        self.df = pd.read_excel(self.file, engine='openpyxl')
        self.sekolah = pd.read_excel(self.file, usecols='A', engine='openpyxl')
    
    def create_table(self, parent, title, columns, rows): #membuat table untuk nilai, normalisasi, dan normalisasi terbobot
        self.lbl = tk.Label(parent, text=title, font=("Family", 20, "bold"))
        self.lbl.pack(pady=10)

        tree = ttk.Treeview(parent)
        tree.pack(fill=tk.X, expand=False, padx=10, pady=10)

        tree['columns'] = list(columns)
        tree['show'] = 'headings'

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, anchor='w', stretch=True, width=90, minwidth=75)

        df_rows = rows
        for row in df_rows:
            tree.insert("", "end", values=list(row))
        
        return tree
    
    def create_table_d(self, parent, title, columns, rows): #membuat table untuk jarak ideal dan hasil
        self.nilai_lbl = tk.Label(parent, text=title, font=("Family", 20, "bold"))
        self.nilai_lbl.pack(pady=10)

        tree = ttk.Treeview(parent)
        tree.pack(fill=tk.X, expand=False, padx=10, pady=10)

        tree['columns'] = list(columns)
        tree['show'] = 'headings'

        self.column_list = columns

        for col in self.column_list:
            tree.heading(col, text=col)
            tree.column(col, anchor='w', stretch=True, width=90, minwidth=75)

        df_rows = rows
        for row in df_rows:
            tree.insert("", "end", values=list(row))
        
        return tree
    
    def window_bobot(self): #membuat window untuk memasukkan bobot
        self.bobot_wd = tk.Toplevel(self.root)
        self.bobot_wd.geometry('240x500')
        self.bobot_wd.title("Bobot")

        self.akred_lbl = tk.Label(self.bobot_wd, text="Akreditasi", font=("Family", 14, "bold"))
        self.akred_lbl.pack()
        self.benefit_lbl = tk.Label(self.bobot_wd, text="(Benefit)")
        self.benefit_lbl.pack()
        self.akred_inp = tk.Entry(self.bobot_wd, width=5, justify="center")
        self.akred_inp.pack(pady=5)

        self.fasilitas_lbl = tk.Label(self.bobot_wd, text="Fasilitas", font=("Family", 14, "bold"))
        self.fasilitas_lbl.pack()
        self.benefit_lbl = tk.Label(self.bobot_wd, text="(Benefit)")
        self.benefit_lbl.pack()
        self.fasilitas_inp = tk.Entry(self.bobot_wd, width=5, justify="center")
        self.fasilitas_inp.pack(pady=5)

        self.prestasi_lbl = tk.Label(self.bobot_wd, text="Prestasi", font=("Family", 14, "bold"))
        self.prestasi_lbl.pack()
        self.benefit_lbl = tk.Label(self.bobot_wd, text="(Benefit)")
        self.benefit_lbl.pack()
        self.prestasi_inp = tk.Entry(self.bobot_wd, width=5, justify="center")
        self.prestasi_inp.pack(pady=5)

        self.biaya_lbl = tk.Label(self.bobot_wd, text="Biaya", font=("Family", 14, "bold"))
        self.biaya_lbl.pack()
        self.cost_lbl = tk.Label(self.bobot_wd, text="(Cost)")
        self.cost_lbl.pack()
        self.biaya_inp = tk.Entry(self.bobot_wd, width=5, justify="center")
        self.biaya_inp.pack(pady=5)

        self.jarak_lbl = tk.Label(self.bobot_wd, text="Jarak", font=("Family", 14, "bold"))
        self.jarak_lbl.pack()
        self.cost_lbl = tk.Label(self.bobot_wd, text="(Cost)")
        self.cost_lbl.pack()
        self.jarak_inp = tk.Entry(self.bobot_wd, width=5, justify="center")
        self.jarak_inp.pack(pady=5)

        self.proses_btn = tk.Button(self.bobot_wd, text="Proses", command=self.show_all_table)
        self.proses_btn.pack(pady=30, ipadx=10, ipady=20)

    def get_bobot(self): #mendapatkan nilai bobot dari entry
        akred = int(self.akred_inp.get())
        fasil = int(self.fasilitas_inp.get())
        pres = int(self.prestasi_inp.get())
        biaya = int(self.biaya_inp.get())
        jarak = int(self.jarak_inp.get())
        self.weights = [akred, fasil, pres, biaya, jarak]
        return self.weights
        

    def show_all_table(self): #memunculkan window dan seluruh tabel
        self.get_bobot()
        # self.bobot_wd.destroy()

        self.table_window = tk.Toplevel(self.root)
        self.table_window.geometry('800x700')
        self.table_window.title("Table")

        self.frame = tk.Frame(self.table_window)
        self.frame.pack(fill='both', expand=True)

        self.canvas = tk.Canvas(self.frame)
        self.canvas.pack(side='left', fill='both', expand=True)

        self.scrollbar = tk.Scrollbar(self.frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.content_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.content_frame, anchor='nw')

        self.read_file()
        self.topsis_calculate()
        self.create_table(self.content_frame, "Tabel Nilai", self.df.columns, self.df.to_numpy().tolist())
        self.create_table(self.content_frame, "Tabel Normalisasi", self.df.columns, self.normal)
        self.create_table(self.content_frame, "Tabel Normalisasi Terbobot", self.df.columns, self.weighted)
        

    
        self.column_table_d = ['Sekolah', 'D+', 'D-']
        self.create_table_d(self.content_frame, "Tabel Jarak Ideal", self.column_table_d, self.d)

        self.column_table_V = ['Sekolah', 'V', 'Ranking']
        self.create_table_d(self.content_frame, "Tabel Hasil dan Ranking", self.column_table_V, self.v)

        self.content_frame.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

        def on_frame_configure(event): #untuk scrol pada bagian tabel
            self.canvas.config(scrollregion=self.canvas.bbox("all"))

        self.content_frame.bind("<Configure>", on_frame_configure)

        def on_mouse_wheel(event): 
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        self.canvas.bind_all("<MouseWheel>", on_mouse_wheel)  # Windows and Linux


def main():
    root = tk.Tk()
    app = Topsis_GUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
