import tkinter as tk
from tkinter import ttk
import json
from tkinter import messagebox

class SubWindow(tk.Toplevel):
    def __init__(self, window):
        super().__init__(window)
        self.protocol('WM_DELETE_WINDOW', window.close)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.calendario = []
        self.id = None
        self.frame1 = tk.Frame(self)
        self.frame2 = tk.Frame(self)

        self.lbl = tk.Label(self.frame1, text='salones/Laboratorios')
        self.lbl.grid(column=0, row=0)
        self.entry = tk.Entry(self.frame1)
        self.entry.grid(column=0, row=1)
        self.btn = ttk.Button(self.frame1,text="Agregar salon/Laboratorio", command=lambda: self.agregar_clase())
        self.btn.grid(column=1, row=1)

        self.listbox = tk.Listbox(self.frame1)
        self.listbox.grid(column=0, row=2, columnspan=3, sticky="nsew")

        self.btn = ttk.Button(self.frame1, text="Guardar en base de datos",command=lambda: self.guardar_BD())
        self.btn.grid(column=2, row=1)

        self.frame1.grid(column=0, row=0, padx=20, pady=20,sticky="nsew")

    def agregar_clase(self):
        self.listbox.insert(tk.END, self.entry.get())
        self.entry.delete(0, tk.END)

    def guardar_BD(self):
        try:
            self.data = {}
            try:
                with open('./BD.json') as file:
                    self.data = json.load(file)
            except:
                self.data['Salones']= {}
            if self.exist(self.data, self.grupo):
                MsgBox = tk.messagebox.askquestion (parent=self, title= 'Registro ya guardado',message='Estas seguro de sobrescribir la información guardada',icon = 'warning')
                if MsgBox == 'no':
                    return False
                else:
                    pass
            self.data['Salones'] = self.listbox.get(0, tk.END)
            
            with open('./BD.json', 'w') as file:
                json.dump(self.data, file, indent=4)
            messagebox.showinfo(parent=self,title ='Guardado con exito',message="Guardado con exito")
        except:
            messagebox.showerror(parent=self,title ='Error',message="No se pudo guardar")

    def exist(self, array, grupo):
        for x in array['Salones']:
            if x.upper() == grupo.upper():
                return True
        return False

            