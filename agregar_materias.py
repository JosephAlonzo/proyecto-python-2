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
        self.rowconfigure(1, weight=1)
        self.materias = []
        self.grupo = None
        self.data = []
        self.index = 0
        self.id = None
        self.frame1 = tk.Frame(self)
        self.frame2 = tk.Frame(self)

        self.lbl = tk.Label(self.frame1, text='Materia')
        self.lbl.grid(column=0, row=0)
        self.entry = tk.Entry(self.frame1, state="readonly")
        self.entry.grid(column=0, row=1)

        self.btn = ttk.Button(self.frame1, text="Agregar",command=lambda: self.agregar())
        self.btn.grid(column=0, row=2)

        self.crear_tabla()
        self.actualizar_tabla()

        self.combo = ttk.Combobox(self, state="readonly")
        
        self.combo.bind('<<ComboboxSelected>>', self.modified)    
        self.combo.grid(column=0, row=0)
        self.agregar_valores_combo()

        self.frame1.grid(column=0, row=1, padx=20, pady=20,sticky="nsew")
        self.frame2.grid(column=0, row=2, padx=20, pady=20,sticky="nsew")

    def agregar(self):
        materia = self.entry.get()
        self.materias.append(materia.upper())
        self.guardar_BD('Guardar')
        self.limpiar_entries()
        self.actualizar_tabla()
    
    def guardar_BD(self, accion, value=None):
        try:
            self.data = {}
            try:
                with open('./BD.json') as file:
                    self.data = json.load(file)
                    self.data['Clases']
            except:
                self.data['Clases'] = {}
                self.data['Clases'][self.grupo]= []

            self.data['Clases'][self.grupo] = self.materias
            
            with open('./BD.json', 'w') as file:
                json.dump(self.data, file, indent=4)
            if accion == 'Guardar':
                messagebox.showinfo(parent=self,title ='Guardado con exito',message="Guardado con exito")
            elif accion == 'Editar':
                messagebox.showinfo(parent=self,title ='Editado con exito',message="Editado con exito")
                self.btn.config(
                    text='Agregar',
                    command= lambda: self.agregar()
                    )
            elif accion == 'Eliminar':
                messagebox.showinfo(parent=self,title ='Eliminado con exito',message="Eliminado con exito")
        except:
            messagebox.showerror(parent=self,title ='Error',message="No se pudo completar la accion")

    def exist(self, array, grupo):
        try:
            for x in array['Grupos']:
                if x.upper() == grupo.upper():
                    return True
            return False
        except:
            return False

    def select_item(self, event):
        obj = self.tabla.item(self.tabla.focus())
        col = self.tabla.identify_column(event.x)
        print(col)
        if col == '#2':
            self.index = int(obj['text'])
            grupo = obj['values'][0].split('-')
            self.entry.delete(0, tk.END)
            self.entry.insert(0, grupo[0])
            self.btn.config(
                text='Editar',
                command= lambda: self.editar()
            )
        elif col == '#3':
            self.index = int(obj['text'])
            answer = tk.messagebox.askquestion(parent= self, message="¿desea eliminar el registro?", title="Eliminar registro",icon = 'warning')
            if answer == 'yes':
                self.eliminar()

    def eliminar(self):
        self.materias.pop(self.index)
        self.guardar_BD('Eliminar')
        self.actualizar_tabla()

    def editar(self):
        materia = self.entry.get()
        self.materias[self.index] = materia.upper()
        self.guardar_BD('Editar')
        self.limpiar_entries()
        self.actualizar_tabla()

    def actualizar_tabla(self):
        try:
            array = self.get_materias()
            self.materias = array
            self.tabla.delete(*self.tabla.get_children())
            i = 0
            for dato in array:
                self.tabla.insert("", 'end',text=i, values=( dato, 'Editar', 'Eliminar' ) )
                i+=1
        except:
            self.materias = []
            pass

    def limpiar_entries(self):
        self.entry.delete(0, tk.END)

    def crear_tabla(self):
        self.tabla = ttk.Treeview(self.frame2, style="mystyle.Treeview")
        self.tabla.grid(column = 0, row=3)

        self.tabla["columns"] = ["materia", "editar", "eliminar"]
        self.tabla["show"] = "headings"

        self.tabla.column("materia",minwidth=0,width=100,anchor=tk.CENTER)
        self.tabla.column("editar",minwidth=0,width=100,anchor=tk.CENTER)
        self.tabla.column("eliminar",minwidth=0,width=100,anchor=tk.CENTER)

        self.tabla.heading("materia", text="Materia",anchor=tk.CENTER)
        self.tabla.heading("editar", text="Editar",anchor=tk.CENTER)
        self.tabla.heading("eliminar", text="Eliminar",anchor=tk.CENTER)

        self.tabla.grid(column=0, row=0,sticky="nsew")
        self.tabla.bind('<ButtonRelease-1>',self.select_item)

    def get_materias(self):
        try:
            with open('./BD.json') as file:
                self.data = json.load(file)
                return self.data['Clases'][self.grupo]
        except:
            pass

    def agregar_valores_combo(self):
        try:
            with open('./BD.json') as file:
                self.data = json.load(file)
                #return self.data['Clases']
                self.combo["values"] = ["Seleccionar un grupo"]
                keys = []
                for key in self.data['Grupos']:
                    keys.append(
                        key.upper()
                    )
                self.data = self.data['Grupos']
                self.combo["values"] = keys
        except:
            pass        

    def modified (self, event) :
        if self.combo.get() != None and self.combo.get() != 'Seleccionar un grupo' and self.combo.get() != "":
            self.grupo = self.combo.get() 
            self.entry.config(state='normal')
            self.actualizar_tabla()
