import tkinter as tk
from tkinter import ttk
import json
from tkinter import messagebox
from tkcalendar import Calendar, DateEntry
import datetime
import time
from datetime import date,timedelta

class SubWindow(tk.Toplevel):
    def __init__(self, window):
        super().__init__(window)
        self.protocol('WM_DELETE_WINDOW', window.close)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.dias = []
        self.data = []
        self.index = 0
        self.frame1 = tk.Frame(self)
        self.frame2 = tk.Frame(self)

        self.lbl = tk.Label(self.frame1, text='Días inhábiles')
        self.lbl.grid(column=0, row=0)
        self.entry = tk.Entry(self.frame1)
        self.entry.grid(column=0, row=1)
        self.entry.bind("<Button-1>", self.tkinterCalendar)

        self.btn = ttk.Button(self.frame1, text="Agregar",command=lambda: self.agregar())
        self.btn.grid(column=0, row=2)

        self.crear_tabla()
        self.actualizar_tabla()

        self.frame1.grid(column=0, row=0, padx=20, pady=20,sticky="nsew")
        self.frame2.grid(column=0, row=1, padx=20, pady=20,sticky="nsew")


    def agregar(self):
        fecha = self.entry.get()
        self.dias.append(fecha)
        self.guardar_BD('Guardar')
        self.limpiar_entries()
        self.actualizar_tabla()
    
    def guardar_BD(self, accion):
        try:
            self.data = {}
            try:
                with open('./BD.json') as file:
                    self.data = json.load(file)
                    self.data['Dias']
            except:
                self.data['Dias']= []
            self.data['Dias'] = self.dias
            
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
            self.entry.delete(0, tk.END)
            self.entry.insert(0, obj['values'][0])
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
        self.dias.pop(self.index)
        self.guardar_BD('Eliminar')
        self.actualizar_tabla()
    def editar(self):
        salon = self.entry.get() 
        self.dias[self.index] = salon
        self.guardar_BD('Editar')
        self.limpiar_entries()
        self.actualizar_tabla()

    def actualizar_tabla(self):
        try:
            array = self.get_salones()
            self.dias = array
            self.tabla.delete(*self.tabla.get_children())
            i = 0
            for dato in array:
                self.tabla.insert("", 'end',text=i, values=( dato, 'Editar', 'Eliminar' ) )
                i+=1
        except:
            self.dias = []
            pass

    def limpiar_entries(self):
        self.entry.delete(0, tk.END)

    def crear_tabla(self):
        self.tabla = ttk.Treeview(self.frame2, style="mystyle.Treeview")
        self.tabla.grid(column = 0, row=3)

        self.tabla["columns"] = ["salon", "editar", "eliminar"]
        self.tabla["show"] = "headings"

        self.tabla.column("salon",minwidth=0,width=100,anchor=tk.CENTER)
        self.tabla.column("editar",minwidth=0,width=100,anchor=tk.CENTER)
        self.tabla.column("eliminar",minwidth=0,width=100,anchor=tk.CENTER)

        self.tabla.heading("salon", text="Días inhábiles",anchor=tk.CENTER)
        self.tabla.heading("editar", text="Editar",anchor=tk.CENTER)
        self.tabla.heading("eliminar", text="Eliminar",anchor=tk.CENTER)

        self.tabla.grid(column=0, row=0,sticky="nsew")
        self.tabla.bind('<ButtonRelease-1>',self.select_item)

    def get_salones(self):
        try:
            with open('./BD.json') as file:
                self.data = json.load(file)
                return self.data['Dias']
        except:
            pass

    def tkinterCalendar(self,event):
        def print_sel():
            f = (cal.selection_get()).weekday()
            if f < 5:
                print(cal.selection_get())
                event.widget.delete(0, tk.END)
                event.widget.insert(0, cal.selection_get())
                top.destroy()
            else:
                tk.messagebox.showerror(parent=top, message="Por favor elija un dia valido", title="Seleccionar fecha")
    
        top = tk.Toplevel(self)
        now = datetime.datetime.now()
        day =now.day
        month =now.month
        year =now.year

        cal = Calendar(top,
                    font="Arial 14", selectmode='day',
                    cursor="hand1", year=year, month=month, day=day)
        cal.pack(fill="both", expand=True)
        ttk.Button(top, text="ok", command=print_sel).pack()