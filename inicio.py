import tkinter as tk
from tkinter import ttk
import json
from tkinter import messagebox
import agregar_salones as third_window
import agregar_grupo as four_window
import agregar_dias_inhabiles as five_window
import agregar_materias as six_window
import hacer_excel as seven_window

class Application(ttk.Frame):
    def __init__(self, window):
        super().__init__(window)
        self._second_window = None
        self.calendario = []
        self.index = None
        self.grupo = None
        self.data = []
        self.horarios = []
        self.frame1 = tk.Frame(window)
        self.frame2 = tk.Frame(window)

        #################### Creacion de las opciones del menu Inicio ##########################
        menubar = tk.Menu(window)
        window.config(menu=menubar)
        menu = tk.Menu(menubar, tearoff=0)
        menu.add_command(label="Grupos", command=lambda: self.new_window3())
        menu.add_command(label="Materias", command=lambda: self.new_window5())
        menu.add_command(label="Salones/Laboratorios", command=lambda: self.new_window2())
        menu.add_command(label="Días inhábiles", command=lambda: self.new_window4())
        menu.add_command(label="Hacer excel", command=lambda: self.new_window6())

        menubar.add_cascade(label="Acciones", menu=menu)
        #################### Creacion de las opciones del menu Fin ##########################


        window.columnconfigure(0, weight=1)
        window.rowconfigure(1, weight=2)
        window.rowconfigure(2, weight=10)

        #################### Crea la tabla ##########################
        self.iniciarInterfaz()

        self.btn = ttk.Button(self.frame1, text="Agregar", command= lambda: self.agregar() ) 
        self.btn.grid(column=5,row=4,sticky="nsew", pady=10)
        self.btn.config(
                    state='disabled'
                )

        self.frame1.grid(column=1, row=1, padx=20, pady=20,sticky="nsew")
        self.frame2.grid(column=0, row=2, padx=20, pady=20,sticky="nsew",columnspan=2)

        self.grid(sticky="nsew")

    #################### Evento cuando se cambia la seleccion en el combobox  ##########################
    def modified (self, event) :
        if self.combo.get() != None and self.combo.get() != 'Seleccionar un grupo' and self.combo.get() != "":
            self.grupo = self.combo.get() 
            grado = self.combo.get().split('-')
            grado = grado[0]
            fallo = False
            #################### Activar combosbox si hay un grupo elegido  ##########################
            try:
                for i in range(len(self.days)):
                    if i > 0:
                        self.entriesClases[i].config(state='readonly')
                        self.entriesClases[i]["values"] = self.data['Clases'][self.grupo]
                    else:
                        self.entriesClases[i].config(state='readonly')
                        self.entriesClases[i]["values"] = self.get_lista_horarios(grado)
            except:
                fallo = True
                messagebox.showinfo(parent=self,title ='Sin materias asignadas',message="Revise la sección: \nAcciones>Materias \nPara añadir nuevas materias")
            try:
                for i in range(len(self.days)):
                    if i > 0:
                        self.entriesSalones[i].config(state='readonly')
                        self.entriesSalones[i]["values"] = self.data['Salones']
            except:
                messagebox.showinfo(parent=self,title ='Sin salones/laboratorios asignadas',message="Revise la sección: \nAcciones>Salones/Laboratorios \nPara añadir nuevas materias")
            self.actualizar_tabla()
            #################### Si no hay grupo seleccionado o el grupo no tiene materias  ##########################
            if fallo :
                self.btn.config(
                    state='disabled'
                )
                for i in range(len(self.days)):
                    if i > 0:
                        self.entriesClases[i].config(state='disabled')
                        self.entriesSalones[i].config(state='disabled')
                    else:
                        self.entriesClases[i].config(state='disabled')
            else:
                self.btn.config(
                    state='normal'
                )
    
    def select_item(self, event):
        obj = self.tabla.item(self.tabla.focus())
        col = self.tabla.identify_column(event.x)
        print(col)
        #################### Editar ##########################
        if col == '#7':
            self.index = int(obj['text'])
            for i in range(len(self.entriesClases)):
                if i > 0:
                    info = obj['values'][i].split('|')
                    self.entriesClases[i].config(
                        state='normal'
                    )
                    self.entriesSalones[i].config(
                        state='normal'
                    )
                    self.entriesClases[i].delete(0, tk.END)
                    self.entriesClases[i].insert(0, info[0] )
                    self.entriesSalones[i].delete(0, tk.END)
                    self.entriesSalones[i].insert(0, info[1])

                    self.entriesClases[i].config(
                        state='readonly'
                    )
                    self.entriesSalones[i].config(
                        state='readonly'
                    )
                else:
                    self.entriesClases[i].config(
                        state='normal'
                    )
                    self.entriesClases[i].delete(0, tk.END)
                    self.entriesClases[i].insert(0, obj['values'][i])
                    self.entriesClases[i].config(
                        state='readonly'
                    )

                
            self.btn.config(
                text='Editar',
                command= lambda: self.editar()
                )
        elif col == '#8':
            answer = tk.messagebox.askquestion(message="¿desea eliminar el registro?", title="Eliminar registro",icon = 'warning')
            if answer == 'yes':
                self.index = int(obj['text'])
                count = self.index * 5
                for i in range(5):
                    self.calendario.pop(count)
                self.guardar_BD('Eliminar')
                self.actualizar_tabla()
    
    def agregar(self):
        informacion = []
        dias = ['Lu', 'Ma', 'Mi', 'Ju', 'Vi']
        for i in range(len(self.entriesClases)):
            valor = self.entriesClases[i].get()
            if i == 0:
                horario = valor
                row = self.getRow(horario)
                continue
            if self.entriesClases[i].get() != '':
                informacion= {
                        'clase': self.entriesClases[i].get(),
                        'lab': self.entriesSalones[i].get(),
                        'horario': horario,
                        'dia' : dias[i-1],
                        'row': row,
                        'column':i-1
                    }
            else:
                informacion= {
                        'clase': 'Modulo libre',
                        'lab': '',
                        'horario': horario,
                        'dia' : dias[i-1],
                        'row': row,
                        'column':i-1
                    }
            self.calendario.append(informacion)
        
        self.guardar_BD('Guardar')
        self.limpiar_entries()
        self.actualizar_tabla()

    def getRow(self, horario):
        i = 0
        for x in self.horarios:
            if x == horario:
                return i
            i+=1
    def editar(self):
        informacion = []
        count = self.index * 5
        dias = ['Lu', 'Ma', 'Mi', 'Ju', 'Vi']
        for i in range(len(self.entriesClases)):
            if i == 0:
                horario = self.entriesClases[i].get()
                row = self.getRow(horario)
                continue
            informacion = {
                    'clase': self.entriesClases[i].get(),
                    'lab': self.entriesSalones[i].get(),
                    'horario': horario,
                    'dia' : dias[i-1],
                    'row': row,
                    'column':i-1
                }
            self.calendario[count] = informacion
            count += 1
        self.guardar_BD('Editar')
        self.limpiar_entries()
        self.actualizar_tabla()

    def limpiar_entries(self):
        for i in range(len(self.entriesClases)):
            self.entriesClases[i].config(
                state = 'normal'
            )
            self.entriesClases[i].delete(0, tk.END)
            self.entriesClases[i].config(
                state = 'readonly'
            )
            if i > 0:
                self.entriesSalones[i].config(
                    state = 'normal'
                )
                self.entriesSalones[i].delete(0, tk.END)
                self.entriesSalones[i].config(
                state = 'readonly'
                )
            

    def actualizar_tabla(self):
        try:
            array = self.get_calendario()
            self.calendario = array
            self.tabla.delete(*self.tabla.get_children())
            i = 0
            t = int(len(array)/5)
            datos = [ [] for i in range(t)]
            horarios = [ [] for i in range(t)]
            n = 0
            for i in range(len(array)):
                if i % 5 == 0 and i>0:
                    n+=1
                value = array[i]['clase']+'|'+array[i]['lab']
                horario = array[i]['horario']
                datos[n].append(
                    value
                )
                horarios[n] = horario
                
            for i in range(t):
                self.tabla.insert("", 'end',text=i, values=(horarios[i], datos[i][0], datos[i][1], datos[i][2], datos[i][3], datos[i][4], 'Editar', 'Eliminar' ) )
        except:
            self.calendario = []
            pass

    def guardar_BD(self, accion):
        try:
            self.data = {}
            try:
                with open('./BD.json') as file:
                    self.data = json.load(file)
                    self.data['Calendario']
            except:
                self.data['Calendario']= {}
                self.data['Calendario'][self.grupo] = []
            self.data['Calendario'][self.grupo] = self.calendario

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
            messagebox.showerror(title ='Error',message="No se pudo completar la accion")

    def get_calendario(self):
        try:
            with open('./BD.json') as file:
                self.data = json.load(file)
                return self.data['Calendario'][self.grupo]
        except:
            pass
    
    def new_window2(self):
        if self._second_window is not None:
            return
        self._second_window = third_window.SubWindow(self)

    def new_window3(self):
        if self._second_window is not None:
            return
        self._second_window = four_window.SubWindow(self)

    def new_window4(self):
        if self._second_window is not None:
            return
        self._second_window = five_window.SubWindow(self)

    def new_window5(self):
        if self._second_window is not None:
            return
        self._second_window = six_window.SubWindow(self)

    def new_window6(self):
        if self._second_window is not None:
            return
        self._second_window = seven_window.SubWindow(self)


    def close(self):
        if self._second_window is not None:
            self._second_window.destroy()
            self._second_window = None
            self.agregar_valores_combo()
            
    def agregar_valores_combo(self):
        try:
            with open('./BD.json') as file:
                self.data = json.load(file)
                #return self.data['Clases']
                self.combo["values"] = ["Seleccionar un grupo"]
                self.combo["values"] = self.data['Grupos']
        except:
            pass
    
    def get_lista_horarios(self, grado):
        self.horarios = []
        if int(grado) < 7:
            self.horarios = [
                '7:00-7:50',
                '7:50-8:40',
                '9:00-9:50',
                '9:50-10:40',
                '11:00-11:50',
                '11:00-11:50',
                '11:50-12:40',
                '11:50-12:40',
                '12:40-13:30',
                '13:30-14:20'
            ]
        else:
            self.horarios = [
                '16:00-16:50',
                '16:50-17:40',
                '17:40-16:30',
                '16:30-19:20',
                '19:20-20:10',
                '20:10-21:00'
            ]
        return self.horarios

    def iniciarInterfaz(self):
        self.tabla = ttk.Treeview(self.frame2, style="mystyle.Treeview")
        self.tabla.grid(column = 0, row=0)

        self.days = ["Horario", "lunes","martes","miercoles","jueves","viernes"]
        self.tabla["columns"] = ["horario","lunes","martes","miercoles","jueves","viernes", "editar", "eliminar"]
        self.tabla["show"] = "headings"

        self.tabla.column("horario",minwidth=0,width=100,anchor=tk.CENTER)
        self.tabla.column("lunes",minwidth=0,width=100,anchor=tk.CENTER)
        self.tabla.column("martes",minwidth=0,width=100,anchor=tk.CENTER)
        self.tabla.column("miercoles",minwidth=0,width=100,anchor=tk.CENTER)
        self.tabla.column("jueves",minwidth=0,width=100,anchor=tk.CENTER)
        self.tabla.column("viernes",minwidth=0,width=100,anchor=tk.CENTER)
        self.tabla.column("editar",minwidth=0,width=100,anchor=tk.CENTER)
        self.tabla.column("eliminar",minwidth=0,width=100,anchor=tk.CENTER)

        self.tabla.heading("horario", text="Horario",anchor=tk.CENTER)
        self.tabla.heading("lunes", text="Lunes",anchor=tk.CENTER)
        self.tabla.heading("martes", text="Martes",anchor=tk.CENTER)
        self.tabla.heading("miercoles", text="Miercoles",anchor=tk.CENTER)
        self.tabla.heading("jueves", text="Jueves",anchor=tk.CENTER)
        self.tabla.heading("viernes", text="Viernes",anchor=tk.CENTER)
        self.tabla.heading("editar", text="Editar",anchor=tk.CENTER)
        self.tabla.heading("eliminar", text="Eliminar",anchor=tk.CENTER)

        self.tabla.grid(column=0, row=0,sticky="nsew")
        self.tabla.bind('<ButtonRelease-1>',self.select_item)

        self.entriesClases = ['' for x in range(len(self.days)) ]
        self.entriesSalones = ['' for x in range(len(self.days)) ]
        for i in range(len(self.days)):
            self.lbl = tk.Label(self.frame1, text=self.days[i])
            self.lbl.grid(column=i, row=0)
            self.entriesClases[i] = ttk.Combobox(self.frame1)
            self.entriesClases[i].grid(column=i, row=1)
            self.entriesClases[i].config(state='disabled')
            if i > 0:
                self.lbl2 = tk.Label(self.frame1, text="Laboratorio/Salon")
                self.lbl2.grid(column=i, row=2)
                self.entriesSalones[i] = ttk.Combobox(self.frame1)
                self.entriesSalones[i].grid(column=i, row=3)
                self.entriesSalones[i].config(state='disabled')

        self.combo = ttk.Combobox(window, state="readonly")
        
        # self.combo["values"]  = ['Seleccionar un grupo']
        self.combo.bind('<<ComboboxSelected>>', self.modified)    
        self.combo.grid(column=1, row=0)
        self.agregar_valores_combo()
        


window = tk.Tk()
app = Application(window)
app.mainloop()
