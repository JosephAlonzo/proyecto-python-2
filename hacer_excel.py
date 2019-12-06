import tkinter as tk
from tkinter import ttk
import json
from tkinter import messagebox
from tkcalendar import Calendar, DateEntry
import datetime
import time
from datetime import date,timedelta
from openpyxl import load_workbook
import sys
import os
import operator

class SubWindow(tk.Toplevel):
    def __init__(self, window):
        super().__init__(window)
        self.protocol('WM_DELETE_WINDOW', window.close)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.materias = []
        self.calendario = []
        self.lista = []
        self.grupo = None
        self.data = []
        self.dias_sin_clases= []
        self.get_dias()
        self.listaMateriasNoAplicadas = []
        self.index = 0
        self.id = None
        self.frame1 = tk.Frame(self)
        self.frame2 = tk.Frame(self)

        self.lbl = tk.Label(self.frame1, text='Periodos 1')
        self.lbl.grid(column=0, row=0)
        self.entry = tk.Entry(self.frame1)
        self.entry.grid(column=0, row=1)
        self.entry.bind("<Button-1>", self.abrir_calendario)

        self.lbl = tk.Label(self.frame1, text='Periodos 2')
        self.lbl.grid(column=0, row=2)
        self.entry2 = tk.Entry(self.frame1)
        self.entry2.grid(column=0, row=3)
        self.entry2.bind("<Button-1>", self.abrir_calendario)

        self.lbl = tk.Label(self.frame1, text='Periodos 3')
        self.lbl.grid(column=0, row=4)
        self.entry3 = tk.Entry(self.frame1)
        self.entry3.grid(column=0, row=5)
        self.entry3.bind("<Button-1>", self.abrir_calendario)

        self.lbl = tk.Label(self.frame1, text='ordinario')
        self.lbl.grid(column=0, row=6)
        self.entry4 = tk.Entry(self.frame1)
        self.entry4.grid(column=0, row=7)
        self.entry4.bind("<Button-1>", self.abrir_calendario)

        self.lbl = tk.Label(self.frame1, text='extraordinario 1')
        self.lbl.grid(column=0, row=8)
        self.entry5 = tk.Entry(self.frame1)
        self.entry5.grid(column=0, row=9)
        self.entry5.bind("<Button-1>", self.abrir_calendario)

        self.lbl = tk.Label(self.frame1, text='extraordinario 2')
        self.lbl.grid(column=0, row=10)
        self.entry6 = tk.Entry(self.frame1)
        self.entry6.grid(column=0, row=11)
        self.entry6.bind("<Button-1>", self.abrir_calendario)


        self.btn = ttk.Button(self.frame1, text="Hacer excel", command=lambda: self.crear_excel())
        self.btn.grid(column=0, row=12)

       
        self.combo = ttk.Combobox(self, state="readonly")
        
        self.combo.bind('<<ComboboxSelected>>', self.modified)    
        self.combo.grid(column=0, row=0)
        self.agregar_valores_combo()

        self.frame1.grid(column=0, row=1, padx=20, pady=20,sticky="nsew")

        self.entry.config(state='disabled')
        self.entry2.config(state='disabled')
        self.entry3.config(state='disabled')
        self.entry4.config(state='disabled')
        self.entry5.config(state='disabled')
        self.entry6.config(state='disabled')

    
    def crear_excel(self):
        try:
            entry = self.entry.get().split('-')
            dia= date( int(entry[0]), int(entry[1]), int(entry[2]))
            dia1 = dia - datetime.timedelta(days=1)
            parcial1 = self.armarExcel(dia1)

            entry = self.entry2.get().split('-')
            dia= date( int(entry[0]), int(entry[1]), int(entry[2]))
            dia2 = dia - datetime.timedelta(days=1)
            parcial2 = self.armarExcel(dia2)

            entry = self.entry3.get().split('-')
            dia= date( int(entry[0]), int(entry[1]), int(entry[2]))
            dia3 = dia - datetime.timedelta(days=1)
            parcial3 = self.armarExcel(dia3)

            entry = self.entry4.get().split('-')
            dia= date( int(entry[0]), int(entry[1]), int(entry[2]))
            dia4 = dia - datetime.timedelta(days=1)
            parcial4 = self.armarExcel(dia4)

            entry = self.entry5.get().split('-')
            dia= date( int(entry[0]), int(entry[1]), int(entry[2]))
            dia5 = dia - datetime.timedelta(days=1)
            parcial5 = self.armarExcelExtras(dia5)

            entry = self.entry6.get().split('-')
            dia= date( int(entry[0]), int(entry[1]), int(entry[2]))
            dia6 = dia - datetime.timedelta(days=1)
            parcial6 = self.armarExcelExtras(dia6)

            archivo = load_workbook('template.xlsx')
            # leer y prepara para escribir
            datosTemplate = archivo.active
            
            # Celdas para llenar datos
            datosTemplate['B2'] = self.grupo #la columna de grado y grupo
            datosTemplate['B3'] = 'TIC-ITI Sep-Dic 2019' # Carrera y periodo sep-dic,etc..

            datosTemplate['B8'] = self.cambiar_formato(dia1 + datetime.timedelta(days=1))  #fecha primer parcial
            datosTemplate['C8'] = self.cambiar_formato(dia2 + datetime.timedelta(days=1))  #fecha segundo parcial
            datosTemplate['D8'] = self.cambiar_formato(dia3 + datetime.timedelta(days=1))  #fecha tercer parcial
            datosTemplate['E8'] = self.cambiar_formato(dia4 + datetime.timedelta(days=1))   #fecha ordinarios
            datosTemplate['F8'] = self.cambiar_formato(dia5 + datetime.timedelta(days=1))     #fecha primer extra
            datosTemplate['G8'] = self.cambiar_formato(dia6 + datetime.timedelta(days=1))   #fecha segundo extra
            c= 8
        
            for x in self.lista:
                if x[0].lower() != 'modulo libre':
                    datosTemplate['A'+ str(c+1)] = x[0]
                    c+=1

            celdas = ['B', 'C', 'D', 'E','F','G']

            c = 9
            for i in range(len(self.lista)):
                if self.lista[i][0].lower() == 'ingles':
                    texto = 'De acuerdo a la coordinación de Inglés'
                    datosTemplate[celdas[0]+str(c)] = texto
                    datosTemplate[celdas[1]+str(c)] = texto
                    datosTemplate[celdas[2]+str(c)] = texto
                    datosTemplate[celdas[3]+str(c)] = texto
                    datosTemplate[celdas[4]+str(c)] = texto
                    datosTemplate[celdas[5]+str(c)] = texto
                elif self.lista[i][0].lower() == 'modulo libre':
                    continue
                else:
                    for x in parcial1:
                        if self.lista[i][0] == x['clase']:
                            celda = celdas[0]+str(c)
                            datosTemplate[celda] =       self.cambiar_formato2(x['fecha'])+ '\n' + x['hora'] + '\n' +  x['lab'].lower()
                            break
                            
                    for x in parcial2:
                        if self.lista[i][0]  == x['clase']:
                            celda = celdas[1]+str(c)
                            datosTemplate[celda] =       self.cambiar_formato2(x['fecha'])+ '\n' + x['hora'] + '\n' +  x['lab'].lower()
                            break

                    for x in parcial3: 
                        if self.lista[i][0]  == x['clase']:
                            celda = celdas[2]+str(c)
                            datosTemplate[celda] =       self.cambiar_formato2(x['fecha'])+ '\n' + x['hora'] + '\n' +  x['lab'].lower()
                            break

                    for x in parcial4:
                        if self.lista[i][0]  == x['clase']:
                            celda = celdas[3]+str(c)
                            datosTemplate[celda] =       self.cambiar_formato2(x['fecha'])+ '\n' + x['hora'] + '\n' +  x['lab'].lower()
                            break

                    for x in parcial5:
                        if self.lista[i][0] == x['clase']:
                            celda = celdas[4]+str(c)
                            datosTemplate[celda] =      self.cambiar_formato2(x['fecha']) + '\n' + 'Pendiente horario y laboratorio'
                            break
                    for x in parcial6:
                        if self.lista[i][0] == x['clase']:
                            celda = celdas[5]+str(c)
                            datosTemplate[celda] =      self.cambiar_formato2(x['fecha']) + '\n' + 'Pendiente horario y laboratorio'
                            break
                c+=1
            if len(self.listaMateriasNoAplicadas) > 0:
                mensaje = "no se pudieron asignar las siguientes materias: "
                for clase in self.listaMateriasNoAplicadas:
                    mensaje += ('\n' + clase)
                messagebox.showwarning(parent=self, title='materrias no aplicadas', message=mensaje)

            messagebox.showinfo(parent=self, title='Se creo correctamente', message="El excel creo correctamente")
            # Formato final para guardar
            archivo.save("calendario.xlsx")

            path = os.path.dirname(sys.argv[0])
            if not path:
                path = '.'
            print(path)
            os.system('start excel.exe "%s\\calendario"' % (path, ))

        except:
            messagebox.showerror(parent=self, title='Error', message="Ocurrio un error")


    def cambiar_formato(self, date):
        months = ("Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic")
        day = date.day
        month = months[date.month - 1]
        formato = "{}-{}".format(day, month)

        return formato

    def cambiar_formato2(self, fecha):
        dias = ['Lunes','Martes','Miercoles','Jueves','Viernes']
        months = ("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
        f = fecha.weekday()
        f = dias[f]
        day = fecha.day
        month = months[fecha.month - 1]
        formato = "{} {} de {}".format(f, day, month)
        return formato

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
            self.entry2.config(state='normal')
            self.entry3.config(state='normal')
            self.entry4.config(state='normal')
            self.entry5.config(state='normal')
            self.entry6.config(state='normal')
            self.formato_calendario()
            self.lista = self.get_materias()
    
    def get_calendario(self):
        try:
            with open('./BD.json') as file:
                self.data = json.load(file)
                return self.data['Calendario'][self.grupo]
        except:
            pass
    
    def get_materias(self):
        try:
            with open('./BD.json') as file:
                self.data = json.load(file)
                return self.data['Clases'][self.grupo]
        except:
            pass

    def formato_calendario(self):
        tmp = self.get_calendario()
        x = int( len(tmp)/5)
        calendario =[
          ['' for i in range(5)] for j in range(x)  
        ] 
        
        for i in range(len(tmp)):
            for row in range(len(calendario)):
                for col in  range(len(calendario[row])):
                    if row == int(tmp[i]['row']) and col == int(tmp[i]['column']):
                        calendario[row][col] = tmp[i]
                        break

        self.calendario = calendario
        
        
    def get_dias(self):
        try:
            with open('./BD.json') as file:
                self.data = json.load(file)
                self.dias_sin_clases = self.data['Dias']
        except:
            pass

    def abrir_calendario(self,event):
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

    def validar_dia(self, dia2):
        if dia2.weekday() >= 5:
            return True
        try:
            for dia in self.dias_sin_clases:
                x = str(dia).split('-')
                comparacion = date( int(x[0]), int(x[1]), int(x[2]))
                if dia2 == comparacion:
                    return True
            return False
        except:
            return False
    
    def armarExcel(self, dia):
        excel = []
        datos = self.calcular_fecha_inicio(dia)
        dia_inicio = datos[0] 
        semana = 0 
        f = dia_inicio.weekday()
        self.lista = self.getLisOrderedByPriority(self.calendario, dia)
        ultimaMateria = ''
        count = 0
        while len(self.lista) > 0:
            if ultimaMateria == self.lista[0][0]:
                if self.lista[0][1] == 0 and count == 3:
                    self.listaMateriasNoAplicadas.append(
                        self.lista[0][0]
                    )
                    self.lista.pop(0)
                    count = 0
                elif count > 0:
                    newLista = []
                    for x in range(len(self.lista)):
                        if x == 0:
                            newLista.append((self.lista[x][0], 0))
                        else:
                            newLista.append((self.lista[x][0], self.lista[x][1]))
                    self.lista = newLista
                    count += 1
                else:
                    count = 1
            if len(self.lista) > 0:
                ultimaMateria = self.lista[0][0]
            for col in range(len(self.calendario[0])):
                if len(self.lista) < 1:
                    break
                if self.lista[0][0].lower() == 'ingles' or self.lista[0][0].lower() == 'tutorias' or self.lista[0][0].lower() == 'modulo libre':
                    self.lista.pop(0)
                if col < f and semana == 0:
                    continue
                if self.validar_dia(dia_inicio):
                    dia_inicio += datetime.timedelta(days=1)
                    if dia_inicio > dia:
                        break
                    continue
                if self.validar_limite_de_examenes(excel, dia):
                    dia_inicio += datetime.timedelta(days=1)
                    if dia_inicio > dia:
                        break
                    continue
                try:
                    clase = self.lista[0][0]
                except:
                    break
                for row in range(len(self.calendario[col])):
                    c = self.calendario[row][col]['clase']
                    if clase == c:
                        hora = self.calendario[row][col]['horario'].split('-')
                        excel.append({
                            'clase': self.calendario[row][col]['clase'],
                            'lab': self.calendario[row][col]['lab'],
                            'fecha': dia_inicio,
                            'hora': hora[0]
                        })
                        self.lista.pop(0)
                        break
                dia_inicio += datetime.timedelta(days=1)
                if dia_inicio > dia:
                    break
            semana += 1
            dia_inicio += datetime.timedelta(days=2)
            if semana > datos[1]:
                semana = 0
                dia_inicio = datos[0]
                
        self.lista = self.getLisOrderedByPriority(self.calendario, dia)
        return excel
            

    def armarExcelExtras(self, dia):
        excel = []
        count = 0
        fecha_inicio = self.calcular_fecha_inicio_extras(dia)
        
        for x in range(len(self.lista) ):
            excel.append({
                'clase': self.lista[x][0],
                'fecha': fecha_inicio[0]
            })

        return excel
            

    def calcular_fecha_inicio(self, fecha):
        fecha_limite = fecha
        fecha_inicio = fecha
        i=0
        count = 0
        while count < 5:
            dia = fecha_limite - datetime.timedelta(days=i)
            if self.validar_dia(dia) == False:
                count += 1
            if count < 5:
                fecha_inicio -= datetime.timedelta(days=1)
            i+=1
        
        f1 = fecha_limite.isocalendar()[1]
        f2 = fecha_inicio.isocalendar()[1] 
        numero = f1 - f2 
        retorno = [fecha_inicio,numero]
        
        return retorno
    
    def validar_limite_de_examenes(self, array, dia):
        count = 0 
        for x in array:
            if x['fecha'] == dia:
                count +=1
            if count >= 2:
                return True
        return False


    def calcular_fecha_inicio_extras(self, fecha):
        fecha_limite = fecha
        fecha_inicio = fecha
        i=0
        count = 0
        while count < 1:
            dia = fecha_limite - datetime.timedelta(days=i)
            if self.validar_dia(dia) == False:
                count += 1
            if count < 1:
                fecha_inicio -= datetime.timedelta(days=1)
            i+=1
        
        f1 = fecha_limite.isocalendar()[1]
        f2 = fecha_inicio.isocalendar()[1] 
        numero = f1 - f2 
        retorno = [fecha_inicio,numero]
        
        return retorno

    def getLisOrderedByPriority(self, array, dia):
        listaMaterias = []
        listaMaterias2 = []
        datos = self.calcular_fecha_inicio(dia)
        dia_inicio = datos[0]
        f = dia_inicio.weekday()
        for semana in range(datos[1] + 1):
            for col in range(len(array[0])):
                if col < f and semana == 0:
                    continue
                for row in range(len(array)):
                    if array[row][col] != '':
                        materia = array[row][col]['clase']
                        if materia[0:8].lower() != 'tutorias':
                            try:
                                if array[row+1][col] != '':
                                    nextMateria = array[row+1][col]['clase']
                                else:
                                    nextMateria = ''
                            except:
                                nextMateria = ''
                            if materia == nextMateria:
                                listaMaterias.append( materia )
                                if len(array) < row+1:
                                    row+=1
                            if listaMaterias2.count(materia) < 1:
                                listaMaterias2.append(materia)

        for item in range(len(listaMaterias2)):
            listaMaterias2[item] = ( listaMaterias2[item], listaMaterias.count(listaMaterias2[item]) )
             
        listaMaterias2.sort(key = operator.itemgetter(1))
        
        return listaMaterias2

    
