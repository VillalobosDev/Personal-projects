from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import tkinter
import random
import pymysql
import csv
from datetime import datetime
import numpy as np
import ctypes

window=tkinter.Tk()
window.title("Catastro | CRUD")
window.geometry("1080x640")
my_tree=ttk.Treeview(window,show='headings',height=20)
style=ttk.Style()

myappid = 'mycompany.myproduct.subproduct.version' # arbitrary string
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
window.iconbitmap(r'C:\Users\Nelson\Desktop\Serv COM & Proyecto\Code\CatastroCRUD\img\img.ico')




placeholderArray=['','','','','','','','']
numeric='1234567890'
alpha='ABCDEFGHIJKLMNOPQRSTUVWXYZ'

def connection():
    conn=pymysql.connect(
        host='localhost',
        user='root',
        password='',
        db='regdb'
    )
    return conn

conn=connection()
cursor=conn.cursor()

for i in range(0,9):
    if i < len(placeholderArray):
        placeholderArray[i]=tkinter.StringVar()
    else:
        placeholderArray.append(tkinter.StringVar())

def read(): 
    cursor.connection.ping()
    sql="SELECT `cedula`, `contribuyente`, `nombreinmueble`, `rif`, `sector`, `uso`, `codcatastral`, `fechaliquidacion`  FROM reg ORDER BY `fechaliquidacion` DESC"
    cursor.execute(sql)
    results=cursor.fetchall()
    conn.commit()
    conn.close()
    return results

def refreshTable(): 
    # Clear the existing items in the tree
    for data in my_tree.get_children():
        my_tree.delete(data)
        
    # Insert new items
    for array in read():
        # Check if the iid already exists
        if not my_tree.exists(array):
            my_tree.insert(parent='', index='end', iid=array, text="", values=(array), tag="orow")
    
    my_tree.tag_configure('orow', background="#EEEEEE")
    my_tree.pack()

def setph(word,num):
    for ph in range(0,9):
        if ph == num:
            placeholderArray[ph].set(word)

def save():
    cedula              =   str(cedulaEntry.get())
    contribuyente       =   str(contribuyenteEntry.get())
    nombreinmueble      =   str(nombreinmuebleEntry.get())
    rif                 =   str(rifEntry.get())
    sector              =   str(sectorEntry.get())
    uso                 =   str(usoEntry.get())
    codcatastral        =   str(codcatastralEntry.get())
    fechaliquilacion    =   str(fechaliquidacionEntry.get())
    
    if not(cedula and cedula.strip()) or not(contribuyente and contribuyente.strip()) or not(nombreinmueble and nombreinmueble.strip()) or not(rif and rif.strip()) or not(sector and sector.strip()) or not(uso and uso.strip()) or not(codcatastral and codcatastral.strip()) or not(fechaliquilacion and fechaliquilacion.strip()):
        messagebox.showwarning("","Llena todos los formularios")
        return
    try:
        cursor.connection.ping()
        sql = f"INSERT INTO reg (`cedula`, `contribuyente`, `nombreinmueble`, `rif`, `sector`,`uso`,`codcatastral`,`fechaliquidacion`) VALUES ('{cedula}','{contribuyente}','{nombreinmueble}','{rif}','{sector}','{uso}','{codcatastral}','{fechaliquilacion}')"
        cursor.execute(sql)
        conn.commit()
        conn.close()
        for num in range(0,9):
            setph('',(num))
    except Exception as e:
        print(e)
        messagebox.showwarning("","Se produjo un error: "+str(e))
        return
    refreshTable()

def update():
    selectedItemId = ''
    
    # Check if an item is selected
    if not my_tree.selection():
        messagebox.showwarning("", "Por favor selecciona un elemento para actualizar.")
        return
    
    try:
        selectedItem = my_tree.selection()[0]
        item_values = my_tree.item(selectedItem)['values']
        
        if not item_values or len(item_values) < 8:  # Check if item has the expected number of values
            messagebox.showwarning("", "El elemento seleccionado no tiene valores completos.")
            return
        
        selectedItemId = str(item_values[0])
    except Exception as err:
        messagebox.showwarning("", "Un error se produjo: " + str(err))
        return
    
    cedula              =   str(cedulaEntry.get())
    contribuyente       =   str(contribuyenteEntry.get())
    nombreinmueble      =   str(nombreinmuebleEntry.get())
    rif                 =   str(rifEntry.get())
    sector              =   str(sectorEntry.get())
    uso                 =   str(usoEntry.get())
    codcatastral        =   str(codcatastralEntry.get())
    fechaliquidacion    =   str(fechaliquidacionEntry.get())
    
    try:
        cursor.connection.ping()
        sql = """UPDATE reg 
                 SET `cedula` = %s, `contribuyente` = %s, `nombreinmueble` = %s, 
                     `rif` = %s, `sector` = %s, `uso` = %s, `codcatastral` = %s, 
                     `fechaliquidacion` = %s 
                 WHERE `cedula` = %s"""
        
        cursor.execute(sql, (cedula, contribuyente, nombreinmueble, rif, sector, uso, codcatastral, fechaliquidacion, selectedItemId))
        conn.commit()
        conn.close()
        
        for num in range(0,9):
            setph('', num)
        
        refreshTable()
        
    except Exception as err:
        messagebox.showwarning("", "Un error se produjo: " + str(err))
        return

# def update():
#     selectedItemId = ''
#     try:
#         selectedItem = my_tree.selection()[0]
#         item_values = my_tree.item(selectedItem)['values']
        
#         if not item_values or len(item_values) < 8:  # Check if item has the expected number of values
#             messagebox.showwarning("", "El elemento seleccionado no tiene valores completos.")
#             return
        
#         selectedItemId = str(item_values[0])
#     except Exception as err:
#         messagebox.showwarning("", "Un error se produjo: " + str(err))
#         return
    
#     cedula              =   str(cedulaEntry.get())
#     contribuyente       =   str(contribuyenteEntry.get())
#     nombreinmueble      =   str(nombreinmuebleEntry.get())
#     rif                 =   str(rifEntry.get())
#     sector              =   str(sectorEntry.get())
#     uso                 =   str(usoEntry.get())
#     codcatastral        =   str(codcatastralEntry.get())
#     fechaliquidacion    =   str(fechaliquidacionEntry.get())
    
#     try:
#         cursor.connection.ping()
#         sql = """UPDATE reg 
#                  SET `cedula` = %s, `contribuyente` = %s, `nombreinmueble` = %s, 
#                      `rif` = %s, `sector` = %s, `uso` = %s, `codcatastral` = %s, 
#                      `fechaliquidacion` = %s 
#                  WHERE `cedula` = %s"""
        
#         cursor.execute(sql, (cedula, contribuyente, nombreinmueble, rif, sector, uso, codcatastral, fechaliquidacion, selectedItemId))
#         conn.commit()
#         conn.close()
        
#         for num in range(0,9):
#             setph('', num)
        
#         refreshTable()
        
#     except Exception as err:
#         messagebox.showwarning("", "Un error se produjo: " + str(err))
#         return

# def update():
#     selectedItemId = ''
#     try:
#         selectedItem = my_tree.selection()[0]
#         selectedItemId = str(my_tree.item(selectedItem)['values'][0])
#     except Exception as err:
#         messagebox.showwarning(f"", "Un error se produjo: " + str(err))
#         return
    
#     cedula              =   str(cedulaEntry.get())
#     contribuyente       =   str(contribuyenteEntry.get())
#     nombreinmueble      =   str(nombreinmuebleEntry.get())
#     rif                 =   str(rifEntry.get())
#     sector              =   str(sectorEntry.get())
#     uso                 =   str(usoEntry.get())
#     codcatastral        =   str(codcatastralEntry.get())
#     fechaliquidacion    =   str(fechaliquidacionEntry.get())
    
#     try:
#         cursor.connection.ping()
#         sql = """UPDATE reg 
#                  SET `cedula` = %s, `contribuyente` = %s, `nombreinmueble` = %s, 
#                      `rif` = %s, `sector` = %s, `uso` = %s, `codcatastral` = %s, 
#                      `fechaliquidacion` = %s 
#                  WHERE `cedula` = %s"""
        
#         cursor.execute(sql, (cedula, contribuyente, nombreinmueble, rif, sector, uso, codcatastral, fechaliquidacion, selectedItemId))
#         conn.commit()
#         conn.close()
        
#         for num in range(0,9):
#             setph('', num)
        
#         refreshTable()
        
#     except Exception as err:
#         messagebox.showwarning("", "Un error se produjo: " + str(err))
#         return

def delete():
    try:
        if(my_tree.selection()[0]):
            decision = messagebox.askquestion("", "Seguro de eliminar los datos seleccionados?")
            if(decision != 'yes'):
                return
            else:
                selectedItem = my_tree.selection()[0]
                cedula = str(my_tree.item(selectedItem)['values'][0])
                try:
                    cursor.connection.ping()
                    sql=f"DELETE FROM reg WHERE `cedula` = '{cedula}' "
                    cursor.execute(sql)
                    conn.commit()
                    conn.close()
                    messagebox.showinfo("","Se elimino el registro correctamente")
                except:
                    messagebox.showinfo("","Un error se produjo")
                refreshTable()
    except:
        messagebox.showwarning("", "Por favor selecciona fila")

def select():
    try:
        selectedItem = my_tree.selection()[0]
        cedula           = str(my_tree.item(selectedItem)['values'][0])
        contribuyente    = str(my_tree.item(selectedItem)['values'][1])
        nombreinmueble   = str(my_tree.item(selectedItem)['values'][2])
        rif              = str(my_tree.item(selectedItem)['values'][3])
        sector           = str(my_tree.item(selectedItem)['values'][4])
        uso              = str(my_tree.item(selectedItem)['values'][5])
        codcatastral     = str(my_tree.item(selectedItem)['values'][6])
        fechaliquidacion = str(my_tree.item(selectedItem)['values'][7])

        setph(cedula,0)
        setph(contribuyente,1)
        setph(nombreinmueble,2)
        setph(rif,3)
        setph(sector,4)
        setph(uso,5)
        setph(codcatastral,6)
        setph(fechaliquidacion,7)
    except:
        messagebox.showwarning("", "Porfavor selecciona una columna")


def find():
    try:
        conn = connection()  # Assuming you have a function to establish connection
        cursor = conn.cursor()
        
        cedula = str(cedulaEntry.get()).strip()
        contribuyente = str(contribuyenteEntry.get()).strip()
        nombreinmueble = str(nombreinmuebleEntry.get()).strip()
        rif = str(rifEntry.get()).strip()
        sector = str(sectorEntry.get()).strip()
        uso = str(usoEntry.get()).strip()
        codcatastral = str(codcatastralEntry.get()).strip()
        fechaliquidacion = str(fechaliquidacionEntry.get()).strip()
        
        fields = {
            'cedula': cedula,
            'contribuyente': contribuyente,
            'nombreinmueble': nombreinmueble,
            'rif': rif,
            'sector': sector,
            'uso': uso,
            'codcatastral': codcatastral,
            'fechaliquidacion': fechaliquidacion
        }

        for field, value in fields.items():
            if value:
                sql = f"SELECT `cedula`, `contribuyente`, `nombreinmueble`, `rif`, `sector`, `uso`, `codcatastral`, `fechaliquidacion` FROM reg WHERE `{field}` LIKE %s"
                cursor.execute(sql, (f'%{value}%',))
                result = cursor.fetchall()
                
                if result and len(result[0]) > 0:
                    for num in range(0, min(9, len(result[0]))):  # Ensure we don't go out of range
                        setph(result[0][num], (num))
                    conn.commit()
                    return
                else:
                    messagebox.showwarning("", "No se encontro el registro")
                    return

        conn.close()
    except pymysql.err.InterfaceError as e:
        print("InterfaceError:", e)
        messagebox.showwarning("", "An error occurred with the database interface.")
    except pymysql.err.OperationalError as e:
        print("OperationalError:", e)
        messagebox.showwarning("", "Operational error occurred.")
    except Exception as e:
        print("Error:", e)
        messagebox.showwarning("", f"An error occurred: {e}")


def clear():
    for num in range(0,9):
        setph('',(num))

def exportExcel():
    cursor.connection.ping()
    sql=f"SELECT `cedula`, `contribuyente`, `nombreinmueble`, `rif`, `sector`, `uso`, `codcatastral`, `fechaliquidacion` FROM reg ORDER BY `fechaliquidacion` ASC"
    cursor.execute(sql)
    dataraw=cursor.fetchall()
    fechaliquidacion = str(datetime.now())
    fechaliquidacion = fechaliquidacion.replace(' ', '_')
    fechaliquidacion = fechaliquidacion.replace(':', '-')
    dateFinal = fechaliquidacion[0:16]
    with open("stocks_"+dateFinal+".csv",'a',newline='') as f:
        w = csv.writer(f, dialect='excel')
        for record in dataraw:
            w.writerow(record)
    print("saved: stocks_"+dateFinal+".csv")
    conn.commit()
    conn.close()
    messagebox.showinfo("","Excel exportado exitosamente")

frame=tkinter.Frame(window,bg="#235441")
frame.pack()

btnColor="#368051"

manageFrame=tkinter.LabelFrame(frame,borderwidth=4)
manageFrame.grid(row=0,column=0,sticky="w",padx=[20,500],pady=20,ipadx=[6])

saveBtn     =   Button(manageFrame,text="Guardar",width=10,borderwidth=3,bg=btnColor,fg='white',command=save)
updateBtn   =   Button(manageFrame,text="Actualizar",width=10,borderwidth=3,bg=btnColor,fg='white',command=update)
deleteBtn   =   Button(manageFrame,text="Eliminar",width=10,borderwidth=3,bg=btnColor,fg='white',command=delete)
selectBtn   =   Button(manageFrame,text="Seleccionar",width=10,borderwidth=3,bg=btnColor,fg='white',command=select)
findBtn     =   Button(manageFrame,text="Buscar",width=10,borderwidth=3,bg=btnColor,fg='white',command=find)
clearBtn    =   Button(manageFrame,text="Limpiar",width=10,borderwidth=3,bg=btnColor,fg='white',command=clear)
exportBtn   =   Button(manageFrame,text="Exportar excel",width=15,borderwidth=3,bg=btnColor,fg='white',command=exportExcel)

saveBtn.grid    (row=0,column=0,padx=5,pady=5)
updateBtn.grid  (row=0,column=1,padx=5,pady=5)
deleteBtn.grid  (row=0,column=2,padx=5,pady=5)
selectBtn.grid  (row=0,column=3,padx=5,pady=5)
findBtn.grid    (row=0,column=4,padx=5,pady=5)
clearBtn.grid   (row=0,column=5,padx=5,pady=5)
exportBtn.grid  (row=0,column=6,padx=5,pady=5)

entriesFrame=tkinter.LabelFrame(frame,borderwidth=4)
entriesFrame.grid(row=1,column=0,sticky="w",padx=[20,200],pady=[0,20],ipadx=[6])

cedulaLabel             =       Label(entriesFrame,text="Cedula",anchor="w",width=20)
contribuyenteLabel      =       Label(entriesFrame,text="Contribuyente",anchor="w",width=20)
nombreinmuebleLabel     =       Label(entriesFrame,text="Nombre inmueble",anchor="w",width=20)
rifLabel                =       Label(entriesFrame,text="Rif",anchor="w",width=20)
sectorLabel             =       Label(entriesFrame,text="Sector",anchor="w",width=20)
usoLabel                =       Label(entriesFrame,text="Uso",anchor="w",width=20)
codcatastralLabel       =       Label(entriesFrame,text="Codigo catastral",anchor="w",width=20)
fechaliquidacionLabel   =       Label(entriesFrame,text="Fecha liquidacion",anchor="w",width=20)


cedulaLabel.grid            (row=0,column=0,padx=10)
contribuyenteLabel.grid     (row=1,column=0,padx=10)
nombreinmuebleLabel.grid    (row=2,column=0,padx=10)
rifLabel.grid               (row=3,column=0,padx=10)
sectorLabel.grid            (row=4,column=0,padx=10)
usoLabel.grid               (row=5,column=0,padx=10)
codcatastralLabel.grid      (row=6,column=0,padx=10)
fechaliquidacionLabel.grid  (row=7,column=0,padx=10)


# categoryArray=['Networking Tools','Computer Parts','Repair Tools','Gadgets']

cedulaEntry             =   Entry(entriesFrame,width=50,textvariable=placeholderArray[0])
contribuyenteEntry      =   Entry(entriesFrame,width=50,textvariable=placeholderArray[1])
nombreinmuebleEntry     =   Entry(entriesFrame,width=50,textvariable=placeholderArray[2])
rifEntry                =   Entry(entriesFrame,width=50,textvariable=placeholderArray[3])
sectorEntry             =   Entry(entriesFrame,width=50,textvariable=placeholderArray[4])
usoEntry                =   Entry(entriesFrame,width=50,textvariable=placeholderArray[5])
codcatastralEntry       =   Entry(entriesFrame,width=50,textvariable=placeholderArray[6])
fechaliquidacionEntry   =   Entry(entriesFrame,width=50,textvariable=placeholderArray[7])

cedulaEntry.grid            (row=0,column=2,padx=5,pady=5)
contribuyenteEntry.grid     (row=1,column=2,padx=5,pady=5)
nombreinmuebleEntry.grid    (row=2,column=2,padx=5,pady=5)
rifEntry.grid               (row=3,column=2,padx=5,pady=5)
sectorEntry.grid            (row=4,column=2,padx=5,pady=5)
usoEntry.grid               (row=5,column=2,padx=5,pady=5)
codcatastralEntry.grid      (row=6,column=2,padx=5,pady=5)
fechaliquidacionEntry.grid  (row=7,column=2,padx=5,pady=5)


# generateIdBtn=Button(entriesFrame,text="GENERATE ID",borderwidth=3,bg=btnColor,fg='white',command=generateRand)
# generateIdBtn.grid(row=0,column=3,padx=5,pady=5)

style.configure(window)
my_tree['columns']=("Cedula","Contribuyente","Nombre inmueble","Rif","Sector","Uso","Codigo Catastral","Fecha liquidacion")
my_tree.column("#0",width=0,stretch=NO)
my_tree.column("Cedula",anchor=W,width=70)
my_tree.column("Contribuyente",anchor=W,width=150)
my_tree.column("Nombre inmueble",anchor=W,width=150)
my_tree.column("Rif",anchor=W,width=100)
my_tree.column("Sector",anchor=W,width=150)
my_tree.column("Uso",anchor=W,width=150)
my_tree.column("Codigo Catastral",anchor=W,width=150)
my_tree.column("Fecha liquidacion",anchor=W,width=150)

my_tree.heading("Cedula",text="Cedula",anchor=W)
my_tree.heading("Contribuyente",text="Contribuyente",anchor=W)
my_tree.heading("Nombre inmueble",text="Nombre inmueble",anchor=W)
my_tree.heading("Rif",text="Rif",anchor=W)
my_tree.heading("Sector",text="Sector",anchor=W)
my_tree.heading("Uso",text="Uso",anchor=W)
my_tree.heading("Codigo Catastral",text="Codigo Catastral",anchor=W)
my_tree.heading("Fecha liquidacion",text="Fecha liquidacion",anchor=W)
my_tree.tag_configure('orow',background="#EEEEEE")
my_tree.pack()

refreshTable()

window.resizable(False,False)
window.mainloop()
