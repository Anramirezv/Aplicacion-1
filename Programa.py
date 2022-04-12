from tkinter import *
from tkinter import ttk
from datetime import date, datetime
from xlwt import *
from xlwt.Workbook import Workbook
from openpyxl import workbook
from openpyxl import load_workbook
from datetime import datetime
dtnow=datetime.now()
dt=dtnow.strftime("%d/%m/%Y %H:%M:%S")
print(dt)
#py installer --clean --onefile --windowed PrintoApp.py
codigos={}
Inicio_turno={}
fli_prod={}
rech_rot={}
cambio=["2.2","2.5","2.6"]
num_ord={}
nueva_refer={}
comments={}

class product:
    #Ventana principal
    
    def __init__(self,window):
        self.wind=window
        self.wind.title("Pintura 1 PrintoGlass")
        frame=LabelFrame(self.wind, text="Registro de códigos. ") #Frame contenedor
        frame.grid(row=1, column=0, columnspan=2, pady=10)   #pad y distancia entre objetos. 
        #boton
        Label(frame, text="Registro evento").grid(row=1, column=0)#input
        self.name=Entry(frame)
        self.name.grid(row=1, column=1)
        sn=self.name
        print(sn)
        #boton producto
        ttk.Button(frame, text="Guardar evento ", command= self.add_product).grid(row=4, columnspan=2, sticky=W+E)#boton  #sticky ancho del boton
        self.tree = ttk.Treeview(height=10,columns=2) #tabla que muestra lo que se ve
        self.tree.grid(row=4, column=0, columnspan=2)
        self.tree.heading("#0", text= "", anchor=CENTER)
        self.tree.heading("#1", text="", anchor=CENTER)
        #botones turnos
        ttk.Button(text="Finalizar turno  ", command=self.Reporte_final).grid(row=7, columnspan=2, sticky=W+E)
        ttk.Button(text="Registrar inicio de turno  ", command=self.inicio).grid(row=0, columnspan=2, sticky=W+E)
        """ttk.Button(text="Agregar comentario  ", command=self.add_comment).grid(row=6, columnspan=2, sticky=W+E)"""
        self.message=Label(text= "", fg= "red")
        self.message.grid(row=5, column=0, columnspan=2, sticky= W+E)
         #ventana
    def inicio(self):
        self.edit_wind=Toplevel()
        self.edit_wind.title("Registro inicio de turno")        
        self.message1=Label(self.edit_wind ,text="Inicio de turno ")#label
        self.message1.grid(row=1, column=1)
        self.message2=Label(self.edit_wind ,text="Referencia a trabajar: ")
        self.message2.grid(row=2, column=0, sticky=W+E)
        self.message3=Label(self.edit_wind,text="Código trabajador: ")
        self.message3.grid(row=4, column=0, sticky=W+E)
        self.mr5=Label(self.edit_wind, text= "", fg= "red")
        self.mr5.grid(row=5, columnspan=2, sticky=W+E)
        self.m2=Entry(self.edit_wind)#entry
        self.m2.grid(row=2, column=1)
        self.m3=Entry(self.edit_wind)
        self.m3.grid(row=4, column=1)
        Button(self.edit_wind,text="Guardar registro ", command=self.inicio_turno).grid(row=6, columnspan=4, sticky=W+E)#button   
        print(Inicio_turno)

    def add_comment(self):
        self.edit_wind3=Toplevel()
        self.edit_wind3.title("registro comentarios")
        Label(self.edit_wind3, text="Comentario: ").grid(row=1, column=0)
        self.comment=Text(self.edit_wind3,width=30, height=6 ).grid(row=1, column=1)
        self.mr5=Label(self.edit_wind3, text= "", fg= "red").grid(row=2, columnspan=2, sticky=W+E)
        ttk.Button(self.edit_wind3, text="Guardar comentario  ", command=self.Save_comm).grid(row=3, columnspan=2, sticky=W+E)     

    def Save_comm(self):
        comments[dt]=self.comment.get()
        self.mr5["text"]="Comentario registrado. Ya puede cerrar la ventana "
        print("Comentario guardado")

    def inicio_turno(self):
        if self.validation2():       
            it=self.m3.get()
            Inicio_turno[it]=self.m2.get()
            print(Inicio_turno)
            self.mr5["text"]="Información registrada. Ya puede cerrar la ventana. "                    
        else:
            print("name and reference is required")
        
    def guardarPlanilla(self,nombreArchivo):
            self.wb.save(nombreArchivo)
            print ("Generado")  

    def guardar_turno(self):
        dtnow=datetime.now()
        dt=dtnow.strftime("%H:%M:%S")
        fp=self.m4.get()
        fli_prod[fp]=self.m5.get()
        rt=self.m6.get()
        rech_rot[rt]=self.m7.get()
        num_ord[dt]=self.m8.get()
        print(Inicio_turno)
        print(fli_prod)
        print(rech_rot)
        print(num_ord)
        print("Datos guardados")
        print(Inicio_turno)
        self.mr4["text"]="Resitros de turno registrados "

    def generar_excel(self):
        Num=len(codigos)
        Num2=len(fli_prod)
        Num3=len(rech_rot)
        Num4=len(num_ord)
        num5=len(Inicio_turno)
        Num6=len(nueva_refer)
        fw=Workbook()
        ws1=fw.add_sheet("firsT")
        ws1.write(0,0,"Fecha")
        ws1.write(0,1,"Código")
        ws1.write(0,3,"Referencia")
        ws1.write(0,4,"Código trabajadores")
        ws1.write(0,5,"Flint")
        ws1.write(0,6,"Total produccion")
        ws1.write(0,7,"Cantidad rechazo")
        ws1.write(0,8,"Cantidad rotura")
        ws1.write(0,9,"Numero de orden")
        ws1.write(0,10,"Comentarios")
        ws1.write(0,12,dt)
        i=2
        j=1
        k=1
        l=1
        m=1
        n=1
        x=1
        for referencia in nueva_refer.items():
            ws1.write(i, 3,referencia )
            i+=1
            if i ==Num6+1:
                break
        for fecha, codigo in codigos.items():
            ws1.write(j,0,fecha)
            ws1.write(j,1,codigo)
            j+=1
            if j == Num+1:
                break
        for fecha, comment in comments.items():
            ws1.write(k,12,fecha)
            ws1.write(k,13,comment)
            k+=1
            if k == Num3+1:
                break

        for referencia, codigot in Inicio_turno.items():
            ws1.write(l,4,referencia)
            ws1.write(l,3,codigot)
            l+=1
            if l == num5+1:
                break
        for flint, produccion in fli_prod.items():
            ws1.write(m,5,flint)
            ws1.write(m,6,produccion)
            m+=1
            if m == Num2+1:
                break
        for rechazo, rotura in rech_rot.items():
            ws1.write(n,7,rechazo)
            ws1.write(n,8,rotura)
            n+=1
            if n == 6:
                break
        for orden, hora in num_ord.items():
            ws1.write(x, 10, orden)
            ws1.write(x,9, hora)
            x+=1
            if x == Num4+1:
                break
        fw.save("Reporte_pintura_1.xls")
        print("Guardado")

    def Reporte_final(self):
        self.edit_wind4=Toplevel()
        self.edit_wind4.title("Datos finales de turno")

        self.message10=Label(self.edit_wind4, text="Datos finales turno")
        self.message10.grid(row=1, column=0, sticky=W+E)
        self.message4=Label(self.edit_wind4, text="Flint: ")
        self.message4.grid(row=2, column=0, sticky=W+E)
        self.message5=Label(self.edit_wind4, text="Total producción: ")
        self.message5.grid(row=3, column=0, sticky=W+E)
        self.message6=Label(self.edit_wind4, text="Cantidad rechazo: ")
        self.message6.grid(row=4, column=0, sticky=W+E)
        self.message7=Label(self.edit_wind4, text="Cantidad rotura: ")
        self.message7.grid(row=5, column=0, sticky=W+E)
        self.message8=Label(self.edit_wind4, text="Número de orden: ")
        self.message8.grid(row=6, column=0, sticky=W+E)
        self.mr4=Label(self.edit_wind4, text= "", fg= "red")
        self.mr4.grid(row=7, column=0, sticky=W+E)
        self.m4=Entry(self.edit_wind4)
        self.m4.grid(row=2, column=1)
        self.m5=Entry(self.edit_wind4)
        self.m5.grid(row=3, column=1)
        self.m6=Entry(self.edit_wind4)
        self.m6.grid(row=4, column=1)
        self.m7=Entry(self.edit_wind4)
        self.m7.grid(row=5, column=1)
        self.m8=Entry(self.edit_wind4)
        self.m8.grid(row=6, column=1)
        Button(self.edit_wind4,text="Finalizar turno ", command=self.generar_excel).grid(row=9, columnspan=4, sticky=W+E)#button
        Button(self.edit_wind4,text="Registrar información ", command=self.guardar_turno).grid(row=8, columnspan=4, sticky=W+E)#button    
           
    def ref_nueva(self):
        dtnow=datetime.now()
        dt=dtnow.strftime("%H:%M:%S")
        nueva_refer[dt]=self.m12.get()
        fp=self.m14.get()
        fli_prod[fp]=self.m15.get()
        rt=self.m16.get()
        rech_rot[rt]=self.m17.get()
        num_ord[dt]=self.m18.get()
        print(Inicio_turno)
        print(fli_prod)
        print(rech_rot)
        print(num_ord)
        print("Datos guardados")
        print(Inicio_turno)
        self.mr4["text"]="Resitros de turno registrados "
        self.m14.delete(0,END)
        self.m15.delete(0,END)
        self.m16.delete(0,END)
        self.m17.delete(0,END)
        self.m18.delete(0,END) 
        print(nueva_refer)

    def validation(self):
        return len(self.name.get()) != 0

    def validation2(self):
        return len(self.m2.get()) != 0 and len(self.m3.get()) != 0    

    def add_product(self):
        dtnow=datetime.now()
        dt=dtnow.strftime("%H:%M:%S")
        if self.validation():
            print(self.name.get())
            if self.name.get() in cambio:
                self.edit_wind2=Toplevel()
                self.edit_wind2.title("Cambio de referencia")
                self.message2=Label(self.edit_wind2 ,text="Referencia a trabajar: ")
                self.message2.grid(row=1, column=0, sticky=W+E)
                self.message4=Label(self.edit_wind2, text="Flint: ")
                self.message4.grid(row=2, column=0, sticky=W+E)
                self.message5=Label(self.edit_wind2, text="Total producción: ")
                self.message5.grid(row=3, column=0, sticky=W+E)
                self.message6=Label(self.edit_wind2, text="Cantidad rechazo: ")
                self.message6.grid(row=4, column=0, sticky=W+E)
                self.message7=Label(self.edit_wind2, text="Cantidad rotura: ")
                self.message7.grid(row=5, column=0, sticky=W+E)
                self.message8=Label(self.edit_wind2, text="Número de orden: ")
                self.message8.grid(row=6, column=0, sticky=W+E)
                self.mr4=Label(self.edit_wind2, text= "", fg= "red")
                self.mr4.grid(row=7, column=0, sticky=W+E)
                self.m12=Entry(self.edit_wind2)#entry
                self.m12.grid(row=1, column=1)
                self.m14=Entry(self.edit_wind2)
                self.m14.grid(row=2, column=1)
                self.m15=Entry(self.edit_wind2)
                self.m15.grid(row=3, column=1)
                self.m16=Entry(self.edit_wind2)
                self.m16.grid(row=4, column=1)
                self.m17=Entry(self.edit_wind2)
                self.m17.grid(row=5, column=1)
                self.m18=Entry(self.edit_wind2)
                self.m18.grid(row=6, column=1)
                ttk.Button(self.edit_wind2, text="Registrar nueva referencia ", command=self.ref_nueva).grid(row=8, columnspan=2, sticky=W+E)
            codigos[dt]=self.name.get()
            print(codigos)
            self.message["text"]="Código {} registrado ".format(self.name.get())
            self.name.delete(0,END)
        else:
            print("name and price is required")
if __name__=="__main__":
    window=Tk()
    application=product(window)
    window.mainloop()
