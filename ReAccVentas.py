import sqlite3
import pyodbc
from tkinter import ttk
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import *

class Principal:
	def __init__(self, master,tk):
		self.master = master
		self.master.title("Registro de Acciones - Ventas")
		self.master.iconbitmap('C:/Registro Acciones Ventas/223.ico')
		self.master.resizable(0,0)
		self.master.geometry("350x300")

		self.frameUser = Frame(self.master, width=400, height=250,bg="#A5B1B8",relief="groove", borderwidth=5)
		self.frameUser.pack(fill=BOTH,side="top",expand=YES)
		
		self.IniciarSesion=Label(self.frameUser,bg="#A5B1B8",text="Iniciar Sesion",fg="black",height=2,font=("Ubuntu",20))
		self.Usuario=Label(self.frameUser,bg="#A5B1B8",text="Ingrese Usuario:",fg="black",height=2)
		self.Contraseña=Label(self.frameUser,bg="#A5B1B8",text="Ingrese Contraseña:",fg="black",height=2)
		self.datoUs=StringVar()
		self.datoPw=StringVar()
		self.entradaUs=Entry(self.frameUser,textvariable=self.datoUs)
		self.entradaPw=Entry(self.frameUser,textvariable=self.datoPw,show="*")
		self.entradaUs.bind("<Return>", self.verificacion)
		self.entradaPw.bind("<Return>", self.verificacion)
		#entradaUs.bind("<Return>", verificacion)
		#entradaPw.bind("<Return>", verificacion)
		self.botonIngresar=Button(self.frameUser, text="Ingresar",bg="#02475B",fg="#42D5FF",width=10,height=0,command=self.verificacion)
		self.botonIngresar.bind("<Return>", self.verificacion)
		self.botonCamCon=Button(self.frameUser, text="Cambiar Contraseña",bg="#02475B",fg="#42D5FF",width=15,height=0)

		self.IniciarSesion.grid(row=0,column=0,columnspan=2,pady=10,padx=80)
		self.Usuario.grid(row=1,column=0,pady=10)
		self.Contraseña.grid(row=2,column=0,pady=10)
		self.entradaUs.grid(row=1,column=1)
		self.entradaPw.grid(row=2,column=1)
		self.botonIngresar.grid(row=3,column=0,columnspan=2,pady=10)
		self.botonCamCon.grid(row=4,column=0,columnspan=2,pady=10)

		self.RSTelemarketer = ""
		self.privi = None
		self.user = ""

	def verificacion(self):

		self.BD=Connection('DatabaseReAcc.db')

		self.user= self.entradaUs.get()
		self.BD.cursor.execute("SELECT contrasena,RazonSocial,intentos,estado,privilegios FROM UsersVentas where Usuario = '%s'"% (self.user,))
		self.datos=self.BD.cursor.fetchall()
		print(self.datos)
		if self.datos!=[]:
			if self.datos[0][3]:
				self.privi=self.datos[0][4]
				self.password=self.entradaPw.get()

				self.RSTelemarketer=self.datos[0][1]
				if self.password==self.datos[0][0]:
					self.RealizarOP=RealizarOperacion(self.privi,self.RSTelemarketer)
					self.BD.cursor.execute("UPDATE UsersVentas SET intentos = %i WHERE Usuario = '%s'"% (0,self.user,))
					self.BD.conn.commit()

					self.BD.cursor.execute("UPDATE EstadoUsers SET Estado = 'Disponible' WHERE Usuario = '%s'"% (self.user,))
					self.BD.conn.commit()

					self.master.destroy()
				else:
					self.BD.cursor.execute("UPDATE UsersVentas SET intentos = %i WHERE Usuario = '%s'"% (self.datos[0][2]+1,self.user,))
					self.BD.conn.commit()
					if self.datos[0][2]==2:
						self.BD.cursor.execute("UPDATE UsersVentas SET estado = 0 WHERE Usuario = '%s'"% (self.user,))
						self.BD.conn.commit()					
					messagebox.showerror("No se pudo Iniciar Sesión","Contraseña Invalida")
			else:
				messagebox.showerror("No se pudo Iniciar Sesión","Usuario deshabilitado")
		else:
			messagebox.showerror("No se pudo Iniciar Sesión","Usuario no existente")

		self.BD.conn.close() 

class Connection:
	def __init__(self,BDD):
		self.conn = sqlite3.connect(BDD)
		self.cursor = self.conn.cursor()

class RealizarOperacion():
	def __init__(self,privilegios,RSTelemarketer):
		self.raiz=Tk()
		self.raiz.focus_set()
		self.raiz.grab_set()
		#self.raiz.protocol("WM_DELETE_WINDOW", Salir)
		self.version = "1.0"
		self.raiz.title("Registro de Acciones - Ventas         "+RSTelemarketer+"       V:"+self.version)
		self.raiz.resizable(0,0)
		self.raiz.geometry("800x550")
		self.raiz.config(bg="#A5B1B8")

		self.miFrame1=Frame(self.raiz, width=800, height=100,bg="#A5B1B8",relief="groove", borderwidth=5)
		self.miFrame1.pack(fill=BOTH,side="top")
		self.miFrame2=Frame(self.raiz, width=800, height=200,bg="#A5B1B8",relief="groove", borderwidth=5)
		self.miFrame2.pack(fill=BOTH,side="top",expand=YES)
		self.miFrame3=Frame(self.raiz, width=800, height=100,bg="#A5B1B8",relief="groove", borderwidth=5)
		self.miFrame3.pack(fill=BOTH,side="top",expand=YES)

		self.privilegios = privilegios

		self.Titulo=Label(self.miFrame1,bg="#A5B1B8", text="Registro de Actividades en curso",fg="#323638",font=("Ubuntu",30))
		self.Titulo.pack(fill=BOTH,side="top")

		self.ID=0
		self.Nombre=""
		self.Telefono=""
		self.Email=""
		self.HoraIni=""
		self.HoraFin=""
		self.FechaOp=""
		self.ReaCom=0


		self.Idclte=Label(self.miFrame2,bg="#A5B1B8",text="Ingrese Id:",fg="black",height=2)
		self.datoID=StringVar()
		self.entradaID=Entry(self.miFrame2,textvariable=self.datoID,state="disabled")
		#self.entradaID.bind("<Return>", filtra2)
		self.botonID=Button(self.miFrame2, text="Filtrar",bg="#02475B",fg="#42D5FF",width=5,height=0,state="disabled")
		#self.botonID.bind("<Return>", filtra2)

		self.Idclte.grid(row=0,column=0,padx=5)
		self.entradaID.grid(row=0,column=1,padx=5)
		self.botonID.grid(row=0,column=2,padx=5)


		self.Encabezado=ttk.Treeview(self.miFrame2, columns=[1,2,3,4], show="headings",height=1)
		self.Encabezado.column("#1", minwidth = 50, width=50, stretch=NO)
		self.Encabezado.column("#2", minwidth = 250, width=250, stretch=NO)
		self.Encabezado.column("#3", minwidth = 100, width=100, stretch=NO)
		self.Encabezado.column("#4", minwidth = 250, width=250, stretch=NO)
		self.Encabezado.heading(1, text = "ID", anchor = W)
		self.Encabezado.heading(2, text = "Nombre", anchor = W)
		self.Encabezado.heading(3, text = "Telefono", anchor = W)
		self.Encabezado.heading(4, text = "E-mail", anchor = W)

		self.Encabezado.grid(row=1,column=0,columnspan=4,padx=50,pady=10)


		self.CorreoN=Label(self.miFrame3,bg="#A5B1B8",text="¿Correo Nuevo?",fg="black",height=2)
		self.datoCN=StringVar()
		self.entradaCN=Entry(self.miFrame3,textvariable=self.datoCN,width=30,state="disabled")

		self.NumeroW=Label(self.miFrame3,bg="#A5B1B8",text="Numero WhatsApp",fg="black",height=2)
		self.datoNW=StringVar()
		self.entradaNW=Entry(self.miFrame3,textvariable=self.datoNW,width=20,state="disabled")


		self.CorreoN.grid(row=0,column=0,columnspan=2)
		self.entradaCN.grid(row=1,column=1,padx=5)
		self.NumeroW.grid(row=2,column=0,columnspan=2)
		self.entradaNW.grid(row=3,column=1,padx=5)

		self.TipoLl=Label(self.miFrame3,bg="#A5B1B8",text="Tipo de Llamada",fg="black",height=2)
		self.datoTLl=IntVar()
		self.botonRea=Radiobutton(self.miFrame3,text="Realizada",variable=self.datoTLl,value=0,state="disabled",bg="#A5B1B8",fg="black",pady=2)
		self.botonRec=Radiobutton(self.miFrame3,text="Recibida",variable=self.datoTLl,value=1,state="disabled",bg="#A5B1B8",fg="black",pady=2)

		self.TipoLl.grid(row=2,column=2,padx=50)
		self.botonRea.grid(row=3,column=2,padx=50)
		self.botonRec.grid(row=4,column=2,padx=50)

		self.ComenTxt = Label(self.miFrame3,bg="#A5B1B8",text="Observación: ",fg="black",height=2)

		self.Comentarios = Text(self.miFrame3,height = 5, width = 25,state="disabled")

		self.ComenTxt.grid(row=0,column=3,padx=50)
		self.Comentarios.grid(row=1,column=3,padx=50)

		self.OptionList = ["Carga de Pedido",
			"Gestión de reclamo","Comunicación de despacho","Pauta de horario","Cliente nuevo","Comunicación para otro sector","Gestión WhatsApp","Gestión por Instagram",
			"Gestión por Facebook","Gestión Web","Gestión PreVenta","Gestión Autorizada por Supervisor","Saliente ocupado","Saliente no atiende","Saliente numero erróneo"]
		self.MotivoLl=Label(self.miFrame3,bg="#A5B1B8",text="Seleccione Motivo",fg="black",height=2)
		self.botonIni=Button(self.miFrame3, text="Iniciar Operacion",bg="#02475B",fg="#42D5FF",width=20,height=0)
		self.botonFin=Button(self.miFrame3, text="Finalizar Operacion",bg="#02475B",fg="#42D5FF",width=20,height=0,state="disabled")
		self.datoMot=StringVar()
		self.ListaMotivo= ttk.Combobox(self.miFrame3,width=31,textvariable = NO,state="disabled",values=["Carga de Pedido",
			"Gestión de reclamo","Comunicación de despacho","Pauta de horario","Cliente nuevo","Comunicación para otro sector","Gestión WhatsApp","Gestión por Instagram",
			"Gestión por Facebook","Gestión Web","Gestión PreVenta","Gestión Autorizada por Supervisor","Saliente ocupado","Saliente no atiende","Saliente numero erróneo"])
		self.ListaMotivo.current(0)
		self.tiempo=Label(self.miFrame3,bg="#A5B1B8", fg='red', font=("","12"))
		self.tiempo.grid(row=7,column=0,columnspan=2)

		#datoMot.set(OptionList[0])
		#ListaMotivo = OptionMenu(miFrame3,datoMot,*OptionList)
		#readonly
		self.MotivoLl.grid(row=0,column=2)
		self.ListaMotivo.grid(row=1,column=2,padx=50)

		self.botonIni.grid(row=6,column=2,pady=50,padx=50)
		self.botonFin.grid(row=6,column=3,pady=50)

raizUser = tk.Tk()
InicioSesion = Principal(raizUser,tk)
raizUser.mainloop()