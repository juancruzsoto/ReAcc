import sqlite3
import pyodbc
from tkinter import ttk
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import *
from datetime import datetime, timedelta
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
		#self.entradaUs.bind("<Return>", self.acceder)
		#self.entradaPw.bind("<Return>", self.acceder)
		self.entradaUs.bind("<Return>", self.verificacion)
		self.entradaPw.bind("<Return>", self.verificacion)
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

	def verificacion(self,*args):

		self.BD=Connection('DatabaseReAcc.db')

		self.user= self.entradaUs.get()
		self.BD.cursor.execute("SELECT contrasena,RazonSocial,intentos,estado,privilegios FROM UsersVentas where Usuario = '%s'"% (self.user,))
		self.datos=self.BD.cursor.fetchall()
		if self.datos!=[]:
			if self.datos[0][3]:
				self.privi=self.datos[0][4]
				self.password=self.entradaPw.get()

				self.RSTelemarketer=self.datos[0][1]
				if self.password==self.datos[0][0]:
					self.RealizarOP=RealizarOperacion(self.privi,self.RSTelemarketer,self.user,self.master)
					self.BD.cursor.execute("UPDATE UsersVentas SET intentos = %i WHERE Usuario = '%s'"% (0,self.user,))
					self.BD.conn.commit()

					self.BD.cursor.execute("UPDATE EstadoUsers SET Estado = 'Disponible' WHERE Usuario = '%s'"% (self.user,))
					self.BD.conn.commit()

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

class RealizarOperacion(Principal):
	def __init__(self,privilegios,RSTelemarketer,user,master):
		self.raiz=tk.Toplevel(master)
		self.raiz.focus_set()
		self.raiz.grab_set()
		self.version = "1.0"
		self.EnOperacion = False
		self.raiz.protocol("WM_DELETE_WINDOW", self.Salir())
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

		self.user=user
		#self.privilegios = privilegios
		self.BD=Connection('DatabaseReAcc.db')

		self.proceso2=None
		self.tiempo2=Label(self.raiz,bg="#A5B1B8", fg='red', font=("","12"))
		self.iniciarCrono()

		self.barramenu=Menu(self.raiz,background="#02475B",relief="groove",fg="black")
		self.raiz.config(menu=self.barramenu)

		self.archivoSalir=Menu(self.barramenu,tearoff=0)
		self.archivoSalir.add_command(label="Cambiar de Usuario")
		self.archivoSalir.add_command(label="Salir",command=self.Salir())

		self.archivoProcesos=Menu(self.barramenu,tearoff=0)

		if privilegios:
			self.archivoProcesos.add_command(label="Exportar Registros")
			self.archivoProcesos.add_command(label="Visualizar Estados")
			self.archivoProcesos.add_command(label="Reporte Horas")
			self.archivoProcesos.add_command(label="Modificar Usuario")
			self.archivoProcesos.add_command(label="Agregar Usuario")
		else:
			self.archivoProcesos.add_command(label="Exportar Registros",state="disabled")
			self.archivoProcesos.add_command(label="Visualizar Estados",state="disabled")
			self.archivoProcesos.add_command(label="Reporte Horas",state="disabled")
			self.archivoProcesos.add_command(label="Modificar Usuario",state="disabled")
			self.archivoProcesos.add_command(label="Agregar Usuario",state="disabled")



		self.barramenu.add_cascade(label="Procesos",menu=self.archivoProcesos)
		self.barramenu.add_cascade(label="Salir",menu=self.archivoSalir)

		self.Titulo=Label(self.miFrame1,bg="#A5B1B8", text="Registro de Actividades en curso",fg="#323638",font=("Ubuntu",30))
		self.Titulo.pack(fill=BOTH,side="top")

		'''
		self.ID=0
		self.Nombre=""
		self.Telefono=""
		self.Email=""
		self.HoraIni=""
		self.HoraFin=""
		self.FechaOp=""
		self.ReaCom=0'''


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
		self.botonIni=Button(self.miFrame3, text="Iniciar Operacion",bg="#02475B",fg="#42D5FF",width=20,height=0,command=self.iniciarOp)
		self.botonFin=Button(self.miFrame3, text="Finalizar Operacion",bg="#02475B",fg="#42D5FF",width=20,height=0,state="disabled")
		self.datoMot=StringVar()
		self.ListaMotivo= ttk.Combobox(self.miFrame3,width=31,textvariable = NO,state="disabled",values=["Carga de Pedido",
			"Gestión de reclamo","Comunicación de despacho","Pauta de horario","Cliente nuevo","Comunicación para otro sector","Gestión WhatsApp","Gestión por Instagram",
			"Gestión por Facebook","Gestión Web","Gestión PreVenta","Gestión Autorizada por Supervisor","Saliente ocupado","Saliente no atiende","Saliente numero erróneo"])
		self.ListaMotivo.current(0)
		self.proceso=None
		self.tiempo=Label(self.miFrame3,bg="#A5B1B8", fg='red', font=("","12"))
		self.tiempo.grid(row=7,column=0,columnspan=2)
		self.proceso=None

		#datoMot.set(OptionList[0])
		#ListaMotivo = OptionMenu(miFrame3,datoMot,*OptionList)
		#readonly
		self.MotivoLl.grid(row=0,column=2)
		self.ListaMotivo.grid(row=1,column=2,padx=50)

		self.botonIni.grid(row=6,column=2,pady=50,padx=50)
		self.botonFin.grid(row=6,column=3,pady=50)

	def iniciarOp(self):

		self.tiempo2.after_cancel(self.proceso2)
		self.EnOperacion=True

		self.BD.cursor.execute("UPDATE EstadoUsers SET Estado = 'Ocupado' WHERE Usuario = '%s'"% (self.user,))
		self.BD.conn.commit()

		self.entradaID["state"]="normal"
		self.botonID["state"]="normal"
		self.entradaCN["state"]="normal"
		self.entradaNW["state"]="normal"
		self.botonRec["state"]="normal"
		self.botonRea["state"]="normal"
		self.ListaMotivo["state"]="readonly"
		self.botonFin["state"]="normal"
		self.botonIni["state"]="disabled"
		self.Comentarios["state"]="normal"

		now = datetime.now()

		self.HoraIni=self.FormatoTiempo(str(now.hour),str(now.minute),str(now.second))

		self.iniciarCrono()

	#Registra el tiempo que lleva en actividad o libre
	def iniciarCrono(self,h=0,m=0,s=0):

		if s >= 60:
			s=0
			m=m+1
			if m >= 60:
				m=0
				h=h+1
				if h >= 24:
					h=0

		crono= self.FormatoTiempo(str(h),str(m),str(s))

		if self.EnOperacion:
			self.tiempo['text'] = crono

			self.BD.cursor.execute("UPDATE EstadoUsers SET ActEnCurso = '%s' WHERE Usuario = '%s'"% (self.ListaMotivo.get(),self.user,))

			self.BD.cursor.execute("UPDATE EstadoUsers SET TiempoEnCurso = '%s' WHERE Usuario = '%s'"% (crono,self.user,))
			self.BD.conn.commit()
		    # iniciamos la cuenta progresiva de los segundos
			self.proceso=self.tiempo.after(1000, self.iniciarCrono, (h), (m), (s+1))
		else:
			crono="-"+crono
		    #etiqueta que muestra el cronometro en pantalla
			self.tiempo2['text'] = crono

			self.BD.cursor.execute("UPDATE EstadoUsers SET TiempoEnCurso = '%s' WHERE Usuario = '%s'"% (crono,self.user,))
			self.BD.conn.commit()
		    # iniciamos la cuenta progresiva de los segundos
			self.proceso2=self.tiempo2.after(1000, self.iniciarCrono, (h), (m), (s+1))
				

	def FormatoTiempo(self,h,m,s):

		if len(h)==1:
			hora="0"+h+":"
		else:
			hora=h+":"

		if len(m)==1:
			hora=hora+"0"+m+":"
		else:
			hora=hora+m+":"
		if len(s)==1:
			hora=hora+"0"+s
		else:
			hora=hora+s

		return hora

	def Salir(self):

		if self.EnOperacion:
			if (messagebox.askyesno(message="¿Esta seguro que desea cerrar?", title="Confirmar Acción",parent=self.raiz)):
				self.BD.cursor.execute("UPDATE EstadoUsers SET Estado = 'Desconectado' WHERE Usuario = '%s'"% (self.user,))
				self.BD.cursor.execute("UPDATE EstadoUsers SET ActEnCurso = ' ' WHERE Usuario = '%s'"% (self.user,))
				self.BD.cursor.execute("UPDATE EstadoUsers SET TiempoEnCurso = ' ' WHERE Usuario = '%s'"% (self.user,))
				self.BD.conn.commit()
				self.raiz.destroy()
			else:
				self.BD.cursor.execute("UPDATE EstadoUsers SET Estado = 'Desconectado' WHERE Usuario = '%s'"% (self.user,))
				self.BD.cursor.execute("UPDATE EstadoUsers SET ActEnCurso = ' ' WHERE Usuario = '%s'"% (self.user,))
				self.BD.cursor.execute("UPDATE EstadoUsers SET TiempoEnCurso = ' ' WHERE Usuario = '%s'"% (self.user,))
				self.BD.conn.commit()
				self.raiz.destroy()

raizUser = tk.Tk()
InicioSesion = Principal(raizUser,tk)
raizUser.mainloop()

