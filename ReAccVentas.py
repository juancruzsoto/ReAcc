import sqlite3
import pyodbc
from tkinter import ttk
import tkinter as tk
#from tkinter import tkentrycomplete
from tkinter import messagebox
from tkinter import filedialog
from tkinter import *
from datetime import datetime, timedelta
import calendar
from tkcalendar import Calendar, DateEntry
import xlsxwriter
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk)

class Principal:
	def __init__(self, master,tk):
		self.master = master
		self.master.title("Registro de Acciones - Ventas")
		self.master.iconbitmap('IconoRA.ico')
		self.master.resizable(0,0)
		self.master.geometry("350x300")

		self.frameUser = Frame(self.master, width=400, height=250,bg="#9BBCD1",relief="groove", borderwidth=5)
		self.frameUser.pack(fill=BOTH,side="top",expand=YES)
		
		self.IniciarSesion=Label(self.frameUser,bg="#9BBCD1",text="Iniciar Sesion",fg="black",height=2,font=("Ubuntu",20))
		self.Usuario=Label(self.frameUser,bg="#9BBCD1",text="Ingrese Usuario:",fg="black",height=2)
		self.Contraseña=Label(self.frameUser,bg="#9BBCD1",text="Ingrese Contraseña:",fg="black",height=2)
		self.datoUs=StringVar()
		self.datoPw=StringVar()
		self.entradaUs=Entry(self.frameUser,textvariable=self.datoUs)
		self.entradaPw=Entry(self.frameUser,textvariable=self.datoPw,show="*")
		self.entradaUs.bind("<Return>", self.verificacion)
		self.entradaPw.bind("<Return>", self.verificacion)
		self.botonIngresar=Button(self.frameUser, text="Ingresar",bg="#2D373D",fg="#42D5FF",width=10,height=0,command=self.verificacion)
		self.botonIngresar.bind("<Return>", self.verificacion)
		self.botonCamCon=Button(self.frameUser, text="Cambiar Contraseña",bg="#2D373D",fg="#42D5FF",width=15,height=0,command=lambda:CambiarContrasena(master))

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

	def GetRaiz(self):
		return self.IniciarSesion

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

class CambiarContrasena(Principal):

	def __init__(self,master):
		self.raizCamCon=tk.Toplevel(master)
		self.raizCamCon.title("Cambiar de Contraseña")
		self.raizCamCon.resizable(0,0)
		self.raizCamCon.geometry("400x210")
		self.frameCamCon=Frame(self.raizCamCon, width=400, height=250,bg="#9BBCD1",relief="groove", borderwidth=5)
		self.frameCamCon.pack(fill=BOTH,side="top",expand=YES)	

		self.BD = Connection('DatabaseReAcc.db')

		self.datoUscc=StringVar()
		self.datoPwcc=StringVar()
		self.datoPwcn=StringVar()
		self.UsuarioCC=Label(self.frameCamCon,bg="#9BBCD1",text="Ingrese Usuario:",fg="black",height=2)
		self.ContraseñaV=Label(self.frameCamCon,bg="#9BBCD1",text="Ingrese Contraseña Actual:",fg="black",height=2)
		self.ContraseñaN=Label(self.frameCamCon,bg="#9BBCD1",text="Ingrese Contraseña Nueva:",fg="black",height=2)
		self.entradaUscc=Entry(self.frameCamCon,textvariable=self.datoUscc)
		self.entradaPwcc=Entry(self.frameCamCon,textvariable=self.datoPwcc,show="*")
		self.entradaPwcn=Entry(self.frameCamCon,textvariable=self.datoPwcn,show="*")

		self.UsuarioCC.grid(row=1,column=0,padx=30,pady=4)
		self.ContraseñaV.grid(row=2,column=0,padx=30,pady=4)
		self.ContraseñaN.grid(row=3,column=0,padx=30,pady=4)
		self.entradaUscc.grid(row=1,column=1)
		self.entradaPwcc.grid(row=2,column=1)	
		self.entradaPwcn.grid(row=3,column=1)	

		self.botonReaCam=Button(self.frameCamCon, text="Realizar Cambio",bg="#2D373D",fg="#42D5FF",width=15,height=0,command=lambda:self.validacion())
		self.botonReaCam.grid(row=4,column=0,columnspan=2,pady=10)


	def validacion(self):
		self.BD.cursor.execute("SELECT contrasena FROM UsersVentas where Usuario = '%s'"% (self.entradaUscc.get(),))
		datos=self.BD.cursor.fetchall()

		password=self.entradaPwcc.get()

		if datos!=[]:
			if password==datos[0][0]:
				try:
					self.BD.cursor.execute("UPDATE UsersVentas SET contrasena = '%s' WHERE Usuario = '%s'"% (self.entradaPwcn.get(),self.entradaUscc.get(),))
					self.BD.conn.commit()
					messagebox.showinfo("Operacion Realizada con Éxito","La contraseña se cambio correctamente",parent=self.raizCamCon)
					self.raizCamCon.destroy()
				except:
					messagebox.showerror("No se pudo Cambiar Contraseña","Error al acceder a la Base de Datos",parent=self.raizCamCon)
			else:
				messagebox.showerror("No se pudo Cambiar Contraseña","Contraseña Actual Invalida",parent=self.raizCamCon)
		else:
			messagebox.showerror("No se pudo Cambiar Contraseña","Usuario no existente",parent=self.raizCamCon)



class Connection:
	def __init__(self,BDD):
		self.conn = sqlite3.connect(BDD)
		self.cursor = self.conn.cursor()

class RealizarOperacion():
	def __init__(self,privilegios,RSTelemarketer,user,master):
		self.raiz=tk.Toplevel(master)
		#self.raiz.focus_set()
		#self.raiz.grab_set()
		self.version = "1.0"
		self.EnOperacion = False
		self.raiz.protocol("WM_DELETE_WINDOW", lambda:self.Salir(master,False))
		self.raiz.title("Registro de Acciones - Ventas         "+RSTelemarketer+"       V:"+self.version)
		self.raiz.resizable(0,0)
		self.raiz.geometry("800x550+300+40")
		self.raiz.config(bg="#9BBCD1")

		self.miFrame1=Frame(self.raiz, width=800, height=100,bg="#9BBCD1",relief="groove", borderwidth=5)
		self.miFrame1.pack(fill=BOTH,side="top")
		self.miFrame2=Frame(self.raiz, width=800, height=200,bg="#9BBCD1",relief="groove", borderwidth=5)
		self.miFrame2.pack(fill=BOTH,side="top",expand=YES)
		self.miFrame3=Frame(self.raiz, width=800, height=100,bg="#9BBCD1",relief="groove", borderwidth=5)
		self.miFrame3.pack(fill=BOTH,side="top",expand=YES)

		self.user=user
		self.BD=Connection('DatabaseReAcc.db')

		self.proceso2=None
		self.tiempo2=Label(self.raiz,bg="#9BBCD1", fg='red', font=("","12"))
		self.iniciarCrono()

		self.barramenu=Menu(self.raiz,background="#9BBCD1",relief="groove",fg="black")
		self.raiz.config(menu=self.barramenu)

		self.archivoSalir=Menu(self.barramenu,tearoff=0)
		self.archivoSalir.add_command(label="Cambiar de Usuario",command=lambda:self.Salir(master,True))
		self.archivoSalir.add_command(label="Salir",command=lambda:self.Salir(master,False))

		self.archivoProcesos=Menu(self.barramenu,tearoff=0)

		if privilegios:
			self.archivoProcesos.add_command(label="Exportar Registros",command= lambda:ExportarRegistros(self.raiz,RSTelemarketer))
			self.archivoProcesos.add_command(label="Visualizar Estados",command=lambda:EstadoUsers(self.raiz,RSTelemarketer))
			self.archivoProcesos.add_command(label="Reporte Horas",command=lambda:ReporteHoras(self.raiz,RSTelemarketer))
			self.archivoProcesos.add_command(label="Modificar Usuario",command= lambda: ModificarUsuario(self.raiz,RSTelemarketer))
			self.archivoProcesos.add_command(label="Agregar Usuario",command= lambda:AltaUsuario(self.raiz))
		else:
			self.archivoProcesos.add_command(label="Exportar Registros",state="disabled")
			self.archivoProcesos.add_command(label="Visualizar Estados",state="disabled")
			self.archivoProcesos.add_command(label="Reporte Horas",state="disabled")
			self.archivoProcesos.add_command(label="Modificar Usuario",state="disabled")
			self.archivoProcesos.add_command(label="Agregar Usuario",state="disabled")

		self.barramenu.add_cascade(label="Procesos",menu=self.archivoProcesos)
		self.barramenu.add_cascade(label="Salir",menu=self.archivoSalir)

		self.Titulo=Label(self.miFrame1,bg="#9BBCD1", text="Registro de Actividades en curso",fg="#323638",font=("Ubuntu",30))
		self.Titulo.pack(fill=BOTH,side="top")

		self.ID=0
		self.Nombre=""
		self.Telefono=""
		self.Email=""
		self.HoraIni=""
		self.HoraFin=""
		self.FechaOp=""
		self.ReaCom=0


		self.Idclte=Label(self.miFrame2,bg="#9BBCD1",text="Ingrese Id:",fg="black",height=2)
		self.datoID=StringVar()
		self.entradaID=Entry(self.miFrame2,textvariable=self.datoID,state="disabled")
		self.entradaID.bind("<Return>", self.filtrarCliente)
		self.botonID=Button(self.miFrame2, text="Filtrar",bg="#2D373D",fg="#42D5FF",width=5,height=0,state="disabled",command=self.filtrarCliente)
		self.botonID.bind("<Return>", self.filtrarCliente)

		self.Idclte.grid(row=0,column=0,padx=5)
		self.entradaID.grid(row=0,column=1,padx=5)
		self.botonID.grid(row=0,column=2,padx=5)


		self.Encabezado=ttk.Treeview(self.miFrame2, columns=[1,2,3,4], show="headings",height=1)
		self.Encabezado.column("#1", minwidth = 50, width=50, stretch=NO)
		self.Encabezado.column("#2", minwidth = 250, width=250, stretch=NO)
		self.Encabezado.column("#3", minwidth = 120, width=120, stretch=NO)
		self.Encabezado.column("#4", minwidth = 250, width=250, stretch=NO)
		self.Encabezado.heading(1, text = "ID", anchor = W)
		self.Encabezado.heading(2, text = "Nombre", anchor = W)
		self.Encabezado.heading(3, text = "Telefono", anchor = W)
		self.Encabezado.heading(4, text = "E-mail", anchor = W)

		self.Encabezado.grid(row=1,column=0,columnspan=4,padx=50,pady=10)


		self.CorreoN=Label(self.miFrame3,bg="#9BBCD1",text="¿Correo Nuevo?",fg="black",height=2)
		self.datoCN=StringVar()
		self.entradaCN=Entry(self.miFrame3,textvariable=self.datoCN,width=30,state="disabled")

		self.NumeroW=Label(self.miFrame3,bg="#9BBCD1",text="Numero WhatsApp",fg="black",height=2)
		self.datoNW=StringVar()
		self.entradaNW=Entry(self.miFrame3,textvariable=self.datoNW,width=20,state="disabled")


		self.CorreoN.grid(row=0,column=0,columnspan=2)
		self.entradaCN.grid(row=1,column=1,padx=5)
		self.NumeroW.grid(row=2,column=0,columnspan=2)
		self.entradaNW.grid(row=3,column=1,padx=5)

		self.TipoLl=Label(self.miFrame3,bg="#9BBCD1",text="Tipo de Llamada",fg="black",height=2)
		self.datoTLl=IntVar()
		self.botonRea=Radiobutton(self.miFrame3,text="Realizada",variable=self.datoTLl,value=0,state="disabled",bg="#9BBCD1",fg="black",pady=2)
		self.botonRec=Radiobutton(self.miFrame3,text="Recibida",variable=self.datoTLl,value=1,state="disabled",bg="#9BBCD1",fg="black",pady=2)

		self.TipoLl.grid(row=2,column=2,padx=50)
		self.botonRea.grid(row=3,column=2,padx=50)
		self.botonRec.grid(row=4,column=2,padx=50)

		self.ComenTxt = Label(self.miFrame3,bg="#9BBCD1",text="Observación: ",fg="black",height=2)

		self.Comentarios = Text(self.miFrame3,height = 5, width = 25,state="disabled")

		self.ComenTxt.grid(row=0,column=3,padx=50)
		self.Comentarios.grid(row=1,column=3,padx=50)

		self.OptionListCliente = ["Carga de Pedido","Gestión de reclamo","Comunicación de despacho","Pauta de horario"]
		self.OptionList = ["Carga de Pedido",
			"Gestión de reclamo","Comunicación de despacho","Pauta de horario","Cliente nuevo","Comunicación para otro sector","Gestión WhatsApp","Gestión por Instagram",
			"Gestión por Facebook","Gestión Web","Gestión PreVenta","Gestión Autorizada por Supervisor","Saliente ocupado","Saliente no atiende","Saliente numero erróneo"]
		self.MotivoLl=Label(self.miFrame3,bg="#9BBCD1",text="Seleccione Motivo",fg="black",height=2)
		self.botonIni=Button(self.miFrame3, text="Iniciar Operacion",bg="#2D373D",fg="#42D5FF",width=20,height=0,command=self.iniciarOp)
		self.botonFin=Button(self.miFrame3, text="Finalizar Operacion",bg="#2D373D",fg="#42D5FF",width=20,height=0,state="disabled",command=lambda:self.registrarLlamada(RSTelemarketer))
		self.datoMot=StringVar()
		self.ListaMotivo= ttk.Combobox(self.miFrame3,width=31,textvariable = NO,state="disabled",values=["Carga de Pedido",
			"Gestión de reclamo","Comunicación de despacho","Pauta de horario","Cliente nuevo","Comunicación para otro sector","Gestión WhatsApp","Gestión por Instagram",
			"Gestión por Facebook","Gestión Web","Gestión PreVenta","Gestión Autorizada por Supervisor","Saliente ocupado","Saliente no atiende","Saliente numero erróneo"])
		self.ListaMotivo.current(0)
		self.proceso=None
		self.tiempo=Label(self.miFrame3,bg="#9BBCD1", fg='red', font=("","12"))
		self.tiempo.grid(row=7,column=0,columnspan=2)
		self.proceso=None

		self.MotivoLl.grid(row=0,column=2)
		self.ListaMotivo.grid(row=1,column=2,padx=50)

		self.botonIni.grid(row=6,column=2,pady=50,padx=50)
		self.botonFin.grid(row=6,column=3,pady=50)

		InicioSesion.entradaUs["state"]="disabled"
		InicioSesion.entradaPw["state"]="disabled"
		InicioSesion.botonIngresar["state"]="disabled"
		master.iconify()

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

		self.HoraIni=now.strftime("%H:%M:%S")

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

		crono= datetime.strptime("%s:%s:%s"%(h,m,s,), '%H:%M:%S').strftime("%H:%M:%S")

		if self.EnOperacion:
			self.tiempo['text'] = crono

			self.BD.cursor.execute("UPDATE EstadoUsers SET ActEnCurso = '%s' WHERE Usuario = '%s'"% (self.ListaMotivo.get(),self.user,))

			self.BD.cursor.execute("UPDATE EstadoUsers SET TiempoEnCurso = '%s' WHERE Usuario = '%s'"% (crono,self.user,))
			self.BD.conn.commit()
		    # iniciamos la cuenta progresiva de los segundos
			self.proceso=self.tiempo.after(1000, self.iniciarCrono, (h), (m), (s+1))
		else:
			crono="-"+str(crono)
		    #etiqueta que muestra el cronometro en pantalla
			self.tiempo2['text'] = crono

			self.BD.cursor.execute("UPDATE EstadoUsers SET TiempoEnCurso = '%s' WHERE Usuario = '%s'"% (crono,self.user,))
			self.BD.conn.commit()
		    # iniciamos la cuenta progresiva de los segundos
			self.proceso2=self.tiempo2.after(1000, self.iniciarCrono, (h), (m), (s+1))
				

	def filtrarCliente(self,*args):

		try:

			self.ID=int(self.entradaID.get())

			self.BD.cursor.execute("SELECT * FROM Clientes where IDCliente = %s"% (self.ID,))
			
			datos=self.BD.cursor.fetchall()
			self.Nombre=datos[0][1]
			self.Telefono=datos[0][2]
			self.Email=datos[0][3]

			self.Encabezado.insert("", 0,values=(self.ID,self.Nombre,self.Telefono,self.Email))
		except:
			messagebox.showerror("Operación Fallida","Verifique que el ID ingresado sea correcto",parent=self.raiz)	

	def registrarLlamada(self,RSTelemarketer):

		if self.ID!=0 or self.ListaMotivo.get() not in self.OptionListCliente:

			if (messagebox.askyesno(message="¿Realizo Compra?", title="",parent=self.raiz)):
				self.ReaCom=1
			else:
				self.ReaCom=0

			if self.datoTLl.get()==0:
				TipLL="Realizada"
			else:
				TipLL="Recibida"

			self.EnOperacion=False

			now = datetime.now()

			self.HoraFin=now.strftime("%H:%M:%S")

			self.tiempo.after_cancel(self.proceso)
			FechaOp=now.strftime("%Y-%m-%d 00:00:00")

			try:
				self.BD.cursor.execute("INSERT INTO LlamadasVtas(IDCliente,RazonSocial,CorreoNuevo,NumeroWp,TipoLlamada,MotivoLlamada,Telemarketer,Fecha,horaInicio,horaFin,duracionOp,RealizoCompra,Observaciones) VALUES(%s,'%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')" % 
						(self.ID,self.Nombre,self.entradaCN.get(),self.entradaNW.get(),TipLL,self.ListaMotivo.get(),RSTelemarketer,FechaOp,self.HoraIni,self.HoraFin,self.tiempo['text'],self.ReaCom,self.Comentarios.get(1.0,tk.END)[0:100]))
				self.BD.conn.commit()

				self.BD.cursor.execute("SELECT MAX(IDLlamada) FROM LlamadasVtas")
				datos=self.BD.cursor.fetchall()

				self.BD.cursor.execute("UPDATE EstadoUsers SET UltimaActividad = '%s' WHERE Usuario = '%s'"% (self.ListaMotivo.get(),self.user,))
				self.BD.cursor.execute("UPDATE EstadoUsers SET HoraActFin = '%s' WHERE Usuario = '%s'"% (self.HoraFin,self.user,))
				self.BD.cursor.execute("UPDATE EstadoUsers SET Estado = 'Disponible' WHERE Usuario = '%s'"% (self.user,))
				self.BD.cursor.execute("UPDATE EstadoUsers SET ActEnCurso = ' ' WHERE Usuario = '%s'"% (self.user,))
				self.BD.cursor.execute("UPDATE EstadoUsers SET TiempoEnCurso = ' ' WHERE Usuario = '%s'"% (self.user,))
				self.BD.conn.commit()

				self.iniciarCrono()
				messagebox.showinfo("Operación Exitosa","La actividad fue registrada correctamente con ID '%s'" % (str(datos[0][0]),),parent=self.raiz)
			except:
				messagebox.showerror("Operación Fallida","No se pudo acceder a la Base de Datos para completar el registro",parent=self.raiz)
			self.Nombre=""

			self.entradaID["state"]="disabled"
			self.botonID["state"]="disabled"
			self.entradaCN["state"]="disabled"
			self.entradaNW["state"]="disabled"
			self.botonRec["state"]="disabled"
			self.botonRea["state"]="disabled"
			self.ListaMotivo["state"]="disabled"
			self.ListaMotivo.current(0)
			self.botonFin["state"]="disabled"
			self.botonIni["state"]="normal"
			self.Comentarios.delete(1.0,tk.END)
			self.Comentarios["state"]="disabled"

			self.Encabezado.insert("", 0,values=("","","",""))
			self.datoID.set("")
			self.datoCN.set("")
			self.datoNW.set("")
			self.ID=0
		else:
			messagebox.showerror("Operación Fallida","Por favor, primero filtre un Cliente",parent=self.raiz)

	def Salir(self,master,Bool):

		if not self.EnOperacion or messagebox.askyesno(message="¿Esta seguro que desea cerrar?", title="Confirmar Acción",parent=self.raiz):
			self.BD.cursor.execute("UPDATE EstadoUsers SET Estado = 'Desconectado' WHERE Usuario = '%s'"% (self.user,))
			self.BD.cursor.execute("UPDATE EstadoUsers SET ActEnCurso = ' ' WHERE Usuario = '%s'"% (self.user,))
			self.BD.cursor.execute("UPDATE EstadoUsers SET TiempoEnCurso = ' ' WHERE Usuario = '%s'"% (self.user,))
			self.BD.conn.commit()

			if Bool:
				InicioSesion.entradaUs["state"]="normal"
				InicioSesion.entradaPw["state"]="normal"
				InicioSesion.botonIngresar["state"]="normal"
				InicioSesion.datoUs.set("")
				InicioSesion.datoPw.set("")
				InicioSesion.datoUs.set("")
				master.deiconify()
				self.raiz.destroy()
			else:
				master.destroy()

class ModificarUsuario:

	def __init__(self,raiz,RSTelemarketer):
		self.raizResetPass=Toplevel(raiz)
		#self.raizResetPass.focus_set()
		#self.raizResetPass.grab_set()
		#self.raizResetPass.iconify()
		self.raizResetPass.title("Modificar Usuarios            "+RSTelemarketer)
		self.raizResetPass.resizable(0,0)
		self.raizResetPass.geometry("600x500")
		self.frameResetPass=Frame(self.raizResetPass, width=600, height=80,bg="#9BBCD1",relief="groove", borderwidth=5)
		self.frameResetPass.pack(fill=BOTH,side="top")
		self.frameResetPass2=Frame(self.raizResetPass, width=600, height=420,bg="#9BBCD1",relief="groove", borderwidth=5)
		self.frameResetPass2.pack(fill=BOTH,side="top",expand=YES)

		self.Titulo=Label(self.frameResetPass,bg="#9BBCD1", text="Modificar Usuarios",fg="#323638",font=("Ubuntu",20))
		self.Titulo.pack(fill=BOTH,side="top")

		self.EncabezadoM=ttk.Treeview(self.frameResetPass2, columns=[1,2,3,4], show="headings",height=5)
		self.EncabezadoM.column("#1", minwidth = 100, width=100, stretch=NO)
		self.EncabezadoM.column("#2", minwidth = 100, width=100, stretch=NO)
		self.EncabezadoM.column("#3", minwidth = 250, width=250, stretch=NO)
		self.EncabezadoM.column("#4", minwidth = 100, width=100, stretch=NO)
		self.EncabezadoM.heading(1, text = "Usuario", anchor = W)
		self.EncabezadoM.heading(2, text = "Contraseña", anchor = W)
		self.EncabezadoM.heading(3, text = "Razon Social", anchor = W)
		self.EncabezadoM.heading(4, text = "Activo", anchor = W)

		self.EncabezadoM.grid(row=0,column=0,columnspan=4,padx=10,pady=20)

		self.scrolvert = Scrollbar(self.frameResetPass2, command = self.EncabezadoM.yview)
		self.scrolvert.grid(row=0, column=4, sticky="nsew")
		self.EncabezadoM.config(yscrollcommand=self.scrolvert.set)
		self.EncabezadoM.bind('<<TreeviewSelect>>',self.infoUsers)

		self.UsuarioR=Label(self.frameResetPass2,bg="#9BBCD1",text="Usuario:",fg="black",height=2)
		self.ContraseñaR=Label(self.frameResetPass2,bg="#9BBCD1",text="Contraseña:",fg="black",height=2)
		self.RazonSocialR=Label(self.frameResetPass2,bg="#9BBCD1",text="Razon Social:",fg="black",height=2)
		self.Horarios=Label(self.frameResetPass2,bg="#9BBCD1",text="Horarios",fg="black",height=3)
		self.LunAVie=Label(self.frameResetPass2,bg="#9BBCD1",text="Lunes a Viernes:",fg="black",height=2)
		self.Sabado=Label(self.frameResetPass2,bg="#9BBCD1",text="Sabado:",fg="black",height=2)
		self.CheckVar1 = IntVar()
		self.CheckVar2 = IntVar()
		self.C1 = Checkbutton(self.frameResetPass2, text = "Activo",bg="#9BBCD1", variable = self.CheckVar1, onvalue = 1, offvalue = 0)
		self.C2 = Checkbutton(self.frameResetPass2, text = "Privilegios",bg="#9BBCD1", variable = self.CheckVar2, onvalue = 1, offvalue = 0, height=5, width = 20)
		self.datosUsers=""
		self.datoUsr=StringVar()
		self.datoPwr=StringVar()
		self.datoRsr=StringVar()
		self.datoLaV=StringVar()
		self.datoSab=StringVar()
		self.entradaUsr=Entry(self.frameResetPass2,textvariable=self.datoUsr)
		self.entradaPwr=Entry(self.frameResetPass2,textvariable=self.datoPwr)
		self.entradaRsr=Entry(self.frameResetPass2,textvariable=self.datoRsr)
		self.entradaLaV=Entry(self.frameResetPass2,textvariable=self.datoLaV)
		self.entradaSab=Entry(self.frameResetPass2,textvariable=self.datoSab)

		self.botonModificar=Button(self.frameResetPass2, text="Modificar Usuario",bg="#2D373D",fg="#42D5FF",width=20,height=0,command=self.RealizarModificacion)


		self.UsuarioR.grid(row=1,column=0,padx=5)
		self.ContraseñaR.grid(row=1,column=2,padx=5)
		self.entradaUsr.grid(row=1,column=1,padx=5)
		self.entradaPwr.grid(row=1,column=3,padx=5)
		self.RazonSocialR.grid(row=2,column=0,padx=5)
		self.entradaRsr.grid(row=2,column=1,padx=5)
		self.C1.grid(row=2,column=2,padx=5)
		self.C2.grid(row=2,column=3,padx=5)
		self.Horarios.grid(row=3,column=0)
		self.LunAVie.grid(row=3,column=1,padx=5)
		self.entradaLaV.grid(row=4,column=1,padx=5)
		self.Sabado.grid(row=3,column=3,padx=5)
		self.entradaSab.grid(row=4,column=3,padx=5)


		self.botonModificar.grid(row=5,column=0,columnspan=4,pady=35,padx=15)

		self.BD=Connection('DatabaseReAcc.db')

		self.BD.cursor.execute("SELECT Usuario,contrasena,RazonSocial,estado FROM UsersVentas")
		datos=self.BD.cursor.fetchall()
		datos2=[]
		datosExp=datos
		for y in range(len(datos)):
			for x in datos[y]:
				datos2.append(x)
			if datos2[3]:
				datos2.append("Si")
			else:
				datos2.append("No")
			datos2.pop(3)
			self.EncabezadoM.insert("", 0,values=datos2)
			datos2=[]

	def RealizarModificacion(self):

		try:
			self.BD.cursor.execute("UPDATE UsersVentas SET Usuario = '%s' WHERE Usuario = '%s'"% (self.entradaUsr.get(),self.datosUsers[0],))
			self.BD.cursor.execute("UPDATE UsersVentas SET contrasena = '%s' WHERE Usuario = '%s'"% (self.entradaPwr.get(),self.datosUsers[0],))
			self.BD.cursor.execute("UPDATE UsersVentas SET RazonSocial = '%s' WHERE Usuario = '%s'"% (self.entradaRsr.get(),self.datosUsers[0],))
			self.BD.cursor.execute("UPDATE UsersVentas SET estado = %i WHERE Usuario = '%s'"% (self.CheckVar1.get(),self.datosUsers[0],))
			self.BD.cursor.execute("UPDATE UsersVentas SET privilegios = %i WHERE Usuario = '%s'"% (self.CheckVar2.get(),self.datosUsers[0],))
			self.BD.cursor.execute("UPDATE UsersVentas SET intentos = 0 WHERE Usuario = '%s'"% (self.datosUsers[0],))
			self.BD.cursor.execute("UPDATE EstadoUsers SET Activo = %i WHERE Usuario = '%s'"% (self.CheckVar1.get(),self.datosUsers[0],))
			self.BD.cursor.execute("UPDATE UsersVentas SET horasLaV = %i WHERE Usuario = '%s'"% (int(self.entradaLaV.get()),self.datosUsers[0],))
			self.BD.cursor.execute("UPDATE UsersVentas SET horasSab = %i WHERE Usuario = '%s'"% (int(self.entradaSab.get()),self.datosUsers[0],))

			self.BD.conn.commit()

			self.BD.cursor.execute("SELECT Usuario,contrasena,RazonSocial,estado FROM UsersVentas")
			datos=self.BD.cursor.fetchall()
			datos2=[]
			datosExp=datos
			self.EncabezadoM.delete(*self.EncabezadoM.get_children())
			for y in range(len(datos)):
				for x in datos[y]:
					datos2.append(x)
				if datos2[3]:
					datos2.append("Si")
				else:
					datos2.append("No")
				datos2.pop(3)
				self.EncabezadoM.insert("", 0,values=datos2)
				datos2=[]

			messagebox.showinfo("Modificación Realizada con Éxito","Los datos se registraron correctamente",parent=self.raizResetPass)
		except:
			messagebox.showerror("No se pudo realizar la Modificación","Error al acceder a la Base de Datos",parent=self.raizResetPass)


	def infoUsers(self,event):

		item=event.widget.selection()
		self.datosUsers=self.EncabezadoM.item(item[0])["values"]

		self.datoUsr.set(self.datosUsers[0])
		self.datoPwr.set(self.datosUsers[1])
		self.datoRsr.set(self.datosUsers[2])


		self.BD.cursor.execute("SELECT estado,privilegios,horasLaV,horasSab FROM UsersVentas WHERE Usuario= '%s'" % (self.datosUsers[0],))
		datos=self.BD.cursor.fetchall()

		if datos[0][0]:
			self.CheckVar1.set(1)
		else:
			self.CheckVar1.set(0)

		if datos[0][1]:
			self.CheckVar2.set(1)
		else:
			self.CheckVar2.set(0)

		self.datoLaV.set(datos[0][2])
		self.datoSab.set(datos[0][3])

class AltaUsuario:

	def __init__(self,raiz):
		self.raizAltaUser=tk.Toplevel(raiz)
		#self.raizAltaUser.focus_set()
		#self.raizAltaUser.grab_set()
		self.raizAltaUser.title("Alta de Usuarios")
		self.raizAltaUser.resizable(0,0)
		self.raizAltaUser.geometry("400x300")
		self.frameAltaUser=Frame(self.raizAltaUser, width=400, height=80,bg="#9BBCD1",relief="groove", borderwidth=5)
		self.frameAltaUser.pack(fill=BOTH,side="top")
		self.frameAltaUser2=Frame(self.raizAltaUser, width=400, height=220,bg="#9BBCD1",relief="groove", borderwidth=5)
		self.frameAltaUser2.pack(fill=BOTH,side="top",expand=YES)

		self.Titulo=Label(self.frameAltaUser,bg="#9BBCD1", text="Alta de Usuario",fg="#323638",font=("Ubuntu",20))
		self.Titulo.pack(fill=BOTH,side="top")
		
		self.UsuarioAU=Label(self.frameAltaUser2,bg="#9BBCD1",text="Ingrese Usuario Nuevo:",fg="black",height=2)
		self.ContraseñaAU=Label(self.frameAltaUser2,bg="#9BBCD1",text="Contraseña por Default:",fg="black",height=2)
		self.RazonSocialAU=Label(self.frameAltaUser2,bg="#9BBCD1",text="Ingrese Razon Social:",fg="black",height=2)
		self.datoUsau=StringVar()
		self.datoPwau=StringVar()
		self.datoRsau=StringVar()
		self.datoPwau.set("1234")
		self.entradaUsau=Entry(self.frameAltaUser2,textvariable=self.datoUsau)
		self.entradaPwau=Entry(self.frameAltaUser2,textvariable=self.datoPwau)
		self.entradaRsau=Entry(self.frameAltaUser2,textvariable=self.datoRsau)

		self.UsuarioAU.grid(row=1,column=0,padx=45,pady=4)
		self.ContraseñaAU.grid(row=2,column=0,padx=45,pady=4)
		self.RazonSocialAU.grid(row=3,column=0,padx=45,pady=4)
		self.entradaUsau.grid(row=1,column=1)
		self.entradaPwau.grid(row=2,column=1)	
		self.entradaRsau.grid(row=3,column=1)

		self.BD=Connection('DatabaseReAcc.db')

		self.botonReaAlt=Button(self.frameAltaUser2, text="Dar de alta",bg="#2D373D",fg="#42D5FF",width=15,height=0,command=self.DarAlta)
		self.botonReaAlt.grid(row=4,column=0,columnspan=2,pady=10,padx=10)

	def DarAlta(self):

		try:
			self.BD.cursor.execute("INSERT INTO UsersVentas(Usuario,contrasena,RazonSocial,intentos,estado,privilegios,horasLaV,horasSab) values('%s','%s','%s',0,1,0,0,0)"% (self.entradaUsau.get(),self.entradaPwau.get(),self.entradaRsau.get(),))
			self.BD.cursor.execute("INSERT INTO EstadoUsers(Usuario,RazonSocial,UltimaActividad,HoraActFin,ActEnCurso,TiempoEnCurso,Estado,Activo) values('%s','%s',' ',' ',' ',' ','Desconectado',1)"% (self.entradaUsau.get(),self.entradaRsau.get(),))

			self.BD.conn.commit()
			messagebox.showinfo("Operacion Realizada con Éxito","El Usuario se dio de alta correctamente",parent=self.raizAltaUser)
			self.datoUsau.set("")
			self.datoRsau.set("")
		except:
			messagebox.showerror("No se pudo dar de Alta Usuario","Error al acceder a la Base de Datos",parent=self.raizAltaUser)

class ExportarRegistros:

	def __init__(self,raiz,RSTelemarketer):
		self.raizExpReg=tk.Toplevel(raiz)
		#self.raizExpReg.focus_set()
		#self.raizExpReg.grab_set()
		self.raizExpReg.title("Exportar Registros            "+RSTelemarketer)
		self.raizExpReg.resizable(0,0)
		self.raizExpReg.geometry("1300x650")
		self.frameExpReg=Frame(self.raizExpReg, width=1300, height=80,bg="#9BBCD1",relief="groove", borderwidth=5)
		self.frameExpReg.pack(fill=BOTH,side="top")
		self.frameExpReg2=Frame(self.raizExpReg, width=1300, height=570,bg="#9BBCD1",relief="groove", borderwidth=5)
		self.frameExpReg2.pack(fill=BOTH,side="top",expand=YES)

		self.Titulo=Label(self.frameExpReg,bg="#9BBCD1", text="Exportar Registro de Llamadas",fg="#323638",font=("Ubuntu",20))
		self.Titulo.pack(fill=BOTH,side="top")
			
		self.TextFI=Label(self.frameExpReg2,bg="#9BBCD1", text="Desde: ",fg="#323638",font=("Ubuntu",10))
		self.TextFI.place(x=50,y=9)

		now = datetime.now()

		self.calendar1 = DateEntry(self.frameExpReg2, selectmode="day", locale="es_AR", year=now.year, month=now.month, day=now.day)
		self.calendar1.place(x=100,y=10)

		self.TextFF=Label(self.frameExpReg2,bg="#9BBCD1", text="Hasta: ",fg="#323638",font=("Ubuntu",10))
		self.TextFF.place(x=200,y=9)

		self.calendar2 = DateEntry(self.frameExpReg2, selectmode="day", locale="es_AR", year=now.year, month=now.month, day=now.day)
		self.calendar2.place(x=250,y=10)


		self.botonBusReg=Button(self.frameExpReg2, text="Buscar",bg="#2D373D",fg="#42D5FF",width=15,height=0,command=self.filtrado2)
		self.botonBusReg.grid(row=0,column=1,pady=8,padx=200)	

		self.botonExpReg=Button(self.frameExpReg2, text="Exportar",bg="#2D373D",fg="#42D5FF",width=15,height=0,command= self.Exportan2)
		self.botonExpReg.grid(row=2,column=1,columnspan=3)	

		self.BD=Connection('DatabaseReAcc.db')

		self.EncabezadoER=ttk.Treeview(self.frameExpReg2, columns=[1,2,3,4,5,6,7,8,9,10,11,12,13,14], show="headings",height=20)
		self.EncabezadoER.column("#1", minwidth = 80, width=80, stretch=NO)
		self.EncabezadoER.column("#2", minwidth = 60, width=60, stretch=NO)
		self.EncabezadoER.column("#3", minwidth = 180, width=180, stretch=NO)
		self.EncabezadoER.column("#4", minwidth = 70, width=70, stretch=NO)
		self.EncabezadoER.column("#5", minwidth = 80, width=80, stretch=NO)
		self.EncabezadoER.column("#6", minwidth = 100, width=100, stretch=NO)
		self.EncabezadoER.column("#7", minwidth = 100, width=100, stretch=NO)
		self.EncabezadoER.column("#8", minwidth = 100, width=100, stretch=NO)
		self.EncabezadoER.column("#9", minwidth = 67, width=67, stretch=NO)
		self.EncabezadoER.column("#10", minwidth = 70, width=70, stretch=NO)
		self.EncabezadoER.column("#11", minwidth = 65, width=65, stretch=NO)
		self.EncabezadoER.column("#12", minwidth = 65, width=65, stretch=NO)
		self.EncabezadoER.column("#13", minwidth = 65, width=65, stretch=NO)
		self.EncabezadoER.column("#14", minwidth = 150, width=150,stretch=NO)

		self.EncabezadoER.heading(1, text = "IDLlamada", anchor = W)
		self.EncabezadoER.heading(2, text = "IDCliente", anchor = W)
		self.EncabezadoER.heading(3, text = "Razon Social", anchor = W)
		self.EncabezadoER.heading(4, text = "E-mail Nuevo", anchor = W)
		self.EncabezadoER.heading(5, text = "Numero WhatsApp", anchor = W)
		self.EncabezadoER.heading(6, text = "Tipo Llamada", anchor = W)
		self.EncabezadoER.heading(7, text = "Motivo Llamada", anchor = W)
		self.EncabezadoER.heading(8, text = "Telemarketer", anchor = W)
		self.EncabezadoER.heading(9, text = "Fecha", anchor = W)
		self.EncabezadoER.heading(10, text = "Hora Inicio", anchor = W)
		self.EncabezadoER.heading(11, text = "Hora Fin", anchor = W)
		self.EncabezadoER.heading(12, text = "Duración", anchor = W)
		self.EncabezadoER.heading(13, text = "Compro", anchor = W)
		self.EncabezadoER.heading(14, text="Observaciones",anchor=W)

		self.EncabezadoER.grid(row=1,column=0,columnspan=6,padx=5,pady=10)

		self.scrolvert = Scrollbar(self.frameExpReg2, command = self.EncabezadoER.yview)
		self.scrolvert.grid(row=1, column=6, sticky="nsew")
		self.EncabezadoER.config(yscrollcommand=self.scrolvert.set)

		self.datosExp = ""

	def filtrado2(self):

		if self.calendar1.get_date()>self.calendar2.get_date():
			messagebox.showerror("Fechas Erroneas","Por favor elija un rango de fecha valido",parent=self.raizExpReg)
			return

		self.EncabezadoER.delete(*self.EncabezadoER.get_children())
		try:

			self.BD.cursor.execute("SELECT * FROM (SELECT * FROM LlamadasVtas Where Fecha >= '%s') Where Fecha <= '%s'" % (str(self.calendar1.get_date()),str(self.calendar2.get_date())+" 00:00:00.000",))
			
			datos = self.BD.cursor.fetchall()

			datos2=[]
			self.datosExp=datos
			for y in range(len(datos)):
				for x in datos[y]:
					datos2.append(x)
				if datos2[12]:
					datos2.insert(13, 'Si')
				else:
					datos2.insert(13, 'No')
				datos2.pop(12)
				self.EncabezadoER.insert("", 0,values=datos2)
				datos2=[]
		except:
			messagebox.showerror("Operación Fallida","Error al Obtener Registros")	

	def Exportan2(self):

		f = filedialog.asksaveasfile(mode='a',defaultextension=".xlsx", filetypes=[("Ficheros Excel","*.xlsx")])
		if f is None: # asksaveasfile return `None` if dialog closed with "cancel".
			return

		try:
			libro = xlsxwriter.Workbook(f.name)
			hoja = libro.add_worksheet()
			Cabecera=["IDLlamada","IDCliente","Razon Social","E-mail","Numero WhatsApp","Tipo Llamada","Motivo Llamada","Telemarketer","Fecha","Hora Inicio","Hora Fin","Duración","Compro","Observaciones"]
			
			for x in range(len(Cabecera)):
				hoja.write(0, x, Cabecera[x])

			for y in range(len(self.datosExp)):
				for x in range(8):
					hoja.write(y+1, x, self.datosExp[y][x])
				hoja.write(y+1, 8, datetime.strptime(str(self.datosExp[y][8]), '%Y-%m-%d %H:%M:%S'))
				hoja.write(y+1, 9, datetime.strptime(str(self.datosExp[y][9]), '%H:%M:%S'))
				hoja.write(y+1, 10, datetime.strptime(str(self.datosExp[y][10]), '%H:%M:%S'))
				hoja.write(y+1, 11, datetime.strptime(str(self.datosExp[y][11]), '%H:%M:%S'))
				if self.datosExp[y][12]:
					hoja.write(y+1, 12, "Si")
				else:
					hoja.write(y+1, 12, "No")
				hoja.write(y+1,13,self.datosExp[y][13])
			libro.close()
			messagebox.showinfo("Operación Realizada con Éxito","El archivo se exportó correctamente",parent=self.raizExpReg)
		except:
			messagebox.showerror("Operación Fallida","No se pudo exportar el archivo",parent=self.raizExpReg)

class EstadoUsers:

	def __init__(self,raiz,RSTelemarketer):
		self.raizEstUs=tk.Toplevel(raiz)
		#self.raizEstUs.focus_set()
		#self.raizEstUs.grab_set()
		self.raizEstUs.title("Estado de Usuarios            "+RSTelemarketer)
		self.raizEstUs.resizable(0,0)
		self.raizEstUs.geometry("1150x400")
		self.frameEstUs=Frame(self.raizEstUs, width=1150, height=80,bg="#9BBCD1",relief="groove", borderwidth=5)
		self.frameEstUs.pack(fill=BOTH,side="top")
		self.frameEstUs2=Frame(self.raizEstUs, width=1150, height=420,bg="#9BBCD1",relief="groove", borderwidth=5)
		self.frameEstUs2.pack(fill=BOTH,side="top",expand=YES)

		self.Titulo=Label(self.frameEstUs,bg="#9BBCD1", text="Ver Estado de Usuarios",fg="#323638",font=("Ubuntu",20))
		self.Titulo.pack(fill=BOTH,side="top")

		self.BD=Connection('DatabaseReAcc.db')

		self.style = ttk.Style()
		self.style.map('Treeview', foreground=self.fixed_map('foreground'), background=self.fixed_map('background'),fieldbackground=self.fixed_map('fieldbackground'))

		self.EncabezadoAE=ttk.Treeview(self.frameEstUs2, columns=[1,2,3,4,5,6,7], show="headings",height=8)
		self.EncabezadoAE.column("#1", minwidth = 80, width=80, stretch=NO)
		self.EncabezadoAE.column("#2", minwidth = 180, width=180, stretch=NO)
		self.EncabezadoAE.column("#3", minwidth = 190, width=190, stretch=NO)
		self.EncabezadoAE.column("#4", minwidth = 180, width=180, stretch=NO)
		self.EncabezadoAE.column("#5", minwidth = 190, width=190, stretch=NO)
		self.EncabezadoAE.column("#6", minwidth = 180, width=180, stretch=NO)
		self.EncabezadoAE.column("#7", minwidth = 120, width=120, stretch=NO)

		self.EncabezadoAE.heading(1, text = "Usuario", anchor = W)
		self.EncabezadoAE.heading(2, text = "Razon Social", anchor = W)
		self.EncabezadoAE.heading(3, text = "Ultima Actividad", anchor = W)
		self.EncabezadoAE.heading(4, text = "Hora Actividad Finalizada", anchor = W)
		self.EncabezadoAE.heading(5, text = "Actividad Realizando", anchor = W)
		self.EncabezadoAE.heading(6, text = "Tiempo en Actividad", anchor = W)
		self.EncabezadoAE.heading(7, text = "Estado", anchor = W)
		
		self.EncabezadoAE.grid(row=1,column=0,columnspan=5,padx=5,pady=10)

		self.ActualizarEstados()

	def fixed_map(self,option):
		
		return [elm for elm in self.style.map('Treeview', query_opt=option) if
			elm[:2] != ('!disabled', '!selected')]

	def ActualizarEstados(self):

		self.BD.cursor.execute("SELECT Usuario,RazonSocial,UltimaActividad,HoraActFin,ActEnCurso,TiempoEnCurso,Estado FROM EstadoUsers WHERE Activo=1")
		datos=self.BD.cursor.fetchall()
		datos2=[]
		datosExp=datos

		self.EncabezadoAE.delete(*self.EncabezadoAE.get_children())
		self.EncabezadoAE.tag_configure('white', foreground='#EEFFFF')
		self.EncabezadoAE.tag_configure("red",background="#E50000")
		self.EncabezadoAE.tag_configure("green",background="dark green")
		self.EncabezadoAE.tag_configure("gray", background="lightgray")

		for y in range(len(datos)):
			for x in datos[y]:
				datos2.append(x)

			if datos2[6]=='Disponible':
				self.EncabezadoAE.insert("",tk.END,tag=("green","white"),values=datos2)
			else:
				if datos2[6]=='Ocupado':
					self.EncabezadoAE.insert("",tk.END,tag=("red","white"),values=datos2)
				else:
					self.EncabezadoAE.insert("",tk.END,tag="gray",values=datos2)
			datos2=[]
			
		procesoact=self.Titulo.after(1000, self.ActualizarEstados)



class ReporteHoras:

	def __init__(self,raiz,RSTelemarketer):
		self.raizRepHor=tk.Toplevel(raiz)
		#self.raizRepHor.focus_set()
		#self.raizRepHor.grab_set()
		self.raizRepHor.title("Reporte de Horas            "+RSTelemarketer)
		self.raizRepHor.resizable(0,0)
		self.raizRepHor.geometry("1200x650")
		self.frameRepHor=Frame(self.raizRepHor, width=1200, height=80,bg="#9BBCD1",relief="groove", borderwidth=5)
		self.frameRepHor.pack(fill=BOTH,side="top")
		self.frameRepHor2=Frame(self.raizRepHor, width=1200, height=570,bg="#9BBCD1",relief="groove", borderwidth=5)
		self.frameRepHor2.pack(fill=BOTH,side="top",expand=YES)

		self.BD=Connection('DatabaseReAcc.db')
		self.meses=["Enero","Febrero","Marzo",
			"Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]

		Titulo=Label(self.frameRepHor,bg="#9BBCD1", text="Reporte de Horas Realizadas",fg="#323638",font=("Ubuntu",20))
		Titulo.pack(fill=BOTH,side="top")

		self.fig = Figure(figsize=(9, 8), dpi=60)
		self.fig.add_subplot(111).plot([], [], marker = 'o')
		self.fig.set_facecolor('#9BBCD1')

		self.fig2 = Figure(figsize=(7, 6), dpi=80)
		self.fig2.add_subplot(111).pie([100,0], labels=["Tiempo Productivo","Tiempo Improductivo"],colors= ["Green","Red"],autopct="%0.1f %%")
		self.fig2.set_facecolor('#9BBCD1')
		#fig2.add_subplot(111).axis("equal")


		self.canvas = FigureCanvasTkAgg(self.fig, master=self.frameRepHor2)  # A tk.DrawingArea.
		self.canvas.draw()
		#canvas.get_tk_widget().grid(row=0,column=0,pady=5)

		self.toolbar = NavigationToolbar2Tk(self.canvas, self.frameRepHor2)
		self.toolbar.update()
		#toolbar.pack(side="bottom",anchor=W)
		self.canvas.get_tk_widget().place(x=10,y=50)


		self.canvas2 = FigureCanvasTkAgg(self.fig2, master=self.frameRepHor2)  # A tk.DrawingArea.
		self.canvas2.draw()
		#canvas.get_tk_widget().grid(row=0,column=0,pady=5)

		self.toolbar2 = NavigationToolbar2Tk(self.canvas2, self.frameRepHor2)
		self.toolbar2.update()
		#toolbar2.pack(side="bottom",anchor=W)
		self.canvas2.get_tk_widget().place(x=570,y=50)
		self.toolbar.place(x=0,y=550)
		self.toolbar2.place(x=570,y=550)

		self.botonGenRep=Button(self.frameRepHor2, text="Generar Reporte",bg="#2D373D",fg="#42D5FF",width=15,height=0,command=lambda:self.GenerarReporteH())
		self.botonGenRep.place(x=650,y=500)

		self.DatCabecera = []

		self.botonExpor=Button(self.frameRepHor2, text="Exportar Reporte",bg="#2D373D",state="disabled",fg="#42D5FF",width=15,height=0,command=lambda:self.ExportarReporteH())
		self.botonExpor.place(x=800,y=500)
	
		self.ListaTmk= ttk.Combobox(self.frameRepHor2,width=25,textvariable = NO,state="readonly",values=self.ObtenerTelemarketers())
		self.ListaTmk.current(0)
		self.ListaTmk.place(x=30,y=50)
		self.prueb=StringVar()

		self.ListaAno= ttk.Combobox(self.frameRepHor2,width=15,state="readonly",values=self.ObtenerYears())
		self.ListaAno.place(x=450,y=50)
		self.ListaAno.current(0)

		self.ListaMes= ttk.Combobox(self.frameRepHor2,width=15,state="readonly",values=self.ObtenerMonths())
		self.ListaMes.place(x=300,y=50)
		self.ListaMes.current(0)

		self.SelecTmk=Label(self.frameRepHor2,bg="#9BBCD1", text="Seleccione Telemarketer",fg="#323638",font=("Ubuntu",10))
		self.SelecTmk.place(x=30,y=25)
		self.SelecMes=Label(self.frameRepHor2,bg="#9BBCD1", text="Mes:",fg="#323638",font=("Ubuntu",10))
		self.SelecMes.place(x=300,y=25)
		self.SelecAno=Label(self.frameRepHor2,bg="#9BBCD1", text="Año:",fg="#323638",font=("Ubuntu",10))
		self.SelecAno.place(x=450,y=25)


	def ObtenerTelemarketers(self):
		self.BD.cursor.execute("SELECT RazonSocial FROM UsersVentas where estado = 1 and privilegios = 0  order by RazonSocial")
		datos=self.BD.cursor.fetchall()
		salida=[]

		for x in range(len(datos)):
			salida.append(datos[x][0])

		return salida

	def ObtenerYears(self):
		#self.BD.cursor.execute("SELECT YEAR(Min(Fecha)), YEAR(Max(Fecha)) FROM LlamadasVtas")
		self.BD.cursor.execute("SELECT Min(Fecha), Max(Fecha) FROM LlamadasVtas")
		datos=self.BD.cursor.fetchall()[0]
		salida=[]
		for x in range(len(datos)):
			#salida.append(datos[x])
			salida.append(int(datetime.strptime(datos[x], '%Y-%m-%d %H:%M:%S').strftime("%Y")))

		return set(salida)

	def ObtenerMonths(self):
		#self.BD.cursor.execute("SELECT MONTH(Max(Fecha)) FROM LlamadasVtas WHERE YEAR(Fecha) = '%s'"% (year,))
		self.BD.cursor.execute("SELECT Min(Fecha), Max(Fecha) FROM LlamadasVtas")
		datos=self.BD.cursor.fetchall()[0]

		salida=self.meses[int(datetime.strptime(datos[0], '%Y-%m-%d %H:%M:%S').strftime("%m"))-1:int(datetime.strptime(datos[1], '%Y-%m-%d %H:%M:%S').strftime("%m"))-1]

		return salida

	def GenerarReporteH(self):

		meses=["01","02","03","04","05","06","07","08","09","10","11","12"]

		self.BD.cursor.execute("SELECT * FROM (SELECT Telemarketer,duracionOP,Fecha FROM LlamadasVtas WHERE Fecha >= '%s-%s-01') WHERE Fecha < '%s-%s-01'"% (self.ListaAno.get(),meses[self.meses.index(self.ListaMes.get())],self.ListaAno.get(),meses[self.meses.index(self.ListaMes.get())+1],))

		datos=self.BD.cursor.fetchall()
		dataimp=[]

		for x in range(len(datos)):
			if self.ListaTmk.get() in str(datos[x][0]):
				dataimp.append(datos[x])

		marca=str(dataimp[0][2])
		horasacum=datetime.strptime('00:00:00', '%H:%M:%S')

		diasT=[]
		horasT=[]
		horasLaborales=0

		self.BD.cursor.execute("SELECT horasLaV, horasSab FROM UsersVentas WHERE RazonSocial = '%s'"% (self.ListaTmk.get(),))
		dataHoras=self.BD.cursor.fetchall()

		horasLaV=dataHoras[0][0]
		horasSab=dataHoras[0][1]

		for x in range(len(dataimp)):

			if marca!=str(dataimp[x][2]):
				diasT.append(marca[8:10])
				horasT.append(horasacum)
				marca=str(dataimp[x][2])
				horasacum=datetime.strptime('00:00:00', '%H:%M:%S')

				if calendar.weekday(int(self.ListaAno.get()), int(meses[self.meses.index(self.ListaMes.get())]), int(marca[8:10])) != 5:
					horasLaborales+=horasLaV
				else:
					horasLaborales+=horasSab

			horasacum= horasacum + timedelta(seconds=int(dataimp[x][1][6:8]), minutes=int(dataimp[x][1][3:5]), hours=int(dataimp[x][1][0:2]))

		for x in range(len(horasT)):
			horasT[x]=int(str(horasT[x])[12])+int(str(horasT[x])[14:16])/60

		self.fig.clear() 
		self.fig2.clear()
		if diasT!=[]:
			
			self.fig.add_subplot(111).plot(diasT, horasT, marker = 'o')
			self.canvas.draw()

			self.fig2.add_subplot(111).pie([sum(horasT),horasLaborales - sum(horasT)], labels=["Tiempo Productivo","Tiempo Improductivo"],colors= ["Green","Red"], autopct="%1.1f %%")
			self.canvas2.draw()

			self.DatCabecera = [self.ListaTmk.get(),self.ListaMes.get(),self.ListaAno.get(),diasT,horasT]
			self.botonExpor["state"]="normal"
		else:
			self.botonExpor["state"]="disabled"
			messagebox.showerror("No existen datos","Este Telemarketer no trabajo en el mes de %s"%(self.ListaMes.get(),))

	def ExportarReporteH(self):
		f = filedialog.asksaveasfile(mode='a',defaultextension=".xlsx", filetypes=[("Ficheros Excel","*.xlsx")])
		if f is None: # asksaveasfile return `None` if dialog closed with "cancel".
			return

		try:
			libro = xlsxwriter.Workbook(f.name)
			hoja = libro.add_worksheet()
			
			Cabecera= ["Telemarketer","Mes","Año"]

			Data=["Dia","Horas"]
			
			for x in range(len(Cabecera)):
				hoja.write(0, x, Cabecera[x])

			for x in range(3):
				hoja.write(1, x, self.DatCabecera[x])

			for x in range(2):
				hoja.write(3, x, Data[x])

			for y in range(len(self.DatCabecera[3])):
				for x in range(2):
					hoja.write(y+4, x, self.DatCabecera[x+3][y])

			libro.close()
		except:
			messagebox.showerror("No se pudo exportar el Reporte","Se generó un error en la exportación de los datos")


if __name__ == '__main__':
	raizUser = tk.Tk()
	InicioSesion = Principal(raizUser,tk)
	raizUser.mainloop()

