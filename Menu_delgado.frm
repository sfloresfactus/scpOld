VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Menu 
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Control de Producción"
   ClientHeight    =   4125
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8160
   Icon            =   "Menu_delgado.frx":0000
   ScaleHeight     =   4125
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3750
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   7320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Menu Mnu 
      Caption         =   "Archivos &Maestros"
      Index           =   1
      Begin VB.Menu Mnu1 
         Caption         =   "Mnu1"
         Index           =   1
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "&Gerencia"
      Index           =   2
      Begin VB.Menu Mnu2 
         Caption         =   "Mnu2"
         Index           =   1
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "&Planificación"
      Index           =   3
      Begin VB.Menu Mnu3 
         Caption         =   "Mnu3"
         Index           =   1
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "&Inspección"
      Index           =   4
      Begin VB.Menu Mnu4 
         Caption         =   "Mnu4"
         Index           =   1
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "&Despacho"
      Index           =   5
      Begin VB.Menu Mnu5 
         Caption         =   "Mnu5"
         Index           =   1
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "&Adquisiciones"
      Index           =   6
      Begin VB.Menu Mnu6 
         Caption         =   "Mnu6"
         Index           =   1
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "I&nformes"
      Index           =   7
      Begin VB.Menu Mnu7 
         Caption         =   "Mnu7"
         Index           =   1
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "&Utilitarios"
      Index           =   8
      Begin VB.Menu Mnu8 
         Caption         =   "Mnu8"
         Index           =   1
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "Mantenciones"
      Index           =   9
      Begin VB.Menu Mnu9 
         Caption         =   "Mnu9"
         Index           =   1
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "Arco Sumergido"
      Index           =   10
      Begin VB.Menu Mnu10 
         Caption         =   "Mnu10"
         Index           =   1
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "A&dministración"
      Index           =   11
      Begin VB.Menu Mnu11 
         Caption         =   "Mnu11"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private i As Integer, j As Integer
'///////////// rutinas para cerrar programa /////////////////////////
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'////////////////////////////////////////////////////////////////////

Private Sub Form_Activate()

If Usuario.ObrasTerminadas Then
    StatusBar.Panels(1).Text = "OBRAS TERMINADAS"
Else
    StatusBar.Panels(1).Text = "OBRAS EN PROCESO"
End If

If Usuario.Nv_Activas Then
    StatusBar.Panels(2).Text = "NV: Solo Activas"
Else
    StatusBar.Panels(2).Text = "NV: Todas"
End If

StatusBar.Panels(3).Text = "Usuario: " & Usuario.nombre

'StatusBar.Panels(3).Text = Drive_Server & " " &path_mdb

End Sub
Private Sub Form_Load()

'Dim mf As Date
'mf = #3/21/2007#
'For i = 1 To 41
'    mf = mf + 7
'    Debug.Print mf, fecha2semana(mf)
'Next
'End

'ShortCut "copia_scp.exe", "Actualiza Scp" ' crea acceso directo a copia_scp.exe
'ShortCut "scp.exe", "Scp"
'ShortCut "scc.exe", "Scc"

c128_Poblar
densidad_poblar

If testeo Then
    MsgBox "c128_poblar"
End If

Crea_Repo

If testeo Then
    MsgBox "crea_repo"
End If
'MsgBox "2"

Mnu(1).Caption = "Archivos &Maestros"
Mnu1(1).Caption = "&Clientes"
Load Mnu1(2)
Mnu1(2).Caption = "Pro&veedores"
Load Mnu1(3)
Mnu1(3).Caption = "&Locales de Proveedores"
Load Mnu1(4)
Mnu1(4).Caption = "C&lasificación de Proveedores"
Load Mnu1(5)
Mnu1(5).Caption = "Pro&ductos"
Load Mnu1(6)
Mnu1(6).Caption = "&Tipos de Producto"
Load Mnu1(7)
Mnu1(7).Caption = "&Contratistas"
Load Mnu1(8)
Mnu1(8).Caption = "C&hoferes"
Load Mnu1(9)
Mnu1(9).Caption = "&Patentes"
Load Mnu1(10)
Mnu1(10).Caption = "T&rabajadores"
Load Mnu1(11)
Mnu1(11).Caption = "ChekList->Areas"
Load Mnu1(12)
Mnu1(12).Caption = "Tablas" ' por ejemplo:
Load Mnu1(13)
Mnu1(13).Caption = "Cuentas Contables"
Load Mnu1(14)
Mnu1(14).Caption = "Centros de Costo"

Mnu(2).Caption = "&Gerencia"
Mnu2(1).Caption = "&Notas de Venta"
Load Mnu2(2)
Mnu2(2).Caption = "-"
Load Mnu2(3)
Mnu2(3).Caption = "Informe &Notas de Venta"
Load Mnu2(4)
Mnu2(4).Caption = "-"
Load Mnu2(5)
Mnu2(5).Caption = "&Clasificación de Contratistas"
Load Mnu2(6)
Mnu2(6).Caption = "-"
Load Mnu2(7)
Mnu2(7).Caption = "&Usuarios Web"
Load Mnu2(8)
Mnu2(8).Caption = "&Actualizar datos en la Web"

Mnu(3).Caption = "&Planificación"
Mnu3(1).Caption = "&Recepción Planos"
Load Mnu3(2)
Mnu3(2).Caption = "&Digitar Planos"
Load Mnu3(3)
Mnu3(3).Caption = "&Importación Masiva Planos (IPE)"
Load Mnu3(4)
Mnu3(4).Caption = "Importación Pie&zas (Mettal)" ' nueva 11/08/10
'Mnu3(4).Caption = "&Exportación Masiva Planos"
Load Mnu3(5)
Mnu3(5).Caption = "&Historico de Revisiones de Planos"
Load Mnu3(6)
Mnu3(6).Caption = "Importar Planos desde Excel"
Load Mnu3(7)
Mnu3(7).Caption = "&Cambiar Plano a otra Obra"
Load Mnu3(8)
Mnu3(8).Caption = "-"
Load Mnu3(9)
Mnu3(9).Caption = "OT &Fabricación"
Load Mnu3(10)
Mnu3(10).Caption = "OT &Especial"
Load Mnu3(11)
Mnu3(11).Caption = "-"
Load Mnu3(12)
Mnu3(12).Caption = "Informe Recepción Planos"
Load Mnu3(13)
Mnu3(13).Caption = "Informe &Planos" ' General y Detalle
Load Mnu3(14)
Mnu3(14).Caption = "Informe OT &Fabricación"
Load Mnu3(15)
Mnu3(15).Caption = "Informe OT &Especial"
Load Mnu3(16)
Mnu3(16).Caption = "Informe de Piezas"
Load Mnu3(17)
Mnu3(17).Caption = "Informe OTf x Planos" ' General y Detalle

Mnu(4).Caption = "&Inspección"

'Mnu4(1).Caption = "ITO &Reproceso"
Mnu4(1).Caption = "-"
'Mnu4(1).Enabled = False
Load Mnu4(2)
Mnu4(2).Caption = "ITO &Fabricación"
Load Mnu4(3)
Mnu4(3).Caption = "ITO &Pintura"
Load Mnu4(4)
Mnu4(4).Caption = "ITO G&ranallado"
Load Mnu4(5)
Mnu4(5).Caption = "ITO Granallado Especial"
Load Mnu4(6)
Mnu4(6).Caption = "ITO Producción Pin&tura"
Load Mnu4(7)
Mnu4(7).Caption = "ITO Producción Pintura Especial"

Load Mnu4(8)
Mnu4(8).Caption = "ITO &Especial"
'Mnu4(8).visible = False
'Mnu4(8).Enabled = False

Load Mnu4(9)
Mnu4(9).Caption = "-"
Load Mnu4(10)
Mnu4(10).Caption = "Informe ITO &Fabricación"
Load Mnu4(11)
Mnu4(11).Caption = "Informe ITO Reproceso" 'Galvanizado" '&Arenado"
Load Mnu4(12)
Mnu4(12).Caption = "Informe ITO &Pintura"
Load Mnu4(13)
Mnu4(13).Caption = "Informe ITO Granallado"
Load Mnu4(14)
Mnu4(14).Caption = "Informe ITO Granallado Especial"
Load Mnu4(15)
Mnu4(15).Caption = "Informe ITO Produccion Pintura"
Load Mnu4(16)
Mnu4(16).Caption = "Informe ITO Produccion Pintura Especial"
Load Mnu4(17)
Mnu4(17).Caption = "Informe ITO &Especial"
Load Mnu4(18)
Mnu4(18).Caption = "Informe ITOf x OTf" ' General y Detalle
Load Mnu4(19)
Mnu4(19).Caption = "Informe de Pintura"
Load Mnu4(20)
Mnu4(20).Caption = "Granallado Pendiente" ' para juan clavero y jimmy diaz 14/07/12
Load Mnu4(21)
' para rodrigo nuñez 14/05/13
' ahora es de sRuz 01/07/13
Mnu4(21).Caption = "Protocolo Pintura (Listado de Elementos Pintados)"
'Load Mnu4(22)
'Mnu4(22).Caption = "Protocolo Pintura (Condiciones Ambientales)"

Mnu(5).Caption = "&Despacho"
Mnu5(1).Caption = "&Guía de Despacho"
Load Mnu5(2)
Mnu5(2).Caption = "&Digita Nº Factura x Guía"
Load Mnu5(3)
'Mnu5(3).Caption = "&Servicios Externos"
Mnu5(3).Caption = "&Bultos" ' 24/10/06
Load Mnu5(4)
Mnu5(4).Caption = "&Informe de Guías de Despacho"
Load Mnu5(5)
Mnu5(5).Caption = "Informe &Correlativo de Guías"
Load Mnu5(6)
Mnu5(6).Caption = "Informe de Bultos" ' 23/01/13
Load Mnu5(7)
Mnu5(7).Caption = "-"
Load Mnu5(8)
Mnu5(8).Caption = "&Facturas de Venta"
Load Mnu5(9)
Mnu5(9).Caption = "Informe de &Facturas"

Load Mnu5(10)
Mnu5(10).Caption = "-"
Load Mnu5(11)
Mnu5(11).Caption = "&Servicios Externos"
Load Mnu5(12)
Mnu5(12).Caption = "&Informe Servicios Externos"
Load Mnu5(13)
Mnu5(13).Caption = "-"
Load Mnu5(14)
Mnu5(14).Caption = "Entradas de Pintura"
Load Mnu5(15)
Mnu5(15).Caption = "Salidas de Pintura"
Load Mnu5(16)
Mnu5(16).Caption = "Informe Entradas de Pintura"
Load Mnu5(17)
Mnu5(17).Caption = "Informe Salidas de Pintura"

'Load Mnu5(9)
'Mnu5(9).Caption = "-"
'Load Mnu5(10)
'Mnu5(10).Caption = "Guías de &Ingreso"
'Mnu5(10).Caption = "Informe Servicios &Externos"

Mnu(6).Caption = "&Adquisiciones"
Mnu6(1).Caption = "&Orden de Compra"
Load Mnu6(2)
Mnu6(2).Caption = "Orden de Compra &Especial"
Load Mnu6(3)
Mnu6(3).Caption = "&Genera OC Especial (desde Excel)" ' nueva 05/08/05
'Mnu6(3).Caption = "-"
Load Mnu6(4)
Mnu6(4).Caption = "&Recepción de Materiales"
Load Mnu6(5)
'Mnu6(5).Caption = "&Certificados por Recibir"
Mnu6(5).Caption = "-"
Load Mnu6(6)
Mnu6(6).Caption = "-"
Load Mnu6(7)
Mnu6(7).Caption = "&Informe de Órdenes de Compra"
Load Mnu6(8)
Mnu6(8).Caption = "Informe Recepción de &Materiales"
Load Mnu6(9)
Mnu6(9).Caption = "Informe &Productos por Recibir"
Load Mnu6(10)
Mnu6(10).Caption = "Historial Oc"
Load Mnu6(11)
Mnu6(11).Caption = "Informe Certificados por Recibir"
Load Mnu6(12)
Mnu6(12).Caption = "Busca OC Especial"
'Load Mnu6(11)
'Mnu6(11).Caption = "-"
'Load Mnu6(12)
'Mnu6(12).Caption = "Corte Archivo &Histórico"
'Load Mnu6(13)
'Mnu6(13).Caption = "Escoge Archivo"
'Load Mnu6(14)
'Mnu6(14).Caption = "-"
Load Mnu6(13)
Mnu6(13).Caption = "&Vale de Consumo"
Load Mnu6(14)
Mnu6(14).Caption = "Informe Vale de Consumo x Contratista"
Load Mnu6(15)
Mnu6(15).Caption = "Informe Vale de Consumo x Trabajador"
Load Mnu6(16)
Mnu6(16).Caption = "Informe Vale de Consumo x Fecha"
Load Mnu6(17)
Mnu6(17).Caption = "Digita Vales Consumo Facturados"
Load Mnu6(18)
Mnu6(18).Caption = "Toma de Inventario"
Load Mnu6(19)
Mnu6(19).Caption = "-"
Load Mnu6(20)
Mnu6(20).Caption = "Digitación Chek List"
Load Mnu6(21)
Mnu6(21).Caption = "Informe Check List"
Load Mnu6(22)
Mnu6(22).Caption = "Informe Observaciones Check List"
Load Mnu6(23)
Mnu6(23).Caption = "Informe Final Check List"
Load Mnu6(24)
Mnu6(24).Caption = "Informe OC Dinámico"

Mnu(7).Caption = "I&nformes"
Mnu7(1).Caption = "Movimientos de una &Marca"
Load Mnu7(2)
Mnu7(2).Caption = "&Avances de Pago"
Load Mnu7(3)
Mnu7(3).Caption = "Piezas &Pendientes Ordenadas x Plano"

Load Mnu7(4) ' nueva 22/04/05
Mnu7(4).Caption = "Piezas &Pendientes Ordenadas x Descr"

Load Mnu7(5)
Mnu7(5).Caption = "Piezas por &Fabricar" ' por asignar
Load Mnu7(6)
Mnu7(6).Caption = "Piezas en Fa&bricación" ' por Recibir"
Load Mnu7(7)
Mnu7(7).Caption = "Piezas Fabricadas en &Negro" ' nuevo sep 2005
Load Mnu7(8)
Mnu7(8).Caption = "Piezas Fabricadas (Detalle)" ' 18/11/06
Load Mnu7(9)
'Mnu7(9).Caption = "Piezas en Galvanizado" ' 04/02/07
Mnu7(9).Caption = "Piezas en Reproceso" ' solo cmabio de nombre
Load Mnu7(10)
'Mnu7(10).Caption = "Piezas &Pintadas o Galvanizadas" ' por Despachar
Mnu7(10).Caption = "Piezas &Pintadas o Reprocesadas"
Load Mnu7(11)
'Mnu7(9).Caption = "Eti&quetas Piezas Pintadas o Galvanizadas" ' por Despachar"
Mnu7(11).Caption = "Fabricacion y Despacho"
Load Mnu7(12)
Mnu7(12).Caption = "Piezas &Despachadas"
Load Mnu7(13)
Mnu7(13).Caption = "Piezas &Despachadas (Detalle)"
Load Mnu7(14)
Mnu7(14).Caption = "&Bono de Producción"
Load Mnu7(15)
Mnu7(15).Caption = "Producción Mensual"
Load Mnu7(16)
'Mnu7(16).Caption = "-"
Mnu7(16).Caption = "ITO Fab Valorizada"
Load Mnu7(17)
Mnu7(17).Caption = "&Informe Kilos por Obra"
Load Mnu7(18)
Mnu7(18).Caption = "&Informe General NV"
Load Mnu7(19)
Mnu7(19).Caption = "&NV x Cliente"
Load Mnu7(20)
Mnu7(20).Caption = "Registro NO Conformidad"
Load Mnu7(21)
Mnu7(21).Caption = "Informe NO Conformidad"
Load Mnu7(22)
Mnu7(22).Caption = "Impresion de Etiquetas" ' para contratistas
'Load Mnu7(22)
'Mnu7(22).Caption = "Informe NO Conformidades Abiertas"
'Load Mnu7(18)
'Mnu7(18).Caption = "-"
'Load Mnu7(19)
'Mnu7(19).Caption = "&Kardex"
'Load Mnu7(20)
'Mnu7(20).Caption = "Inventario &Valorizado"

Mnu(8).Caption = "&Configuración"
Mnu8(1).Caption = "Configuración &Impresora"
Load Mnu8(2)
Mnu8(2).Caption = "&Selección de Obras"
Load Mnu8(3)
Mnu8(3).Caption = "&DesActivación de Obras"
Load Mnu8(4)
Mnu8(4).Caption = "&Parámetros"
Load Mnu8(5)
Mnu8(5).Caption = "&Cambio Clave Usuario SCP"
Load Mnu8(6)
Mnu8(6).Caption = "&Cambio Clave Usuario NC"

Mnu(9).Caption = "Mantenciones"
Mnu9(1).Caption = "Centros de Costo"
Load Mnu9(2)
Mnu9(2).Caption = "OT Mantención"
Load Mnu9(3)
Mnu9(3).Caption = "Informe OT Mantención"

Mnu(10).Caption = "Control de Vigas"
Mnu10(1).Caption = "Digitacion Arco Sumergido"
Load Mnu10(2)
Mnu10(2).Caption = "-"
Load Mnu10(3)
Mnu10(3).Caption = "Informe Arco Sumergido"
Load Mnu10(4)
Mnu10(4).Caption = "Resumen de Piezas por Turno"
Load Mnu10(5)
Mnu10(5).Caption = "Informe Bono Arco Sumergido"
Load Mnu10(6)
Mnu10(6).Caption = "Tabla Bono Arco Sumergido"
Load Mnu10(7)
Mnu10(7).Caption = "-"
Load Mnu10(8)
Mnu10(8).Caption = "Importación de Vigas"

Mnu(11).Caption = "&Utiles"
Mnu11(1).Caption = "&Compactar Archivos"
Load Mnu11(2)
Mnu11(2).Caption = "Recalcula &Planos"
Load Mnu11(3)
Mnu11(3).Caption = "Recalcula &Documentos"
Load Mnu11(4)
Mnu11(4).Caption = "&Usuarios"
Load Mnu11(5)
Mnu11(5).Caption = "U&tilitario IP"
'Load Mnu11(6)
'Mnu11(6).Caption = "Envia email"
'Mnu11(6).Caption = "&Importa Lista de Materiales" ' para Sebastian 22/08/06
Load Mnu11(6)
Mnu11(6).Caption = "&Imprime Etiquetas Pascua Lama" ' JLGonzalez-RodrigoNuñez 02/03/11

Load Mnu11(7)
Mnu11(7).Caption = "&Clave Acceso para GD Especial"

'MsgBox "3"

Login.Exitoso = False

'MsgBox "4"

Main

If testeo Then
    MsgBox "main"
End If

'MsgBox "5"

'Me.Show 1
If Login.Exitoso = False Then
    Unload Me
    Exit Sub
End If

'MsgBox "6"

Privilegios

If testeo Then
    MsgBox "privilegios"
End If

'PrivilegiosImprimir

'cierra programa bat
'EndAllInstances "cmd.exe"
'EndAllInstances "scp"

If False Then

    Track_Registrar "TEST", 4, "Tst"
    
End If

' cheque scp.ini
Dim FechaIni As String
FechaIni = ReadIniValue(Path_Local & "scp.ini", "Default", "version")
If Len(FechaIni) < 6 Then
    MsgBox "Version scp.ini obsoleta, ver: " & FechaIni, vbCritical, "Error"
    Unload Me
    Exit Sub
End If
' ultima version 14/11/16
If FechaIni < "161114" Then
    MsgBox "Version scp.ini obsoleta, ver: " & FechaIni, vbCritical, "Error"
    Unload Me
    Exit Sub
End If

' carga centros de costo en arreglo publico
centrosCostoTotal = centrosCostoCargar(aCeCo)

' carga cuentas contables en arreglo publico
cuentasContablesTotal = cuentasContablesCargar(aCuCo)

scpNew_NVLeer
scpNew_CentroCostoLeer

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim prg As String, cs As String, p1 As Integer, p2 As Integer
'prg = App.EXEName
prg = "scp"

cs = CnxSqlServer_scp0.ConnectionString
p1 = InStr(1, cs, "Data Source")
p2 = InStr(p1, cs, ";")
If p2 > p1 Then
    cs = Mid(cs, p1 + 12, p2 - p1 - 12)
End If

If KeyCode = vbKeyF4 Then

'    MsgBox "Drive: " & Drive_Server & vbLf & "Path: " & Path_Mdb & vbLf & Archivo_Fecha(App.Path & "\", prg & ".exe") & vbLf & "Usuario: " & Usuario.nombre & vbLf & "App Path: " & App.Path & vbLf & _
'     vbLf & "Servidor SQL:" & vbLf & _
'     "Proveedor: " & CnxSqlServer.Provider & vbLf & _
'     "Base de Datos SCP: " & CnxSqlServer.DefaultDatabase & vbLf & _
'     "Base de Datos SCP New: " & CnxSqlServerScpNew.DefaultDatabase & vbLf & _
'     "Data Source: " & cs
    
    MsgBox "Drive: " & Drive_Server & vbLf & _
    "Path: " & Path_Mdb & vbLf & _
     Archivo_Fecha(App.Path & "\", prg & ".exe") & vbLf & _
    "Usuario: " & Usuario.nombre & vbLf & "App Path: " & App.Path & vbLf & _
     vbLf & "Servidor SQL:" & vbLf & _
     "Proveedor: " & CnxSqlServer_scp0.Provider & vbLf & _
     "Base de Datos SCP: " & CnxSqlServer_scp0.DefaultDatabase & vbLf & _
     "Data Source: " & cs

'    Debug.Print CnxSqlServer.Attributes
'    Debug.Print CnxSqlServer.CommandTimeout
'    Debug.Print CnxSqlServer.ConnectionString
'    Debug.Print CnxSqlServer.ConnectionTimeout
'    Debug.Print CnxSqlServer.CursorLocation
'    Debug.Print CnxSqlServer.DefaultDatabase ' me interesa
'    Debug.Print CnxSqlServer.IsolationLevel
'    Debug.Print CnxSqlServer.Mode
'    Debug.Print CnxSqlServer.Properties.Count
'    Debug.Print CnxSqlServer.Provider ' me interesa
'    Debug.Print CnxSqlServer.State
'    Debug.Print CnxSqlServer.Version
    
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
DataBases_Cerrar
'If Usuario.indice <> 0 Then Login_Registrar Usuario.indice
End Sub
Private Sub Mnu1_Click(Index As Integer)
' Archivos Maestros

Privilegios_Setear 1, Index

Select Case Index
Case 1

    'OCcertificadosRecibidos
    'GdDescuadrados 3202

    'planoMarcaModificar
    'procedimientoAlmacendo5
    'proveedoresAgregar
    'contratistasAgregar

'    Email_Generar "sflores@emdelgado.cl", "2847 Prueba", "OC10449", "B", "PL6", "PLANCHA10MM", "10", "1,1"
'    Email_Generar "sflores@factus.cl", "2847 Prueba", "OC10449", "B", "PL6", "PLANCHA10MM", "10", "1,1"
'    Email_Generar "sflores.factus@gmail.com", "2847 Prueba", "OC10449", "B", "PL6", "PLANCHA10MM", "10", "1,1"

'     Archivos2Nc

'Word_Leer.Show 1

'    Repo_ProdxDesc "GD"
'    Repo_OCResumen

'    Repo_Kardex "PLACA CCMT", "", "01/01/06"

'PascuaLama_Etiquetas_Importar.Show 1

' scp_util.ProdxDesp

    ' deja todos los documentos con certificado recibido
    'Dim db As Database
    'Set db = OpenDatabase(Madq_file)
    'db.Execute "UPDATE documentos SET certificadoRecibido=1"

'Dim a(99) As Integer

'certificadosNoRecibidos

   'RepoEnProceso

   MousePointer = vbHourglass
   Load Clientes
   MousePointer = vbDefault
   Clientes.Show 1

'conectar_sqlserver.Show 1

Case 2
    MousePointer = vbHourglass
    Load Proveedores
    MousePointer = vbDefault
    Proveedores.Show 1
Case 3
    MousePointer = vbHourglass
    Load ProvLoc_Gral
    MousePointer = vbDefault
    ProvLoc_Gral.Show 1
Case 4 ' clasificacion de proveedores
    MousePointer = vbHourglass
'    Load Productos
    MousePointer = vbDefault
'    Productos.Show 1
Case 5
    MousePointer = vbHourglass
    Load Productos
    MousePointer = vbDefault
    Productos.Show 1
Case 6
    MousePointer = vbHourglass
    Load TipoProducto
    MousePointer = vbDefault
    TipoProducto.Show 1
Case 7
    MousePointer = vbHourglass
    Load sql_contratistas
    MousePointer = vbDefault
    sql_contratistas.Show 1
Case 8
    MousePointer = vbHourglass
    Tablas.Tipo = "CHOFER"
    Load Tablas
    MousePointer = vbDefault
    Tablas.Show 1
Case 9
    MousePointer = vbHourglass
    Tablas.Tipo = "PATENTE"
    Load Tablas
    MousePointer = vbDefault
    Tablas.Show 1
    
Case 10
    
    MousePointer = vbHourglass
    Load Trabajadores
    MousePointer = vbDefault
    Trabajadores.Show 1
    
Case 11

    MousePointer = vbHourglass
    Load Areas
    MousePointer = vbDefault
    Areas.Show 1

Case 12

    MousePointer = vbHourglass
'    sql_TablasMaestras.TipoTabla = "CCO" ' centros de costo
    sql_TablasMaestras.TipoTabla = "GER" ' gerencia (para no conformidades)
    Load sql_TablasMaestras
    MousePointer = vbDefault
    sql_TablasMaestras.Show 1

Case 13

    MousePointer = vbHourglass
    sql_TablasMaestras.TipoTabla = "CUCO" ' cuentas contables
    Load sql_TablasMaestras
    MousePointer = vbDefault
    sql_TablasMaestras.Show 1

Case 14

    MousePointer = vbHourglass
    sql_TablasMaestras.TipoTabla = "CECO" ' centros de costo
    Load sql_TablasMaestras
    MousePointer = vbDefault
    sql_TablasMaestras.Show 1

End Select
End Sub
Private Sub Mnu2_Click(Index As Integer)
' Gerencia

ReportDef

mpro_file = Movs_Path(Empresa.rut, Usuario.ObrasTerminadas)
Privilegios_Setear 2, Index

Select Case Index
Case 1



'    MousePointer = vbHourglass
'    Load Empresa_Escoger
'    MousePointer = vbDefault
'    Empresa_Escoger.Show 1
'    If Empresa_Escoger.Eligio Then
        MousePointer = vbHourglass
        Load Nv
        MousePointer = vbDefault
        Nv.Show 1
'    End If
Case 3

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "NV"
    Report_Def.Op_Cliente = True
    
'    Report_Def.Op_Proveedor = True
'    Report_Def.Op_ProveedorTipo = True
'    Report_Def.Op_Numero = True
    
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault

'    MousePointer = vbHourglass
'    Load Empresa_Escoger
'    MousePointer = vbDefault
'    Empresa_Escoger.Show 1
'    If Empresa_Escoger.Eligio Then

        
'    End If

Case 5

    MousePointer = vbHourglass
    Load SubC_Clasificacion
    MousePointer = vbDefault
    SubC_Clasificacion.Show 1
    
Case 7

    ' usuarios web
    MousePointer = vbHourglass
'    Load UsuariosWeb
    MousePointer = vbDefault
'    UsuariosWeb.Show 1
    
Case 8

    If Usuario.ObrasTerminadas Then
    
        MsgBox "Debe Seleccinar Obras En Proceso", , "ATENCIÓN"
        
    Else

        MousePointer = vbHourglass
'        Load Scp2Inet
        MousePointer = vbDefault
'        Scp2Inet.Show 1
    
    End If
    
End Select
End Sub
Private Sub Mnu3_Click(Index As Integer)

' Planificación

ReportDef

mpro_file = Movs_Path(Empresa.rut, Usuario.ObrasTerminadas)
Privilegios_Setear 3, Index

Select Case Index
Case 1
    MousePointer = vbHourglass
    Load PlanosRecepcion
    MousePointer = vbDefault
    PlanosRecepcion.Show 1
Case 2
    MousePointer = vbHourglass
    Load Plano_Dig
    MousePointer = vbDefault
    Plano_Dig.Show 1
Case 3
    Planos_ImportacionMasiva.Show 1
Case 4
    piezas_importar.Show 1
Case 5
'    Planos_ExportacionMasiva.Show 1

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "REVISIONES DE PLANOS"
    
    Report_Def.Op_NotaVenta = True
'    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault

Case 6
    MousePointer = vbHourglass
    Load planosEnExcelImportar
    MousePointer = vbDefault
    planosEnExcelImportar.Show 1
Case 7
    MousePointer = vbHourglass
    Load PlanoMover
    MousePointer = vbDefault
    PlanoMover.Show 1
Case 9
    MousePointer = vbHourglass
    Load OT_Fabricacion2
    MousePointer = vbDefault
    OT_Fabricacion2.Show 1
Case 51
    OT_Opcion.Titulo = "OT Arenado"
    OT_Opcion.Show 1
    Select Case OT_Opcion.TipoOT
    Case "N" ' ot arenado normal
        MousePointer = vbHourglass
'        OT_Arenado.xPlano = False
'        Load OT_Arenado
        MousePointer = vbDefault
'        OT_Arenado.Show 1
    Case "P" ' ot arenado x plano
        MousePointer = vbHourglass
'        OT_Arenado.xPlano = True
'        Load OT_Arenado
        MousePointer = vbDefault
'        OT_Arenado.Show 1
    End Select
Case 61
    OT_Opcion.Titulo = "OT Pintura"
    OT_Opcion.Show 1
    Select Case OT_Opcion.TipoOT
    Case "N" ' ot pintura normal
        MousePointer = vbHourglass
'        OT_Pintura.xPlano = False
'        Load OT_Pintura
        MousePointer = vbDefault
'        OT_Pintura.Show 1
    Case "P" ' ot pintura x plano
        MousePointer = vbHourglass
'        OT_Pintura.xPlano = True
'        Load OT_Pintura
        MousePointer = vbDefault
'        OT_Pintura.Show 1
    End Select
Case 10

    MousePointer = vbHourglass
    Load OT_Especial
    MousePointer = vbDefault
    OT_Especial.Show 1
    
Case 12 ' Recepcion de Planos

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "RECEPCION DE PLANOS"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 13 ' Planos (General) o (Detalle)

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "PLANOS"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_TipoRepo = True
    'Report_Def.Op_Plano = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 14 ' OTf (General o Detalle)

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "OT FABRICACIÓN"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Fecha = True
    Report_Def.Op_TipoRepo = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 111 ' OTa

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "OT ARENADO"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 121 ' OTp

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "OT PINTURA"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 15 ' OT especial

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "OT ESPECIAL"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 16

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "Informe de Piezas"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
'    Report_Def.Op_TipoRepo = True
'    Report_Def.Op_Plano = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 17

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "OTf x Plano"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_TipoRepo = True
    'Report_Def.Op_Plano = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
End Select
End Sub
Private Sub Mnu4_Click(Index As Integer)
' Inspección

ReportDef

mpro_file = Movs_Path(Empresa.rut, Usuario.ObrasTerminadas)
Privilegios_Setear 4, Index

Select Case Index

Case 1 ' reproceso ( va despues de ot, pero es opcional)
    MousePointer = vbHourglass
    sql_ito_pg.TipoDoc = "IRP" ' nuevo tipo
    Load sql_ito_pg
    MousePointer = vbDefault
    sql_ito_pg.Show 1
Case 2
    MousePointer = vbHourglass
    Load ITO_Fabricacion
    MousePointer = vbDefault
    ITO_Fabricacion.Show 1
'Case 2 ' galvanizado
'    MousePointer = vbHourglass
'    ITO_PG.TipoDoc = "G"
'    Load ITO_PG
'    MousePointer = vbDefault
'    ITO_PG.Show 1
Case 3
    MousePointer = vbHourglass
    ITO_PG.TipoDoc = "P" ' pintura
    Load ITO_PG
    MousePointer = vbDefault
    ITO_PG.Show 1
Case 4
    MousePointer = vbHourglass
    ITO_PG.TipoDoc = "R" ' granallado
    Load ITO_PG
    MousePointer = vbDefault
    ITO_PG.Show 1
Case 5
    MousePointer = vbHourglass
    ITO_PG_Esp.TipoDoc = "S" ' granallado especial
    Load ITO_PG_Esp
    MousePointer = vbDefault
    ITO_PG_Esp.Show 1
Case 6
    MousePointer = vbHourglass
    ITO_PG.TipoDoc = "T" ' produccion pintura
    Load ITO_PG
    MousePointer = vbDefault
    ITO_PG.Show 1
Case 7
    MousePointer = vbHourglass
    ITO_PG_Esp.TipoDoc = "U" ' produccion pintura especial
    Load ITO_PG_Esp
    MousePointer = vbDefault
    ITO_PG_Esp.Show 1
Case 8
    MousePointer = vbHourglass
    Load ITO_Especial
    MousePointer = vbDefault
    ITO_Especial.Show 1
    
Case 10 ' ITO Fab

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "ITO FABRICACIÓN"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_CentroCosto = True ' 24/09/15
    Report_Def.Op_Contratista = True
    Report_Def.Op_Fecha = True
    Report_Def.Op_BotonRango = True ' 22/21
    Report_Def.Op_TipoRepo = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 11 ' ITO galvanizado

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "ITO REPROCESO"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 12 ' ITO Pin

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "ITO PINTURA"

    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Fecha = True
    Report_Def.Op_TipoRepo = True

    Report_Def.Show 1
    MousePointer = vbDefault

Case 13 ' ITO granallado

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "ITO GRANALLADO"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Fecha = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Operador = True
    Report_Def.Op_TipoGranalla = True
    Report_Def.Op_Maquina = True
    Report_Def.Op_Turno = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 14 ' ITO granallado especial

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "ITO GRANALLADO ESPECIAL"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Fecha = True
    Report_Def.Op_Operador = True
    Report_Def.Op_TipoGranalla = True
    Report_Def.Op_Maquina = True
    Report_Def.Op_Turno = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 15 ' ITO Produccion pintura

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "ITO PRODUCCION PINTURA"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Fecha = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Operador = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 16 ' ITO Produccion pintura especial

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "ITO PRODUCCION PINTURA ESPECIAL"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Fecha = True
    Report_Def.Op_Operador = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 17 ' ITO Esp

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "ITO ESPECIAL"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 18 ' ITOs x OTs General y Detalle

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "ITOf x OTf"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_TipoRepo = True
    Report_Def.Op_OT = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 19

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "INFORME DE PINTURA"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    'Report_Def.Op_Plano = True ' comentada 09/10/15
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 20

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "GRANALLADO PENDIENTE"
    
    Report_Def.Op_NotaVenta = True
    
    Report_Def.Show 1
    MousePointer = vbDefault

Case 21

    MousePointer = vbHourglass
    Load protocoloPintura
    MousePointer = vbDefault
    protocoloPintura.Show 1
    
End Select
End Sub
Private Sub Mnu5_Click(Index As Integer)
' Despacho

ReportDef

mpro_file = Movs_Path(Empresa.rut, Usuario.ObrasTerminadas)
Privilegios_Setear 5, Index

Select Case Index
Case 1
'    MousePointer = vbHourglass
'    Load Empresa_Escoger
'    MousePointer = vbDefault
'    Empresa_Escoger.Show 1
'    If Empresa_Escoger.Eligio Then
        MousePointer = vbHourglass
        Load GD
        MousePointer = vbDefault
        GD.Show 1
'    End If
Case 2

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "DIGITA FACTURAS X GUIA"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Cliente = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 3
    MousePointer = vbHourglass
'    Load SE
    Load Bultos
    MousePointer = vbDefault
'    SE.Show 1
    Bultos.Show 1
Case 4 ' GD

    MousePointer = vbHourglass
    Report_Def.Titulo = "GUÍAS DE DESPACHO"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Cliente = True
    Report_Def.Op_Fecha = True
    Report_Def.Op_TipoGD = True
    Report_Def.Op_TipoRepo = True ' agregado 01/06/2016
    
    Report_Def.Show 1
    MousePointer = vbDefault
        
'    End If

Case 5 ' correlativos GD

    MousePointer = vbHourglass
    Report_Def.Titulo = "CORRELATIVO DE GUÍAS"
    
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 6 ' Informe de Bultos

    MousePointer = vbHourglass
    Report_Def.Titulo = "BULTOS"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 8 'facturas de venta
'    MousePointer = vbHourglass
'    Load Factura
'    MousePointer = vbDefault
'    Factura.Show 1
Case 9 ' informe facturas

    MousePointer = vbHourglass
    Report_Def.Titulo = "FACTURAS DE VENTA"
    
    Report_Def.Op_Cliente = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 11

    MousePointer = vbHourglass
    Load SE
    MousePointer = vbDefault
    SE.Show 1
    
Case 12 ' informe servicio externos

    MousePointer = vbHourglass
    Report_Def.Titulo = "SERVICIOS EXTERNOS"
    
    Report_Def.Op_Cliente = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 14

    MousePointer = vbHourglass
    Load ValeBodega
    MousePointer = vbDefault
    ValeBodega.Show 1
    
End Select

End Sub
Private Sub Mnu6_Click(Index As Integer)
' Adquisiciones

ReportDef

mpro_file = Movs_Path(Empresa.rut, Usuario.ObrasTerminadas)
Privilegios_Setear 6, Index

Select Case Index
Case 1 ' OC
    MousePointer = vbHourglass
    Load Oc
    MousePointer = vbDefault
    Oc.Show 1
Case 2 ' OC Esp
    MousePointer = vbHourglass
    Load OC_Esp
    MousePointer = vbDefault
    OC_Esp.Show 1
Case 3
    MousePointer = vbHourglass
    Load oce_generar
    MousePointer = vbDefault
    oce_generar.Show 1
Case 4 ' recepción de materiales
    MousePointer = vbHourglass
    Load RecepcionMateriales
    MousePointer = vbDefault
    RecepcionMateriales.Show 1
Case 5 ' digita certificados pendientes
    
    MousePointer = vbHourglass
    Report_Def.Titulo = "RECEPCION DE CERTIFICADOS"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Proveedor = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 7 ' informe OC

    MousePointer = vbHourglass
    Report_Def.Titulo = "ÓRDENES DE COMPRA"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Producto = True
    Report_Def.Op_Proveedor = True
    Report_Def.Op_ProveedorTipo = True
    Report_Def.Op_Fecha = True
    Report_Def.Op_Numero = True
    Report_Def.Op_TipoRepo = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
        
Case 8 ' informe RM

    MousePointer = vbHourglass
    Report_Def.Titulo = "RECEPCIÓN DE MATERIALES"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Producto = True
    Report_Def.Op_Proveedor = True
    Report_Def.Op_ProveedorTipo = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 9 ' Productos por Recibir

    MousePointer = vbHourglass
    Report_Def.Titulo = "PRODUCTOS POR RECIBIR"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Producto = True
    Report_Def.Op_Proveedor = True
    Report_Def.Op_ProveedorTipo = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 10 ' historial de una oc
    MousePointer = vbHourglass
    Load Oc_Mov
    MousePointer = vbDefault
    Oc_Mov.Show 1
    
Case 11 ' inf certificados pendientes

    MousePointer = vbHourglass
    Report_Def.Titulo = "CERTIFICADOS POR RECIBIR"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Proveedor = True
    Report_Def.Op_Producto = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 110 ' recalcula oc
    MousePointer = vbHourglass
'    Load OC_Recalcular
    MousePointer = vbDefault
'    OC_Recalcular.Show 1
Case 112 ' FECHA DE CORTE
    MousePointer = vbHourglass
'    Load Oc_Corte
    MousePointer = vbDefault
'    Oc_Corte.Show 1
    
Case 12

    MousePointer = vbHourglass
    Load OCe_Buscar
    MousePointer = vbDefault
    
    OCe_Buscar.Show 1

Case 13

    MousePointer = vbHourglass
    Load ValeConsumo
    MousePointer = vbDefault
    
    ValeConsumo.Show 1

Case 14 ' informe vales de consumo (o salidas)

    MousePointer = vbHourglass
    Report_Def.Titulo = "VALES DE CONSUMO X CONTRATISTA"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Producto = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 15 ' informe vales de consumo (o salidas)

    MousePointer = vbHourglass
    Report_Def.Titulo = "VALES DE CONSUMO X TRABAJADOR"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Trabajador = True
    Report_Def.Op_Producto = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 16 ' informe vales de consumo (o salidas)

    MousePointer = vbHourglass
    Report_Def.Titulo = "VALES DE CONSUMO X FECHA"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Producto = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
        
Case 17

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "DIGITA FACTURAS X VALE CONSUMO"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 18 ' digita toma de inventario

'    MousePointer = vbHourglass
'    Load InventarioTomar
'    MousePointer = vbDefault
'    InventarioTomar.Show 1
    
Case 20 ' digita

'    MousePointer = vbHourglass
'    Load ChkLst_Mov
'    MousePointer = vbDefault
'    ChkLst_Mov.Show 1
    
Case 21

'    MousePointer = vbHourglass
'    Report_Def.Titulo = "CHECK LIST"
    
'    Report_Def.Op_Fecha = True
'    Report_Def.Op_chklstArea = True
'    Report_Def.Op_chklstResponsableArea = True
    
'    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 22

    MousePointer = vbHourglass
    Report_Def.Titulo = "CHECK LIST OBSERVACIONES"
    
    Report_Def.Op_Fecha = True
    Report_Def.Op_chklstArea = True
    Report_Def.Op_chklstResponsableArea = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 23

    MousePointer = vbHourglass
    Report_Def.Titulo = "CHECK LIST FINAL"
    
    Report_Def.Op_Fecha = True
    Report_Def.Op_chklstArea = True
    Report_Def.Op_chklstResponsableArea = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
        
Case 24 ' informe OC dinamico

    MousePointer = vbHourglass
    Report_Def.Titulo = "ÓRDENES DE COMPRA DINÁMICO"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Producto = True
    Report_Def.Op_Proveedor = True
    Report_Def.Op_ProveedorTipo = True
    Report_Def.Op_Fecha = True
    Report_Def.Op_Numero = True
    Report_Def.Op_TipoRepo = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
        
End Select
End Sub
Private Sub Mnu7_Click(Index As Integer)
' Informes

ReportDef

mpro_file = Movs_Path(Empresa.rut, Usuario.ObrasTerminadas)
Privilegios_Setear 7, Index

Select Case Index
Case 1 ' mov de marca
    MousePointer = vbHourglass
    Load Marca_Mov
    MousePointer = vbDefault
    Marca_Mov.Show 1
Case 2 'avances de pago

    MousePointer = vbHourglass
    Report_Def.Titulo = "AVANCES DE PAGO"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 3 ' piezas pendientes ordenadas por plano

    MousePointer = vbHourglass
    Report_Def.Titulo = "PIEZAS PENDIENTES ORDENADAS POR PLANO"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_PP = True
    'Report_Def.Op_Plano = True ' comentada 09/10/15
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 4 ' piezas pendientes ordenadas por descripcion

    MousePointer = vbHourglass
    Report_Def.Titulo = "PIEZAS PENDIENTES ORDENADAS POR DESCRIPCION"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_PP = True
    'Report_Def.Op_Plano = True ' la comente el 09/10/15
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 5 ' piezas por fabricar  ( asignar )

    MousePointer = vbHourglass
    Report_Def.Titulo = "PIEZAS POR FABRICAR" 'ASIGNAR"
    
    Report_Def.Op_NotaVenta = True
    'Report_Def.Op_Plano = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 6 ' piezas en fabricacion ( por recibir )

    MousePointer = vbHourglass
    Report_Def.Titulo = "PIEZAS EN FABRICACIÓN" 'POR RECIBIR"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Fecha = True
    'Report_Def.Op_Plano = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 7 ' piezas fabricadas en negro ( por despachar )

    MousePointer = vbHourglass
    Report_Def.Titulo = "PIEZAS FABRICADAS EN NEGRO" ' (ANTIGUO) POR DESPACHAR"
    
    Report_Def.Op_NotaVenta = True
    'Report_Def.Op_Plano = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 8 ' piezas fabricadas (detalle) 18/11/06

    MousePointer = vbHourglass
    Report_Def.Titulo = "PIEZAS FABRICADAS (DETALLE)"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 9 ' piezas en galvanizado

    MousePointer = vbHourglass
'    Report_Def.Titulo = "PIEZAS EN GALVANIZADO"
    Report_Def.Titulo = "PIEZAS EN REPROCESO"
    
    Report_Def.Op_NotaVenta = True
    'Report_Def.Op_Plano = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 10 ' piezas pint o galv

    MousePointer = vbHourglass
'    Report_Def.Titulo = "PIEZAS PINTADAS O GALVANIZADAS" ' por despachar
    Report_Def.Titulo = "PIEZAS PINTADAS O REPROCESADAS"
    
    Report_Def.Op_NotaVenta = True
    'Report_Def.Op_Plano = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 99 ' etiquetas piezas pint o galv

    MousePointer = vbHourglass
    Report_Def.Titulo = "ETIQUETAS PIEZAS PINTADAS O GALVANIZADAS" ' POR DESPACHAR"
    
    Report_Def.Op_NotaVenta = True
    'Report_Def.Op_Plano = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 11 ' informe fabricacion y despacho

    MousePointer = vbHourglass
    Report_Def.Titulo = "FABRICACION Y DESPACHO"
    
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 12 ' piezas despachadas

    MousePointer = vbHourglass
    Report_Def.Titulo = "PIEZAS DESPACHADAS"
    
    Report_Def.Op_NotaVenta = True
'    Report_Def.Op_Fecha = False 'para este caso no se puede true
    'Report_Def.Op_Plano = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 13 ' piezas despachadas nueva 10/10/06

    MousePointer = vbHourglass
    Report_Def.Titulo = "PIEZAS DESPACHADAS (DETALLE)"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 14 ' bono de producción

    MousePointer = vbHourglass
    Report_Def.Titulo = "BONO DE PRODUCCIÓN"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 15 ' producción mensual

    MousePointer = vbHourglass
    Report_Def.Titulo = "PRODUCCIÓN MENSUAL"
    
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 16 ' ito fab valorizada

    MousePointer = vbHourglass
    Report_Def.Titulo = "ITO FAB VALORIZADA"
    
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 17 ' junio 2004

    MousePointer = vbHourglass
    Report_Def.Titulo = "INFORME KILOS POR OBRA"
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 18 ' 04/06/04

    MousePointer = vbHourglass
    Report_Def.Titulo = "INFORME GENERAL NV"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Fecha = True 'aunque no pesca
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 19 ' 31/08/06

    MousePointer = vbHourglass
    Report_Def.Titulo = "NV X CLIENTE"
    
    Report_Def.Op_Cliente = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    MousePointer = vbDefault
    
Case 20 ' 22/11/06

'    inc_dig.Show 1
    sql_noconformidad.Show 1
    
Case 21 ' informe de no conformidad

    ' si escoge todas las gerencia => (resumen)
    ' si escoge UNA gerencia => (detalle)

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "INFORME NO CONFORMIDAD"

    Report_Def.Op_Fecha = True
'    Report_Def.Op_NCArea = True ' gerencia
    Report_Def.Op_TipoRepo = True
    Report_Def.Op_Trabajador = True
    Report_Def.Frame_TR.Caption = "Emisor"
        
    Report_Def.Show 1
    MousePointer = vbDefault

Case 22 ' impresion de etiquetas

    ' solo para contratistas
    ' cuando terminan una pieza deben imprimir etiqueta de la pieza y adherirla a la misma
    ' desde 06/11/12
    MousePointer = vbHourglass
    EtiquetasContratistasImprimir.Show 1
    MousePointer = vbDefault
    
End Select
End Sub
Private Sub Mnu8_Click(Index As Integer)
' Utilitarios

mpro_file = Movs_Path(Empresa.rut, Usuario.ObrasTerminadas)
Privilegios_Setear 8, Index

Select Case Index
Case 1
    MousePointer = vbHourglass
    Load PrinterConfig
    MousePointer = vbDefault
    PrinterConfig.Show 1
Case 2
    MousePointer = vbHourglass
    Load ObrasActivar
    MousePointer = vbDefault
    ObrasActivar.Show 1
Case 3
    MousePointer = vbHourglass
    Load ObrasDesActivar
    MousePointer = vbDefault
    ObrasDesActivar.Show 1
Case 4
    MousePointer = vbHourglass
    Load Parametros
    MousePointer = vbDefault
    Parametros.Show 1
Case 5
    MousePointer = vbHourglass
    Load ClaveCambiar
    MousePointer = vbDefault
    ClaveCambiar.Show 1
Case 6
    MousePointer = vbHourglass
    Load ClaveCambiarNC
    MousePointer = vbDefault
    ClaveCambiarNC.Show 1
End Select
End Sub
Private Sub Mnu9_Click(Index As Integer)

ReportDef

mpro_file = Movs_Path(Empresa.rut, Usuario.ObrasTerminadas)
Privilegios_Setear 9, Index

Select Case Index

Case 1

    MousePointer = vbHourglass
    TablasMaestras.TipoTabla = "CCO"
    Load TablasMaestras
    MousePointer = vbDefault
    TablasMaestras.Show 1
    
Case 2

    MousePointer = vbHourglass
    Load OT_Mantencion
    MousePointer = vbDefault
    OT_Mantencion.Show 1
    
Case 3 ' OT Mantencion

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "OT MANTENCION"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Contratista = True
    Report_Def.Op_Fecha = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
End Select
End Sub
Private Sub Mnu10_Click(Index As Integer)

ReportDef

mpro_file = Movs_Path(Empresa.rut, Usuario.ObrasTerminadas)
Privilegios_Setear 11, Index

Select Case Index

Case 1

    MousePointer = vbHourglass
    Load As_Mantencion
    MousePointer = vbDefault
    As_Mantencion.Show 1

Case 3

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "ARCO SUMERGIDO"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Operador = True
    Report_Def.Op_Fecha = True
    Report_Def.Op_Turno = True
    Report_Def.Op_TipoPieza = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault

Case 4

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "RESUMEN PIEZAS X TURNO"
    
    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Operador = True
    Report_Def.Op_Fecha = True
'    Report_Def.Op_Turno = True
'    Report_Def.Op_TipoPieza = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 5

    MousePointer = vbHourglass
    
    Report_Def.Titulo = "BONO ARCO SUMERGIDO"
    
'    Report_Def.Op_NotaVenta = True
    Report_Def.Op_Operador = True
    Report_Def.Op_Fecha = True
'    Report_Def.Op_Turno = True
'    Report_Def.Op_TipoPieza = True
    
    Report_Def.Show 1
    
    MousePointer = vbDefault
    
Case 6

    MousePointer = vbHourglass
    Load As_TablaBono
    MousePointer = vbDefault
    As_TablaBono.Show 1

Case 8

    MousePointer = vbHourglass
'    Load importacion_vigas
    MousePointer = vbDefault
'    importacion_vigas.Show 1

End Select

End Sub
Private Sub Mnu11_Click(Index As Integer)
' Administracion del Sistema

mpro_file = Movs_Path(Empresa.rut, Usuario.ObrasTerminadas)
Privilegios_Setear 11, Index

Select Case Index
Case 1
    MousePointer = vbHourglass
'    Load Compactar
    MousePointer = vbDefault
'    Compactar.Show 1
Case 2
    MousePointer = vbHourglass
    BasePlano_Recalcula
    MousePointer = vbDefault
Case 3
    MousePointer = vbHourglass
'    Load Docs_Recalcular
    MousePointer = vbDefault
'    Docs_Recalcular.Show 1
'Case 4
'    MarcasSinPlano.Show 1
Case 4
    MousePointer = vbHourglass
    Load Usuarios
    MousePointer = vbDefault
    Usuarios.Show 1
Case 5
    MousePointer = vbHourglass
'    Load Usr_Activos
    Load ip_utiles
    MousePointer = vbDefault
    ip_utiles.Show 1
'    Usr_Activos.Show 1
Case 66 ' envia email
    Dim msg As String
    msg = "<HTML><BODY><B>Hola</B> que tal</BODY></HTML>"
'    msg = "Scp" & vbCrLf
'    msg = msg & "Prueba de envio de mensaje desde VB, Scp" & vbCrLf
'    msg = msg & Date & " " & Time & vbCrLf
'    frmCorreoEnviar.Destinatario = "sflores@factus.cl"
''    frmCorreoEnviar.Destinatario = "sflores.factus@gmail.com"
'    frmCorreoEnviar.Asunto = "Scp: Atención !!!"
'    frmCorreoEnviar.Mensaje = msg
'    frmCorreoEnviar.Enviar
Case 6
    PascuaLama_Etiquetas_Imprimir.Show 1
Case 7
    MousePointer = vbHourglass
    Load Accesos
    MousePointer = vbDefault
    Accesos.Show 1
End Select
End Sub
Private Sub Privilegios()
Dim j As Integer, suma As Integer
' (des)habilita opciones de menu segun clave usuario

' menu
' habilita opcion de menu principal ssi hay opciones habilitadas en submenu
For i = 1 To Menu_Columnas
    suma = 0
    For j = 1 To Menu_Filas
        suma = suma + Val(Privi(i, j))
    Next
    Mnu(i).visible = suma
Next

' submenus
On Error GoTo Error
For j = 1 To Menu_Filas

'    Mnu1(j).Enabled = Val(Privi(1, j))
    
    Mnu1(j).visible = Val(Privi(1, j))
    Mnu2(j).visible = Val(Privi(2, j))
    Mnu3(j).visible = Val(Privi(3, j))
    Mnu4(j).visible = Val(Privi(4, j))
    Mnu5(j).visible = Val(Privi(5, j))
    Mnu6(j).visible = Val(Privi(6, j))
    Mnu7(j).visible = Val(Privi(7, j))
    Mnu8(j).visible = Val(Privi(8, j))
    Mnu9(j).visible = Val(Privi(9, j))
    Mnu10(j).visible = Val(Privi(10, j))
    Mnu11(j).visible = Val(Privi(11, j))
'    Mnu12(j).visible = Val(Privi(12, j))
    
Next
Exit Sub
Error:
Resume Next
End Sub
Private Sub Privilegios_Setear(f, c)
Dim p As Long
'readonly
p = InStr(1, mpro_file, "ScpHist", 1)
If p <> 0 Then
    ' el ScpHist es siempre de sólo lectura
    Usuario.ReadOnly = True
Else
    Usuario.ReadOnly = IIf(Val(Privi(f, c)) = 1, True, False)
End If
Usuario.AccesoTotal = IIf(Val(Privi(f, c)) = 3, True, False)
End Sub
Private Sub ReportDef()

Report_Def.Op_NotaVenta = False
Report_Def.Op_CentroCosto = False
Report_Def.Op_Contratista = False
Report_Def.Op_Cliente = False
Report_Def.Op_Trabajador = False
Report_Def.Op_Proveedor = False
Report_Def.Op_ProveedorTipo = False
Report_Def.Op_Numero = False
Report_Def.Op_Producto = False
Report_Def.Op_Fecha = False
Report_Def.Op_BotonRango = False
Report_Def.Op_TipoRepo = False
Report_Def.Op_Plano = False
Report_Def.Op_OT = False
Report_Def.Op_TipoGD = False
Report_Def.Op_chklstArea = False
Report_Def.Op_chklstResponsableArea = False
Report_Def.Op_Operador = False
Report_Def.Op_TipoGranalla = False
Report_Def.Op_Maquina = False
Report_Def.Op_TipoPieza = False
Report_Def.Op_Turno = False
Report_Def.Op_PP = False

End Sub
Private Sub PrivilegiosImprimir()
' imprime privilegios de usarios scp
Dim Dbc As Database, RsUs As Recordset, txt As String, largo As Integer, c As Integer, txtOpcion As String

'Set Dbc = OpenDatabase(Syst_file, False, False, ";pwd=eml")
'Set RsUs = Dbc.OpenRecordset("Usuarios")
'RsUs.Index = "Nombre"

On Error Resume Next

With RsUs
Do While Not .EOF

'    If UCase(!nombre) <> "ALEJANDRO" Then GoTo Siguiente
'    If UCase(!nombre) <> "ERWIN" Then GoTo Siguiente

    Debug.Print " "
    Debug.Print !nombre
    For i = 4 To 4
    
        txt = Mnu(i).Caption
        txt = Replace(txt, "&")
'        Debug.Print ""
'        Debug.Print txt
'        Debug.Print "--------------------"
        
        
        txt = RsUs("menu" & PadL(Trim(i), 2, "0"))
        largo = Len(txt)

        For c = 1 To 1 'largo
        
            txtOpcion = ""
            
            Select Case i
            Case 1
                txtOpcion = Mnu1(c).Caption
            Case 2
                txtOpcion = Mnu2(c).Caption
            Case 3
                txtOpcion = Mnu3(c).Caption
            Case 4
                txtOpcion = Mnu4(c).Caption
            Case 5
                txtOpcion = Mnu5(c).Caption
            Case 6
                txtOpcion = Mnu6(c).Caption
            Case 7
                txtOpcion = Mnu7(c).Caption
            Case 8
                txtOpcion = Mnu8(c).Caption
            Case 9
                txtOpcion = Mnu9(c).Caption
            Case 10
                txtOpcion = Mnu10(c).Caption
            End Select

            If txtOpcion <> "" And txtOpcion <> "-" And Mid(txt, c, 1) <> "" And Mid(txt, c, 1) <> "0" And Mid(txt, c, 1) <> "1" Then
            
                txtOpcion = Replace(txtOpcion, "&")
                
                Debug.Print txtOpcion; "|";
                Debug.Print Mid(txt, c, 1); "|"
            
            End If
            
        Next
        
    Next
Siguiente:
    .MoveNext
    
Loop

End With
End Sub
'////////////////////////////////////////////////////////////////////
' tres rutina necesarias para cerrar programa *.bat ejecutado para llamar a scp
'////////////////////////////////////////////////////////////////////
Public Function EndAllInstances(ByVal WindowCaption As String) As Boolean
'*********************************************
'PURPOSE: ENDS ALL RUNNING INSTANCES OF A PROCESS
'THAT CONTAINS ANY PART OF THE WINDOW CAPTION

'INPUT: ANY PART OF THE WINDOW CAPTION

'RETURNS: TRUE IF SUCCESSFUL (AT LEASE ONE PROCESS WAS STOPPED,
'FALSE OTHERWISE)

'EXAMPLE EndProcess "Notepad"

'NOTES:
'1. THIS IS DESIGNED TO TERMINATE THE PROCESS IMMEDIATELY,
'   THE APP WILL NOT RUN THROUGH IT'S NORMAL SHUTDOWN PROCEDURES
'   E.G., THERE WILL BE NO DIALOG BOXES LIKE "ARE YOU SURE
'   YOU WANT TO QUIT"

'2. BE CAREFUL WHEN USING:
'   E.G., IF YOU CALL ENDPROCESS("A"), ANY PROCESS WITH A
'   WINDOW THAT HAS THE LETTER "A" IN ITS CAPTION WILL BE
'   TERMINATED

'3. AS WRITTEN, ALL THIS CODE MUST BE PLACED WITHIN
'   A FORM MODULE

'***********************************************
Dim hWnd As Long
Dim hInst As Long
Dim hProcess As Long
Dim lProcessID
Dim bAns As Boolean
Dim lExitCode As Long
Dim lRet As Long

On Error GoTo ErrorHandler

If Trim(WindowCaption) = "" Then Exit Function
Do
hWnd = FindWin(WindowCaption)
If hWnd = 0 Then Exit Do
hInst = GetWindowThreadProcessId(hWnd, lProcessID)
'Get handle to process
hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0&, lProcessID)
If hProcess <> 0 Then
    'get exit code
    GetExitCodeProcess hProcess, lExitCode
        If lExitCode <> 0 Then
                'bye-bye
            lRet = TerminateProcess(hProcess, lExitCode)
            If bAns = False Then bAns = lRet > 0
        End If
End If
Loop

EndAllInstances = bAns
ErrorHandler:

End Function

Private Function FindWin(WinTitle As String) As Long

Dim lhWnd As Long, sAns As String
Dim sTitle As String

lhWnd = Me.hWnd
sTitle = LCase(WinTitle)

Do

   DoEvents
      If lhWnd = 0 Then Exit Do
        sAns = LCase$(GetCaption(lhWnd))
             

       If InStr(sAns, sTitle) Then

          FindWin = lhWnd
          Exit Do
       Else
         FindWin = 0
       End If

       lhWnd = GetNextWindow(lhWnd, 2)

Loop

End Function

Private Function GetCaption(lhWnd As Long) As String

Dim sAns As String, lLen As Long

   lLen& = GetWindowTextLength(lhWnd)
   sAns = String(lLen, 0)
   Call GetWindowText(lhWnd, sAns, lLen + 1)
   GetCaption = sAns

End Function
Private Sub ShortCut(Aplicacion As String, Descripcion As String)

On Error Resume Next
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
Dim copiaScp As String
'copiaScp = "copia_scp.exe"

Dim MyShortcut, MyDesktop, DesktopPath

DesktopPath = WSHShell.specialfolders("AllUsersDesktop")

Set MyShortcut = WSHShell.CreateShortcut(DesktopPath & "\" & Descripcion & ".lnk")

MyShortcut.targetpath = WSHShell.ExpandEnvironmentStrings(AppPath & Aplicacion)
MyShortcut.WorkingDirectory = WSHShell.ExpandEnvironmentStrings(AppPath)
MyShortcut.WindowStyle = 4
MyShortcut.IconLocation = WSHShell.ExpandEnvironmentStrings(AppPath & Aplicacion) & ", 0"
MyShortcut.save
Err.Clear

End Sub
Public Function AppPath() As String
    
    Dim sAns As String
    sAns = App.Path
    If Right(App.Path, 1) <> "\" Then sAns = sAns & "\"
    AppPath = sAns

End Function
