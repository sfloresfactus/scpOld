VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Trabajadores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trabajadores"
   ClientHeight    =   8910
   ClientLeft      =   3135
   ClientTop       =   3090
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar"
            Object.Tag             =   "[Agregando]"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "[Modificando]"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "[Eliminando]"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Listar"
            Object.Tag             =   "[Listando]"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Deshacer"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox correo 
      Height          =   300
      Left            =   1440
      TabIndex        =   16
      Top             =   2760
      Width           =   4335
   End
   Begin TabDlg.SSTab Tab_Datos 
      Height          =   4935
      Left            =   240
      TabIndex        =   17
      Top             =   3240
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   8705
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Seguridad"
      TabPicture(0)   =   "Trabajadores.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Familia"
      TabPicture(1)   =   "Trabajadores.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fg"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Otros"
      TabPicture(2)   =   "Trabajadores.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Chk_Granalla"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Chk_Pintor"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame_chklst"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "FrameNC"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.Frame FrameNC 
         Height          =   1215
         Left            =   360
         TabIndex        =   53
         Top             =   600
         Width           =   5415
         Begin VB.TextBox clave0 
            Height          =   300
            Left            =   4080
            TabIndex        =   60
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox clave2 
            Height          =   285
            Left            =   4080
            TabIndex        =   58
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox clave1 
            Height          =   300
            Left            =   1560
            TabIndex        =   55
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox ChkEmisor 
            Caption         =   "Emisor NO Conformidad"
            Height          =   255
            Left            =   360
            TabIndex        =   54
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label lbl 
            Caption         =   "Clave Actual"
            Height          =   255
            Index           =   21
            Left            =   3000
            TabIndex        =   59
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lbl 
            Caption         =   "Clave Nueva"
            Height          =   255
            Index           =   20
            Left            =   3000
            TabIndex        =   57
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lbl 
            Caption         =   "Clave Nueva"
            Height          =   255
            Index           =   19
            Left            =   480
            TabIndex        =   56
            Top             =   720
            Width           =   1095
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fg 
         Height          =   4095
         Left            =   -74760
         TabIndex        =   52
         Top             =   600
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   7223
         _Version        =   393216
      End
      Begin VB.Frame Frame_chklst 
         Caption         =   "Check List"
         Height          =   855
         Left            =   360
         TabIndex        =   49
         Top             =   1920
         Width           =   5415
         Begin VB.CheckBox chklst_evaluador 
            Caption         =   "Evaluador"
            Height          =   255
            Left            =   3000
            TabIndex        =   51
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chklst_responsable 
            Caption         =   "Responsable Area"
            Height          =   255
            Left            =   360
            TabIndex        =   50
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.CheckBox Chk_Pintor 
         Caption         =   "Pintor"
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   3000
         Width           =   855
      End
      Begin VB.CheckBox Chk_Granalla 
         Caption         =   "Granalla"
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Elementos de Seguridad"
         Height          =   4095
         Left            =   -74760
         TabIndex        =   18
         Top             =   480
         Width           =   5415
         Begin VB.TextBox dato3 
            Height          =   300
            Left            =   1440
            TabIndex        =   23
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox dato4 
            Height          =   300
            Left            =   3720
            TabIndex        =   25
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox dato6 
            Height          =   300
            Left            =   1440
            TabIndex        =   29
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox dato7 
            Height          =   300
            Left            =   3720
            TabIndex        =   31
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox dato1 
            Height          =   300
            Left            =   1440
            TabIndex        =   20
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox dato2 
            Height          =   300
            Left            =   3720
            TabIndex        =   46
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox dato5 
            Height          =   300
            Left            =   1440
            TabIndex        =   26
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox dato8 
            Height          =   300
            Left            =   1440
            TabIndex        =   33
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox dato9 
            Height          =   300
            Left            =   1440
            TabIndex        =   35
            Top             =   2160
            Width           =   1815
         End
         Begin VB.TextBox dato10 
            Height          =   300
            Left            =   1440
            TabIndex        =   37
            Top             =   2520
            Width           =   1815
         End
         Begin VB.TextBox dato11 
            Height          =   300
            Left            =   1440
            TabIndex        =   39
            Top             =   2880
            Width           =   1815
         End
         Begin VB.TextBox dato12 
            Height          =   300
            Left            =   1440
            TabIndex        =   41
            Top             =   3240
            Width           =   1815
         End
         Begin VB.TextBox dato13 
            Height          =   300
            Left            =   1440
            TabIndex        =   43
            Top             =   3600
            Width           =   1815
         End
         Begin VB.Label lbl 
            Caption         =   "Talla Chaqueta"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   22
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo Guante"
            Height          =   255
            Index           =   6
            Left            =   2640
            TabIndex        =   24
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lbl 
            Caption         =   "Nº Calzado"
            Height          =   255
            Index           =   5
            Left            =   2760
            TabIndex        =   30
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label lbl 
            Caption         =   "Tipo Calzado"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   28
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label lbl 
            Caption         =   "Talla Pantalon"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   27
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lbl 
            Caption         =   "Tipo Lente"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tapon Oido"
            Height          =   255
            Left            =   2760
            TabIndex        =   21
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lbl 
            Caption         =   "Color Casco"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   32
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label lbl 
            Caption         =   "Polera"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   34
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label lbl 
            Caption         =   "T.Agua"
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   36
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label lbl 
            Caption         =   "Bota"
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   38
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label lbl 
            Caption         =   "Parka"
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   40
            Top             =   3240
            Width           =   1095
         End
         Begin VB.Label lbl 
            Caption         =   "Termico"
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   42
            Top             =   3600
            Width           =   1095
         End
      End
   End
   Begin VB.TextBox Cargo 
      Height          =   300
      Left            =   1440
      TabIndex        =   14
      Top             =   2400
      Width           =   4335
   End
   Begin VB.CheckBox Activo 
      Caption         =   "Activo"
      Height          =   255
      Left            =   240
      TabIndex        =   44
      Top             =   8400
      Width           =   855
   End
   Begin VB.ComboBox CbSexo 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox CbSeccion 
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox apmaterno 
      Height          =   300
      Left            =   1440
      TabIndex        =   8
      Top             =   1680
      Width           =   4335
   End
   Begin VB.TextBox appaterno 
      Height          =   300
      Left            =   1440
      TabIndex        =   6
      Top             =   1320
      Width           =   4335
   End
   Begin VB.TextBox nombres 
      Height          =   300
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox Codigo 
      Height          =   300
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   4680
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3000
      MousePointer    =   14  'Arrow and Question
      Picture         =   "Trabajadores.frx":0054
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   300
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   6000
      Top             =   600
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trabajadores.frx":0156
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trabajadores.frx":0268
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trabajadores.frx":037A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trabajadores.frx":048C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trabajadores.frx":059E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trabajadores.frx":06B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "correo"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Cargo"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Se&xo"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label seccion 
      Caption         =   "Sección"
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Ap. &Materno"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&Nombres"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Ap. &Paterno"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&RUT Trabajador"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Menu MenuPop 
      Caption         =   ""
      Begin VB.Menu Menu 
         Caption         =   "Agregar Familiar"
         Index           =   1
      End
      Begin VB.Menu Menu 
         Caption         =   "Modificar Familiar"
         Index           =   2
      End
      Begin VB.Menu Menu 
         Caption         =   "Eliminar Familiar"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Trabajadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Obj As String, Objs As String, nuevo As Boolean, Accion As String, old_accion As String
Private Db As Database, Rs As Recordset, RsGf As Recordset
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnListar As Button, btnGrabar As Button, btnDesHacer As Button
Private m_Rut As String, a_Seccion(1, 21) As String, i As Integer
Private n_filas As Integer, n_columnas As Integer
Private m_ClaveOLD As String
Private Sub Inicializa()

Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnListar = Toolbar.Buttons(4)
Set btnGrabar = Toolbar.Buttons(5)
Set btnDesHacer = Toolbar.Buttons(6)

Obj = "TRABAJADOR"
Objs = "TRABAJADORES"

Accion = ""
old_accion = ""

btnBuscar.visible = False
btnBuscar.ToolTipText = "Busca " & StrConv(Obj, vbProperCase)

Codigo.MaxLength = 10
nombres.MaxLength = 30
appaterno.MaxLength = 30
apmaterno.MaxLength = 30
CbSeccion.ListIndex = -1
Cargo.MaxLength = 30

dato1.MaxLength = 10
dato2.MaxLength = 10
dato3.MaxLength = 2
dato4.MaxLength = 10
dato5.MaxLength = 2
dato6.MaxLength = 10
dato7.MaxLength = 2
dato8.MaxLength = 10
dato9.MaxLength = 10
dato10.MaxLength = 10
dato11.MaxLength = 10
dato12.MaxLength = 10
dato13.MaxLength = 10

correo.MaxLength = 30
clave0.MaxLength = 10
clave1.MaxLength = 10
clave2.MaxLength = 10

a_Seccion(0, 1) = "ADQ"
a_Seccion(1, 1) = "Adquisicion"
a_Seccion(0, 2) = "ARS"
a_Seccion(1, 2) = "Arco Sumergido"
a_Seccion(0, 3) = "BOD"
a_Seccion(1, 3) = "Bodega"
a_Seccion(0, 4) = "CON"
a_Seccion(1, 4) = "Contabilidad"
a_Seccion(0, 5) = "CHO"
a_Seccion(1, 5) = "Chofer"
a_Seccion(0, 6) = "DES"
a_Seccion(1, 6) = "Despacho"
a_Seccion(0, 7) = "GER"
a_Seccion(1, 7) = "Gerencia"
a_Seccion(0, 8) = "GRA"
a_Seccion(1, 8) = "Granalla"
a_Seccion(0, 9) = "GRU"
a_Seccion(1, 9) = "Gruero"
a_Seccion(0, 10) = "GUI"
a_Seccion(1, 10) = "Guillotina"
a_Seccion(0, 11) = "INS"
a_Seccion(1, 11) = "Inspección"
a_Seccion(0, 12) = "MEL"
a_Seccion(1, 12) = "Mantencion Electrica"
a_Seccion(0, 13) = "MON"
a_Seccion(1, 13) = "Montaje"
a_Seccion(0, 14) = "OPE"
a_Seccion(1, 14) = "Operaciones"
a_Seccion(0, 15) = "OXI"
a_Seccion(1, 15) = "Oxicorte"
a_Seccion(0, 16) = "PPL"
a_Seccion(1, 16) = "Patio Plancha"
a_Seccion(0, 17) = "PIN"
a_Seccion(1, 17) = "Pintura"
a_Seccion(0, 18) = "PLA"
a_Seccion(1, 18) = "Plasma"
a_Seccion(0, 19) = "PMA"
a_Seccion(1, 19) = "Prep. Material"
a_Seccion(0, 20) = "PRE"
a_Seccion(1, 20) = "Prevensión"
a_Seccion(0, 21) = "PRO"
a_Seccion(1, 21) = "Producción"

'CbSeccion.AddItem " "
For i = 1 To 21
    CbSeccion.AddItem a_Seccion(1, i)
Next

CbSexo.AddItem "Masculino"
CbSexo.AddItem "Femenino"

n_filas = 10 '3
n_columnas = 5

Detalle_Config

clave0.PasswordChar = "*"
clave1.PasswordChar = "*"
clave2.PasswordChar = "*"

MenuPop.visible = False

End Sub

Private Sub ChkEmisor_Click()
If ChkEmisor.Value = 1 Then
    clave0.Enabled = True
    clave1.Enabled = True
    clave2.Enabled = True
Else
    clave0.Text = ""
    clave1.Text = ""
    clave2.Text = ""
    clave0.Enabled = False
    clave1.Enabled = False
    clave2.Enabled = False
End If
End Sub

Private Sub Form_Load()

Inicializa

Me.Caption = "MANTENCIÓN DE " & Objs
Botones_Enabled True, True, True, True, False, False

' abre archivos
Set Db = OpenDatabase(data_file)
Set Rs = Db.OpenRecordset("trabajadores")
Rs.Index = "RUT"

Set RsGf = Db.OpenRecordset("Grupo Familiar")
RsGf.Index = "Rut-Rut Carga"

Campos_Limpiar

nuevo = False

End Sub
Private Sub Botones_Enabled(btn_Agregar As Boolean, btn_Modificar As Boolean, btn_Eliminar As Boolean, btn_Listar As Boolean, btn_Grabar As Boolean, btn_DesHacer As Boolean)
Dim i As Integer
btnAgregar.Enabled = btn_Agregar
btnModificar.Enabled = btn_Modificar
btnEliminar.Enabled = btn_Eliminar
btnListar.Enabled = btn_Listar
btnGrabar.Enabled = btn_Grabar
btnDesHacer.Enabled = btn_DesHacer

For i = 1 To 6
    Toolbar.Buttons(i).Value = tbrUnpressed
Next

End Sub
Private Sub Campos_Enabled(Si As Boolean)

Codigo.Enabled = Si
nombres.Enabled = Si
appaterno.Enabled = Si
apmaterno.Enabled = Si
CbSexo.Enabled = Si
CbSeccion.Enabled = Si
Cargo.Enabled = Si

Tab_Datos.Enabled = Si
Tab_Datos.Tab = 0

dato1.Enabled = Si
dato2.Enabled = Si
dato3.Enabled = Si
dato4.Enabled = Si
dato5.Enabled = Si
dato6.Enabled = Si
dato7.Enabled = Si
dato8.Enabled = Si
dato9.Enabled = Si
dato10.Enabled = Si
dato11.Enabled = Si
dato12.Enabled = Si
dato13.Enabled = Si

correo.Enabled = Si
ChkEmisor.Enabled = Si
clave0.Enabled = Si
clave1.Enabled = Si
clave2.Enabled = Si

chklst_responsable.Enabled = Si
chklst_evaluador.Enabled = Si

Chk_Pintor.Enabled = Si
Chk_Granalla.Enabled = Si

activo.Enabled = Si

End Sub
Private Sub Campos_Limpiar()

Codigo.Text = ""
nombres.Text = ""
appaterno.Text = ""
apmaterno.Text = ""
CbSexo.ListIndex = 0
'CbSeccion.Text = " "
CbSeccion.ListIndex = -1
Cargo.Text = ""

dato1.Text = ""
dato2.Text = ""
dato3.Text = ""
dato4.Text = ""
dato5.Text = ""
dato6.Text = ""
dato7.Text = ""
dato8.Text = ""
dato9.Text = ""
dato10.Text = ""
dato11.Text = ""
dato12.Text = ""
dato13.Text = ""

correo.Text = ""

ChkEmisor.Value = 0
clave0.Text = ""
clave1.Text = ""
clave2.Text = ""

activo.Value = 0

chklst_responsable.Value = 0
chklst_evaluador.Value = 0

Chk_Pintor.Value = 0
Chk_Granalla.Value = 0

Campos_Enabled False

End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Arch As String
Accion = Button.ToolTipText
Select Case Accion
Case "Agregar"
    Campos_Enabled False
    Codigo.Enabled = True
    Codigo.SetFocus
    nuevo = True
    old_accion = Accion
Case "Modificar"
    Codigo.Enabled = True
    nombres.Enabled = False
    Codigo.SetFocus
    nuevo = False
    old_accion = Accion
    btnBuscar.visible = True
Case "Eliminar"
    Codigo.Enabled = True
    nombres.Enabled = False
    Codigo.SetFocus
'    nuevo = False
    btnBuscar.visible = True
Case "Listar"
    MousePointer = vbHourglass
    cr.WindowTitle = Objs
    cr.WindowMaxButton = False
    cr.WindowMinButton = False
    cr.WindowState = crptMaximized
    cr.DataFiles(0) = data_file & ".MDB"
    cr.Formulas(0) = "RAZON SOCIAL=""" & Empresa.Razon & """"
    cr.ReportSource = crptReport
    cr.ReportFileName = Drive_Server & Path_Rpt & "trabajadores.rpt"
    cr.Action = 1
    MousePointer = vbDefault
Case "Grabar"
    If Valida(nuevo) Then
        Registro_Grabar nuevo
    Else
        Exit Sub
    End If
Case "Deshacer"
    Campos_Limpiar
    btnBuscar.visible = False
End Select

Select Case Button.Index
Case 5  ' btnGrabar
    Campos_Limpiar
    
    Codigo.Enabled = True
    Codigo.SetFocus
    btnGrabar.Value = tbrUnpressed
    btnGrabar.Enabled = False
    
Case 4 To 6 ' btnDesHacer
    Botones_Enabled True, True, True, True, False, False
    Me.Caption = "MANTENCIÓN DE " & Objs
    
Case 8 ' listado dos
    MousePointer = vbHourglass
    cr.WindowTitle = "Grupo Familiar"
    cr.WindowMaxButton = False
    cr.WindowMinButton = False
    cr.WindowState = crptMaximized
    cr.DataFiles(0) = data_file & ".MDB"
    cr.DataFiles(1) = data_file & ".MDB"
    cr.Formulas(0) = "RAZON SOCIAL=""" & Empresa.Razon & """"
    cr.ReportSource = crptReport
    cr.ReportFileName = Drive_Server & Path_Rpt & "grupofamiliar.rpt"
    cr.Action = 1
    MousePointer = vbDefault
Case Else
    Botones_Enabled False, False, False, False, False, True
    Me.Caption = "MANTENCIÓN DE " & Objs & " " & Button.Tag
End Select

End Sub
Private Sub btnBuscar_Click()

Search.Muestra data_file, "Trabajadores", "RUT", "ApPaterno]+' '+[ApMaterno]+' '+[Nombres", Obj, Objs

Codigo.Text = Search.Codigo
nombres.Text = Search.Descripcion

If Codigo.Text <> "" Then
    Rs.Seek "=", Codigo.Text
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        Codigo.SetFocus
    Else
        After_Enter
    End If
End If
End Sub
Private Sub Codigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If Codigo.Text = "" Then
        Beep
    Else
        After_Enter
    End If
End If
End Sub
Private Sub After_Enter()

If Rut_Verifica(Codigo.Text) = False Then
    MsgBox "RUT no Válido"
    Exit Sub
End If

m_Rut = Rut_Formato(Codigo.Text)

Select Case Accion
Case "Agregar"
    Rs.Seek "=", m_Rut ' Codigo
    If Rs.NoMatch Then
        Campos_Enabled True
        Codigo.Enabled = False
        nombres.SetFocus
        btnGrabar.Enabled = True
    Else
        Registro_Leer
        MsgBox Obj & " YA EXISTE"
        Campos_Limpiar
        Codigo.Enabled = True
        Codigo.SetFocus
    End If
Case "Modificar"
    Rs.Seek "=", m_Rut ' Codigo
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        Codigo.SetFocus
    Else
        Campos_Enabled True
        Codigo.Enabled = False
        nombres.SetFocus
        Registro_Leer
        btnGrabar.Enabled = True
        btnBuscar.visible = False
    End If
Case "Eliminar"
    Rs.Seek "=", m_Rut ' Codigo
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        Codigo.SetFocus
    Else
        Campos_Enabled False
        Registro_Leer
        If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo, "Atención") = vbYes Then
            Rs.Delete
        End If
        btnBuscar.visible = True
        Campos_Limpiar
        Codigo.Enabled = True
        Codigo.SetFocus
    End If
End Select
End Sub
Private Sub nombres_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub appaterno_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub apmaterno_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Function Valida(nuevo As Boolean) As Boolean

Valida = False

If IsObjBlanco(nombres, "NOMBRE", btnGrabar) Then Exit Function
If IsObjBlanco(appaterno, "APELLIDO PATERNO", btnGrabar) Then Exit Function

'Valida = True

'Exit Function

If ChkEmisor.Value = 1 Then

    'If clave0.Text = "" Then
    '    MsgBox "Debe Digitar Clave Actual"
    '    clave0.SetFocus
    '    btnGrabar.Value = tbrUnpressed
    '    Exit Function
    'End If
    
    If clave0.Text <> m_ClaveOLD Then
        MsgBox "Clave Actual INCORRECTA"
        clave0.SetFocus
        btnGrabar.Value = tbrUnpressed
        Exit Function
    End If
    
    If clave1.Text = "" Then
        MsgBox "Debe Digitar Clave Nueva"
        clave1.SetFocus
        btnGrabar.Value = tbrUnpressed
        Exit Function
    End If
    
    If clave2.Text = "" Then
        MsgBox "Debe Digitar Clave Nueva"
        clave2.SetFocus
        btnGrabar.Value = tbrUnpressed
        Exit Function
    End If
    
    If clave1.Text <> clave2.Text Then
        MsgBox "Clave Nueva Debe ser igual"
        clave2.SetFocus
        btnGrabar.Value = tbrUnpressed
        Exit Function
    End If
    
End If

Valida = True

End Function
Private Sub Registro_Grabar(nuevo As Boolean)

With Rs

    If nuevo Then
        .AddNew
        !rut = m_Rut 'Codigo
    Else
        .Edit
    End If
    
    ![nombres] = nombres.Text
    !appaterno = appaterno.Text
    !apmaterno = apmaterno.Text
    !clase1 = a_Seccion(0, CbSeccion.ListIndex + 1)
    !Sexo = Left(CbSexo.Text, 1)
    !Cargo = Cargo.Text
    
    !dato1 = dato1.Text
    !dato2 = dato2.Text
    !dato3 = dato3.Text
    !dato4 = dato4.Text
    !dato5 = dato5.Text
    !dato6 = dato6.Text
    !dato7 = dato7.Text
    !dato8 = dato8.Text
    !dato9 = dato9.Text
    !dato10 = dato10.Text
    !dato11 = dato11.Text
    !dato12 = dato12.Text
    !dato13 = dato13.Text
    
    !dato20 = correo.Text
    
    !emisor_nc = ChkEmisor.Value
    !emisor_nc_clave = clave1.Text

    !chklst_responsable = chklst_responsable.Value
    !chklst_evaluador = chklst_evaluador.Value
    
    !tipo4 = Chk_Pintor.Value
    !tipo5 = Chk_Granalla.Value
    
    !activo = activo.Value
    
    .Update
    
End With

Campos_Limpiar
Accion = old_accion

End Sub
Private Sub Registro_Leer()
With Rs

nombres.Text = !nombres
appaterno.Text = !appaterno
apmaterno.Text = !apmaterno

For i = 1 To 21
    If a_Seccion(0, i) = !clase1 Then
        CbSeccion.ListIndex = i - 1
    End If
Next

CbSexo.ListIndex = IIf(!Sexo = "F", 1, 0)

Cargo.Text = NoNulo(!Cargo)

'lentes_tipo.Text = NoNulo(!lentes_tipo)
dato1.Text = NoNulo(!dato1)
dato2.Text = NoNulo(!dato2)
dato3.Text = NoNulo(!dato3)
dato4.Text = NoNulo(!dato4)
dato5.Text = NoNulo(!dato5)
dato6.Text = NoNulo(!dato6)
dato7.Text = NoNulo(!dato7)
dato8.Text = NoNulo(!dato8)
dato9.Text = NoNulo(!dato9)
dato10.Text = NoNulo(!dato10)
dato11.Text = NoNulo(!dato11)
dato12.Text = NoNulo(!dato12)

correo.Text = NoNulo(!dato20)

ChkEmisor.Value = IIf(!emisor_nc, 1, 0)
m_ClaveOLD = NoNulo(!emisor_nc_clave)

'    Set Campo(11) = Td.CreateField("dato1", dbText, 20) ' tipos de lentes
'    Set Campo(12) = Td.CreateField("dato2", dbText, 20) ' tapones
'    Set Campo(13) = Td.CreateField("dato3", dbText, 20) ' talla chaqueta
'    Set Campo(14) = Td.CreateField("dato4", dbText, 20) ' tipos de guantes
'    Set Campo(15) = Td.CreateField("dato5", dbText, 20) ' talla pantalon
'    Set Campo(16) = Td.CreateField("date6", dbText, 20) ' tipo de calzado
'    Set Campo(17) = Td.CreateField("dato7", dbText, 20) ' numero del calzado
'    Set Campo(18) = Td.CreateField("dato8", dbText, 20) ' color casco
'    Set Campo(19) = Td.CreateField("dato9", dbText, 20) ' polera
'    Set Campo(20) = Td.CreateField("dato10", dbText, 20) ' t.agua
'    Set Campo(21) = Td.CreateField("dato11", dbText, 20) ' bota
'    Set Campo(22) = Td.CreateField("dato12", dbText, 20) ' parka
'    Set Campo(23) = Td.CreateField("dato13", dbText, 20) ' termico

chklst_responsable.Value = IIf(!chklst_responsable, 1, 0)
chklst_evaluador.Value = IIf(!chklst_evaluador, 1, 0)

Chk_Pintor.Value = IIf(!tipo4, 1, 0)
Chk_Granalla.Value = IIf(!tipo5, 1, 0)

' lee grupo familiar
GrupoFamiliar_Leer

activo.Value = IIf(!activo, 1, 0)

End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Rs.Close
Db.Close
End Sub
'////////////////////////////////////////////////////////////////////
Private Sub Detalle_Config()
Dim i As Integer, ancho As Integer, c_totales As Integer

'fg.Left = 100
fg.WordWrap = True
'fg.RowHeight(0) = 450
fg.Rows = n_filas + 1
fg.Cols = n_columnas + 1

fg.TextMatrix(0, 0) = ""
fg.TextMatrix(0, 1) = "RUT"
fg.TextMatrix(0, 2) = "Nombre"
fg.TextMatrix(0, 3) = "F.Nacim"
fg.TextMatrix(0, 4) = "Par."
fg.TextMatrix(0, 5) = "Sexo"

fg.ColWidth(0) = 300
fg.ColWidth(1) = 1000
fg.ColWidth(2) = 2900
fg.ColWidth(3) = 800
fg.ColWidth(4) = 500
fg.ColWidth(5) = 500

'fg.ColAlignment(1) = 0
'fg.ColAlignment(2) = 0

'Ancho = fg.ColWidth(n_columnas)

ancho = 350 ' con scroll vertical
For i = 0 To n_columnas
    ancho = ancho + fg.ColWidth(i)
Next

fg.Width = ancho
'Me.Width = Ancho + fg.Left * 2

For i = 1 To n_filas
    fg.TextMatrix(i, 0) = i
Next

'For i = 1 To n_filas
'    fg.Row = i
'    fg.col = n_columnas '6
'    fg.CellForeColor = vbRed
'Next

''''''txtEdit.Text = ""

'fg.Enabled = False

End Sub
Private Sub fg_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = 2 Then
    Dim Fila As Integer, m_Rut As String
    Fila = fg.RowSel
    m_Rut = fg.TextMatrix(Fila, 1)
'    If Trim(m_rut) <> "" Then
'        nd = CtaCte.TextMatrix(fila, 3)
    SubMenu m_Rut
'    End If
End If

End Sub
Private Sub SubMenu(m_Rut As String)
'muestra menu
If m_Rut = "" Then
    menu(1).Enabled = True
    menu(2).Enabled = False
    menu(3).Enabled = False
Else
    menu(1).Enabled = True
    menu(2).Enabled = True
    menu(3).Enabled = True
End If
PopupMenu MenuPop
End Sub
Private Sub menu_Click(Index As Integer)
Select Case Index
Case 1 ' agregar
    MousePointer = vbHourglass
    Familiares.ruttrabajador = Codigo.Text
    Familiares.Accion = "agregar"
    Familiares.RutCarga = ""
    Familiares.RecSet = RsGf
    Load Familiares
    MousePointer = vbDefault
    Familiares.Show 1
    GrupoFamiliar_Leer
Case 2 ' editar
    MousePointer = vbHourglass
    Familiares.ruttrabajador = Codigo.Text
    Familiares.Accion = "editar"
    Familiares.RutCarga = fg.TextMatrix(fg.Row, 1)
    Familiares.RecSet = RsGf
    Load Familiares
    MousePointer = vbDefault
    Familiares.Show 1
    GrupoFamiliar_Leer
Case 3 ' eliminar
    If MsgBox("¿ Seguro que desea Eliminar ?", vbYesNo) = vbYes Then
        ' elimina
        RsGf.Seek "=", Codigo.Text, fg.TextMatrix(fg.Row, 1)
        If Not RsGf.NoMatch Then
            RsGf.Delete
            GrupoFamiliar_Leer
        End If
    End If
End Select
End Sub
Private Sub GrupoFamiliar_Leer()
' lee grupo familiar del trabajador
Dim NF As Integer, i As Integer
With RsGf

NF = 0
.Seek ">=", Codigo.Text
If Not .NoMatch Then
    Do While Not .EOF
        If !rut <> Codigo.Text Then Exit Do
        ' puebla fg
        NF = NF + 1
        
        If NF >= fg.Rows Then
            fg.Rows = fg.Rows + 1
            fg.TextMatrix(NF, 0) = NF
        End If
        
        fg.TextMatrix(NF, 1) = ![Rut Carga]
        fg.TextMatrix(NF, 2) = !nombres & " " & ![paterno] & " " & ![materno]
        fg.TextMatrix(NF, 3) = ![Fecha Nacimiento]
        fg.TextMatrix(NF, 4) = ![Parentesco]
        fg.TextMatrix(NF, 5) = ![Sexo]
        
        .MoveNext
        
    Loop
End If

For i = NF + 1 To fg.Rows - 1
    fg.TextMatrix(i, 1) = ""
    fg.TextMatrix(i, 2) = ""
    fg.TextMatrix(i, 3) = ""
    fg.TextMatrix(i, 4) = ""
    fg.TextMatrix(i, 5) = ""
Next

End With

End Sub
