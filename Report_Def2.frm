VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Report_Def 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Informe"
   ClientHeight    =   10935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame_PP 
      Caption         =   "PIEZAS PENDIENTES"
      Height          =   855
      Left            =   120
      TabIndex        =   102
      Top             =   9480
      Width           =   5895
      Begin VB.OptionButton OpPP 
         Caption         =   "&Todas"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   104
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OpPP 
         Caption         =   "Solo &Pendientes"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   103
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame_CC 
      Caption         =   "CENTRO COSTO"
      Height          =   855
      Left            =   120
      TabIndex        =   98
      Top             =   1200
      Width           =   5895
      Begin VB.ComboBox ComboCC 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   360
         Width           =   3135
      End
      Begin VB.OptionButton OpCC 
         Caption         =   "&Uno"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   100
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton OpCC 
         Caption         =   "&Todos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   99
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame_NCGerencia 
      Caption         =   "Gerencia"
      Height          =   855
      Left            =   120
      TabIndex        =   93
      Top             =   8880
      Width           =   5895
      Begin VB.OptionButton OpNCGerencia 
         Caption         =   "Todas"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   96
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OpNCGerencia 
         Caption         =   "Una"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   95
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox CbNCGerencia 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame_TProv 
      Caption         =   "TIPO PROVEEDOR"
      Height          =   855
      Left            =   120
      TabIndex        =   76
      Top             =   2040
      Width           =   5895
      Begin VB.OptionButton OpTProv 
         Caption         =   "Todos"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   79
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OpTProv 
         Caption         =   "Un Tipo"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   78
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox CbClasif 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame_TPza 
      Caption         =   "TIPO DE PIEZA"
      Height          =   855
      Left            =   120
      TabIndex        =   89
      Top             =   0
      Width           =   5895
      Begin VB.OptionButton OpTPza 
         Caption         =   "Toda&s"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   92
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OpTPza 
         Caption         =   "Un&a"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   91
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox CbTipoPieza 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame_Turno 
      Caption         =   "TURNO"
      Height          =   855
      Left            =   120
      TabIndex        =   85
      Top             =   8880
      Width           =   5895
      Begin VB.OptionButton OpTurno 
         Caption         =   "&Noche"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   88
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OpTurno 
         Caption         =   "&Todos"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   87
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OpTurno 
         Caption         =   "&Dia"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   86
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame_Numero 
      Caption         =   "NUMERO"
      Height          =   855
      Left            =   120
      TabIndex        =   80
      Top             =   2400
      Width           =   5895
      Begin VB.TextBox Numero_Ini 
         Height          =   300
         Left            =   1200
         TabIndex        =   82
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Numero_Fin 
         Height          =   300
         Left            =   3240
         TabIndex        =   81
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblNumero 
         Caption         =   "&Desde"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   84
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblNumero 
         Caption         =   "&Hasta"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   83
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame_Prov 
      Caption         =   "PROVEEDOR"
      Height          =   855
      Left            =   120
      TabIndex        =   71
      Top             =   1560
      Width           =   5895
      Begin VB.CommandButton btnPrvBuscar 
         Height          =   300
         Left            =   2160
         Picture         =   "Report_Def2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   360
         Width           =   300
      End
      Begin VB.OptionButton OpProv 
         Caption         =   "U&no"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   73
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton OpProv 
         Caption         =   "T&odos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   72
         Top             =   360
         Width           =   855
      End
      Begin VB.Label PrvDescripcion 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2520
         TabIndex        =   75
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame_Maquina 
      Caption         =   "MAQUINA"
      Height          =   735
      Left            =   120
      TabIndex        =   67
      Top             =   9240
      Width           =   5895
      Begin VB.OptionButton OpMaq 
         Caption         =   "Automatica"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   70
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OpMaq 
         Caption         =   "Manual"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   69
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OpMaq 
         Caption         =   "Todas"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   68
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame_TipoGranalla 
      Caption         =   "TIPO GRANALLADO"
      Height          =   735
      Left            =   120
      TabIndex        =   63
      Top             =   8640
      Width           =   5895
      Begin VB.ComboBox CbTipoGranalla 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OpTGr 
         Caption         =   "Uno"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   65
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OpTGr 
         Caption         =   "Todos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   64
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame_OP 
      Caption         =   "OPERADOR"
      Height          =   855
      Left            =   120
      TabIndex        =   58
      Top             =   7800
      Width           =   5895
      Begin VB.CommandButton btnOpBuscar 
         Height          =   300
         Left            =   2160
         Picture         =   "Report_Def2.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   360
         Width           =   300
      End
      Begin VB.OptionButton OpOp 
         Caption         =   "U&no"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   60
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OpOp 
         Caption         =   "T&odos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   59
         Top             =   360
         Width           =   855
      End
      Begin VB.Label OpNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2520
         TabIndex        =   62
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame_RA 
      Caption         =   "RESPONSABLE AREA"
      Height          =   855
      Left            =   120
      TabIndex        =   53
      Top             =   7200
      Width           =   5895
      Begin VB.CommandButton btnRABuscar 
         Height          =   300
         Left            =   2160
         Picture         =   "Report_Def2.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   360
         Width           =   300
      End
      Begin VB.OptionButton OpRA 
         Caption         =   "U&no"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   55
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OpRA 
         Caption         =   "T&odos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   54
         Top             =   360
         Width           =   855
      End
      Begin VB.Label RANombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2520
         TabIndex        =   57
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame_chklstArea 
      Caption         =   "AREA"
      Height          =   855
      Left            =   120
      TabIndex        =   47
      Top             =   6600
      Width           =   5895
      Begin VB.ComboBox CbchklstAreas 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton Op_clArea 
         Caption         =   "Una"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   49
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Op_clArea 
         Caption         =   "Todas"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   48
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame_TG 
      Caption         =   "TIPO GUIA"
      Height          =   855
      Left            =   120
      TabIndex        =   43
      Top             =   6000
      Width           =   5895
      Begin VB.ComboBox ComboTG 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton OpTG 
         Caption         =   "Un Tipo"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   45
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OpTG 
         Caption         =   "Todas"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   44
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame_OT 
      Caption         =   "OT"
      Height          =   855
      Left            =   120
      TabIndex        =   39
      Top             =   5400
      Width           =   5895
      Begin VB.ComboBox ComboOT 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   240
         Width           =   2775
      End
      Begin VB.OptionButton OpOT 
         Caption         =   "&Una"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   41
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OpOT 
         Caption         =   "&Todas"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   40
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame_PL 
      Caption         =   "PLANO"
      Height          =   855
      Left            =   120
      TabIndex        =   35
      Top             =   4800
      Width           =   5895
      Begin VB.ComboBox ComboPl 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   360
         Width           =   2775
      End
      Begin VB.OptionButton OpPl 
         Caption         =   "Un&o"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   37
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OpPl 
         Caption         =   "Todo&s"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   36
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame_TI 
      Caption         =   "TIPO DE INFORME"
      Height          =   855
      Left            =   120
      TabIndex        =   32
      Top             =   4200
      Width           =   5895
      Begin VB.OptionButton OpTI 
         Caption         =   "D&etalle"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   34
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OpTI 
         Caption         =   "&General"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   33
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame_Fecha 
      Caption         =   "FECHA"
      Height          =   855
      Left            =   120
      TabIndex        =   27
      Top             =   3600
      Width           =   5895
      Begin VB.CommandButton btnFechaRango 
         Caption         =   "22 / 21"
         Height          =   255
         Left            =   4560
         TabIndex        =   97
         Top             =   360
         Width           =   855
      End
      Begin MSMask.MaskEdBox Fecha_Fin 
         Height          =   300
         Left            =   3240
         TabIndex        =   31
         Top             =   360
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   327680
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Fecha_Ini 
         Height          =   300
         Left            =   1200
         TabIndex        =   29
         Top             =   360
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   327680
         PromptChar      =   "_"
      End
      Begin VB.Label lblFecha 
         Caption         =   "&Hasta"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   30
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblFecha 
         Caption         =   "&Desde"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame_PR 
      Caption         =   "PRODUCTO"
      Height          =   855
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   5895
      Begin VB.OptionButton OpPrd 
         Caption         =   "T&odos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OpPrd 
         Caption         =   "U&no"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   24
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton btnPrdBuscar 
         Height          =   300
         Left            =   2160
         Picture         =   "Report_Def2.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   300
      End
      Begin VB.Label PrdDescripcion 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2520
         TabIndex        =   26
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame_TR 
      Caption         =   "TRABAJADOR"
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   5895
      Begin VB.OptionButton OpTr 
         Caption         =   "T&odos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OpTr 
         Caption         =   "U&no"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton btnTrBuscar 
         Height          =   300
         Left            =   2160
         Picture         =   "Report_Def2.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   300
      End
      Begin VB.Label TrNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2520
         TabIndex        =   21
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame_Cl 
      Caption         =   "CLIENTE"
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   5895
      Begin VB.CommandButton btnClBuscar 
         Height          =   300
         Left            =   2160
         Picture         =   "Report_Def2.frx":050A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   300
      End
      Begin VB.OptionButton OpCl 
         Caption         =   "U&no"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OpCl 
         Caption         =   "T&odos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label ClRazon 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2520
         TabIndex        =   16
         Top             =   360
         Width           =   3255
      End
   End
   Begin Crystal.CrystalReport CR 
      Left            =   5280
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowMaxButton =   0   'False
      WindowMinButton =   0   'False
      WindowState     =   2
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3240
      TabIndex        =   52
      Top             =   10440
      Width           =   1095
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1920
      TabIndex        =   51
      Top             =   10440
      Width           =   1095
   End
   Begin VB.Frame Frame_SC 
      Caption         =   "CONTRATISTA"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   5895
      Begin VB.OptionButton OpSc 
         Caption         =   "T&odos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OpSc 
         Caption         =   "U&no"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton btnScBuscar 
         Height          =   300
         Left            =   2160
         Picture         =   "Report_Def2.frx":060C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   300
      End
      Begin VB.Label ScRazon 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2520
         TabIndex        =   11
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame_NV 
      Caption         =   "NOTA DE VENTA"
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5895
      Begin VB.TextBox Nv 
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OpNV 
         Caption         =   "&Todas"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OpNV 
         Caption         =   "&Una"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox ComboNv 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Label lblInforme 
      Caption         =   "Informe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label lbl 
      Caption         =   "Informe :"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Report_Def"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Titulo As String

Private m_Opcion_Nv As Boolean
Private m_Opcion_CC As Boolean

Private m_Opcion_Contratista As Boolean
Private m_Opcion_ContratistaObligatorio As Boolean ' para ayd

Private m_Opcion_Cliente As Boolean

Private m_Opcion_Proveedor As Boolean ' 02/11/07
Private m_Opcion_ProveedorTipo As Boolean ' 02/11/07
Private m_Opcion_Numero As Boolean ' 02/11/07

Private m_Opcion_Trabajador As Boolean
Private m_Opcion_Producto As Boolean

Private m_Opcion_Fecha As Boolean
' boton para rango de fechas (de corte), 22/21
Private m_Opcion_BotonRango As Boolean

Private m_Opcion_TipoInforme As Boolean
Private m_Opcion_Plano As Boolean
Private m_Opcion_OT As Boolean
Private m_Opcion_TipoGD As Boolean
Private m_Opcion_chklstArea As Boolean
Private m_Opcion_chklstResponsableArea As Boolean
Private m_Opcion_Operador As Boolean ' arco sumergido
Private m_Opcion_TipoGranalla As Boolean
Private m_Opcion_Maquina As Boolean
Private m_Opcion_Turno As Boolean ' para ...  y arco sumergido
Private m_Opcion_TipoPieza As Boolean ' para arco sumergido
Private m_Opcion_NCArea As Boolean ' areas de no conformidad
' para piezas pendientes, "T"odoas o Solo "P"endientes
' solicitado por Renan 15/12/16
Private m_Opcion_PP As Boolean

Private DbD As Database, RsCl As Recordset, RsAreas As Recordset
Private Dbm As Database
'Private RsNVc As Recordset
Private RsPc As Recordset
Private RsOTc As Recordset

' variables de salida
Private NV_Numero As Double
Private NV_nombre As String
Private NV_Area As Integer

Private ccCodigo As String
Private ccDescripcion As String

Private Contratista_Rut As String
Private Contratista_Razon As String
Private Cliente_RUT As String
Private Cliente_Razon As String
Private Trabajador_RUT As String
Private Trabajador_Nombre As String
Private Proveedor_RUT As String
Private Proveedor_Nombre As String
Private Proveedor_Tipo As String
Private Numero As Double
Private Producto_Codigo As String
Private ResponsableArea_RUT As String
Private ResponsableArea_Nombre As String
Private Plano As String, Revision As String
Private Operador_RUT As String
Private Operador_Nombre As String
Private OT_Numero As Double
Private NC_Area As String
'Private Prd_Codigo As String
Private d As Variant

Private Type Condizion
    NotaVenta As String
    Contratista As String
    Cliente As String
    Trabajador As String
    Proveedor As String
    ProveedorTipo As String
    ResponsableArea As String
    Operador As String
    Plano As String
    FechaInicial As String
    FechaFinal As String
    Numero As String
    OT As String
    Producto As String
End Type
Private Condicion As Condizion

' 0: numero,  1: nombre obra
Private a_Nv(2999, 1) As String, i As Integer
Private a_CC(2999, 1) As String

Private m_top As Integer ' top de cada frame
Private Numero_Frames As Integer ' numero de frames o filtros
Private m_TipoGuia As String
Private a_Areas(99) As String, m_NvArea As Integer
Private qry As String
'///////////
Public Property Let Titulo(ByVal New_Titulo As String)
m_Titulo = New_Titulo
End Property
Public Property Let Op_NotaVenta(ByVal New_Opcion As Boolean)
m_Opcion_Nv = New_Opcion
End Property
Public Property Let Op_CentroCosto(ByVal New_Opcion As Boolean)
m_Opcion_CC = New_Opcion
End Property
Public Property Let Op_Contratista(ByVal New_Opcion As Boolean)
m_Opcion_Contratista = New_Opcion
End Property
Public Property Let Op_ContratistaObligatorio(ByVal New_Opcion As Boolean)
m_Opcion_ContratistaObligatorio = New_Opcion
End Property
Public Property Let Op_Cliente(ByVal New_Opcion As Boolean)
m_Opcion_Cliente = New_Opcion
End Property
Public Property Let Op_Trabajador(ByVal New_Opcion As Boolean)
m_Opcion_Trabajador = New_Opcion
End Property
Public Property Let Op_Proveedor(ByVal New_Opcion As Boolean)
m_Opcion_Proveedor = New_Opcion
End Property
Public Property Let Op_ProveedorTipo(ByVal New_Opcion As Boolean)
m_Opcion_ProveedorTipo = New_Opcion
End Property
Public Property Let Op_Numero(ByVal New_Opcion As Boolean)
m_Opcion_Numero = New_Opcion
End Property
Public Property Let Op_Producto(ByVal New_Opcion As Boolean)
m_Opcion_Producto = New_Opcion
End Property
Public Property Let Op_Fecha(ByVal New_Opcion As Boolean)
m_Opcion_Fecha = New_Opcion
End Property
Public Property Let Op_BotonRango(ByVal New_Opcion As Boolean)
m_Opcion_BotonRango = New_Opcion
End Property
Public Property Let Op_TipoRepo(ByVal New_Opcion As Boolean)
' general o detalle
m_Opcion_TipoInforme = New_Opcion
End Property
Public Property Let Op_Plano(ByVal New_Opcion As Boolean)
m_Opcion_Plano = New_Opcion
End Property
Public Property Let Op_OT(ByVal New_Opcion As Boolean)
m_Opcion_OT = New_Opcion
End Property
Public Property Let Op_TipoGD(ByVal New_Opcion As Boolean)
m_Opcion_TipoGD = New_Opcion
End Property
Public Property Let Op_chklstArea(ByVal New_Opcion As Boolean)
m_Opcion_chklstArea = New_Opcion
End Property
Public Property Let Op_chklstResponsableArea(ByVal New_Opcion As Boolean)
m_Opcion_chklstResponsableArea = New_Opcion
End Property
Public Property Let Op_Operador(ByVal New_Opcion As Boolean)
m_Opcion_Operador = New_Opcion
End Property
Public Property Let Op_TipoGranalla(ByVal New_Opcion As Boolean)
m_Opcion_TipoGranalla = New_Opcion
End Property
Public Property Let Op_Maquina(ByVal New_Opcion As Boolean)
m_Opcion_Maquina = New_Opcion
End Property
Public Property Let Op_Turno(ByVal New_Opcion As Integer)
m_Opcion_Turno = New_Opcion
End Property
Public Property Let Op_TipoPieza(ByVal New_Opcion As Integer)
m_Opcion_TipoPieza = New_Opcion
End Property
Public Property Let Op_NCArea(ByVal New_Opcion As Integer)
m_Opcion_NCArea = New_Opcion
End Property
Public Property Let Op_PP(ByVal New_Opcion As Integer)
m_Opcion_PP = New_Opcion
End Property

Private Sub btnFechaRango_Click()
' pone la fecha desde y hasta
Dim dia_ini As Integer, dia_fin As Integer, mes_ini As Integer, mes_fin As Integer, ano_ini As Integer, ano_fin As Integer
'dia = Day(Date)

If Fecha_Ini.Text <> Fecha_Vacia Or Fecha_Fin.Text <> Fecha_Vacia Then

    Fecha_Ini.Text = Fecha_Vacia
    Fecha_Fin.Text = Fecha_Vacia
    
    btnFechaRango.Caption = "22 / 21"
    
Else

    ano_fin = Year(Date)
    ano_ini = ano_fin
    mes_fin = Month(Date)
    
    mes_ini = mes_fin - 1
    If mes_ini = 0 Then ' enero
        mes_ini = 12
        ano_ini = ano_ini - 1
    End If
    
    dia_ini = 22
    dia_fin = 21
    
    Fecha_Ini.Text = Format(dia_ini & "/" & mes_ini & "/" & ano_ini, Fecha_Format)
    Fecha_Fin.Text = Format(dia_fin & "/" & mes_fin & "/" & ano_fin, Fecha_Format)
    
    btnFechaRango.Caption = "Limpiar"
    
End If

End Sub

Private Sub btnPrdBuscar_Click()
Dim Obj As String, Objs As String
MousePointer = vbHourglass
Obj = "Producto"
Objs = "Productos"
Search.Muestra data_file, "Productos", "Codigo", "Descripcion", Obj, Objs

Producto_Codigo = Search.Codigo
PrdDescripcion.Caption = Search.Descripcion

MousePointer = vbDefault

End Sub
Private Sub btnRABuscar_Click()
Dim Obj As String, Objs As String
MousePointer = vbHourglass
Obj = "Trabajador"
Objs = "Trabajadores"
Search.Muestra data_file, "Trabajadores", "RUT", "ApPaterno]+' '+[ApMaterno]+' '+[Nombres", Obj, Objs, "trabajadores.chklst_responsable"
ResponsableArea_RUT = Search.Codigo
ResponsableArea_Nombre = Search.Descripcion
RANombre.Caption = ResponsableArea_Nombre
MousePointer = vbDefault
End Sub

Private Sub btnTrBuscar_Click()

Dim Obj As String, Objs As String
MousePointer = vbHourglass
Obj = "Trabajador"
Objs = "Trabajadores"
Search.Muestra data_file, "Trabajadores", "RUT", "ApPaterno]+' '+[ApMaterno]+' '+[Nombres", Obj, Objs

Trabajador_RUT = Search.Codigo
Trabajador_Nombre = Search.Descripcion
TrNombre.Caption = Trabajador_Nombre

MousePointer = vbDefault
End Sub
Private Sub btnOpBuscar_Click()

Dim Obj As String, Objs As String
MousePointer = vbHourglass
Obj = "Operador"
Objs = "Operadores"
Search.Muestra data_file, "Trabajadores", "RUT", "ApPaterno]+' '+[ApMaterno]+' '+[Nombres", Obj, Objs ', Condicion

Operador_RUT = Search.Codigo
Operador_Nombre = Search.Descripcion
OpNombre.Caption = Operador_Nombre

MousePointer = vbDefault
End Sub
Private Sub btnPrvBuscar_Click()

Dim Obj As String, Objs As String
MousePointer = vbHourglass
Obj = "Proveedor"
Objs = "Proveedores"
Search.Muestra data_file, "Proveedores", "rut", "razon social", Obj, Objs

Proveedor_RUT = Search.Codigo
PrvDescripcion.Caption = Search.Descripcion

MousePointer = vbDefault

End Sub
'///////////
Private Sub Form_Load()

Inicializa

Set Dbm = OpenDatabase(mpro_file)

If m_Opcion_Nv Then

    nvListar False
    
    ' Combo obra
    ComboNv.AddItem " "
    For i = 1 To nvTotal
        a_Nv(i, 0) = aNv(i).Numero
        a_Nv(i, 1) = aNv(i).obra
        ComboNv.AddItem Format(aNv(i).Numero, "0000") & " - " & aNv(i).obra
    Next
    
    'OpNV(0).Value = True
    ComboNv.Text = " "
    
End If

If m_Opcion_CC Then
  
    ' Combo Centro Costo
    ComboCC.AddItem " "
    For i = 0 To scpNew_aCeCo_size - 1
        ComboCC.AddItem scpNew_aCeCo(i, 0) & " - " & scpNew_aCeCo(i, 1)
    Next
    
    OpCC(0).Value = True
    
End If

If m_Opcion_Plano Then

    Set RsPc = Dbm.OpenRecordset("Planos Cabecera")
    RsPc.Index = "NV-Plano"
    
    ComboPl.AddItem " "
    OpPl(0).Value = True
    
End If

Select Case True
Case m_Opcion_OT

    Set RsOTc = Dbm.OpenRecordset("OT Fab Cabecera")
    RsOTc.Index = "Numero"
    
    ComboOT.AddItem " "
    OpOT(0).Value = True
    
Case m_Opcion_Producto

    ' opcion producto
'    Frame_Producto.Visible = m_Opcion_Producto
    OpPrd(0).Value = True
    OpPrd(0).visible = m_Opcion_Producto
    OpPrd(1).visible = m_Opcion_Producto
    btnPrdBuscar.visible = m_Opcion_Producto
    PrdDescripcion.visible = m_Opcion_Producto
    
Case m_Opcion_TipoGD

'    Frame_TipoGuia.Visible = True
'    OpTG(0).Value = True

Case m_Opcion_TipoGranalla

    CbTipoGranalla.AddItem "SP2"
    CbTipoGranalla.AddItem "SP3"
    CbTipoGranalla.AddItem "SP5"
    CbTipoGranalla.AddItem "SP6"
    CbTipoGranalla.AddItem "SP7"
    CbTipoGranalla.AddItem "SP10"
    
    OpTGr(0).Value = True
    OpTGr(0).visible = m_Opcion_TipoGranalla
    OpTGr(1).visible = m_Opcion_TipoGranalla
    CbTipoGranalla.visible = m_Opcion_TipoGranalla

Case m_Opcion_Maquina

    OpMaq(0).Value = True
    OpMaq(0).visible = m_Opcion_Maquina
    OpMaq(1).visible = m_Opcion_Maquina
    OpMaq(2).visible = m_Opcion_Maquina

End Select

End Sub
Private Sub Inicializa()

Operador_RUT = ""

Dim RsCla As Recordset

Set DbD = OpenDatabase(data_file)

lblInforme.Caption = m_Titulo

Numero_Frames = 0

Frame_NV.visible = False
Frame_CC.visible = False
Frame_SC.visible = False
Frame_Cl.visible = False
Frame_TR.visible = False

Frame_Prov.visible = False
Frame_TProv.visible = False
Frame_Numero.visible = False

Frame_PR.visible = False
Frame_Fecha.visible = False
Frame_TI.visible = False
Frame_PL.visible = False
Frame_OT.visible = False
Frame_TG.visible = False
Frame_chklstArea.visible = False
Frame_RA.visible = False
Frame_OP.visible = False

Frame_TipoGranalla.visible = False
Frame_Maquina.visible = False

Frame_TPza.visible = False

Frame_Turno.visible = False

If m_Opcion_Nv Then
    Numero_Frames = Numero_Frames + 1
    Frame_NV.Top = (Numero_Frames - 1) * 960 + 600
    Frame_NV.visible = True
End If

If m_Opcion_CC Then
    Numero_Frames = Numero_Frames + 1
    Frame_CC.Top = (Numero_Frames - 1) * 960 + 600
    Frame_CC.visible = True
End If

' primero veo si el usuario NO es contratista
If Usuario.Tipo <> "C" Then
    If m_Opcion_Contratista Then
        Numero_Frames = Numero_Frames + 1
        Frame_SC.Top = (Numero_Frames - 1) * 960 + 600
        OpSc(0).Value = True
        Frame_SC.visible = True
    End If
End If

If m_Opcion_Cliente Then
    Numero_Frames = Numero_Frames + 1
    Frame_Cl.Top = (Numero_Frames - 1) * 960 + 600
    OpCl(0).Value = True
    Frame_Cl.visible = True
End If

If m_Opcion_Trabajador Then
    Numero_Frames = Numero_Frames + 1
    Frame_TR.Top = (Numero_Frames - 1) * 960 + 600
    OpTr(0).Value = True
    Frame_TR.visible = True
End If

If m_Opcion_Proveedor Then
    Numero_Frames = Numero_Frames + 1
    Frame_Prov.Top = (Numero_Frames - 1) * 960 + 600
    OpProv(0).Value = True
    Frame_Prov.visible = True
End If

If m_Opcion_ProveedorTipo Then

    Numero_Frames = Numero_Frames + 1
    Frame_TProv.Top = (Numero_Frames - 1) * 960 + 600
    OpTProv(0).Value = True
    Frame_TProv.visible = True
    
    Set RsCla = DbD.OpenRecordset("Clasificacion de Proveedores")
    RsCla.Index = "Codigo"
    
    CbClasif.AddItem " "
    With RsCla
    Do While Not .EOF
        CbClasif.AddItem !Codigo
        .MoveNext
    Loop
    End With
        
End If

If m_Opcion_Numero Then
    
    Numero_Frames = Numero_Frames + 1
    Frame_Numero.Top = (Numero_Frames - 1) * 960 + 600
        
    Numero_Ini.Text = ""
    Numero_Fin.Text = ""
    
    Frame_Numero.visible = True
    
End If

If m_Opcion_Producto Then
    Numero_Frames = Numero_Frames + 1
    Frame_PR.Top = (Numero_Frames - 1) * 960 + 600
    Frame_PR.visible = True
End If

If m_Opcion_Fecha Then
    Numero_Frames = Numero_Frames + 1
    Frame_Fecha.Top = (Numero_Frames - 1) * 960 + 600
    Frame_Fecha.visible = True
    Fecha_Ini.Mask = "##/##/##"
    Fecha_Fin.Mask = "##/##/##"
    
    btnFechaRango.visible = m_Opcion_BotonRango
    
End If

If m_Opcion_TipoInforme Then
    Numero_Frames = Numero_Frames + 1
    OpTI(0).Value = True
    Frame_TI.Top = (Numero_Frames - 1) * 960 + 600
    Frame_TI.visible = True
End If

If m_Opcion_Plano Then
    Numero_Frames = Numero_Frames + 1
    Frame_PL.Top = (Numero_Frames - 1) * 960 + 600
    Frame_PL.visible = True
End If

If m_Opcion_OT Then
    Numero_Frames = Numero_Frames + 1
    Frame_OT.Top = (Numero_Frames - 1) * 960 + 600
    Frame_OT.visible = True
End If

If m_Opcion_TipoGD Then
    Numero_Frames = Numero_Frames + 1
    Frame_TG.Top = (Numero_Frames - 1) * 960 + 600
    OpTG(0).Value = True
    Frame_TG.visible = True
    
    ComboTG.AddItem ""
    ComboTG.AddItem "Normal"
    ComboTG.AddItem "Especial"
    ComboTG.AddItem "Galvanizado"
    ComboTG.AddItem "Pernos"
    
End If

If m_Opcion_chklstArea Then

    Numero_Frames = Numero_Frames + 1
    Frame_chklstArea.Top = (Numero_Frames - 1) * 960 + 600
    Op_clArea(0).Value = True
    Frame_chklstArea.visible = True

    Set RsAreas = DbD.OpenRecordset("chklst_Areas")

    With RsAreas
    .Index = "codigo"

    CbchklstAreas.AddItem ""
    a_Areas(0) = ""
    
    i = 0
    Do While Not .EOF
        CbchklstAreas.AddItem !Descripcion
        i = i + 1
        a_Areas(i) = !Codigo
        .MoveNext
    Loop
    .Close
    End With

End If

If m_Opcion_chklstResponsableArea Then
    Numero_Frames = Numero_Frames + 1
    Frame_RA.Top = (Numero_Frames - 1) * 960 + 600
    OpRA(0).Value = True
    Frame_RA.visible = True
End If

If m_Opcion_Operador Then
    Numero_Frames = Numero_Frames + 1
    Frame_OP.Top = (Numero_Frames - 1) * 960 + 600
    OpOp(0).Value = True
    Frame_OP.visible = True
End If

If m_Opcion_TipoGranalla Then
    Numero_Frames = Numero_Frames + 1
    Frame_TipoGranalla.Top = (Numero_Frames - 1) * 960 + 600
    OpTGr(0).Value = True
    Frame_TipoGranalla.visible = True
End If

If m_Opcion_Maquina Then
    Numero_Frames = Numero_Frames + 1
    Frame_Maquina.Top = (Numero_Frames - 1) * 960 + 600
    OpMaq(0).Value = True
    Frame_Maquina.visible = True
End If

If m_Opcion_TipoPieza Then
    
    CbTipoPieza.AddItem " "
    CbTipoPieza.AddItem "VIGA"
    CbTipoPieza.AddItem "TUBULAR"
    CbTipoPieza.AddItem "TUBEST"
    CbTipoPieza.AddItem "PLANCHA"
    
    Numero_Frames = Numero_Frames + 1
    Frame_TPza.Top = (Numero_Frames - 1) * 960 + 600
    OpTPza(0).Value = True
    Frame_TPza.visible = True
    
End If

If m_Opcion_Turno Then
    Numero_Frames = Numero_Frames + 1
    Frame_Turno.Top = (Numero_Frames - 1) * 960 + 600
    OpTurno(0).Value = True
    Frame_Turno.visible = True
End If

If m_Opcion_NCArea Then

    Dim RsMae As ADODB.Recordset, sql As String
    Set RsMae = New ADODB.Recordset
    sql = "SELECT * FROM maestros WHERE tipo='GER'"
    RsMae.Open sql, CnxSqlServer_scp0

    Numero_Frames = Numero_Frames + 1
    Frame_NCGerencia.Top = (Numero_Frames - 1) * 960 + 600
    OpNCGerencia(0).Value = True
    Frame_NCGerencia.visible = True

    CbchklstAreas.AddItem ""
    
    i = 0
    With RsMae
    Do While Not .EOF
        i = i + 1
        a_Areas(i) = "OPE"
        CbNCGerencia.AddItem RsMae!Descripcion
        RsMae.MoveNext
    Loop
    End With
    
End If

If m_Opcion_PP Then
    Numero_Frames = Numero_Frames + 1
    Frame_PP.Top = (Numero_Frames - 1) * 960 + 600
    OpPP(0).Value = True
    Frame_PP.visible = True
   
End If

btnAceptar.Top = Numero_Frames * 960 + 700
btnCancelar.Top = Numero_Frames * 960 + 700

Me.Height = Numero_Frames * 960 + 1800

End Sub

Private Sub OpCC_Click(Index As Integer)
If Index = 0 Then
    ' todos los centros de costo
    MousePointer = vbHourglass
    ComboCC.Text = " "
    ComboCC.Enabled = False
    MousePointer = vbDefault
Else
    ' un centro de costo
    ComboCC.Enabled = True
    ComboCC.SetFocus
End If
End Sub
Private Sub OpNCGerencia_Click(Index As Integer)
If Index = 0 Then
    CbNCGerencia.ListIndex = -1
    CbNCGerencia.Enabled = False
Else
    CbNCGerencia.Enabled = True
End If
End Sub
' viene de otro
'///////////////////////////////////////
Private Sub OpNV_Click(Index As Integer)
If Index = 0 Then
    ' todas las NV
    MousePointer = vbHourglass
    Nv.Text = ""
    Nv.Enabled = False
    ComboNv.Text = " "
    ComboNv.Enabled = False
    MousePointer = vbDefault
    
    If m_Opcion_Plano Then
        Planos_Todos
    End If
    
    If m_Opcion_OT Then
        OT_Todas
    End If
    
Else

    ' una NV
    
    Nv.Enabled = True
    ComboNv.Enabled = True
    Nv.SetFocus
    
'    ComboNv.Enabled = True
'    ComboNv.SetFocus
'    SendKeys "%{Down}"
    
End If

End Sub
Private Sub OpNV_KeyPress(Index As Integer, KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Nv_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Nv_LostFocus()
Dim m_Nv As Integer
m_Nv = Val(Nv.Text)
If m_Nv = 0 Then Exit Sub
' busca nv en combo
i = 1
Do Until a_Nv(i, 0) = ""
    If Val(a_Nv(i, 0)) = m_Nv Then
        ComboNv.ListIndex = i
        Exit Sub
    End If
    i = i + 1
Loop

MsgBox "NV no existe"
Nv.SetFocus

End Sub
Private Sub ComboNV_Click()

If m_Opcion_Plano Then
    If ComboNv.Text <> " " Then
    
        ComboPl.Clear
        ComboPl.AddItem " "
        
        RsPc.Seek ">=", Val(Left(ComboNv.Text, 4)), ""
        If Not RsPc.NoMatch Then
            Do While Not RsPc.EOF
                If Val(Left(ComboNv.Text, 4)) <> RsPc!Nv Then Exit Do
                ComboPl.AddItem RsPc!Plano & " , " & RsPc!Rev
                RsPc.MoveNext
            Loop
        End If
    End If
End If

If m_Opcion_OT Then
'    If ComboOT.Text <> " " Then
    
        ComboOT.Clear
        ComboOT.AddItem " "
        RsOTc.MoveFirst
        Do While Not RsOTc.EOF
            If Val(Left(ComboNv.Text, 4)) = RsOTc!Nv Then
                ComboOT.AddItem Format(RsOTc!Numero, "0000") & " , " & RsOTc!fecha
            End If
            RsOTc.MoveNext
        Loop
'    End If
End If

End Sub
Private Sub OpPrd_Click(Index As Integer)
Producto_Codigo = ""
If Index = 0 Then
    btnPrdBuscar.Enabled = False
    PrdDescripcion.Caption = ""
Else
    btnPrdBuscar.Enabled = True
    btnPrdBuscar_Click
'    Producto_Codigo = "" '?
End If
End Sub
Private Sub OpRA_Click(Index As Integer)
If Index = 0 Then
    btnRABuscar.Enabled = False
    ResponsableArea_RUT = ""
    ResponsableArea_Nombre = ""
    RANombre.Caption = ""
Else
    btnRABuscar.Enabled = True
    btnRABuscar_Click
End If
End Sub
Private Sub OpSc_Click(Index As Integer)
If Index = 0 Then
    btnScBuscar.Enabled = False
    ScRazon.Caption = ""
Else
    btnScBuscar.Enabled = True
    btnScBuscar_Click
End If
End Sub
Private Sub OpCl_Click(Index As Integer)
If Index = 0 Then
    btnClBuscar.Enabled = False
    ClRazon.Caption = ""
Else
    btnClBuscar.Enabled = True
    btnClBuscar_Click
End If
End Sub
Private Sub OpProv_Click(Index As Integer)
Proveedor_RUT = ""
If Index = 0 Then
    btnPrvBuscar.Enabled = False
    PrvDescripcion.Caption = ""
Else
    btnPrvBuscar.Enabled = True
    btnPrvBuscar_Click
End If
End Sub
Private Sub OpTPza_Click(Index As Integer)

If Index = 0 Then
    CbTipoPieza.Text = " "
    CbTipoPieza.Enabled = False
Else
    CbTipoPieza.Enabled = True
    CbTipoPieza.SetFocus
'    SendKeys "%{Down}"
End If

End Sub
Private Sub OpPl_Click(Index As Integer)
If Index = 0 Then
    ComboPl.Text = " "
    ComboPl.Enabled = False
Else
    ComboPl.Enabled = True
    ComboPl.SetFocus
'    SendKeys "%{Down}"
End If
End Sub
Private Sub Planos_Todos()

ComboPl.Clear
ComboPl.AddItem " "

If RsPc.RecordCount > 0 Then
RsPc.MoveFirst
Do While Not RsPc.EOF
    ComboPl.AddItem RsPc!Plano & " , " & RsPc!Rev
    RsPc.MoveNext
Loop
End If

End Sub
Private Sub OT_Todas()

ComboOT.Clear
ComboOT.AddItem " "

RsOTc.MoveFirst
Do While Not RsOTc.EOF
    ComboOT.AddItem Format(RsOTc!Numero, "0000") & " , " & RsOTc!fecha
    RsOTc.MoveNext
Loop

End Sub
Private Sub btnScBuscar_Click()
Dim Obj As String, Objs As String
MousePointer = vbHourglass

Obj = "Contratista"
Objs = "Contratistas"

Dim arreglo(1) As String
arreglo(1) = "razon_social"

sql_Search.Muestra "personas", "RUT", arreglo(), Obj, Objs, "contratista='S'"
Contratista_Rut = sql_Search.Codigo
Contratista_Razon = sql_Search.Descripcion
ScRazon.Caption = sql_Search.Descripcion
    

MousePointer = vbDefault

End Sub
Private Sub btnClBuscar_Click()
Dim Obj As String, Objs As String
MousePointer = vbHourglass
Obj = "Cliente"
Objs = "Clientes"
Search.Muestra data_file, "Clientes", "RUT", "Razon Social", Obj, Objs

Cliente_RUT = Search.Codigo
Cliente_Razon = Search.Descripcion
ClRazon.Caption = Cliente_Razon

MousePointer = vbDefault
End Sub
Private Sub Fecha_Ini_GotFocus()
Fecha_Ini.SelStart = 0
End Sub
Private Sub Fecha_Ini_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
'    SendKeys "{Tab}"
    SendKeysA vbKeyTab, True
'    Fecha_Fin.SetFocus
End If
End Sub
Private Sub Fecha_Ini_LostFocus()
d = Fecha_Valida(Fecha_Ini) ', Fecha_Vacia)
End Sub
Private Sub Fecha_Fin_GotFocus()
Fecha_Fin.SelStart = 0
End Sub
Private Sub Fecha_Fin_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
'    SendKeys "{Tab}"
    SendKeysA vbKeyTab, True
'    btnAceptar.SetFocus
End If
End Sub
Private Sub Fecha_Fin_LostFocus()
d = Fecha_Valida(Fecha_Fin) ', Fecha_Vacia)
End Sub
Private Sub OpOT_Click(Index As Integer)
If Index = 0 Then
    ComboOT.Text = " "
    ComboOT.Enabled = False
Else
    ComboOT.Enabled = True
    ComboOT.SetFocus
'    SendKeys "%{Down}"
End If
End Sub
Private Sub btnAceptar_Click()

If Valida = False Then Exit Sub

Dim p As Integer, formula As String
Dim m_TGr As String, m_Maq As String, m_Tur As Integer
Dim pp As String

MousePointer = vbHourglass

If OpNV(1).Value = True Then
    NV_Numero = Left(ComboNv.Text, 4)
    NV_nombre = Mid(ComboNv.Text, 7)
Else
    NV_Numero = 0
    NV_nombre = ""
End If

If OpCC(1).Value = True Then
    p = ComboCC.ListIndex
    ccCodigo = scpNew_aCeCo(p, 0)
    ccDescripcion = scpNew_aCeCo(p, 1)
Else
    ccCodigo = -1
    ccDescripcion = ""
End If

If OpSc(1).Value = True Then
Else
    Contratista_Rut = ""
    Contratista_Razon = ""
End If

If OpCl(1).Value = True Then
Else
    Cliente_RUT = ""
    Cliente_Razon = ""
End If

If OpTr(1).Value = True Then

Else
    Trabajador_RUT = ""
    Trabajador_Nombre = ""
End If

If OpPl(1).Value = True Then
    p = InStr(1, ComboPl.Text, ",")
    Plano = Trim(Left(ComboPl.Text, p - 1))
    Revision = Mid(ComboPl.Text, p + 1)
Else
    Plano = ""
    Revision = ""
End If

If OpOT(1).Value = True Then
    OT_Numero = Left(ComboOT.Text, 4)
Else
    OT_Numero = 0
End If

If OpTG(1).Value = True Then
    m_TipoGuia = Left(ComboTG.Text, 1)
Else
    m_TipoGuia = ""
End If

Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Rpt")
'Printer.PaperSize = vbPRPSLetter
'Printer.PaperSize = vbPRPSLegal
'vbPRPSLetter    1   Carta, 216 x 279 mm
'Printer.PaperSize = vbPRPSLetterSmall

CR.WindowTitle = m_Titulo
CR.DataFiles(0) = repo_file & ".MDB"
CR.Formulas(0) = "RAZON SOCIAL=""" & Empresa.Razon & """"
CR.Formulas(1) = "TITULO=""" & m_Titulo & """"
CR.ReportSource = crptReport
CR.SelectionFormula = ""

CR.Formulas(2) = ""
CR.Formulas(3) = ""

CR.DataFiles(0) = ""
CR.DataFiles(1) = ""
CR.DataFiles(2) = ""
CR.DataFiles(3) = ""
CR.DataFiles(4) = ""
CR.DataFiles(5) = ""

If Contratista_Rut <> "" Then Contratista_Rut = SqlRutPadL(Contratista_Rut)

Select Case m_Titulo

Case "PLANOS"
    If OpTI(0).Value Then
        'general
        formula = Repo_Planos_General(NV_Numero, Plano)
        CR.SelectionFormula = formula
        CR.DataFiles(0) = mpro_file & ".MDB"
        CR.DataFiles(1) = mpro_file & ".MDB"
        CR.ReportFileName = Drive_Server & Path_Rpt & "PlanosCabecera.Rpt"
        CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
        CR.Action = 1
    Else
        'detalle
        Repo_Planos_Detalle NV_Numero, Plano
        CR.SelectionFormula = ""
        CR.DataFiles(0) = repo_file & ".MDB"
        CR.ReportFileName = Drive_Server & Path_Rpt & "PlanosDetalle.Rpt"
        CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
        CR.Action = 1
    End If

Case "REVISIONES DE PLANOS"

    Repo_Planos_Detalle_Revisiones NV_Numero
    CR.SelectionFormula = ""
    CR.DataFiles(0) = repo_file & ".MDB"
    CR.ReportFileName = Drive_Server & Path_Rpt & "planosdetalle_rev.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "OT FABRICACIÓN"
    
    If Usuario.Tipo = "C" Then
        Contratista_Rut = Usuario.Rut
    End If
    
    If OpTI(0).Value Then
        'general
        Repo_OTf NV_Numero, Contratista_Rut, Fecha_Ini, Fecha_Fin
        CR.DataFiles(0) = repo_file & ".MDB"
        CR.ReportFileName = Drive_Server & Path_Rpt & "OTfc.Rpt"
        CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
        CR.Action = 1
    Else
        ' detalle
        formula = F_Repo_OTfd(NV_Numero, Contratista_Rut, Fecha_Ini, Fecha_Fin)
        CR.SelectionFormula = formula
        CR.DataFiles(0) = mpro_file & ".MDB"
        CR.DataFiles(1) = mpro_file & ".MDB"
        CR.DataFiles(2) = mpro_file & ".MDB"
        CR.DataFiles(3) = mpro_file & ".MDB"
        CR.DataFiles(4) = data_file & ".MDB"
        CR.ReportFileName = Drive_Server & Path_Rpt & "OTfd_v2.Rpt" 'version2
        CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
        CR.Action = 1
    End If
    
Case "OT ARENADO"

    CR.DataFiles(0) = data_file & ".MDB"
    CR.DataFiles(1) = mpro_file & ".MDB"
    CR.DataFiles(2) = mpro_file & ".MDB"

    formula = MiFormula_Resto("OT Are Cabecera.NV", NV_Numero, "OT Are Cabecera.RUT Contratista", Contratista_Rut, "", "", "", "", "OT Are Cabecera.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "", 0, "", "", "")
    CR.SelectionFormula = formula
    
    CR.ReportFileName = Drive_Server & Path_Rpt & "OTac.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "OT PINTURA"

    CR.DataFiles(0) = data_file & ".MDB"
    CR.DataFiles(1) = mpro_file & ".MDB"
    CR.DataFiles(2) = mpro_file & ".MDB"

    formula = MiFormula_Resto("OT Pin Cabecera.NV", NV_Numero, "OT Pin Cabecera.RUT Contratista", Contratista_Rut, "", "", "", "", "OT Pin Cabecera.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "", 0, "", "", "")
    CR.SelectionFormula = formula

    CR.ReportFileName = Drive_Server & Path_Rpt & "OTpc.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "OT ESPECIAL"

    Repo_OTe NV_Numero, Contratista_Rut, Fecha_Ini, Fecha_Fin, "OTe"
    CR.DataFiles(0) = repo_file & ".MDB"
    CR.ReportFileName = Drive_Server & Path_Rpt & "OTe.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "ITO FABRICACIÓN"

    If Usuario.Tipo = "C" Then
        Contratista_Rut = Usuario.Rut
    End If

'    If OpTipo(0).Value Then
    If True Then
        
        'general
        If OpTI(0).Value Then
                
            Repo_ITOc "Fab", NV_Numero, Contratista_Rut, Fecha_Ini, Fecha_Fin ', ""
            
            CR.DataFiles(0) = repo_file & ".MDB"
            
            CR.Formulas(2) = "PERIODO=""" & "Desde: " & Fecha_Ini.Text & " Hasta: " & Fecha_Fin.Text & """"
    '        CR.Formulas(3) = "UNIDADES=""" & "" & """"
            
            CR.ReportFileName = Drive_Server & Path_Rpt & "ITOfc_new.Rpt"
        
        Else
        
            Repo_ITOfd NV_Numero, Contratista_Rut, Fecha_Ini, Fecha_Fin, "FAB"
            
            CR.DataFiles(0) = repo_file & ".MDB"
            
            CR.Formulas(2) = "condicion1=""" & "Desde: " & Fecha_Ini.Text & " Hasta: " & Fecha_Fin.Text & """"
            
            CR.ReportFileName = Drive_Server & Path_Rpt & "ITOfd.Rpt"
            
        End If
    
    Else
    
        CR.Formulas(2) = "PERIODO=""" & "Desde: " & Fecha_Ini.Text & " Hasta: " & Fecha_Fin.Text & """"
'        CR.Formulas(3) = "UNIDADES=""" & "$/Kg" & """"
        
        CR.ReportFileName = Drive_Server & Path_Rpt & "ITOfd.Rpt"
    
    End If
    
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
'Case "ITO GALVANIZADO" ' 04/11/04
Case "ITO REPROCESO" ' solo cambio de nombre

    Repo_ITOc "Gal", NV_Numero, Contratista_Rut, Fecha_Ini, Fecha_Fin ', ""
    
    CR.DataFiles(0) = repo_file & ".MDB"
    
    CR.Formulas(2) = "PERIODO=""" & "Desde: " & Fecha_Ini.Text & " Hasta: " & Fecha_Fin.Text & """"
    CR.Formulas(3) = "UNIDADES=""" & "m2 Total" & """"
    CR.ReportFileName = Drive_Server & Path_Rpt & "ITOfc.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "ITO PINTURA"  ' 04/11/04

    If Operador_RUT = "" Then
        Repo_ITOc "Pin", NV_Numero, Contratista_Rut, Fecha_Ini, Fecha_Fin ', ""
    Else
        Repo_ITOd "Pin", NV_Numero, Contratista_Rut, Fecha_Ini, Fecha_Fin, Operador_RUT, "", "", 0
    End If

    If Operador_RUT = "" Then
    
        If OpTI(0).Value Then
            CR.DataFiles(0) = repo_file & ".MDB"
            CR.Formulas(2) = "PERIODO=""" & "Desde: " & Fecha_Ini.Text & " Hasta: " & Fecha_Fin.Text & """"
            CR.Formulas(3) = "UNIDADES=""" & "m2 Total" & """"
            CR.ReportFileName = Drive_Server & Path_Rpt & "ITOfc.Rpt"
        Else
            Repo_ITOfd NV_Numero, Contratista_Rut, Fecha_Ini, Fecha_Fin, "PYG"
            CR.DataFiles(0) = repo_file & ".MDB"
            CR.Formulas(2) = "condicion1=""" & "Desde: " & Fecha_Ini.Text & " Hasta: " & Fecha_Fin.Text & """"
            CR.ReportFileName = Drive_Server & Path_Rpt & "ITOfd.Rpt"
        End If

    Else

        CR.DataFiles(0) = repo_file & ".MDB"
        CR.ReportFileName = Drive_Server & Path_Rpt & "ITOp_det.Rpt"

    End If
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "ITO GRANALLADO"

    If OpTGr(0).Value Then
        ' todos
        m_TGr = ""
    Else
        m_TGr = CbTipoGranalla.Text
    End If
    
    If OpMaq(0).Value Then
        m_Maq = ""
    End If
    If OpMaq(1).Value Then
        m_Maq = "M"
    End If
    If OpMaq(2).Value Then
        m_Maq = "A"
    End If
    
    m_Tur = 0
    Select Case True
    Case OpTurno(1).Value
        m_Tur = 1
    Case OpTurno(2).Value
        m_Tur = 2
    End Select
    
    Repo_ITOd "Gra", NV_Numero, Contratista_Rut, Fecha_Ini, Fecha_Fin, Operador_RUT, m_TGr, m_Maq, m_Tur
    
    CR.DataFiles(0) = repo_file & ".MDB"
    CR.ReportFileName = Drive_Server & Path_Rpt & "ITOr_det.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "ITO GRANALLADO ESPECIAL"  ' 18/10/07

    If OpTGr(0).Value Then
        ' todos
        m_TGr = ""
    Else
        m_TGr = CbTipoGranalla.Text
    End If
    
    If OpMaq(0).Value Then
        m_Maq = ""
    End If
    If OpMaq(1).Value Then
        m_Maq = "M"
    End If
    If OpMaq(2).Value Then
        m_Maq = "A"
    End If
    
    m_Tur = 0
    Select Case True
    Case OpTurno(1).Value
        m_Tur = 1
    Case OpTurno(2).Value
        m_Tur = 2
    End Select
    
    
    'Repo_ITOd "GraEsp", NV_Numero, Contratista_Rut, Fecha_Ini, Fecha_Fin, Operador_RUT, m_TGr, m_Maq, m_Tur
    Repo_ITOd "GraEsp", NV_Numero, "", Fecha_Ini, Fecha_Fin, Operador_RUT, m_TGr, m_Maq, m_Tur
        
    CR.DataFiles(0) = repo_file & ".MDB"
    CR.ReportFileName = Drive_Server & Path_Rpt & "ITOresp_det.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "ITO PRODUCCION PINTURA"

'    Repo_ITOd "Gra", NV_Numero, Contratista_Rut, Fecha_Ini, Fecha_Fin, Operador_RUT, m_TGr, m_Maq, m_Tur
    Repo_ITOd "pp", NV_Numero, Contratista_Rut, Fecha_Ini, Fecha_Fin, Operador_RUT, "", "", 0
    
    CR.DataFiles(0) = repo_file & ".MDB"
    CR.ReportFileName = Drive_Server & Path_Rpt & "ITOpp_det.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "ITO PRODUCCION PINTURA ESPECIAL"

    'Repo_ITOd "ppesp", NV_Numero, "", Fecha_Ini, Fecha_Fin, Operador_RUT, "", "", 0
    Repo_ITOd "ppesp", NV_Numero, "", Fecha_Ini, Fecha_Fin, Operador_RUT, "", "", 0
    
    CR.DataFiles(0) = repo_file & ".MDB"
    CR.ReportFileName = Drive_Server & Path_Rpt & "ITOppe_det.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "ITO ESPECIAL"

'    Repo_ITOs NV_numero, SubContratista_RUT, Fecha_Ini, Fecha_Fin
    CR.DataFiles(0) = mpro_file & ".MDB"
    CR.DataFiles(1) = mpro_file & ".MDB"
    CR.DataFiles(2) = data_file & ".MDB"
    
    CR.ReportFileName = Drive_Server & Path_Rpt & "ITOe.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "GUÍAS DE DESPACHO"

    If OpTI(0).Value = True Then

        CR.DataFiles(0) = data_file & ".MDB"
        CR.DataFiles(1) = mpro_file & ".MDB"
        CR.DataFiles(2) = mpro_file & ".MDB"

        formula = MiFormula_Resto("GD Cabecera.NV", NV_Numero, "", "", "GD Cabecera.RUT Cliente", Cliente_RUT, "", "", "GD Cabecera.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "", 0, "", "", m_TipoGuia)
        CR.SelectionFormula = formula

        CR.ReportFileName = Drive_Server & Path_Rpt & "Gd.Rpt"
        CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
        CR.Action = 1

    Else

        MsgBox "Vea Informes -> Piezas Depachadas (Detalle)"

    End If

Case "CORRELATIVO DE GUÍAS"

'    If OpTipo(0).Value Then ' general
'    Else ' detalle

    Repo_GDs_Correlativos Fecha_Ini.Text, Fecha_Fin.Text
    
    CR.DataFiles(0) = repo_file & ".MDB"
    CR.Formulas(0) = "RAZON SOCIAL=""" & EmpOC.Razon & """"
    CR.ReportFileName = Drive_Server & Path_Rpt & "GDCorre.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "BULTOS"

    CR.DataFiles(0) = mpro_file & ".MDB"
    CR.DataFiles(1) = mpro_file & ".MDB"
    CR.DataFiles(2) = mpro_file & ".MDB"

    formula = MiFormula_Resto("bultos.NV", NV_Numero, "", "", "", "", "", "", "bultos.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "", 0, "", "", "")
    CR.SelectionFormula = formula
    
    CR.Formulas(2) = "CONDICION1=""" & Condicion.FechaInicial & Condicion.FechaFinal & """"
    
    CR.ReportFileName = Drive_Server & Path_Rpt & "bultos.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "FACTURAS DE VENTA"

    CR.DataFiles(0) = data_file & ".MDB"
    CR.DataFiles(1) = mvta_file & ".MDB"
    
    formula = MiFormula_Resto("", 0, "", "", "FAV Cabecera.RUT Cliente", Cliente_RUT, "", "", "FAV Cabecera.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "", 0, "", "", "")
    CR.SelectionFormula = formula
    
    CR.Formulas(1) = "CONDICION1=""" & Condicion.Cliente & """"
    CR.Formulas(2) = "CONDICION2=""" & Condicion.FechaInicial & """"
    CR.Formulas(3) = "CONDICION3=""" & Condicion.FechaFinal & """"
    
    CR.ReportFileName = Drive_Server & Path_Rpt & "Facturas.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "AVANCES DE PAGO"

'    AvancesPago NV_Numero, NV_nombre, Contratista_Rut, Contratista_Razon, "", "Fab", Fecha_Ini.Text, Fecha_Fin.Text
    
    CR.DataFiles(0) = repo_file & ".MDB"
    CR.DataFiles(1) = mpro_file & ".MDB"
    CR.DataFiles(2) = data_file & ".MDB"
    
    CR.ReportFileName = Drive_Server & Path_Rpt & "AvancedePago.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "PIEZAS PENDIENTES ORDENADAS POR PLANO"
    If (OpPP(0).Value) Then
        pp = "T"
    End If
    If (OpPP(1).Value) Then
        pp = "P"
    End If
    PiezasPendientes NV_Numero, Plano, "", pp

    CR.DataFiles(0) = repo_file & ".MDB"

    CR.ReportFileName = Drive_Server & Path_Rpt & "PiezasPendientes_new.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
Case "PIEZAS PENDIENTES ORDENADAS POR DESCRIPCION"
    
    If (OpPP(0).Value) Then
        pp = "T"
    End If
    If (OpPP(1).Value) Then
        pp = "P"
    End If
    
    PiezasPendientes NV_Numero, Plano, Contratista_Rut, pp

    CR.DataFiles(0) = repo_file & ".MDB"

    CR.ReportFileName = Drive_Server & Path_Rpt & "PiezasPendientes_xdesc_new.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "PIEZAS POR FABRICAR" ' ASIGNAR"

    PiezasxAsignar NV_Numero, Plano
    
    CR.DataFiles(0) = repo_file & ".MDB"
'    cr.ReportFileName = Drive_Server & path_rpt & "PiezasxAsignar.Rpt"
    CR.ReportFileName = Drive_Server & Path_Rpt & "PiezasxFabricar.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "PIEZAS EN FABRICACIÓN" ' POR RECIBIR"

    PiezasxRecibir NV_Numero, Contratista_Rut, Plano, Fecha_Ini, Fecha_Fin
    
'    cr.ReportFileName = Drive_Server & path_rpt & "PiezasxRecibir.Rpt"
    CR.DataFiles(0) = repo_file & ".MDB"

    CR.ReportFileName = Drive_Server & Path_Rpt & "PiezasenFabricacion.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "PIEZAS FABRICADAS EN NEGRO" ' POR DESPACHAR"

    PiezasxDespachar NV_Numero, Plano, True, "ennegro"
    
    CR.DataFiles(0) = repo_file & ".MDB"
'    cr.ReportFileName = Drive_Server & path_rpt & "PiezasxDespachar.Rpt"
    CR.ReportFileName = Drive_Server & Path_Rpt & "PiezasenNegro.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "PIEZAS FABRICADAS (DETALLE)"

    CR.DataFiles(0) = mpro_file & ".MDB"
    CR.DataFiles(1) = mpro_file & ".MDB"
    
    formula = MiFormula_Resto("ito fab detalle.nv", Val(Nv.Text), "", "", "", "", "", "", "ito fab detalle.fecha", Fecha_Ini.Text, Fecha_Fin.Text, "", 0, "", "", "")
    CR.SelectionFormula = formula
    
    CR.Formulas(2) = "CONDICION1=""" & "NV: " & Nv.Text & """"
    
    CR.ReportFileName = Drive_Server & Path_Rpt & "itof_det.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
'Case "PIEZAS EN GALVANIZADO"
Case "PIEZAS EN REPROCESO"

'    PiezasenGalvanizado NV_Numero, SubContratista_RUT, Plano, Fecha_Ini, Fecha_Fin
    PiezasxDespachar NV_Numero, Plano, False, "engalvanizado"
    
    CR.DataFiles(0) = repo_file & ".MDB"
'    cr.ReportFileName = Drive_Server & path_rpt & "PiezasxRecibir.Rpt"
    CR.ReportFileName = Drive_Server & Path_Rpt & "PiezasenGalvanizado.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
'Case "PIEZAS PINTADAS O GALVANIZADAS" ' POR DESPACHAR"
'Case "PIEZAS PINTADAS O EN REPROCESO" ' POR DESPACHAR"
Case "PIEZAS PINTADAS O REPROCESADAS"

    PiezasxDespachar NV_Numero, Plano, True, "pintadas"
    
    CR.DataFiles(0) = repo_file & ".MDB"
    CR.ReportFileName = Drive_Server & Path_Rpt & "PiezasxDespachar.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "ETIQUETAS PIEZAS PINTADAS O GALVANIZADAS" ' POR DESPACHAR"

    PiezasxDespachar NV_Numero, Plano, True, "pintadas"
    
Case "PIEZAS DESPACHADAS"

    PiezasxDespachar NV_Numero, Plano, False, ""
    
    CR.DataFiles(0) = repo_file & ".MDB"
    CR.ReportFileName = Drive_Server & Path_Rpt & "PiezasDespachadas.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "PIEZAS DESPACHADAS (DETALLE) OLD"

    CR.DataFiles(0) = mpro_file & ".MDB"
    CR.DataFiles(1) = mpro_file & ".MDB"
    CR.DataFiles(2) = mpro_file & ".MDB"
    CR.DataFiles(3) = mpro_file & ".MDB"
    
    formula = MiFormula_Resto("gd detalle.nv", Val(Nv.Text), "", "", "", "", "", "", "gd detalle.fecha", Fecha_Ini.Text, Fecha_Fin.Text, "", 0, "", "", "")
    CR.SelectionFormula = formula
    
'    Cr.Formulas(2) = "CONDICION1=""" & "NV: " & Nv.Text & """"
    CR.Formulas(2) = "CONDICION1=""" & "NV: " & ComboNv.Text & """"
    
    CR.ReportFileName = Drive_Server & Path_Rpt & "gd_detOLD.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "PIEZAS DESPACHADAS (DETALLE)"

    ' 25/03/2014
    
    Repo_gd_detalle NV_Numero ', Plano, False, "engalvanizado"

    CR.DataFiles(0) = repo_file & ".MDB"
    
'    Cr.Formulas(2) = "CONDICION1=""" & "NV: " & Nv.Text & """"
    CR.Formulas(2) = "CONDICION1=""" & "NV: " & ComboNv.Text & """"
    
    CR.ReportFileName = Drive_Server & Path_Rpt & "gd_det.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "OTf x Plano" 'gral y deta

    CR.DataFiles(0) = repo_file & ".MDB"

    If OpTI(0).Value Then
        'General
        Repo_OTfxPlanoG NV_Numero, Plano, Contratista_Rut
        CR.ReportFileName = Drive_Server & Path_Rpt & "OTfxPlanoG.Rpt"
    Else
        'Detalle
        Repo_OTfxPlanoD NV_Numero, Plano, Contratista_Rut
        CR.ReportFileName = Drive_Server & Path_Rpt & "OTfxPlanoD.Rpt"
    End If
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "ITOf x OTf" 'gyd

    CR.DataFiles(0) = repo_file & ".MDB"

    If OpTI(0).Value Then
        'General
        Repo_ITOs_de_OTs_General NV_Numero, Contratista_Rut, OT_Numero
    
        CR.ReportFileName = Drive_Server & Path_Rpt & "ITOfxOTg.Rpt"
    
    Else
        'Detalle
        Repo_ITOs_de_OTs_Detalle NV_Numero, Contratista_Rut, OT_Numero
        
        CR.ReportFileName = Drive_Server & Path_Rpt & "ITOfxOTd.Rpt"
    End If
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "BONO DE PRODUCCIÓN"

    Repo_BonoProd NV_Numero, Contratista_Rut, Fecha_Ini.Text, Fecha_Fin.Text
    CR.DataFiles(0) = repo_file & ".MDB"
    CR.Formulas(2) = "PERIODO=""DESDE " & Fecha_Ini.Text & " HASTA " & Fecha_Fin.Text & """"
    CR.ReportFileName = Drive_Server & Path_Rpt & "BonoProd.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "PRODUCCIÓN MENSUAL"

    If Fecha_Ini.Text = "__/__/__" Then
        MsgBox "Debe Digitar Fecha de Inicio"
        Fecha_Ini.SetFocus
    Else

'        Repo_Produccion_Mensual_EML Fecha_Ini.Text, Fecha_Fin.Text
        Repo_Produccion_Mensual Fecha_Ini.Text, Fecha_Fin.Text
        CR.DataFiles(0) = repo_file & ".MDB"
        
'        Cr.ReportFileName = Drive_Server & path_rpt & "ProduccionMensual_6c.Rpt" ' 05/01/06
        CR.ReportFileName = Drive_Server & Path_Rpt & "ProduccionMensual_9c.Rpt" ' 08/11/06
        CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
        CR.Action = 1
    
    End If

Case "INFORME KILOS POR OBRA" ' junio 2004

    Repo_Planos_GralxObra 'NV_Numero
    CR.SelectionFormula = ""
    
    CR.DataFiles(0) = repo_file & ".MDB"
'    CR.ReportFileName = Drive_Server & path_rpt & "planoskilos_new.rpt"
    If Nv_Index = "Numero" Then
        CR.ReportFileName = Drive_Server & Path_Rpt & "planoskilos.rpt"
    Else
        CR.ReportFileName = Drive_Server & Path_Rpt & "planoskilos_xd.rpt" ' 02/03/06
    End If
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "INFORME GENERAL NV" ' junio 2004

    If NV_Numero <> 0 Then

        'general
'        Repo_OTf NV_Numero, SubContratista_RUT, Fecha_Ini, Fecha_Fin
        
'        Repo_OTe NV_Numero, SubContratista_RUT, Fecha_Ini, Fecha_Fin, "OTe"

        CR.DataFiles(0) = repo_file & ".MDB"
        
        Repo_OcxTipo NV_Numero
        
        Repo_GeneralNv NV_Numero, Fecha_Ini.Text, Fecha_Fin.Text
        
        CR.Formulas(2) = "Obra=""" & "Obra: " & ComboNv.Text & """"
    '    Cr.ReportFileName = Drive_Server & path_rpt & "OTfcgral.Rpt"
        CR.ReportFileName = Drive_Server & Path_Rpt & "General_Nv.Rpt"
        CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
        CR.Action = 1
    
    Else
        MsgBox "Debe Escoger Nv"
    End If

Case "VALES DE CONSUMO X CONTRATISTA" ' julio 2004

    formula = MiFormula_Resto("Documentos.NV", NV_Numero, "Documentos.RUT", Contratista_Rut, "", "", "", "", "Documentos.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "", 0, "Documentos.Codigo Producto", Producto_Codigo, "")
    CR.SelectionFormula = formula

    CR.DataFiles(0) = Madq_file & ".MDB"
    CR.DataFiles(1) = mpro_file & ".MDB"
    CR.DataFiles(2) = data_file & ".MDB"
    CR.DataFiles(3) = data_file & ".MDB"
    
    CR.ReportFileName = Drive_Server & Path_Rpt & "vc_contratista.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "VALES DE CONSUMO X TRABAJADOR" ' mayo 2007

    formula = MiFormula_Resto("Documentos.NV", NV_Numero, "Documentos.RUT", Contratista_Rut, "", "", "", "", "Documentos.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "", 0, "Documentos.Codigo Producto", Producto_Codigo, "", Trabajador_RUT)
    CR.SelectionFormula = formula

    CR.DataFiles(0) = Madq_file & ".MDB"
    CR.DataFiles(1) = mpro_file & ".MDB"
    CR.DataFiles(2) = data_file & ".MDB"
    CR.DataFiles(3) = data_file & ".MDB"
    
    CR.ReportFileName = Drive_Server & Path_Rpt & "vc_trabajador.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "VALES DE CONSUMO X FECHA" ' agosto 2004

    formula = MiFormula_Resto("Documentos.NV", NV_Numero, "Documentos.RUT", Contratista_Rut, "", "", "", "", "Documentos.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "", 0, "Documentos.Codigo Producto", Producto_Codigo, "")
    CR.SelectionFormula = formula

    CR.DataFiles(0) = Madq_file & ".MDB"
    CR.DataFiles(1) = mpro_file & ".MDB"
    CR.DataFiles(2) = data_file & ".MDB"
    CR.DataFiles(3) = data_file & ".MDB"
    
    CR.ReportFileName = Drive_Server & Path_Rpt & "vc_fecha.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "DIGITA FACTURAS X GUIA"

    MousePointer = vbHourglass
    DigGdFav.Numero_NotaVenta = NV_Numero
    DigGdFav.Rut_Cliente = Cliente_RUT
    DigGdFav.Fecha_Inicio = Fecha_Ini.Text
    DigGdFav.Fecha_Termino = Fecha_Fin.Text
    Load DigGdFav
    MousePointer = vbDefault
    DigGdFav.Show 1

Case "DIGITA FACTURAS X VALE CONSUMO"

    MousePointer = vbHourglass
    DigVcFav.Numero_NotaVenta = NV_Numero
    DigVcFav.Rut_Contratista = Contratista_Rut
    DigVcFav.Fecha_Inicio = Fecha_Ini.Text
    DigVcFav.Fecha_Termino = Fecha_Fin.Text
    Load DigVcFav
    MousePointer = vbDefault
    DigVcFav.Show 1

Case "RECEPCION DE PLANOS" ' 21/04/05

    formula = MiFormula_Resto("Planos Recepcion.NV", NV_Numero, "Planos Recepcion.RUT Contratista", Contratista_Rut, "", "", "", "", "Planos Recepcion.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "", 0, "", "", "")
    CR.SelectionFormula = formula

    CR.DataFiles(0) = mpro_file & ".MDB"
    CR.DataFiles(1) = mpro_file & ".MDB"
    CR.DataFiles(2) = data_file & ".MDB"
    
    CR.ReportFileName = Drive_Server & Path_Rpt & "planosrecepcion.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "NV"

'    Repo_NV Cliente_RUT, Fecha_Ini.Text, Fecha_Fin.Text
    
    formula = MiFormula_Resto("", 0, "", "", "nv cabecera.rut cliente", Cliente_RUT, "", "", "nv cabecera.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "", 0, "", "", "")
    CR.SelectionFormula = formula

'    cr.WindowTitle = "Notas de Venta"
    
    CR.DataFiles(0) = data_file & ".MDB"
    CR.DataFiles(1) = mpro_file & ".MDB"
    
'    cr.Formulas(0) = "RAZON SOCIAL=""" & Empresa.Razon & """"
'    cr.Formulas(1) = "TITULO=""" & "NOTAS DE VENTA" & """"
'    cr.ReportSource = crptReport
'    cr.WindowState = crptMaximized
'    cr.WindowMaxButton = False
'    cr.WindowMinButton = False

    CR.ReportFileName = Drive_Server & Path_Rpt & "Nv.Rpt"
'    Cr.ReportFileName = Drive_Server & Path_Rpt & "nv_v2.Rpt" ' version 2
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "SERVICIOS EXTERNOS" ' 13/12/05

    formula = MiFormula_Resto("se cabecera.NV", NV_Numero, "", "", "se cabecera.rut_cliente", Cliente_RUT, "", "", "se cabecera.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "", 0, "", "", "")
    CR.SelectionFormula = formula

    CR.DataFiles(0) = mpro_file & ".MDB"
    CR.DataFiles(1) = data_file & ".MDB"
    
    CR.ReportFileName = Drive_Server & Path_Rpt & "sec.rpt" ' serv ext "cabecera"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "ITO FAB VALORIZADA" ' 14/03/06

    Repo_ItoVal NV_Numero, Contratista_Rut, Fecha_Ini.Text, Fecha_Fin.Text

    CR.Formulas(2) = "PERIODO=""" & "Desde: " & Fecha_Ini.Text & " Hasta: " & Fecha_Fin.Text & """"

    CR.DataFiles(0) = repo_file & ".MDB"
    CR.DataFiles(1) = data_file & ".MDB"
    
    CR.ReportFileName = Drive_Server & Path_Rpt & "itoval.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "FABRICACION Y DESPACHO" ' para fco cruces

    Repo_FabricacionyDespacho
    MsgBox "ok despus de repo"

Case "NV X CLIENTE" ' 31/08/06

'    formula = MiFormula("", 0, "", "", "nv cabecera.rut cliente", Cliente_RUT, "", "", "nv cabecera.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "", 0, "", "")
'    Cr.SelectionFormula = formula
    
'    que muestre solo activas Cristina barria para don Erwin 02/10/06
    Repo_NVxCliente Cliente_RUT, Fecha_Ini.Text, Fecha_Fin.Text
    
    CR.DataFiles(0) = repo_file & ".MDB"
    
    CR.Formulas(2) = "cliente=""" & ClRazon.Caption & """"
    CR.SelectionFormula = ""
    CR.ReportFileName = Drive_Server & Path_Rpt & "nvxcliente.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "CHECK LIST"

    If CbchklstAreas.ListIndex > -1 Then
        Repo_ChkLst a_Areas(CbchklstAreas.ListIndex), ResponsableArea_RUT, Fecha_Ini.Text, Fecha_Fin.Text
    Else
        Repo_ChkLst "", ResponsableArea_RUT, Fecha_Ini.Text, Fecha_Fin.Text
    End If

'    CR.SelectionFormula = ""
    CR.ReportFileName = Drive_Server & Path_Rpt & "chklst.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "CHECK LIST OBSERVACIONES"

    If CbchklstAreas.ListIndex > -1 Then
        Repo_ChkLstObs a_Areas(CbchklstAreas.ListIndex), ResponsableArea_RUT, Fecha_Ini.Text, Fecha_Fin.Text
    Else
        Repo_ChkLstObs "", ResponsableArea_RUT, Fecha_Ini.Text, Fecha_Fin.Text
    End If

'    CR.SelectionFormula = ""
    CR.ReportFileName = Drive_Server & Path_Rpt & "chklst_obs.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "CHECK LIST FINAL"

    If CbchklstAreas.ListIndex > -1 Then
        Repo_ChkLstFinal a_Areas(CbchklstAreas.ListIndex), ResponsableArea_RUT, Fecha_Ini.Text, Fecha_Fin.Text
    Else
        Repo_ChkLstFinal "", ResponsableArea_RUT, Fecha_Ini.Text, Fecha_Fin.Text
    End If

'    CR.SelectionFormula = ""
    CR.ReportFileName = Drive_Server & Path_Rpt & "chklstfinal.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "PAGO CONTRATISTAS"

    MousePointer = vbHourglass
    PagoContratistas.Contratista_Rut = Contratista_Rut
    PagoContratistas.Contratista_Nombre = Contratista_Razon
    PagoContratistas.Fecha_Inicio = Fecha_Ini.Text
    PagoContratistas.Fecha_Termino = Fecha_Fin.Text
    Load PagoContratistas
    MousePointer = vbDefault
    PagoContratistas.Show 1

Case "OT MANTENCION"

    formula = MiFormula_Resto("", 0, "ote cabecera.rut contratista", Contratista_Rut, "", "", "", "", "ote cabecera.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "", 0, "", "", "")
    CR.SelectionFormula = formula

    CR.DataFiles(0) = mpro_file & ".MDB"
    CR.DataFiles(1) = data_file & ".MDB"

'    CR.SelectionFormula = ""
    CR.ReportFileName = Drive_Server & Path_Rpt & "otec.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "ARCO SUMERGIDO"

    Dim Texto As String, Logico As Boolean
    qry = "": p = 0
    Filtro_Nv "arco sumergido.NV", NV_Numero, p
    Filtro_Fecha "arco sumergido.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, p
    Filtro_String "arco sumergido.rut operador1", Operador_RUT, p
    Texto = CbTipoPieza.ListIndex
    If Texto <> "0" Then
        Texto = Mid("VTSP", CbTipoPieza.ListIndex, 1)
        Filtro_String "arco sumergido.tipopieza", Texto, p
    End If
    If OpTurno(0).Value = False Then
        If OpTurno(1).Value Then
            Texto = "D"
        Else
            Texto = "N"
        End If
        Filtro_String "arco sumergido.turno", Texto, p
    End If
    
    CR.SelectionFormula = qry
    
    CR.DataFiles(0) = mpro_file & ".MDB"
    CR.DataFiles(1) = mpro_file & ".MDB"
    CR.DataFiles(2) = data_file & ".MDB"
    CR.DataFiles(3) = data_file & ".MDB"

    'Cr.ReportFileName = Drive_Server & Path_Rpt & "as.rpt"
    CR.ReportFileName = Drive_Server & Path_Rpt & "asV2.rpt" ' desde 29/08/13
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "RESUMEN PIEZAS X TURNO"

    Repo_as_PiezasxTurno NV_Numero, Fecha_Ini.Text, Fecha_Fin.Text, Operador_RUT, OpNombre.Caption
    
    CR.DataFiles(0) = repo_file & ".MDB"
    
    If OpNV(0).Value Then
        CR.Formulas(2) = "CONDICION1=""" & "NV: Todas" & """"
    Else
        CR.Formulas(2) = "CONDICION1=""" & "NV: " & ComboNv.Text & """"
    End If
    If OpOp(0).Value Then
        CR.Formulas(3) = "CONDICION2=""" & "Operador: Todos" & """"
    Else
        CR.Formulas(3) = "CONDICION2=""" & "Operador: " & OpNombre.Caption & """"
    End If
    If Fecha_Ini.Text <> "__/__/__" Or Fecha_Fin.Text <> "__/__/__" Then
        CR.Formulas(4) = "CONDICION3=""" & "Periodo: desde " & Fecha_Ini.Text & " hasta " & Fecha_Fin.Text & """"
    Else
        CR.Formulas(4) = "CONDICION3=""" & "" & """"
    End If
    
    CR.ReportFileName = Drive_Server & Path_Rpt & "as_piezasxturno.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "BONO ARCO SUMERGIDO"

    Repo_as_Bono NV_Numero, Fecha_Ini.Text, Fecha_Fin.Text, Operador_RUT, OpNombre.Caption
    
    CR.DataFiles(0) = repo_file & ".MDB"
    
If False Then

    If OpNV(0).Value Then
        CR.Formulas(2) = "CONDICION1=""" & "NV: Todas" & """"
    Else
        CR.Formulas(2) = "CONDICION1=""" & "NV: " & ComboNv.Text & """"
    End If
    If OpOp(0).Value Then
        CR.Formulas(3) = "CONDICION2=""" & "Operador: Todos" & """"
    Else
        CR.Formulas(3) = "CONDICION2=""" & "Operador: " & OpNombre.Caption & """"
    End If

End If

    If Fecha_Ini.Text <> "__/__/__" Or Fecha_Fin.Text <> "__/__/__" Then
        CR.Formulas(2) = "CONDICION1=""" & "Periodo: desde " & Fecha_Ini.Text & " hasta " & Fecha_Fin.Text & """"
    Else
        CR.Formulas(2) = "CONDICION1=""" & "" & """"
    End If

    CR.ReportFileName = Drive_Server & Path_Rpt & "as_bono.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "STOCK"

    If Fecha_Fin.Text = "__/__/__" Then Fecha_Fin.Text = Date
    Stock_Recalcular Fecha_Fin.Text
    
    CR.DataFiles(0) = data_file & ".MDB"

    CR.ReportFileName = Drive_Server & Path_Rpt & "prd_stock.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

'///////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////
Case "ÓRDENES DE COMPRA"
    
    If OpTI(0).Value Then
    
        'general
        CR.DataFiles(0) = data_file & ".MDB"
        CR.DataFiles(1) = Madq_file & ".MDB"
        CR.DataFiles(2) = mpro_file & ".MDB"
        
        formula = MiFormula_Adq("OC Cabecera.NV", NV_Numero, "", "", "OC Cabecera.RUT Proveedor", Proveedor_RUT, CbClasif.Text, "OC Cabecera.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "OC Cabecera.Numero", Val(Numero_Ini.Text), Val(Numero_Fin.Text))
        
        If Val(Numero_Ini.Text) <> 0 Or Val(Numero_Fin.Text) <> 0 Then
            CR.SelectionFormula = formula
            CR.ReportFileName = Drive_Server & Path_Rpt & "occ_xn.rpt"
        Else
            If Len(formula) > 1 Then formula = formula & " AND NOT {OC Cabecera.Nula}"
            CR.SelectionFormula = formula
            CR.ReportFileName = Drive_Server & Path_Rpt & "Oc.Rpt"
        End If
    
    Else
    
        CR.DataFiles(0) = Madq_file & ".MDB"
        CR.DataFiles(1) = Madq_file & ".MDB"
        CR.DataFiles(2) = data_file & ".MDB"
        CR.DataFiles(3) = data_file & ".MDB"
        
        formula = MiFormula_Adq("OC Detalle.NV", NV_Numero, "OC Detalle.Codigo Producto", Producto_Codigo, "OC Detalle.RUT Proveedor", Proveedor_RUT, CbClasif.Text, "OC Detalle.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "OC detalle.Numero", Val(Numero_Ini.Text), Val(Numero_Fin.Text))
'        formula = MiFormula("OC Cabecea.NV", NV_Numero, "", "", "OC Cabecera.RUT Proveedor", Proveedor_RUT, "OC Cabecera.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "OC Cabecera.Numero", Val(Numero_Ini.Text), Val(Numero_Fin.Text))
        
        CR.SelectionFormula = formula
        
        CR.ReportFileName = Drive_Server & Path_Rpt & "Oc_Det.Rpt"
        
    End If
   
'    CR.Destination = crptToFile   ' 1/6/2000
'    CR.PrintFileType = crptHTML30 ' 1/6/2000
'    CR.PrintFileName = "OC2"      ' 1/6/2000
    
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "RECEPCIÓN DE MATERIALES"

    CR.DataFiles(0) = Madq_file & ".MDB"
    CR.DataFiles(1) = Madq_file & ".MDB"
    CR.DataFiles(2) = mpro_file & ".MDB"
    CR.DataFiles(3) = data_file & ".MDB"
    CR.DataFiles(4) = data_file & ".MDB"
    CR.DataFiles(5) = Madq_file & ".MDB"
    
    formula = MiFormula_Adq("Documentos.NV", NV_Numero, "Documentos.Codigo Producto", Producto_Codigo, "Documentos.RUT", Proveedor_RUT, CbClasif.Text, "Documentos.Fecha", Fecha_Ini.Text, Fecha_Fin.Text)
    CR.SelectionFormula = formula
    
    CR.ReportFileName = Drive_Server & Path_Rpt & "rm.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "PRODUCTOS POR RECIBIR"

    CR.DataFiles(0) = data_file & ".MDB"
    CR.DataFiles(1) = data_file & ".MDB"
    CR.DataFiles(2) = Madq_file & ".MDB"
    CR.DataFiles(3) = mpro_file & ".MDB"
    CR.DataFiles(4) = ""

    formula = MiFormula_Adq("OC Detalle.NV", NV_Numero, "OC Detalle.Codigo Producto", Producto_Codigo, "OC Detalle.RUT Proveedor", Proveedor_RUT, CbClasif.Text, "OC Detalle.Fecha", Fecha_Ini.Text, Fecha_Fin.Text)
    formula = IIf(formula = "", "{OC Detalle.Cantidad} > {OC Detalle.Cantidad Recibida}", formula & " AND {OC Detalle.Cantidad} > {OC Detalle.Cantidad Recibida}")
    formula = formula & " AND {OC Detalle.Pendiente}"
    CR.SelectionFormula = formula

    CR.ReportFileName = Drive_Server & Path_Rpt & "ProdxRecibir.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "RECEPCION DE CERTIFICADOS"

    MousePointer = vbHourglass
    DigCertRecib2.Numero_NotaVenta = NV_Numero
    DigCertRecib2.Rut_Proveedor = Proveedor_RUT
    DigCertRecib2.Fecha_Inicio = Fecha_Ini.Text
    DigCertRecib2.Fecha_Termino = Fecha_Fin.Text
    Load DigCertRecib2
    MousePointer = vbDefault
    DigCertRecib2.Show 1

Case "CERTIFICADOS POR RECIBIR OLD"

    CR.DataFiles(0) = Madq_file & ".MDB"
    CR.DataFiles(1) = data_file & ".MDB"
    CR.DataFiles(2) = mpro_file & ".MDB"
    CR.DataFiles(3) = Madq_file & ".MDB"
    ' OLD
    formula = MiFormula_Adq("OC Cabecera.NV", NV_Numero, "", "", "OC Cabecera.RUT Proveedor", Proveedor_RUT, "", "OC Cabecera.Fecha", Fecha_Ini.Text, Fecha_Fin.Text)
    If Len(formula) > 1 Then
        formula = formula & " AND {OC Cabecera.Certificado} AND NOT {OC Cabecera.Certificado Recibido}"
    Else
        formula = "{OC Cabecera.Certificado} AND NOT {OC Cabecera.Certificado Recibido}"
    End If
    ' OLD
    CR.SelectionFormula = formula
        
    CR.ReportFileName = Drive_Server & Path_Rpt & "oc_cert_pend.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1
    
Case "CERTIFICADOS POR RECIBIR"
    
    CR.DataFiles(0) = data_file & ".MDB"
    CR.DataFiles(1) = data_file & ".MDB"
    CR.DataFiles(2) = Madq_file & ".MDB"
    CR.DataFiles(3) = Madq_file & ".MDB"
    CR.DataFiles(4) = Madq_file & ".MDB"

    formula = MiFormula_Adq("Documentos.NV", NV_Numero, "Documentos.Codigo Producto", Producto_Codigo, "Documentos.RUT", Proveedor_RUT, CbClasif.Text, "Documentos.Fecha", Fecha_Ini.Text, Fecha_Fin.Text)
    formula = IIf(formula = "", "NOT {Documentos.certificadoRecibido}", formula & " AND NOT {Documentos.certificadoRecibido}")
    formula = formula & " AND {Documentos.tipo}='RM' AND {OC Cabecera.certificado}"
    CR.SelectionFormula = formula

    CR.ReportFileName = Drive_Server & Path_Rpt & "certificadosXRecibir.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "INFORME DE PINTURA"

    PiezasPendientes NV_Numero, Plano, Contratista_Rut, "T"
    
    CR.DataFiles(0) = repo_file & ".MDB"

    CR.ReportFileName = Drive_Server & Path_Rpt & "repo_pintura.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "ÓRDENES DE COMPRA DINÁMICO"
    
    If OpTI(0).Value Then
    
        'general
        CR.DataFiles(0) = data_file & ".MDB"
        CR.DataFiles(1) = Madq_file & ".MDB"
        CR.DataFiles(2) = mpro_file & ".MDB"
        
        formula = MiFormula_Adq("OC Cabecera.NV", NV_Numero, "", "", "OC Cabecera.RUT Proveedor", Proveedor_RUT, CbClasif.Text, "OC Cabecera.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "OC Cabecera.Numero", Val(Numero_Ini.Text), Val(Numero_Fin.Text))
        
        If Val(Numero_Ini.Text) <> 0 Or Val(Numero_Fin.Text) <> 0 Then
            CR.SelectionFormula = formula
            CR.ReportFileName = Drive_Server & Path_Rpt & "occ_xn.rpt"
        Else
            If Len(formula) > 1 Then formula = formula & " AND NOT {OC Cabecera.Nula}"
            CR.SelectionFormula = formula
            CR.ReportFileName = Drive_Server & Path_Rpt & "Oc.Rpt"
        End If
    
    Else
    
        CR.DataFiles(0) = Madq_file & ".MDB"
        CR.DataFiles(1) = Madq_file & ".MDB"
        CR.DataFiles(2) = data_file & ".MDB"
        CR.DataFiles(3) = data_file & ".MDB"
        CR.DataFiles(4) = ""
        
        formula = MiFormula_Adq("OC Detalle.NV", NV_Numero, "OC Detalle.Codigo Producto", Producto_Codigo, "OC Detalle.RUT Proveedor", Proveedor_RUT, CbClasif.Text, "OC Detalle.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "OC detalle.Numero", Val(Numero_Ini.Text), Val(Numero_Fin.Text))
'        formula = MiFormula("OC Cabecea.NV", NV_Numero, "", "", "OC Cabecera.RUT Proveedor", Proveedor_RUT, "OC Cabecera.Fecha", Fecha_Ini.Text, Fecha_Fin.Text, "OC Cabecera.Numero", Val(Numero_Ini.Text), Val(Numero_Fin.Text))
        
        CR.SelectionFormula = formula
        
        CR.ReportFileName = Drive_Server & Path_Rpt & "Oc_Det_dinamico.Rpt"
        
    End If
    
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "INFORME NO CONFORMIDAD XXX" ' revisar

    If OpNCGerencia(0).Value Then
        'todas las areas
    Else
    
    End If

'    Repo_inc_xa "OPE", "Gerencia Operaciones", Fecha_Ini.Text, Fecha_Fin.Text

    CR.DataFiles(0) = mpro2_file & ".MDB"
    CR.DataFiles(1) = data_file & ".MDB"
    CR.ReportFileName = Drive_Server & Path_Rpt & "nc_resumen.Rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "INFORME NO CONFORMIDAD"

    If OpTI(0).Value Then
        'general
'    If OpNCGerencia(0).Value Then

        Repo_inc_resumen Fecha_Ini.Text, Fecha_Fin.Text
        
'        formula = " AND NOT {inc_xa.fecha emision}"
'        Cr.SelectionFormula = formula

        CR.DataFiles(0) = repo_file & ".MDB"
        CR.ReportFileName = Drive_Server & Path_Rpt & "nc_resumen.rpt"
        CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
        CR.Action = 1

    Else
        ' detalle
'        Repo_inc_xgd "OPE", "Gerencia Operaciones", Fecha_Ini.Text, Fecha_Fin.Text
        Repo_inc_xgd "", "", Fecha_Ini.Text, Fecha_Fin.Text, Trabajador_RUT

        CR.DataFiles(0) = repo_file & ".MDB"
        CR.ReportFileName = Drive_Server & Path_Rpt & "nc_gerenciadetalle_1506.rpt"
        CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
        CR.Action = 1

    End If

Case "INFORME DE PIEZAS"

    Repo_Piezas NV_Numero, Contratista_Rut ', Fecha_Ini, Fecha_Fin
    
    CR.DataFiles(0) = repo_file & ".MDB"
    CR.ReportFileName = Drive_Server & Path_Rpt & "piezas.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

Case "GRANALLADO PENDIENTE"
  
    CR.DataFiles(0) = mpro_file & ".MDB"
    CR.DataFiles(1) = mpro_file & ".MDB"
    
    formula = MiFormula_Resto("planos detalle.nv", NV_Numero, "", "", "", "", "", "", "", "__/__/__", "__/__/__", "", 0, "", "", "")
    CR.SelectionFormula = formula

    CR.ReportFileName = Drive_Server & Path_Rpt & "granalladopendiente.rpt"
    CR.WindowTitle = m_Titulo & " " & CR.ReportFileName
    CR.Action = 1

End Select

MousePointer = vbDefault

End Sub
Private Function Valida() As Boolean
Valida = False

If OpNV(1).Value = True Then
    If Trim(ComboNv.Text) = "" Then
        Beep
        MsgBox "DEBE ESCOGER NOTA DE VENTA"
        ComboNv.SetFocus
        Exit Function
    End If
End If

If OpCC(1).Value = True Then
    If Trim(ComboCC.Text) = "" Then
        Beep
        MsgBox "DEBE ESCOGER CENTRO DE COSTO"
        ComboCC.SetFocus
        Exit Function
    End If
End If

If OpSc(1).Value = True Then
    If ScRazon.Caption = "" Then
        Beep
        MsgBox "DEBE ESCOGER CONTRATISTA"
        btnScBuscar.SetFocus
        Exit Function
    End If
End If

If m_Opcion_ContratistaObligatorio Then

    OpSc(1).Value = True
    
    If ScRazon.Caption = "" Then
        Beep
        MsgBox "DEBE ESCOGER CONTRATISTA"
        btnScBuscar.SetFocus
        Exit Function
    End If
End If

If OpCl(1).Value = True Then
    If ClRazon.Caption = "" Then
        Beep
        MsgBox "DEBE ESCOGER CLIENTE"
        btnClBuscar.SetFocus
        Exit Function
    End If
End If

If OpTr(1).Value = True Then
    If TrNombre.Caption = "" Then
        Beep
        MsgBox "DEBE ESCOGER TRABAJADOR"
        btnTrBuscar.SetFocus
        Exit Function
    End If
End If

If OpProv(1).Value = True Then
    If PrvDescripcion.Caption = "" Then
        Beep
        MsgBox "DEBE ESCOGER PROVEEDOR"
        btnPrvBuscar.SetFocus
        Exit Function
    End If
End If

If OpTProv(1).Value = True Then
    If CbClasif.ListIndex = -1 Then
        Beep
        MsgBox "DEBE ESCOGER TIPO DE PROVEEDOR"
        CbClasif.SetFocus
        Exit Function
    End If
End If

If OpPrd(1).Value = True Then
    If PrdDescripcion.Caption = "" Then
        Beep
        MsgBox "DEBE ESCOGER PRODUCTO"
        btnPrdBuscar.SetFocus
        Exit Function
    End If
End If

GoTo Fin
If m_Opcion_Fecha Then
    If Fecha_Ini = Fecha_Vacia Then
        Beep
        MsgBox "DEBE DIGITAR FECHA INICIAL"
        Fecha_Ini.SetFocus
        Exit Function
    End If
    If Fecha_Fin = Fecha_Vacia Then
        Beep
        MsgBox "DEBE DIGITAR FECHA FINAL"
        Fecha_Fin.SetFocus
        Exit Function
    End If
End If
Fin:

If OpPl(1).Value = True Then
    If Trim(ComboPl.Text) = "" Then
        Beep
        MsgBox "DEBE ESCOGER PLANO"
        ComboPl.SetFocus
        Exit Function
    End If
End If

If OpOT(1).Value = True Then
    If Trim(ComboOT.Text) = "" Then
        Beep
        MsgBox "DEBE ESCOGER OT"
        ComboOT.SetFocus
        Exit Function
    End If
End If

If OpTG(1).Value = True Then
    If Trim(ComboTG.Text) = "" Then
        Beep
        MsgBox "DEBE ESCOGER TIPO DE GUIA"
        ComboTG.SetFocus
        Exit Function
    End If
End If

If Op_clArea(1).Value = True Then
    If Trim(CbchklstAreas.Text) = "" Then
        Beep
        MsgBox "DEBE ESCOGER AREA"
        CbchklstAreas.SetFocus
        Exit Function
    End If
End If

If OpTGr(1).Value Then
    If Trim(CbTipoGranalla.Text) = "" Then
        Beep
        MsgBox "DEBE ESCOGER TIPO DE GRANALLADO"
        CbTipoGranalla.SetFocus
        Exit Function
    End If
End If

If OpTPza(1).Value = True Then
    If CbTipoPieza.ListIndex = 0 Then
        Beep
        MsgBox "DEBE ESCOGER TIPO DE PIEZA"
        CbTipoPieza.SetFocus
        Exit Function
    End If
End If

Valida = True
End Function
'///////////////////////////////////////
Private Sub btnCancelar_Click()
Unload Me
End Sub
Private Function MiFormula_Resto(CampoNV As String, Nv As Double, _
                            CampoContratista As String, RUT_SubC As String, _
                            CampoCliente As String, RUT_Client As String, _
                            CampoPlano As String, Plano As String, _
                            CampoFecha As String, F_Inicial As String, F_Final As String, _
                            CampoOT As String, OT As Double, _
                            CampoProducto As String, Prd_Codigo As String, _
                            Tipo_Guia As String, _
                            Optional RUT_Trabajador As String) As String

Dim p As Integer, qry As String, QryN As String
p = 0: qry = ""

' NOTA DE VENTA
If Nv = 0 Then
    Condicion.NotaVenta = "Nota de Venta: Todas"
Else
    qry = " {" & CampoNV & "} = " & Nv
    p = p + 1
    Condicion.NotaVenta = "Nota de Venta:" & Nv
End If

' CONTRATISTA
' ojo:linea agragada por si entra un contratista vea solo sus datos
' 19/04/10
'RUT_SubC = Usuario.Rut

If RUT_SubC = "" Then
    Condicion.Contratista = "Contratista: Todos"
Else
    QryN = " {" & CampoContratista & "} = '" & RUT_SubC & "'"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
    Condicion.Contratista = "Contratista: " & RUT_SubC
End If

' CLIENTE
If RUT_Client = "" Then
    Condicion.Cliente = "Cliente: Todos"
Else
    QryN = " {" & CampoCliente & "} = '" & RUT_Client & "'"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
    Condicion.Cliente = "Cliente: " & RUT_Client
End If

' PLANO
If Plano = "" Then
    Condicion.Plano = "Plano: Todos"
Else
    QryN = " Plano='" & Plano & "'"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
    Condicion.Plano = "Plano: " & Plano
End If

' FECHA INICIAL
If F_Inicial <> "__/__/__" Then
'If F_Inicial <> "" Then
    QryN = " {" & CampoFecha & "}>=" & FechaCristal(F_Inicial)
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If
Condicion.FechaInicial = "Desde : " & F_Inicial

' FECHA FINAL
If F_Final <> "__/__/__" Then
'If F_Final <> "" Then
    QryN = " {" & CampoFecha & "}<=" & FechaCristal(F_Final)
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If
Condicion.FechaFinal = "Hasta : " & F_Final

' OT
If OT = 0 Then
    Condicion.OT = "OT : Todas"
Else
    QryN = " Número=CDbl(" & OT & ")"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
    Condicion.OT = "OT : " & OT
End If

' PRODUCTO
If Prd_Codigo = "" Then
    Condicion.Producto = "Producto: Todos"
Else
    QryN = " {" & CampoProducto & "} = '" & Prd_Codigo & "'"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
    Condicion.Producto = "Producto: " & Prd_Codigo
End If

' TIPO GUIA
If Tipo_Guia = "" Then
Else
    QryN = " {" & "gd cabecera.tipo" & "} = '" & Tipo_Guia & "'"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If

' TRABAJADOR
If RUT_Trabajador = "" Then
    Condicion.Contratista = "Trabajador: Todos"
Else
    QryN = " {" & "documentos.rut" & "} = '" & RUT_Trabajador & "'"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
    Condicion.Contratista = "Trabajador: " & RUT_SubC
End If

MiFormula_Resto = qry

End Function
Private Function FechaCristal(fecha As String) As String
' fecha en formato dd/mm/aa ya esta validada
Dim f As Date
FechaCristal = ""
f = CDate(fecha)
FechaCristal = "DATE(" & Year(f) & "," & Month(f) & "," & Day(f) & ")"
End Function
Private Sub Filtro_Nv(NvCampo As String, NvNumero As Double, p As Integer)
' variable "qry" viene de afuera de esta funcion, "p" tambien, ambas son modificables
' filtro NOTA DE VENTA
Dim QryN As String

If NvNumero > 0 Then
    
    QryN = " {" & NvCampo & "} = " & NvNumero
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
    
End If

End Sub
Private Sub Filtro_Fecha(CampoFecha As String, F_Inicial As String, F_Final As String, p As Integer)
Dim QryN As String
' FECHA INICIAL
If F_Inicial <> "__/__/__" Then
    QryN = " {" & CampoFecha & "}>=" & FechaCristal(F_Inicial)
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If

' FECHA FINAL
If F_Final <> "__/__/__" Then
    QryN = " {" & CampoFecha & "}<=" & FechaCristal(F_Final)
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If

End Sub
Private Sub Filtro_String(NombreCampo As String, Texto As String, p As Integer)
' sirve de filtro para rut operador, rut cliente, rut contratista, etc
' ej filtro_string "ot cabecera.rut contratista",rut_contratista,p
Dim QryN As String
If Texto <> "" Then
    QryN = " {" & NombreCampo & "} = '" & Texto & "'"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If
End Sub
Private Sub xxx()
' usuarios
' Gerencia Gral
' Javier Delgado     javier     jdv22267+

' Operaciones
' Erwin Mariquez     erwin      2830em+
' Sergio Ruz         sergio     51971r*
' Cristina Barria    cristina   cb4141*
' Sebastian Toro     sebastian  sb6913*
' Tamara Lazo        tamara     tm1313*

' Bodegas
' Hector Castro L.   hector     *hc4969
' Manuel Mellado     manuel     mm1068*

' Carlo Hidalgo      carlos     ch6969*
' Jose Gonzalez      josel      *jlge79

' Adquisiciones
' Claudia Delgado    claudia    0771cdv+
' Claudia Zuñiga     claudia2   3055czg*
' Felix Leiva        felix      6947fld*

' Contabilidad
' Francisco Cruces   francisco  fcm12134+
' Maritza Navas      mari       mn1818*
' Carola Jorquera    carola     cj5295*
' Alejandra Alvarado alejandra  aa1321*
' Heriberto Vilches  heri       hv2113*

' Produccion
' Renan Fernandez    renan      rf0563+
' Danilo Hernandez   danilo     *2143dhc
' Pablo Calderon     pablo      6969pa*
' Livio Figueroa     livio      lv1357*
' Carlos Valdez      carlosv    cv9696*
' Marjorie Rebolledo mayo       mr1611*
' Robinson Vilchez   robinsonv  rv2807*

' Inspeccion
' Frank Berrios      frank      *gg757j
' Mirtha Gutierrez   mirtha     *mth217
' Katherine Jorquera kath       kj1515*
' Francisco Cruz                fc1234+

' Despacho
' Alejandro Martinez jano       jm2683*

' Informatica
' Administrador                 -gf1239g+
' Sergio Flores      sergioflo  -sf1313

End Sub

Private Sub OpTG_Click(Index As Integer)
If Index = 0 Then
    ComboTG.ListIndex = -1
    ComboTG.Enabled = False
Else
    ComboTG.Enabled = True
End If
End Sub

Private Sub OpTr_Click(Index As Integer)
If Index = 0 Then
    btnTrBuscar.Enabled = False
    TrNombre.Caption = ""
Else
    btnTrBuscar.Enabled = True
    btnTrBuscar_Click
End If
End Sub
Private Sub Op_clArea_Click(Index As Integer)
If Index = 0 Then
    CbchklstAreas.ListIndex = -1
    CbchklstAreas.Enabled = False
Else
    CbchklstAreas.Enabled = True
End If
End Sub
Private Sub OpOp_Click(Index As Integer)
If Index = 0 Then
    btnOpBuscar.Enabled = False
    Operador_RUT = ""
    OpNombre.Caption = ""
Else
    btnOpBuscar.Enabled = True
    btnOpBuscar_Click
End If
End Sub
Private Sub OpTProv_Click(Index As Integer)
If Index = 0 Then
    CbClasif.ListIndex = -1
    CbClasif.Enabled = False
Else
    CbClasif.Enabled = True
End If
End Sub
Private Function MiFormula_Adq(CampoNV As String, Nv As Double, _
                            CampoProducto As String, Codigo_Prod As String, _
                            CampoProveedor As String, RUT_Prov As String, _
                            ProveedorTipo As String, _
                            CampoFecha As String, F_Inicial As String, F_Final As String, _
                            Optional CampoNumero As String, Optional N_Inicial As Double, Optional N_Final As Double) _
                            As String
                            
Dim p As Integer, qry As String, QryN As String
p = 0: qry = ""

' NOTA DE VENTA
If Nv = 0 Then
    Condicion.NotaVenta = "Nota de Venta: Todas"
Else
    qry = " {" & CampoNV & "} = " & Nv
    p = p + 1
    Condicion.NotaVenta = "Nota de Venta:" & Nv
End If

' PRODUCTO
If Codigo_Prod = "" Then
    Condicion.Producto = "Producto: Todos"
Else
    QryN = " {" & CampoProducto & "} = '" & Codigo_Prod & "'"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
    Condicion.Producto = "Producto: " & Codigo_Prod
End If

' PROVEEDOR
If RUT_Prov = "" Then
    Condicion.Proveedor = "Proveedor:     Todos """
Else
    QryN = " {" & CampoProveedor & "} = '" & RUT_Prov & "'"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
    Condicion.Proveedor = "Proveedor: " & RUT_Prov
End If

If ProveedorTipo = "" Then
'    Condicion.Proveedor = "Proveedor:     Todos """
Else
    QryN = " {Proveedores.Clasificacion} = '" & ProveedorTipo & "'"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
'    Condicion.Proveedor = "Proveedor: " & RUT_Prov
End If

' FECHA INICIAL
If F_Inicial <> "__/__/__" Then
'If F_Inicial <> "" Then
    QryN = " {" & CampoFecha & "}>=" & FechaCristal(F_Inicial)
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If
Condicion.FechaInicial = "Desde : " & F_Inicial

' FECHA FINAL
If F_Final <> "__/__/__" Then
'If F_Final <> "" Then
    QryN = " {" & CampoFecha & "}<=" & FechaCristal(F_Final)
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If
Condicion.FechaFinal = "Hasta : " & F_Final

' NUMERO INICIAL
If Val(N_Inicial) <> 0 Then
    QryN = " {" & CampoNumero & "}>=" & Val(N_Inicial)
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If

' NUMERO FINAL
If Val(N_Final) <> 0 Then
    QryN = " {" & CampoNumero & "}<=" & Val(N_Final)
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If

MiFormula_Adq = qry

End Function
