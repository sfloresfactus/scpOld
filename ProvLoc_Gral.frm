VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form ProvLoc_Gral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Proveedores"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid DBG 
      Bindings        =   "ProvLoc_Gral.frx":0000
      Height          =   3615
      Left            =   120
      OleObjectBlob   =   "ProvLoc_Gral.frx":0016
      TabIndex        =   0
      ToolTipText     =   "Haga doble click sobre el proveedor"
      Top             =   120
      Width           =   9135
   End
   Begin VB.Data Data_Provee 
      Caption         =   "Data_Provee"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Haga doble click sobre el Proveedor"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   5055
   End
End
Attribute VB_Name = "ProvLoc_Gral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Rut As String, m_Razon As String
Private Sub Form_Load()
Data_Provee.DatabaseName = data_file
Data_Provee.RecordSource = "SELECT Rut,[Razon Social],Direccion FROM Proveedores ORDER BY [Razon Social]"
Data_Provee.Refresh
DBG_Config
End Sub
Private Sub DBG_Config()
Data_Provee.ReadOnly = True
DBG.AllowRowSizing = False
DBG.Columns(0).Width = 1000
DBG.Columns(1).Width = 3000
DBG.Columns(2).Width = 4400
DBG.Columns(0).Alignment = 1
End Sub
Private Sub DBG_DblClick()
m_Rut = DBG.Columns(0).Value
m_Razon = DBG.Columns(1).Value
ProvLoc_Lista.Rut = m_Rut
ProvLoc_Lista.Razon = m_Razon
ProvLoc_Lista.Show 1
End Sub
