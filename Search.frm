VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Search 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de ..."
   ClientHeight    =   5190
   ClientLeft      =   2145
   ClientTop       =   1950
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5190
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin MSDBCtls.DBList DBList 
      Bindings        =   "Search.frx":0000
      DataSource      =   "Data"
      Height          =   4155
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7329
      _Version        =   327680
   End
   Begin VB.Data Data 
      Caption         =   "Data"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   600
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Texto 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba las primeras letras del OBJ que esta buscando"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Tabla As String
Private m_CampoCodigo As String
Private m_CampoDescripcion As String

Private qry As String
Private m_Condicion As String

' output
Private m_Codigo As String
Private m_Descripcion As String
'////////////////////////////////////////////////////////////////////
Public Property Get Codigo() As String
Codigo = m_Codigo
End Property
Public Property Get Descripcion() As String
Descripcion = m_Descripcion
End Property
'////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
m_Codigo = ""
m_Descripcion = ""
End Sub
'Public Sub Muestra(Archivo As String, Tabla As String, Campo_Codigo As String, Campo_Descripcion As String, Obj As String, Objs As String)
Public Sub Muestra(Archivo As String, Tabla As String, Campo_Codigo As String, Campo_Descripcion As String, Obj As String, Objs As String, Optional Condicion As String)

m_Tabla = Tabla
m_CampoCodigo = Campo_Codigo
m_CampoDescripcion = Campo_Descripcion
m_Condicion = Condicion 'viene sin WHERE

'////////////////////////
Caption = "Búsqueda de " + StrConv(Objs, vbProperCase)
lbl(0).Caption = "Escriba las primeras letras del " & StrConv(Obj, vbProperCase) & " que está buscando"

Data.DatabaseName = Archivo
If m_Condicion = "" Then
qry = "SELECT [" & m_CampoCodigo & "] AS COD,[" & m_CampoDescripcion & "] AS DES FROM [" & m_Tabla & "] ORDER BY [" & m_CampoDescripcion & "]"
Else
qry = "SELECT [" & m_CampoCodigo & "] AS COD,[" & m_CampoDescripcion & "] AS DES FROM [" & m_Tabla & "] WHERE " & m_Condicion & " ORDER BY [" & m_CampoDescripcion & "]"
End If
Data.RecordSource = qry
Data.Refresh

DBList.BoundColumn = "COD" ' m_CampoCodigo
DBList.DataField = "DES"
DBList.ListField = "DES"
'//////////////////////

Me.Show 1

End Sub
Private Sub DBList_DblClick()
Seleccion
End Sub
Private Sub DBList_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Seleccion
End Sub
Private Sub Seleccion()
m_Codigo = DBList.BoundText
m_Descripcion = DBList.Text
Unload Me
End Sub
Private Sub Texto_Change()
If Right(Texto.Text, 1) = "'" Then Texto.Text = Left(Texto.Text, Len(Texto.Text) - 1)
If m_Condicion = "" Then
qry = "SELECT [" & m_CampoCodigo & "] AS COD,[" & m_CampoDescripcion & "] AS DES FROM [" & m_Tabla & "] WHERE [" & m_CampoDescripcion & "] LIKE '" & Texto.Text & "*' ORDER BY [" & m_CampoDescripcion & "]"
Else
qry = "SELECT [" & m_CampoCodigo & "] AS COD,[" & m_CampoDescripcion & "] AS DES FROM [" & m_Tabla & "] WHERE " & m_Condicion & " AND [" & m_CampoDescripcion & "] LIKE '" & Texto.Text & "*' ORDER BY [" & m_CampoDescripcion & "]"
End If
Data.RecordSource = qry
Data.Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
'fin
End Sub
