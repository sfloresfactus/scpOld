VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Product_Search 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de Productos"
   ClientHeight    =   4170
   ClientLeft      =   765
   ClientTop       =   1845
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4170
   ScaleWidth      =   7305
   Begin VB.TextBox Codigo 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton btnCancela 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton btnEscoge 
      Caption         =   "&Escoger"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3600
      Width           =   855
   End
   Begin MSDBGrid.DBGrid Productos 
      Bindings        =   "Product_Search.frx":0000
      Height          =   2895
      Left            =   240
      OleObjectBlob   =   "Product_Search.frx":0016
      TabIndex        =   2
      Top             =   600
      Width           =   6855
   End
   Begin VB.Data Data_ProDes 
      Caption         =   "Data_ProDes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Descripcion 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   2010
   End
End
Attribute VB_Name = "Product_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Codigo As String, m_Condicion As String
Public Property Let Condicion(ByVal New_Value As String)
m_Condicion = New_Value
End Property
Public Property Get CodigoP() As String
CodigoP = m_Codigo
End Property
Private Sub Form_Load()
Data_ProDes.DatabaseName = data_file
End Sub
Private Sub Form_Activate()
Descripcion.Text = "X"
Descripcion.Text = ""
End Sub
Private Sub DBG_Config()
Dim c As Integer, d As Integer
c = 1300: d = 5000
Productos.Width = c + d + 550  '550
Me.Width = Productos.Width + 2 * Productos.Left
Descripcion.Width = d
Productos.Columns(0).Width = c
Productos.Columns(1).Width = d
'Productos.Columns(3).NumberFormat = num_Formato
End Sub
Private Sub Codigo_Change()
If Codigo.Text = "'" Then Codigo.Text = ""
Data_ProDes.RecordSource = "SELECT Codigo,Descripcion FROM Productos WHERE Codigo LIKE '" & Codigo.Text & "*'" & m_Condicion & " ORDER BY Codigo"
Data_ProDes.Refresh
DBG_Config
End Sub
Private Sub Descripcion_Change()
If Descripcion.Text = "'" Then Descripcion.Text = ""
Data_ProDes.RecordSource = "SELECT Codigo,Descripcion FROM Productos WHERE Descripcion LIKE '" & Descripcion.Text & "*'" & m_Condicion & " ORDER BY Descripcion"
Data_ProDes.Refresh
DBG_Config
End Sub
Private Sub Productos_DblClick()
'If Productos.SelBookmarks.Count > 0 Then
    m_Codigo = Productos.Columns(0).Value
    Unload Me
'End If
End Sub
Private Sub btnEscoge_Click()
'If Productos.SelBookmarks.Count > 0 Then
    m_Codigo = Productos.Columns(0).Value
    Unload Me
'End If
End Sub
Private Sub btnCancela_Click()
m_Codigo = ""
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
'sale
End Sub
Private Sub Productos_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
'MsgBox "presiono enter"
    m_Codigo = Productos.Columns(0).Value
    Unload Me
End If
End Sub
