VERSION 5.00
Begin VB.Form sql_Search 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de ..."
   ClientHeight    =   4965
   ClientLeft      =   2145
   ClientTop       =   1950
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4965
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List 
      Height          =   3960
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4455
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
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "sql_Search"
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

Private m_ContenidoaMostrar As String
Private m_Orden As String

Private RsTabla As New ADODB.Recordset
Private aCodigo(9999) As String, i As Integer
Private CamposNombres(5) As String, NumeroCampos As Integer
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
Public Sub Muestra(Tabla As String, Campo_Codigo As String, arreglo() As String, Obj As String, Objs As String, Optional Condicion As String)

Dim i As Integer ', m_Descripcion As String

m_Orden = ""
NumeroCampos = UBound(arreglo())
m_CampoDescripcion = arreglo(1)
For i = 1 To NumeroCampos
    CamposNombres(i) = arreglo(i)
    m_Orden = m_Orden & "[" & arreglo(i) & "],"
Next
m_Orden = Left(m_Orden, Len(m_Orden) - 1)

m_Tabla = Tabla
m_CampoCodigo = Campo_Codigo

m_Condicion = Condicion ' viene sin WHERE

'////////////////////////
Caption = "Búsqueda de " + StrConv(Objs, vbProperCase)
lbl(0).Caption = "Escriba las primeras letras del " & StrConv(Obj, vbProperCase) & " que está buscando"

Actualiza True

Me.Show 1

End Sub
Private Sub List_DblClick()
Seleccion
End Sub
Private Sub List_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Seleccion
End Sub
Private Sub Seleccion()
m_Codigo = aCodigo(List.ListIndex)
m_Descripcion = List.Text
Unload Me
End Sub
Private Sub Texto_Change()
If Right(Texto.Text, 1) = "'" Then Texto.Text = Left(Texto.Text, Len(Texto.Text) - 1)
Actualiza False
End Sub
Private Sub Actualiza(Todo As Boolean)
Dim j As Integer
' todo: indica si es la primera vez y debe mostar todo

'If Todo Then
'    qry = "SELECT * FROM [" & m_Tabla & "] WHERE " & m_Condicion & " ORDER BY [" & m_CampoDescripcion & "]"
'Else
'    qry = "SELECT * FROM [" & m_Tabla & "] WHERE " & m_Condicion & " AND [" & m_CampoDescripcion & "] LIKE '" & Texto.Text & "%' ORDER BY [" & m_CampoDescripcion & "]"
'End If

If Todo Then
    qry = "SELECT * FROM [" & m_Tabla & "] WHERE " & m_Condicion & " ORDER BY " & m_Orden
Else
    qry = "SELECT * FROM [" & m_Tabla & "] WHERE " & m_Condicion & " AND " & m_CampoDescripcion & " LIKE '" & Texto.Text & "%' ORDER BY " & m_CampoDescripcion
End If


RsTabla.Open qry, CnxSqlServer_scp0

With RsTabla

List.Clear
i = -1
Do While Not .EOF

    i = i + 1
    
    m_ContenidoaMostrar = ""
    For j = 1 To NumeroCampos
        m_ContenidoaMostrar = m_ContenidoaMostrar & " " & RsTabla(CamposNombres(j))
    Next
    m_ContenidoaMostrar = Trim(m_ContenidoaMostrar)
    
    List.AddItem m_ContenidoaMostrar
    aCodigo(i) = RsTabla(m_CampoCodigo)
    .MoveNext

Loop

.Close

End With

End Sub
Private Sub Form_Unload(Cancel As Integer)
'fin
End Sub
