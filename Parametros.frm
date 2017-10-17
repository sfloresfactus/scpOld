VERSION 5.00
Begin VB.Form Parametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   2625
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnGrabar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtIva 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "Porcentaje IVA"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Iva As Double
Private Sub Form_Load()

txtIva.MaxLength = 3
txtIva.Text = Parametro.Iva

End Sub
Private Sub btnGrabar_Click()

Dim sql As String, Rs As New ADODB.Recordset

m_Iva = Val(txtIva.Text)

If 0 <= m_Iva And m_Iva <= 100 Then
Else
    MsgBox "Porcentaje IVA debe ser entre 0% y 100%"
    txtIva.SetFocus
    Exit Sub
End If

' graba porcentaje iva

sql = "SELECT * FROM parametros" ' tiene un solo registro
Rs_Abrir Rs, sql
sql = "UPDATE parametros SET iva=" & m_Iva
CnxSqlServer_scp0.Execute sql
'    .Close

Parametro.Iva = m_Iva

Unload Me
End Sub
Private Sub txtIva_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
