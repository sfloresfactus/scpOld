VERSION 5.00
Begin VB.Form Accesos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accesos"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   2025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnGrabar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtClaveGDE 
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
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "Clave para Hacer GD Especial"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Accesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private m_Clave As String
Private Sub Form_Load()

txtClaveGDE.MaxLength = 10
'txtClaveGDE.Text = Parametro.Iva

End Sub
Private Sub btnGrabar_Click()

Dim sql As String, Rs As New ADODB.Recordset

If txtClaveGDE.Text = "" Then
    MsgBox "Debe digitar clave"
    txtClaveGDE.SetFocus
    Exit Sub
End If

' graba porcentaje iva

'sql = "SELECT * FROM parametros" ' tiene un solo registro
'Rs_Abrir Rs, sql
sql = "UPDATE acceso SET clave='" & txtClaveGDE.Text & "' WHERE documento='GDE'"
CnxSqlServer_scp0.Execute sql
'    .Close

Unload Me
End Sub
Private Sub txtIva_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
