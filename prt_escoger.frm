VERSION 5.00
Begin VB.Form prt_escoger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Impresión"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame Frame 
      Caption         =   "Impresora"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox CbNombre 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lbl 
         Caption         =   "&Nombre"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "prt_escoger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private i As Integer
Private m_ImpresoraNombre As String
Public Property Let ImpresoraNombre(ByVal New_Value As String)
m_ImpresoraNombre = New_Value
End Property
Public Property Get ImpresoraNombre() As String
ImpresoraNombre = m_ImpresoraNombre
End Property
Private Sub Form_Load()
' lee lista de impresoras
Dim predet As String, impre As String, indice As Integer

predet = Printer.DeviceName

' lista impresoras
indice = 0
For i = 0 To Printers.Count - 1
    impre = Printers(i).DeviceName
'    Printer
    CbNombre.AddItem impre ' rpt
    If impre = predet Then
        indice = i
    End If
Next

CbNombre.ListIndex = indice

End Sub
Private Sub btnAceptar_Click()
If CbNombre.ListIndex = -1 Then
    MsgBox "Debe escoger Impresora"
    CbNombre.SetFocus
    Exit Sub
End If
Impresora_Predeterminada CbNombre.Text
m_ImpresoraNombre = CbNombre.Text
Unload Me
End Sub
Private Sub btnCancelar_Click()
Unload Me
End Sub
