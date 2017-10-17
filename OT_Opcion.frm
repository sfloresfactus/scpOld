VERSION 5.00
Begin VB.Form OT_Opcion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OT"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn 
      Caption         =   "&Cancelar"
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton btn 
      Caption         =   "&Aceptar"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton Op 
      Caption         =   "por &Plano"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.OptionButton Op 
      Caption         =   "&Normal"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "OT_Opcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_TipoOT As String, m_Titulo As String
Public Property Let Titulo(ByVal vNewValue As String)
m_Titulo = vNewValue
End Property
Public Property Get TipoOT() As String
TipoOT = m_TipoOT
End Property
Public Property Let TipoOT(ByVal vNewValue As String)
m_TipoOT = vNewValue
End Property
Private Sub Form_Load()
Me.Caption = m_Titulo
Op(0).Value = True
m_TipoOT = ""
End Sub
Private Sub btn_Click(Index As Integer)
If Index = 0 Then
    m_TipoOT = IIf(Op(0).Value, "N", "P")
Else
    m_TipoOT = ""
End If
Unload Me
End Sub
