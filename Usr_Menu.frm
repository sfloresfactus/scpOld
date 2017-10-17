VERSION 5.00
Begin VB.Form Usr_Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios - Opciones de Menú"
   ClientHeight    =   6075
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Hab 
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   33
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Sm 
      Caption         =   "SubMenu"
      Height          =   375
      Index           =   10
      Left            =   4080
      TabIndex        =   32
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2760
      TabIndex        =   31
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1200
      TabIndex        =   30
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Sm 
      Caption         =   "SubMenu"
      Height          =   375
      Index           =   9
      Left            =   4080
      TabIndex        =   29
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Sm 
      Caption         =   "SubMenu"
      Height          =   375
      Index           =   8
      Left            =   4080
      TabIndex        =   28
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Sm 
      Caption         =   "SubMenu"
      Height          =   375
      Index           =   7
      Left            =   4080
      TabIndex        =   27
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton Sm 
      Caption         =   "SubMenu"
      Height          =   375
      Index           =   6
      Left            =   4080
      TabIndex        =   26
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Sm 
      Caption         =   "SubMenu"
      Height          =   375
      Index           =   5
      Left            =   4080
      TabIndex        =   25
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Sm 
      Caption         =   "SubMenu"
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   24
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Sm 
      Caption         =   "SubMenu"
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   23
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Sm 
      Caption         =   "SubMenu"
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   22
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Sm 
      Caption         =   "SubMenu"
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   21
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Sm 
      Caption         =   "SubMenu"
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   20
      Top             =   120
      Width           =   855
   End
   Begin VB.CheckBox Hab 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   19
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox Hab 
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   18
      Top             =   4560
      Width           =   255
   End
   Begin VB.CheckBox Hab 
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   17
      Top             =   4080
      Width           =   255
   End
   Begin VB.CheckBox Hab 
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   16
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox Hab 
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   15
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox Hab 
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox Hab 
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox Hab 
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox Hab 
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox Hab 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(10)"
      Height          =   255
      Index           =   10
      Left            =   600
      TabIndex        =   34
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(9)"
      Height          =   255
      Index           =   9
      Left            =   600
      TabIndex        =   9
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(8)"
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   8
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(7)"
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   7
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(6)"
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   6
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(5)"
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   5
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(4)"
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(3)"
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   3
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(2)"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(1)"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(0)"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Usr_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' privilegios de usuario
' es decir, opciones del menu que estan disponibles
Option Explicit
Private i As Integer, j As Integer, k As Integer
Private m_Menu(11) As String
'////////////////////////////////////////////////////////////////////
Public Property Get Usr_Menu1() As String
Usr_Menu1 = m_Menu(1)
End Property
Public Property Let Usr_Menu1(ByVal vNewValue As String)
m_Menu(1) = vNewValue
End Property
Public Property Get Usr_Menu2() As String
Usr_Menu2 = m_Menu(2)
End Property
Public Property Let Usr_Menu2(ByVal vNewValue As String)
m_Menu(2) = vNewValue
End Property
Public Property Get Usr_Menu3() As String
Usr_Menu3 = m_Menu(3)
End Property
Public Property Let Usr_Menu3(ByVal vNewValue As String)
m_Menu(3) = vNewValue
End Property
Public Property Get Usr_Menu4() As String
Usr_Menu4 = m_Menu(4)
End Property
Public Property Let Usr_Menu4(ByVal vNewValue As String)
m_Menu(4) = vNewValue
End Property
Public Property Get Usr_Menu5() As String
Usr_Menu5 = m_Menu(5)
End Property
Public Property Let Usr_Menu5(ByVal vNewValue As String)
m_Menu(5) = vNewValue
End Property
Public Property Get Usr_Menu6() As String
Usr_Menu6 = m_Menu(6)
End Property
Public Property Let Usr_Menu6(ByVal vNewValue As String)
m_Menu(6) = vNewValue
End Property
Public Property Get Usr_Menu7() As String
Usr_Menu7 = m_Menu(7)
End Property
Public Property Let Usr_Menu7(ByVal vNewValue As String)
m_Menu(7) = vNewValue
End Property
Public Property Get Usr_Menu8() As String
Usr_Menu8 = m_Menu(8)
End Property
Public Property Let Usr_Menu8(ByVal vNewValue As String)
m_Menu(8) = vNewValue
End Property
Public Property Get Usr_Menu9() As String
Usr_Menu9 = m_Menu(9)
End Property
Public Property Let Usr_Menu9(ByVal vNewValue As String)
m_Menu(9) = vNewValue
End Property
Public Property Get Usr_Menu10() As String
Usr_Menu10 = m_Menu(10)
End Property
Public Property Let Usr_Menu10(ByVal vNewValue As String)
m_Menu(10) = vNewValue
End Property
Public Property Get Usr_Menu11() As String
Usr_Menu11 = m_Menu(11)
End Property
Public Property Let Usr_Menu11(ByVal vNewValue As String)
m_Menu(11) = vNewValue
End Property
'////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
Dim c As Integer
c = 0
For c = 0 To Menu_Columnas - 1
    Hab(c).Value = IIf(Val(m_Menu(c + 1)) > 0, 1, 0)
    Opcion(c).Caption = menu.Mnu(c + 1).Caption
Next
'For c = Menu_Columnas To 11 ' ?
'    Hab(c).visible = False
'    Opcion(c).visible = False
'    Sm(c).visible = False
'Next

End Sub
Private Sub Sm_Click(Index As Integer)
Usr_SubMenu.Usr_Menu = Index + 1
Usr_SubMenu.Usr_SMenu = m_Menu(Index + 1)
Usr_SubMenu.Show 1
m_Menu(Index + 1) = Usr_SubMenu.Usr_SMenu
End Sub
Private Sub btnAceptar_Click()
Unload Me
End Sub
Private Sub btnCancelar_Click()
Unload Me
End Sub
