VERSION 5.00
Begin VB.Form ClaveCambiar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Clave de Usuario"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox clave 
      Height          =   300
      Index           =   2
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox clave 
      Height          =   300
      Index           =   1
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox clave 
      Height          =   300
      Index           =   0
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblCambioClave 
      Caption         =   "Cambio Clave SCP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label asterisco 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   12
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label asterisco 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   11
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label asterisco 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   10
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label UsrNombre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label lbl 
      Caption         =   "Reingrese Nueva Clave"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "Nueva Clave"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Clave Actual"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblUsuario 
      Caption         =   "Usuario:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "ClaveCambiar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ac(9, 1) As String
Private av(9) As String
Private al(9) As Boolean
Private Sub Form_Load()

UsrNombre.Caption = Usuario.nombre

ac(1, 0) = "nombre"
ac(1, 1) = "'"
ac(2, 0) = "clave"
ac(2, 1) = "'"
al(1) = False
al(2) = True
asterisco(0).Caption = ""
asterisco(1).Caption = ""
asterisco(2).Caption = ""

clave(0).PasswordChar = "*"
clave(1).PasswordChar = "*"
clave(2).PasswordChar = "*"

End Sub
Private Function Validar() As Boolean
' valida campos
Validar = False

asterisco(0).Caption = ""
asterisco(1).Caption = ""
asterisco(2).Caption = ""

clave(0).Text = Trim(clave(0).Text)
clave(1).Text = Trim(clave(1).Text)
clave(2).Text = Trim(clave(2).Text)

If clave(0).Text = "" Then
    asterisco(0).Caption = "*"
    MsgBox "Debe Digitar Clave Actual"
    clave(0).SetFocus
    Exit Function
End If

If clave(0).Text <> Usuario.clave Then
    asterisco(0).Caption = "*"
    MsgBox "Clave Actual INCORRECTA !!!"
    clave(0).SetFocus
    Exit Function
End If

If clave(1).Text = "" Then
    asterisco(1).Caption = "*"
    MsgBox "Debe Digitar Nueva Clave"
    clave(1).SetFocus
    Exit Function
End If
If clave(2).Text = "" Then
    asterisco(2).Caption = "*"
    MsgBox "Debe Reingresar Nueva Clave"
    clave(2).SetFocus
    Exit Function
End If

If clave(1).Text <> clave(2).Text Then
    asterisco(1).Caption = "*"
    asterisco(2).Caption = "*"
    MsgBox """Nueva Clave"" debe ser igual a ""Reingrese Nueva Clave"""
    clave(2).SetFocus
    Exit Function
End If

av(1) = Usuario.nombre
av(2) = clave(1).Text

Validar = True

End Function
Private Sub btnAceptar_Click()
If Validar Then
    Registro_Modificar CnxSqlServer_scp0, "usuarios", ac, av, al, 2 ', "nombre='" & UsrNombre.Caption & "'"
    Usuario.clave = clave(1).Text
    Fin
End If
End Sub
Private Sub btnCancelar_Click()
Fin
End Sub
Private Sub clave_KeyPress(Index As Integer, KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Fin()
Unload Me
End Sub
