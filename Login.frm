VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio de sesión"
   ClientHeight    =   1485
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   877.387
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   2267.554
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   339
      Left            =   1170
      TabIndex        =   1
      Top             =   135
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Width           =   1020
   End
   Begin VB.TextBox txtPassword 
      Height          =   339
      IMEMode         =   3  'DISABLE
      Left            =   1170
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   1065
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Usuario:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   180
      Width           =   960
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_Exitoso As Boolean
'Private Db As Database, Rs As Recordset
Private Rs As New ADODB.Recordset, sql As String
'Private Rs As Recordset, sql As String
Public Property Get Exitoso() As Boolean
Exitoso = m_Exitoso
End Property
Public Property Let Exitoso(ByVal vNewValue As Boolean)
m_Exitoso = vNewValue
End Property
Private Sub Form_Load()
'Set Db = OpenDatabase(Syst_file, False, True, ";pwd=eml")
'Set Rs = Db.OpenRecordset("Usuarios")
'Rs.Index = "Nombre"
'Set Rs = cnxSqlServer.OpenRecordset("usuarios")

m_Exitoso = False
txtUserName.MaxLength = 10
txtPassword.MaxLength = 10

End Sub
Private Sub cmdOK_Click()

If txtUserName.Text = "" Then
    MsgBox "Debe digitar Nombre de Usuario, vuelva a intentarlo", , "Inicio de sesión"
    txtUserName.SetFocus
    Exit Sub
End If

'Rs.Seek "=", txtUserName.Text
sql = "SELECT * FROM usuarios WHERE nombre='" & txtUserName.Text & "'" ' sqlserver.scp0
'Set Rs = CnxSqlServer.OpenRecordset(sql)
Rs.Open sql, CnxSqlServer_scp0

'If Rs.NoMatch Then
'If Rs.RecordCount = 0 Then
If Rs.EOF Then
    MsgBox "Nombre de Usuario no válido, vuelva a intentarlo", , "Inicio de sesión"
    txtUserName.SetFocus
'    SendKeysA "{Home}+{End}"
'    SendKeysA
    Rs.Close
    Exit Sub
End If

'comprueba la contraseña correcta
If UCase(txtPassword.Text) = UCase(Trim(Rs!clave)) Then

    Privi_LLena
    
    Usuario.Descripcion = Trim(Rs!Descripcion)
    
    Usuario.Adquis_Actual = True 'default, adquisiciones actual (no historico) 30/03/2000
    Usuario.ObrasTerminadas = IIf(Rs![nv_terminadas] = "S", True, False) '27/01/99
    Usuario.Nv_Activas = IIf(Rs![Nv_Activas] = "S", True, False)
'    mpro_file = Movs_Path(EmpOC.RUT, Usuario.ObrasTerminadas) '27/01/99
    Usuario.Nv_Orden = NoNulo(Rs!Nv_Orden)
    If Usuario.Nv_Orden = "A" Then
        Nv_Index = "Nombre"
    Else
        Nv_Index = "Numero"
    End If
    
'    Rs.Close
'''    Db.Close 'QUEDABA ABIERTA!!!!!!!
'        Load Menu
    Unload IniScreen
    Usuario.nombre = txtUserName.Text
    Usuario.clave = txtPassword.Text
    Usuario.Tipo = NoNulo(Rs!usuario_tipo)
    Usuario.Rut = NoNulo(Rs!usuario_rut)
    
    Usuario.Scc_Mod = IIf(NoNulo(Rs!Scc_Mod) = "S", True, False)

    Rs.Close
    
    Unload Me
'        Menu.Show 1
    m_Exitoso = True ' 01/07/98
'    Login_Registrar 0, Usuario.nombre
    
Else

    MsgBox "Contraseña no válida, vuelva a intentarlo", , "Inicio de sesión"
    txtPassword.SetFocus
    
'    If Not Win7 Then
'        SendKeys "{Home}+{End}"
'    SendKeysA vbKeyHome, True
'    SendKeysA vbKeyEnd, True
    
'    End If
    
    Rs.Close
    
End If
End Sub
Private Sub cmdCancel_Click()
    'establece la variable global a false
    'para indicar un fallo en el inicio de sesión
    Unload Me
    Unload IniScreen
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload IniScreen
End Sub
Private Sub Privi_LLena()
' puebla arreglo de privilegios para menu
Dim i As Integer, j As Integer, k As Integer
Dim m_Privi As String, Valor As String
'm_Privi = Trim(NoNulo(Rs!Privilegios))

For i = 1 To Menu_Columnas
    For j = 1 To Menu_Filas
        m_Privi = Trim(str(i))
        m_Privi = "Menu" & PadL(m_Privi, 2, "0")
        m_Privi = NoNulo(Rs(m_Privi))
        m_Privi = m_Privi & String(20, "0")
        Privi(i, j) = Mid(m_Privi, j, 1)
    Next
Next

End Sub
