VERSION 5.00
Begin VB.Form Usr_Propiedades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades de Usuario"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPrivilegios 
      Caption         =   "&Privilegios"
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox UsrClave1 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox UsrClave0 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox UsrDescripcion 
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox UsrNombre 
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lbl 
      Caption         =   "&Repetir Contraseña"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lbl 
      Caption         =   "C&ontraseña"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "&Descripción"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "&Nombre"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Usr_Propiedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_UsrNombre As String
Private m_UsrDescripcion As String
Private m_UsrClave As String
'Private RsUs As Recordset
Private RsUs As New ADODB.Recordset
'Private RsTemp As New ADODB.Recordset
'
Private m_UsrMenu(12) As String ' solo memoria, no es propiedad
Private sql As String

' para tablas maestras
Private ac(15, 1) As String ' arreglo de campos-comillas para manejo de registros
Private av(15) As String ' arreglo de valores para manejo de registros
Private al(15) As Boolean ' arreglo que indica si el campo se modifica
Private TotalCampos As Integer ' indica el numero total de campos de la tabla
'////////////////////////////
Public Property Get Usr_Nombre() As String
Usr_Nombre = m_UsrNombre
End Property
Public Property Let Usr_Nombre(ByVal vNewValue As String)
m_UsrNombre = vNewValue
End Property

Public Property Get Usr_Descripcion() As String
Usr_Descripcion = m_UsrDescripcion
End Property
Public Property Let Usr_Descripcion(ByVal vNewValue As String)
m_UsrDescripcion = vNewValue
End Property

Public Property Get Usr_Clave() As String
Usr_Clave = m_UsrClave
End Property
Public Property Let Usr_Clave(ByVal vNewValue As String)
m_UsrClave = vNewValue
End Property

'Public Property Get Usr_Privi() As String
'Usr_Privi = m_UsrPrivilegios
'End Property
'Public Property Let Usr_Privi(ByVal vNewValue As String)
'm_UsrPrivilegios = vNewValue
'End Property

'Public Property Get Usr_Recordset() As Recordset
'Set Usr_Recordset = RsUs
'End Property

'Public Property Let Usr_Recordset(ByVal vNewValue As ADODB.Recordset)
'Set RsUs = vNewValue
'End Property
'/////////////////////////////
Private Sub Form_Load()

UsrNombre.MaxLength = 10
UsrDescripcion.MaxLength = 50
UsrClave0.MaxLength = 10
UsrClave1.MaxLength = 10

UsrNombre.Text = m_UsrNombre
UsrDescripcion.Text = m_UsrDescripcion
UsrClave0.Text = m_UsrClave
UsrClave1.Text = m_UsrClave

If Len(m_UsrNombre) > 0 Then
    UsrNombre.Enabled = False
    
    sql = "SELECT * FROM usuarios WHERE nombre='" & UsrNombre.Text & "'"
    
    With RsUs
    
    Rs_Abrir RsUs, sql
    
    If Not .EOF Then
    
        m_UsrMenu(1) = NoNulo(RsUs!menu01)
        m_UsrMenu(2) = NoNulo(RsUs!menu02)
        m_UsrMenu(3) = NoNulo(RsUs!menu03)
        m_UsrMenu(4) = NoNulo(RsUs!menu04)
        m_UsrMenu(5) = NoNulo(RsUs!menu05)
        m_UsrMenu(6) = NoNulo(RsUs!menu06)
        m_UsrMenu(7) = NoNulo(RsUs!menu07)
        m_UsrMenu(8) = NoNulo(RsUs!menu08)
        m_UsrMenu(9) = NoNulo(RsUs!menu09)
        m_UsrMenu(10) = NoNulo(RsUs!menu10)
        m_UsrMenu(11) = NoNulo(RsUs!menu11)
        
    End If
    End With
    
End If

ac(1, 0) = "nombre"
ac(1, 1) = "'"
ac(2, 0) = "clave"
ac(2, 1) = "'"
ac(3, 0) = "descripcion"
ac(3, 1) = "'"
ac(4, 0) = "menu01"
ac(4, 1) = "'"
ac(5, 0) = "menu02"
ac(5, 1) = "'"
ac(6, 0) = "menu03"
ac(6, 1) = "'"
ac(7, 0) = "menu04"
ac(7, 1) = "'"
ac(8, 0) = "menu05"
ac(8, 1) = "'"
ac(9, 0) = "menu06"
ac(9, 1) = "'"
ac(10, 0) = "menu07"
ac(10, 1) = "'"
ac(11, 0) = "menu08"
ac(11, 1) = "'"
ac(12, 0) = "menu09"
ac(12, 1) = "'"
ac(13, 0) = "menu10"
ac(13, 1) = "'"
ac(14, 0) = "menu11"
ac(14, 1) = "'"
ac(15, 0) = "menu12"
ac(15, 1) = "'"

TotalCampos = 15

End Sub
Private Sub btnPrivilegios_Click()
Usr_Menu.Usr_Menu1 = m_UsrMenu(1)
Usr_Menu.Usr_Menu2 = m_UsrMenu(2)
Usr_Menu.Usr_Menu3 = m_UsrMenu(3)
Usr_Menu.Usr_Menu4 = m_UsrMenu(4)
Usr_Menu.Usr_Menu5 = m_UsrMenu(5)
Usr_Menu.Usr_Menu6 = m_UsrMenu(6)
Usr_Menu.Usr_Menu7 = m_UsrMenu(7)
Usr_Menu.Usr_Menu8 = m_UsrMenu(8)
Usr_Menu.Usr_Menu9 = m_UsrMenu(9)
Usr_Menu.Usr_Menu10 = m_UsrMenu(10)
Usr_Menu.Usr_Menu11 = m_UsrMenu(11)
Usr_Menu.Show 1
m_UsrMenu(1) = Usr_Menu.Usr_Menu1
m_UsrMenu(2) = Usr_Menu.Usr_Menu2
m_UsrMenu(3) = Usr_Menu.Usr_Menu3
m_UsrMenu(4) = Usr_Menu.Usr_Menu4
m_UsrMenu(5) = Usr_Menu.Usr_Menu5
m_UsrMenu(6) = Usr_Menu.Usr_Menu6
m_UsrMenu(7) = Usr_Menu.Usr_Menu7
m_UsrMenu(8) = Usr_Menu.Usr_Menu8
m_UsrMenu(9) = Usr_Menu.Usr_Menu9
m_UsrMenu(10) = Usr_Menu.Usr_Menu10
m_UsrMenu(11) = Usr_Menu.Usr_Menu11
End Sub
Private Sub UsrNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then UsrDescripcion.SetFocus
End Sub
Private Sub UsrDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then UsrClave0.SetFocus
End Sub
Private Sub UsrClave0_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then UsrClave1.SetFocus
End Sub
Private Sub UsrClave1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then btnPrivilegios.SetFocus
End Sub
Private Sub btnAceptar_Click()

save:

If Campos_Validar = False Then Exit Sub

'RsUs.Seek "=", UsrNombre.Text
'Set RsUs = CnxSqlServer.OpenRecordset("SELECT * FROM usuarios WHERE nombre='" & UsrNombre.Text & "'")

'sql = "SELECT * FROM usuarios WHERE nombre='" & UsrNombre.Text & "'"
'If RsTemp.Source Then
'RsTemp.Close
'RsUs.Open sql, CnxSqlServer

If UsrNombre.Enabled Then
    ' usuario nuevo

    If Not Registro_Existe("usuarios", "nombre='" & UsrNombre.Text & "'") Then
    
       ' ok
        
'        RsUs.AddNew
'        RsUs!nombre = UsrNombre.Text
        
        sql = "INSERT INTO usuarios ("
        sql = sql & " nombre,"
        sql = sql & " clave,"
        sql = sql & " descripcion,"
        sql = sql & " menu01,"
        sql = sql & " menu02,"
        sql = sql & " menu03,"
        sql = sql & " menu04,"
        sql = sql & " menu05,"
        sql = sql & " menu06,"
        sql = sql & " menu07,"
        sql = sql & " menu08,"
        sql = sql & " menu09,"
        sql = sql & " menu10,"
        sql = sql & " menu11,"
        sql = sql & " menu12,"
        sql = sql & " nv_activas"
        sql = sql & ") VALUES ("
        sql = sql & "'" & UsrNombre.Text & "',"
        sql = sql & "'" & UsrClave0.Text & "',"
        sql = sql & "'" & UsrDescripcion.Text & "',"
        sql = sql & "'" & m_UsrMenu(1) & "',"
        sql = sql & "'" & m_UsrMenu(2) & "',"
        sql = sql & "'" & m_UsrMenu(3) & "',"
        sql = sql & "'" & m_UsrMenu(4) & "',"
        sql = sql & "'" & m_UsrMenu(5) & "',"
        sql = sql & "'" & m_UsrMenu(6) & "',"
        sql = sql & "'" & m_UsrMenu(7) & "',"
        sql = sql & "'" & m_UsrMenu(8) & "',"
        sql = sql & "'" & m_UsrMenu(9) & "',"
        sql = sql & "'" & m_UsrMenu(10) & "',"
        sql = sql & "'" & m_UsrMenu(11) & "',"
        sql = sql & "'" & m_UsrMenu(12) & "',"
        sql = sql & "'N'"
        sql = sql & ")"
        
        av(1) = UsrNombre.Text
        av(2) = UsrClave0.Text
        av(3) = UsrDescripcion.Text
        av(4) = m_UsrMenu(1)
        av(5) = m_UsrMenu(2)
        av(6) = m_UsrMenu(3)
        av(7) = m_UsrMenu(4)
        av(8) = m_UsrMenu(5)
        av(9) = m_UsrMenu(6)
        av(10) = m_UsrMenu(7)
        av(11) = m_UsrMenu(8)
        av(12) = m_UsrMenu(9)
        av(13) = m_UsrMenu(10)
        av(14) = m_UsrMenu(11)
        av(15) = m_UsrMenu(12)
        
        Arreglo_Limpiar al, TotalCampos, True
        
        Registro_Agregar CnxSqlServer_scp0, "usuarios", ac, av, TotalCampos
    
'        Debug.Print sql
'        CnxSqlServer.Execute sql
        
    Else
        MsgBox "USUARIO YA EXISTE"
        UsrNombre.SetFocus
        Exit Sub
    End If
Else

    ' modificacion de usuario
    
    av(1) = UsrNombre.Text
    av(2) = UsrClave0.Text
    av(3) = UsrDescripcion.Text
    av(4) = m_UsrMenu(1)
    av(5) = m_UsrMenu(2)
    av(6) = m_UsrMenu(3)
    av(7) = m_UsrMenu(4)
    av(8) = m_UsrMenu(5)
    av(9) = m_UsrMenu(6)
    av(10) = m_UsrMenu(7)
    av(11) = m_UsrMenu(8)
    av(12) = m_UsrMenu(9)
    av(13) = m_UsrMenu(10)
    av(14) = m_UsrMenu(11)
    av(15) = m_UsrMenu(12)
    
    Arreglo_Limpiar al, TotalCampos, True
    al(1) = False ' es la condicion para modifiar
    
    Registro_Modificar CnxSqlServer_scp0, "usuarios", ac, av, al, TotalCampos ', "nombre='" & UsrNombre.Text & "'"

End If

m_UsrNombre = UsrNombre.Text
m_UsrDescripcion = UsrDescripcion.Text

Unload Me

End Sub
Private Function Campos_Validar() As Boolean
Campos_Validar = False
If UsrNombre.Text = "" Then
    Beep
    MsgBox "NOMBRE NO PUEDE QUEDAR EN BLANCO"
    UsrNombre.SetFocus
    Exit Function
End If
If UsrDescripcion.Text = "" Then
    Beep
    MsgBox "DESCRIPCIÓN NO PUEDE QUEDAR EN BLANCO"
    UsrDescripcion.SetFocus
    Exit Function
End If
If Trim(UsrClave0.Text) = "" Then
    Beep
    MsgBox "CLAVE NO PUEDE QUEDAR EN BLANCO"
    UsrClave0.SetFocus
    Exit Function
End If
If Trim(UsrClave1.Text) = "" Then
    Beep
    MsgBox "CLAVE NO PUEDE QUEDAR EN BLANCO"
    UsrClave1.SetFocus
    Exit Function
End If
If UsrClave0.Text <> UsrClave1.Text Then
    Beep
    MsgBox "CLAVES NO SON IGUALES"
    UsrClave1.SetFocus
    Exit Function
End If
Campos_Validar = True
End Function
Private Sub btnCancelar_Click()
Unload Me
End Sub
