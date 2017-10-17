VERSION 5.00
Begin VB.Form ProvLoc_Propiedades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Local"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox LocContacto 
      Height          =   300
      Left            =   1800
      TabIndex        =   11
      Top             =   2160
      Width           =   4335
   End
   Begin VB.TextBox LocFax 
      Height          =   300
      Left            =   1800
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "C&ancelar"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox LocTelefono 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1800
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox LocCiudad 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox LocComuna 
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox LocDireccion 
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label lbl 
      Caption         =   "C&ontacto"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "&Fax"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "&Teléfono"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "C&iudad"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "&Comuna"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "&Dirección"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "ProvLoc_Propiedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_LocRut As String
Private m_LocId As Integer
Private m_LocDireccion As String
Private m_LocComuna As String
Private m_LocCiudad As String
Private m_LocTelefono As String
Private m_LocFax As String
Private m_LocContacto As String

Private RsLoc As Recordset
'////////////////////////////

Public Property Get Loc_Rut() As String
Loc_Rut = m_LocRut
End Property
Public Property Let Loc_Rut(ByVal vNewValue As String)
m_LocRut = vNewValue
End Property

Public Property Get Loc_Id() As Integer
Loc_Id = m_LocId
End Property
Public Property Let Loc_Id(ByVal vNewValue As Integer)
m_LocId = vNewValue
End Property

Public Property Get Loc_Direccion() As String
Loc_Direccion = m_LocDireccion
End Property
Public Property Let Loc_Direccion(ByVal vNewValue As String)
m_LocDireccion = vNewValue
End Property

Public Property Get Loc_Comuna() As String
Loc_Comuna = m_LocComuna
End Property
Public Property Let Loc_Comuna(ByVal vNewValue As String)
m_LocComuna = vNewValue
End Property

Public Property Get Loc_Ciudad() As String
Loc_Ciudad = m_LocCiudad
End Property
Public Property Let Loc_Ciudad(ByVal vNewValue As String)
m_LocCiudad = vNewValue
End Property

Public Property Get Loc_Telefono() As String
Loc_Telefono = m_LocTelefono
End Property
Public Property Let Loc_Telefono(ByVal vNewValue As String)
m_LocTelefono = vNewValue
End Property

Public Property Get Loc_Fax() As String
Loc_Fax = m_LocFax
End Property
Public Property Let Loc_Fax(ByVal vNewValue As String)
m_LocFax = vNewValue
End Property

Public Property Get Loc_Contacto() As String
Loc_Contacto = m_LocContacto
End Property
Public Property Let Loc_Contacto(ByVal vNewValue As String)
m_LocContacto = vNewValue
End Property

Public Property Get Loc_Recordset() As Recordset
Set Loc_Recordset = RsLoc
End Property
Public Property Let Loc_Recordset(ByVal vNewValue As Recordset)
Set RsLoc = vNewValue
End Property
'/////////////////////////////
Private Sub Form_Load()

LocDireccion.MaxLength = 50
LocComuna.MaxLength = 30
LocCiudad.MaxLength = 30
LocTelefono.MaxLength = 10
LocFax.MaxLength = 10
LocContacto.MaxLength = 50

LocDireccion.Text = m_LocDireccion
LocComuna.Text = m_LocComuna
LocCiudad.Text = m_LocCiudad
LocTelefono.Text = m_LocTelefono
LocFax.Text = m_LocFax
LocContacto.Text = m_LocContacto

End Sub
Private Sub LocDireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then LocComuna.SetFocus
End Sub
Private Sub LocComuna_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then LocCiudad.SetFocus
End Sub
Private Sub LocCiudad_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then LocTelefono.SetFocus
End Sub
Private Sub LocTelefono_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then LocFax.SetFocus
End Sub
Private Sub LocFax_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then LocContacto.SetFocus
End Sub
Private Sub LocContacto_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then btnAceptar.SetFocus
End Sub
Private Sub btnAceptar_Click()

save:
If Campos_Validar = False Then Exit Sub

With RsLoc
If m_LocId = 0 Then
    ' local nuevo
    m_LocId = Local_Nuevo
    RsLoc.AddNew
    RsLoc!rut = m_LocRut
    RsLoc!Codigo = m_LocId
Else
    ' modificacion de local
    RsLoc.Edit
End If

!Direccion = LocDireccion.Text
!Comuna = LocComuna.Text
!Ciudad = LocCiudad.Text
![Telefono 1] = LocTelefono.Text
!Fax = LocFax.Text
!Contacto = LocContacto.Text
.Update
End With

m_LocDireccion = LocDireccion.Text
m_LocComuna = LocComuna.Text
m_LocCiudad = LocCiudad.Text
m_LocTelefono = LocTelefono.Text
m_LocFax = LocFax.Text
m_LocContacto = LocContacto.Text

Unload Me

End Sub
Private Function Campos_Validar() As Boolean
Campos_Validar = False
If LocDireccion.Text = "" Then
    Beep
    MsgBox "DIRECCIÓN NO PUEDE QUEDAR EN BLANCO"
    LocDireccion.SetFocus
    Exit Function
End If
If LocComuna.Text = "" Then
    Beep
    MsgBox "COMUNA NO PUEDE QUEDAR EN BLANCO"
    LocComuna.SetFocus
    Exit Function
End If

Campos_Validar = True
End Function
Private Function Local_Nuevo() As Integer
' entrega nuevo numero de local
Local_Nuevo = 0
With RsLoc
.Seek ">=", m_LocRut, Local_Nuevo
If Not .NoMatch Then
    'encontro algun local
    Do While Not .EOF
        If !rut <> m_LocRut Then Exit Do
        Local_Nuevo = !Codigo
        .MoveNext
    Loop
    Local_Nuevo = Local_Nuevo + 1
Else
    'para cuando el arch este vacio
    Local_Nuevo = 1
End If
End With
End Function
Private Sub btnCancelar_Click()
Unload Me
End Sub
