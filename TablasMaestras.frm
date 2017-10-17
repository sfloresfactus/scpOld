VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form TablasMaestras 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1635
   ClientLeft      =   3135
   ClientTop       =   2805
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar"
            Object.Tag             =   "[Agregando]"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "[Modificando]"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "[Eliminando]"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Listar"
            Object.Tag             =   "[Listando]"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Deshacer"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox descripcion 
      Height          =   300
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox codigo 
      Height          =   300
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   4680
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2880
      MousePointer    =   14  'Arrow and Question
      Picture         =   "TablasMaestras.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   300
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   6960
      Top             =   600
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TablasMaestras.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TablasMaestras.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TablasMaestras.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TablasMaestras.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TablasMaestras.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TablasMaestras.frx":065C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "&Descripci�n"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&C�digo"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "TablasMaestras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Obj As String, Objs As String, nuevo As Boolean, Accion As String, old_accion As String
Private Db As Database, Rs As Recordset
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnListar As Button, btnGrabar As Button, btnDesHacer As Button
Private m_TipoTabla As String
Public Property Let TipoTabla(ByVal New_Opcion As String)
m_TipoTabla = New_Opcion
End Property
Private Sub Form_Load()

Inicializa

Me.Caption = "MANTENCI�N DE " & Objs
Botones_Enabled True, True, True, True, False, False

' abre archivos
Set Db = OpenDatabase(mpro_file)
Set Rs = Db.OpenRecordset("TablasMaestras")
Rs.Index = "Tipo-Codigo"

Campos_Limpiar

nuevo = False

End Sub
Private Sub Inicializa()

Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnListar = Toolbar.Buttons(4)
Set btnGrabar = Toolbar.Buttons(5)
Set btnDesHacer = Toolbar.Buttons(6)

Select Case m_TipoTabla
Case "CCO"
    Obj = "CENTRO DE COSTO"
    Objs = "CENTROS DE COSTO"
End Select

Accion = ""
old_accion = ""

btnBuscar.visible = False
btnBuscar.ToolTipText = "Busca " & StrConv(Obj, vbProperCase)

codigo.MaxLength = 10
descripcion.MaxLength = 50

End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Arch As String
Accion = Button.ToolTipText
Select Case Accion
Case "Agregar"
    Campos_Enabled False
    codigo.Enabled = True
    codigo.SetFocus
    nuevo = True
    old_accion = Accion
Case "Modificar"
    codigo.Enabled = True
    descripcion.Enabled = False
    codigo.SetFocus
    nuevo = False
    old_accion = Accion
    btnBuscar.visible = True
Case "Eliminar"
    codigo.Enabled = True
    descripcion.Enabled = False
    codigo.SetFocus
'    nuevo = False
    btnBuscar.visible = True
Case "Listar"

    MousePointer = vbHourglass
    Cr.WindowTitle = Objs
    Cr.WindowMaxButton = False
    Cr.WindowMinButton = False
    Cr.WindowState = crptMaximized
    Cr.DataFiles(0) = mpro_file & ".MDB"
    Cr.Formulas(0) = "RAZON SOCIAL=""" & Empresa.Razon & """"
    Cr.ReportSource = crptReport
    
    Select Case m_TipoTabla
    Case "CCO"
        Cr.ReportFileName = Drive_Server & Path_Rpt & "centrosdecosto.rpt"
    End Select
    
    Cr.Action = 1
    MousePointer = vbDefault
Case "Grabar"
    If Valida(nuevo) Then
        Registro_Grabar nuevo
    Else
        Exit Sub
    End If
Case "Deshacer"
    Campos_Limpiar
    btnBuscar.visible = False
End Select

Select Case Button.Index
Case 5  ' btnGrabar
    Campos_Limpiar
    
    codigo.Enabled = True
    codigo.SetFocus
    btnGrabar.Value = tbrUnpressed
    btnGrabar.Enabled = False
    
Case 4 To 6 ' btnDesHacer
    Botones_Enabled True, True, True, True, False, False
    Me.Caption = "MANTENCI�N DE " & Objs
Case Else
    Botones_Enabled False, False, False, False, False, True
    Me.Caption = "MANTENCI�N DE " & Objs & " " & Button.Tag
End Select

End Sub
Private Sub Codigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If codigo.Text = "" Then
        Beep
    Else
        After_Enter
    End If
End If
End Sub
Private Sub After_Enter()

Select Case Accion
Case "Agregar"
    Rs.Seek "=", m_TipoTabla, codigo.Text
    If Rs.NoMatch Then
        Campos_Enabled True
        codigo.Enabled = False
        descripcion.SetFocus
        btnGrabar.Enabled = True
    Else
        Registro_Leer
        MsgBox Obj & " YA EXISTE"
        Campos_Limpiar
        codigo.Enabled = True
        codigo.SetFocus
    End If
Case "Modificar"
    Rs.Seek "=", m_TipoTabla, codigo.Text
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        codigo.SetFocus
    Else
        Campos_Enabled True
        codigo.Enabled = False
        descripcion.SetFocus
        Registro_Leer
        btnGrabar.Enabled = True
        btnBuscar.visible = False
    End If
Case "Eliminar"
    Rs.Seek "=", m_TipoTabla, codigo.Text
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        codigo.SetFocus
    Else
        Campos_Enabled False
        Registro_Leer
        If MsgBox("� ELIMINA " & Obj & " ?", vbYesNo, "Atenci�n") = vbYes Then
            Rs.Delete
        End If
        btnBuscar.visible = True
        Campos_Limpiar
        codigo.Enabled = True
        codigo.SetFocus
    End If
End Select
End Sub
Private Sub Campos_Enabled(Si As Boolean)
codigo.Enabled = Si
descripcion.Enabled = Si
End Sub
Private Sub btnBuscar_Click()

Search.Muestra mpro_file, "TablasMaestras", "Codigo", "codigo]+' '+[descripcion", Obj, Objs, "Tipo='" & m_TipoTabla & "'"

codigo = Search.codigo
descripcion = Search.descripcion

If codigo <> "" Then
    Rs.Seek "=", m_TipoTabla, codigo.Text
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        codigo.SetFocus
    Else
        After_Enter
    End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Rs.Close
Db.Close
End Sub
Private Sub Botones_Enabled(btn_Agregar As Boolean, btn_Modificar As Boolean, btn_Eliminar As Boolean, btn_Listar As Boolean, btn_Grabar As Boolean, btn_DesHacer As Boolean)
Dim i As Integer
btnAgregar.Enabled = btn_Agregar
btnModificar.Enabled = btn_Modificar
btnEliminar.Enabled = btn_Eliminar
btnListar.Enabled = btn_Listar
btnGrabar.Enabled = btn_Grabar
btnDesHacer.Enabled = btn_DesHacer

For i = 1 To 6
    Toolbar.Buttons(i).Value = tbrUnpressed
Next

End Sub
Private Function Valida(nuevo As Boolean) As Boolean
Valida = False
If IsObjBlanco(descripcion, "DESCRIPCI�N", btnGrabar) Then Exit Function

'' valida que raz�n social no est� repetida
GoTo Sigue
If nuevo Then
    Rs.Index = "Descripcion"
    Rs.Seek "=", descripcion
    If Not Rs.NoMatch Then
        MsgBox "DESCRIPCI�N YA EXISTE"
        btnGrabar.Value = tbrUnpressed
        Rs.Index = "C�digo"
        descripcion.SetFocus
        Exit Function
    End If
    btnGrabar.Value = tbrUnpressed
    Rs.Index = "C�digo"
End If
''
Sigue:

Valida = True
End Function
Private Sub Registro_Grabar(nuevo As Boolean)
save:
With Rs

If nuevo Then
    .AddNew
    !tipo = m_TipoTabla
    !codigo = UCase(codigo.Text)
Else
    .Edit
End If

!descripcion = descripcion.Text

.Update

End With

Campos_Limpiar
Accion = old_accion

End Sub
Private Sub Registro_Leer()
With Rs
descripcion.Text = NoNulo(!descripcion)
End With
End Sub
Private Sub Campos_Limpiar()
codigo.Text = ""
descripcion.Text = ""
Campos_Enabled False
End Sub
Private Sub Descripcion_KeyPress(KeyAscii As Integer)
'If KeyAscii = vbKeyReturn Then Unidad.SetFocus
End Sub
