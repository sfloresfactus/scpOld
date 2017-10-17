VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Clientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   3180
   ClientLeft      =   3135
   ClientTop       =   2805
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   5820
      _ExtentX        =   10266
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
   Begin VB.TextBox Condiciones 
      Height          =   300
      Left            =   1200
      TabIndex        =   16
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox Telefono2 
      Height          =   300
      Left            =   3600
      TabIndex        =   14
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Telefono1 
      Height          =   300
      Left            =   1200
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Comuna 
      Height          =   300
      Left            =   1200
      TabIndex        =   10
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox Direccion 
      Height          =   300
      Left            =   1200
      TabIndex        =   8
      Top             =   1680
      Width           =   4455
   End
   Begin VB.TextBox Giro 
      Height          =   300
      Left            =   1200
      TabIndex        =   6
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox Razon 
      Height          =   300
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Width           =   4455
   End
   Begin VB.TextBox Codigo 
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   4680
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2760
      MousePointer    =   14  'Arrow and Question
      Picture         =   "Clientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
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
            Picture         =   "Clientes.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Clientes.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Clientes.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Clientes.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Clientes.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Clientes.frx":065C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "C&ondiciones"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Teléfono &2"
      Height          =   255
      Index           =   6
      Left            =   2640
      TabIndex        =   13
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Teléfono &1"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&Comuna"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&Giro"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&Dirección"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Razón &Social"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&RUT Cliente"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Obj As String, Objs As String, nuevo As Boolean, Accion As String, old_accion As String
Private Db As Database, Rs As Recordset
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnListar As Button, btnGrabar As Button, btnDesHacer As Button
Private m_Rut As String
Private Tabla_Nombre As String
Private Sub Inicializa()
Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnListar = Toolbar.Buttons(4)
Set btnGrabar = Toolbar.Buttons(5)
Set btnDesHacer = Toolbar.Buttons(6)

Tabla_Nombre = "clientes"

Obj = "CLIENTE"
Objs = "CLIENTES"

Accion = ""
old_accion = ""

btnBuscar.visible = False
btnBuscar.ToolTipText = "Busca " & StrConv(Obj, vbProperCase)

codigo.MaxLength = 10
Razon.MaxLength = 50
Giro.MaxLength = 60
Direccion.MaxLength = 50
Comuna.MaxLength = 30
Telefono1.MaxLength = 10
Telefono2.MaxLength = 10
Condiciones.MaxLength = 20

End Sub
Private Sub Form_Load()

Inicializa

Me.Caption = "MANTENCIÓN DE " & Objs
Botones_Enabled True, True, True, True, False, False

' abre archivos
Set Db = OpenDatabase(data_file)
Set Rs = Db.OpenRecordset("Clientes")
Rs.Index = "RUT"

Campos_Limpiar

nuevo = False

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
Private Sub Campos_Enabled(Si As Boolean)
codigo.Enabled = Si
Razon.Enabled = Si
Giro.Enabled = Si
Direccion.Enabled = Si
Comuna.Enabled = Si
Telefono1.Enabled = Si
Telefono2.Enabled = Si
Condiciones.Enabled = Si
End Sub
Private Sub Campos_Limpiar()
codigo.Text = ""
Razon.Text = ""
Giro.Text = ""
Direccion.Text = ""
Comuna.Text = ""
Telefono1.Text = ""
Telefono2.Text = ""
Condiciones.Text = ""
Campos_Enabled False
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
    Razon.Enabled = False
    codigo.SetFocus
    nuevo = False
    old_accion = Accion
    btnBuscar.visible = True
Case "Eliminar"
    codigo.Enabled = True
    Razon.Enabled = False
    codigo.SetFocus
'    nuevo = False
    btnBuscar.visible = True
Case "Listar"
    MousePointer = vbHourglass
    Cr.WindowTitle = Objs
    Cr.WindowMaxButton = False
    Cr.WindowMinButton = False
    Cr.WindowState = crptMaximized
    Cr.DataFiles(0) = data_file & ".MDB"
    Cr.Formulas(0) = "RAZON SOCIAL=""" & Empresa.Razon & """"
    Cr.ReportSource = crptReport
    Cr.ReportFileName = Drive_Server & Path_Rpt & "Clientes.Rpt"
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
    Me.Caption = "MANTENCIÓN DE " & Objs
Case Else
    Botones_Enabled False, False, False, False, False, True
    Me.Caption = "MANTENCIÓN DE " & Objs & " " & Button.Tag
End Select

End Sub
Private Sub btnBuscar_Click()

Search.Muestra data_file, "Clientes", "RUT", "Razon Social", Obj, Objs

codigo.Text = Search.codigo
Razon.Text = Search.descripcion

If codigo.Text <> "" Then
  
    Rs.Seek "=", codigo.Text
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        codigo.SetFocus
    Else
        After_Enter
    End If
End If
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

If Rut_Verifica(codigo.Text) = False Then
    MsgBox "RUT no Válido"
    Exit Sub
End If

m_Rut = Rut_Formato(codigo.Text)

Select Case Accion
Case "Agregar"
    Rs.Seek "=", m_Rut ' Codigo
    If Rs.NoMatch Then
        Campos_Enabled True
        codigo.Enabled = False
        Razon.SetFocus
        btnGrabar.Enabled = True
    Else
        Registro_Leer
        MsgBox Obj & " YA EXISTE"
        Campos_Limpiar
        codigo.Enabled = True
        codigo.SetFocus
    End If
Case "Modificar"
    Rs.Seek "=", m_Rut ' Codigo
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        codigo.SetFocus
    Else
        Campos_Enabled True
        codigo.Enabled = False
        Razon.SetFocus
        Registro_Leer
        btnGrabar.Enabled = True
        btnBuscar.visible = False
    End If
Case "Eliminar"
    Rs.Seek "=", m_Rut ' Codigo
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        codigo.SetFocus
    Else
        Campos_Enabled False
        Registro_Leer
        If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo, "Atención") = vbYes Then
            Rs.Delete
        End If
        btnBuscar.visible = True
        Campos_Limpiar
        codigo.Enabled = True
        codigo.SetFocus
    End If
End Select
End Sub
Private Sub Razon_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Giro.SetFocus
End Sub
Private Sub Giro_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Direccion.SetFocus
End Sub
Private Sub Direccion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Comuna.SetFocus
End Sub
Private Sub Comuna_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Telefono1.SetFocus
End Sub
Private Sub Telefono1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Telefono2.SetFocus
End Sub
Private Sub Telefono2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Condiciones.SetFocus
End Sub
Private Sub Condiciones_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Razon.SetFocus
End Sub
Private Function Valida(nuevo As Boolean) As Boolean
Valida = False
If IsObjBlanco(Razon, "RAZÓN SOCIAL", btnGrabar) Then Exit Function
If IsObjBlanco(Giro, "GIRO", btnGrabar) Then Exit Function
If IsObjBlanco(Direccion, "DIRECCIÓN", btnGrabar) Then Exit Function
If IsObjBlanco(Comuna, "COMUNA", btnGrabar) Then Exit Function

'' valida que razón social no esté repetida
If nuevo Then
    Rs.Index = "Razon Social"
    Rs.Seek "=", Razon
    If Not Rs.NoMatch Then
        MsgBox "RAZÓN SOCIAL YA EXISTE"
        btnGrabar.Value = tbrUnpressed
        Rs.Index = "RUT"
        Razon.SetFocus
        Exit Function
    End If
    btnGrabar.Value = tbrUnpressed
    Rs.Index = "RUT"
End If
''

Valida = True
End Function
Private Sub Registro_Grabar(nuevo As Boolean)

With Rs

    If nuevo Then
        .AddNew
        !Rut = m_Rut 'Codigo
    Else
        .Edit
    End If
    
    ![Razon Social] = Razon.Text
    !Giro = Giro.Text
    !Direccion = Direccion.Text
    !Comuna = Comuna.Text
    ![Telefono 1] = Telefono1.Text
    ![Telefono 2] = Telefono2.Text
    !Condiciones = Condiciones.Text
    
    .Update
    
End With

Campos_Limpiar
Accion = old_accion

End Sub
Private Sub Registro_Leer()
With Rs
    Razon.Text = ![Razon Social]
    Giro.Text = !Giro
    Direccion.Text = !Direccion
    Comuna.Text = !Comuna
    Telefono1.Text = NoNulo(![Telefono 1])
    Telefono2.Text = NoNulo(![Telefono 2])
    Condiciones.Text = NoNulo(!Condiciones)
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Rs.Close
Db.Close
End Sub
