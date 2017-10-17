VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Proveedores 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   3135
   ClientTop       =   2805
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6345
      _ExtentX        =   11192
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
   Begin VB.ComboBox CbClasif 
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox Fax 
      Height          =   300
      Left            =   4440
      TabIndex        =   15
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Condiciones 
      Height          =   300
      Left            =   1440
      TabIndex        =   19
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Contacto 
      Height          =   300
      Left            =   1440
      TabIndex        =   17
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Telefono1 
      Height          =   300
      Left            =   1440
      TabIndex        =   13
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Comuna 
      Height          =   300
      Left            =   1440
      TabIndex        =   11
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox Direccion 
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   1680
      Width           =   4695
   End
   Begin VB.TextBox Giro 
      Height          =   300
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Width           =   4695
   End
   Begin VB.TextBox Razon 
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   4695
   End
   Begin VB.TextBox Codigo 
      Height          =   300
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   5280
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2880
      MousePointer    =   14  'Arrow and Question
      Picture         =   "Proveedores.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
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
            Picture         =   "Proveedores.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Proveedores.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Proveedores.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Proveedores.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Proveedores.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Proveedores.frx":065C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "Clasificación"
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   20
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "&Fax"
      Height          =   255
      Index           =   8
      Left            =   3840
      TabIndex        =   14
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "Cond. de &Pago"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "Contacto &Sr"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "&Teléfono"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "&Comuna"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "&Giro"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "&Dirección"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "&Razón Social"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "RUT Proveedor"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Obj As String, Objs As String, nuevo As Boolean, Accion As String, old_accion As String
Dim Db As Database, Rs As Recordset, RsCla As Recordset
Dim btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Dim btnListar As Button, btnGrabar As Button, btnDesHacer As Button
Dim lc(10) As Integer ' largo de los campos
Dim m_Rut As String
Private Sub Campos_Enabled(Si As Boolean)
Codigo.Enabled = Si
Razon.Enabled = Si
Giro.Enabled = Si
Direccion.Enabled = Si
Comuna.Enabled = Si
Telefono1.Enabled = Si
Fax.Enabled = Si
Contacto.Enabled = Si
condiciones.Enabled = Si
CbClasif.Enabled = Si
End Sub
Private Sub btnBuscar_Click()

Search.Muestra data_file, "Proveedores", "RUT", "Razon Social", Obj, Objs

Codigo = Search.Codigo
Razon = Search.Descripcion

If Codigo <> "" Then
    Rs.Seek "=", Codigo
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        Codigo.SetFocus
    Else
        After_Enter
    End If
End If
End Sub
Private Sub Codigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If Codigo.Text = "" Then
        Beep
    Else
        After_Enter
    End If
End If
End Sub
Private Sub After_Enter()

If Rut_Verifica(Codigo) = False Then
    MsgBox "RUT no Válido"
    Exit Sub
End If

m_Rut = Rut_Formato(Codigo)

Select Case Accion
Case "Agregar"
    Rs.Seek "=", m_Rut ' Codigo
    If Rs.NoMatch Then
        Campos_Enabled True
        Codigo.Enabled = False
        Razon.SetFocus
        btnGrabar.Enabled = True
    Else
        Registro_Lee
        MsgBox Obj & " YA EXISTE"
        Campos_Limpiar
        Codigo.Enabled = True
        Codigo.SetFocus
    End If
Case "Modificar"
    Rs.Seek "=", m_Rut ' Codigo
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        Codigo.SetFocus
    Else
        Campos_Enabled True
        Codigo.Enabled = False
        Razon.SetFocus
        Registro_Lee
        btnGrabar.Enabled = True
        btnBuscar.visible = False
    End If
Case "Eliminar"
    Rs.Seek "=", m_Rut ' Codigo
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        Codigo.SetFocus
    Else
        Campos_Enabled False
        Registro_Lee
        If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo, "Atención") = vbYes Then
            Rs.Delete
        End If
        btnBuscar.visible = True
        Campos_Limpiar
        Codigo.Enabled = True
        Codigo.SetFocus
    End If
End Select
End Sub
Private Sub Registro_Lee()
Dim m_Texto As String
With Rs
Razon = ![Razon Social]
Giro = NoNulo(!Giro)
Direccion = NoNulo(!Direccion)
Comuna = NoNulo(!Comuna)
Telefono1 = NoNulo(![Telefono 1])
Fax = NoNulo(!Fax)
Contacto = NoNulo(!Contacto)
condiciones.Text = NoNulo(![Condiciones de Pago])

m_Texto = NoNulo(!Clasificacion)
If m_Texto = "" Then
    CbClasif.Text = " "
Else
    CbClasif.Text = m_Texto
End If

End With
End Sub
Private Sub Inicializa()
Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnListar = Toolbar.Buttons(4)
Set btnGrabar = Toolbar.Buttons(5)
Set btnDesHacer = Toolbar.Buttons(6)

Obj = "PROVEEDOR"
Objs = "PROVEEDORES"

Accion = ""
old_accion = ""

btnBuscar.visible = False
btnBuscar.ToolTipText = "Busca " & StrConv(Obj, vbProperCase)

' largo del los campos
lc(1) = 10 'codigo
lc(2) = 50 'razón social
lc(3) = 30 'giro
lc(4) = 50 'dirección
lc(5) = 30 'comuna
lc(6) = 10 'telefono
lc(7) = 10 'fax
lc(8) = 50 'contacto
lc(9) = 20 'condiciones de pago

Codigo.MaxLength = lc(1)
Razon.MaxLength = lc(2)
Giro.MaxLength = lc(3)
Direccion.MaxLength = lc(4)
Comuna.MaxLength = lc(5)
Telefono1.MaxLength = lc(6)
Fax.MaxLength = lc(7)
Contacto.MaxLength = lc(8)
condiciones.MaxLength = lc(9)

End Sub
Private Sub Form_Load()

Inicializa

Me.Caption = "MANTENCIÓN DE " & Objs
Botones_Enabled True, True, True, True, False, False

' abre archivos
Set Db = OpenDatabase(data_file)
Set Rs = Db.OpenRecordset("Proveedores")
Rs.Index = "RUT"

Set RsCla = Db.OpenRecordset("Clasificacion de Proveedores")
RsCla.Index = "Codigo"

CbClasif.AddItem " "
With RsCla
Do While Not .EOF
    CbClasif.AddItem !Codigo
    .MoveNext
Loop
End With

Campos_Limpiar

nuevo = False

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
If KeyAscii = vbKeyReturn Then Fax.SetFocus
End Sub
Private Sub Fax_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Contacto.SetFocus
End Sub
Private Sub Contacto_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then condiciones.SetFocus
End Sub
Private Sub Condiciones_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Razon.SetFocus
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Arch As String
Accion = Button.ToolTipText
Select Case Accion
Case "Agregar"
    Campos_Enabled False
    Codigo.Enabled = True
    Codigo.SetFocus
    nuevo = True
    old_accion = Accion
Case "Modificar"
    Codigo.Enabled = True
    Razon.Enabled = False
    Codigo.SetFocus
    nuevo = False
    old_accion = Accion
    btnBuscar.visible = True
Case "Eliminar"
    Codigo.Enabled = True
    Razon.Enabled = False
    Codigo.SetFocus
'    nuevo = False
    btnBuscar.visible = True
Case "Listar"
    MousePointer = vbHourglass
    cr.WindowTitle = Objs
    cr.WindowMaxButton = False
    cr.WindowMinButton = False
    cr.WindowState = crptMaximized
    cr.DataFiles(0) = data_file & ".MDB"
    cr.Formulas(0) = "RAZON SOCIAL=""" & Empresa.Razon & """"
    cr.Formulas(1) = ""
    cr.SelectionFormula = "" '24/06/99
    cr.ReportSource = crptReport
    cr.ReportFileName = Drive_Server & Path_Rpt & "Proveedores.Rpt"
    cr.Action = 1
    MousePointer = vbDefault
Case "Grabar"
    If Valida(nuevo) Then
        Registro_Graba nuevo
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
    
    Codigo.Enabled = True
    Codigo.SetFocus
    btnGrabar.Value = tbrUnpressed
    btnGrabar.Enabled = False
    
    ' activa lupa
    If Accion = "Modificar" Then
        btnBuscar.visible = True
    End If
    
Case 4 To 6 ' btnDesHacer
    Botones_Enabled True, True, True, True, False, False
    Me.Caption = "MANTENCIÓN DE " & Objs
Case Else
    Botones_Enabled False, False, False, False, False, True
    Me.Caption = "MANTENCIÓN DE " & Objs & " " & Button.Tag
End Select

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
Private Sub Registro_Graba(nuevo As Boolean)
save:

With Rs
If nuevo Then
    .AddNew
    !rut = m_Rut 'Codigo
Else
    .Edit
End If

![Razon Social] = Razon
!Giro = Giro
!Direccion = Direccion
!Comuna = Comuna
!Ciudad = ""
![Telefono 1] = Telefono1
![Telefono 2] = ""
!Fax = Fax
!Contacto = Contacto
![Condiciones de Pago] = condiciones.Text
!Clasificacion = CbClasif.Text

.Update
End With

Campos_Limpiar
Accion = old_accion

End Sub
Private Sub Campos_Limpiar()
Codigo.Text = ""
Razon.Text = ""
Giro.Text = ""
Direccion.Text = ""
Comuna.Text = ""
Telefono1.Text = ""
Fax.Text = ""
Contacto.Text = ""
condiciones.Text = ""
CbClasif.Text = " "
Campos_Enabled False
End Sub
