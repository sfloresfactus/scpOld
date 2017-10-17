VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Productos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3255
   ClientLeft      =   3135
   ClientTop       =   2805
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6030
      _ExtentX        =   10636
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
   Begin VB.ComboBox CbTipo 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox KgsxMts 
      Height          =   300
      Left            =   1320
      TabIndex        =   13
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox stockminimo 
      Height          =   300
      Left            =   1320
      TabIndex        =   15
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox largo 
      Height          =   300
      Left            =   1320
      TabIndex        =   10
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox unidad 
      Height          =   300
      Left            =   1320
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox descripcion 
      Height          =   300
      Left            =   1320
      TabIndex        =   6
      Top             =   1320
      Width           =   4575
   End
   Begin VB.TextBox codigo 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.CheckBox LargoEspecial 
      Caption         =   "Largo &Especial"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   2040
      Width           =   1575
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
      Left            =   3120
      MousePointer    =   14  'Arrow and Question
      Picture         =   "Productos.frx":0000
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
            Picture         =   "Productos.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Productos.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Productos.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Productos.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Productos.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Productos.frx":065C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "&Tipo Producto"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "&Kgs / Mts"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&Largo"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "&Unidad"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&Stock Mínimo"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "&Descripción"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&Código"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "Productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Obj As String, Objs As String, nuevo As Boolean, Accion As String, old_accion As String
Private Db As Database, Rs As Recordset, RsTp As Recordset
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnListar As Button, btnGrabar As Button, btnDesHacer As Button
Private Sub Form_Load()

Inicializa

Me.Caption = "MANTENCIÓN DE " & Objs
Botones_Enabled True, True, True, True, False, False

' abre archivos
Set Db = OpenDatabase(data_file)
Set Rs = Db.OpenRecordset("Productos")
Rs.Index = "Codigo"

Set RsTp = Db.OpenRecordset("Tipo Producto")
RsTp.Index = "Codigo"

Do While Not RsTp.EOF
    CbTipo.AddItem RsTp![Codigo]
    RsTp.MoveNext
Loop

Campos_Limpiar

nuevo = False

If False Then
    Dim cp As String, des As String
    Dim DbOc As Database
    Dim RsOc As Recordset
    Set DbOc = OpenDatabase(Madq_file)
    cp = "SELECT [Codigo Producto], sum(cantidad) AS cant, sum(cantidad*[precio unitario]) AS precioTotal FROM [OC Detalle] WHERE YEAR(Fecha) = 2015 GROUP BY [codigo producto]"
    Set RsOc = DbOc.OpenRecordset(cp)
    With RsOc
    Do While Not .EOF
        cp = NoNulo(![codigo producto])
        Rs.Seek "=", cp
        des = "*** no encontrado ***"
        If Not Rs.NoMatch Then
            des = Rs![Descripcion]
        End If
        If ![PrecioTotal] >= 10000000 Then
            Debug.Print cp; ";"; des; ";"; ![Cant]; ";"; ![PrecioTotal]
        End If
        .MoveNext
    Loop
    End With
End If

End Sub
Private Sub Inicializa()
Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnListar = Toolbar.Buttons(4)
Set btnGrabar = Toolbar.Buttons(5)
Set btnDesHacer = Toolbar.Buttons(6)

Obj = "PRODUCTO"
Objs = "PRODUCTOS"

Accion = ""
old_accion = ""

btnBuscar.visible = False
btnBuscar.ToolTipText = "Busca " & StrConv(Obj, vbProperCase)


With Codigo
'    .Mask = ">" & String(20, "&")
.MaxLength = 20
'    .PromptInclude = False
End With

With Descripcion
'    .Mask = ">" & String(50, "&")
'    .PromptInclude = False
.MaxLength = 50
End With

With unidad
'    .Mask = ">" & String(3, "&")
'    .PromptInclude = False
.MaxLength = 3
End With

With largo
'    .Mask = "##########"
'    .PromptInclude = False
.MaxLength = 10
End With

With stockminimo
'    .Mask = "##########"
'    .PromptInclude = False
.MaxLength = 10
End With

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
    Descripcion.Enabled = False
    Codigo.SetFocus
    nuevo = False
    old_accion = Accion
    btnBuscar.visible = True
Case "Eliminar"
    Codigo.Enabled = True
    Descripcion.Enabled = False
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
    cr.ReportSource = crptReport
    cr.ReportFileName = Drive_Server & Path_Rpt & "Productos.Rpt"
    cr.Action = 1
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
    
    Codigo.Enabled = True
    Codigo.SetFocus
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

Select Case Accion
Case "Agregar"
    Rs.Seek "=", Codigo
    If Rs.NoMatch Then
        Campos_Enabled True
        Codigo.Enabled = False
        Descripcion.SetFocus
        btnGrabar.Enabled = True
    Else
        Registro_Leer
        MsgBox Obj & " YA EXISTE"
        Campos_Limpiar
        Codigo.Enabled = True
        Codigo.SetFocus
    End If
Case "Modificar"
    Rs.Seek "=", Codigo
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        Codigo.SetFocus
    Else
        Campos_Enabled True
        Codigo.Enabled = False
        Descripcion.SetFocus
        Registro_Leer
        btnGrabar.Enabled = True
        btnBuscar.visible = False
    End If
Case "Eliminar"
    Rs.Seek "=", Codigo
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        Codigo.SetFocus
    Else
        Campos_Enabled False
        Registro_Leer
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
Private Sub Campos_Enabled(Si As Boolean)
Codigo.Enabled = Si
CbTipo.Enabled = Si
Descripcion.Enabled = Si
unidad.Enabled = Si
largo.Enabled = Si
LargoEspecial.Enabled = Si
KgsxMts.Enabled = Si
stockminimo.Enabled = Si
End Sub
Private Sub btnBuscar_Click()

Search.Muestra data_file, "Productos", "Codigo", "Descripcion", Obj, Objs

Codigo = Search.Codigo
Descripcion = Search.Descripcion

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
If IsObjBlanco(Descripcion, "DESCRIPCIÓN", btnGrabar) Then Exit Function
If IsObjBlanco(unidad, "UNIDAD", btnGrabar) Then Exit Function

'' valida que razón social no esté repetida
If nuevo Then
    Rs.Index = "Descripcion"
    Rs.Seek "=", Descripcion
    If Not Rs.NoMatch Then
        MsgBox "DESCRIPCIÓN YA EXISTE"
        btnGrabar.Value = tbrUnpressed
        Rs.Index = "Código"
        Descripcion.SetFocus
        Exit Function
    End If
    btnGrabar.Value = tbrUnpressed
    Rs.Index = "Codigo"
End If
''

Valida = True
End Function
Private Sub Registro_Grabar(nuevo As Boolean)
save:
With Rs

If nuevo Then
    .AddNew
    !Codigo = Codigo.Text
Else
    .Edit
End If

![Tipo Producto] = CbTipo.Text
!Descripcion = Descripcion.Text
![unidad de medida] = unidad.Text
!largo = Val(largo.Text)
![Largo Especial] = IIf(LargoEspecial.Value = 1, True, False)
![KgsxMts] = m_CDbl(KgsxMts.Text)
![Stock Minimo] = Val(stockminimo.Text)
'Rs![Fecha Compra] = Now
'Rs![Cantidad Compra] = 0
'Rs![Precio Compra] = 0
'Rs![Stock 1] = 0

.Update

End With

Campos_Limpiar
Accion = old_accion

End Sub
Private Sub Registro_Leer()
With Rs
On Error Resume Next
CbTipo.Text = NoNulo(![Tipo Producto])
On Error GoTo 0
Descripcion.Text = !Descripcion
unidad.Text = ![unidad de medida]
largo.Text = !largo
LargoEspecial.Value = IIf(![Largo Especial], 1, 0)
KgsxMts.Text = NoNulo_Double(![KgsxMts])
stockminimo.Text = ![Stock Minimo]
End With
End Sub
Private Sub Campos_Limpiar()
Codigo.Text = ""
CbTipo.ListIndex = -1
Descripcion.Text = ""
unidad.Text = ""
largo.Text = ""
LargoEspecial.Value = 0
KgsxMts.Text = ""
stockminimo.Text = ""
Campos_Enabled False
End Sub
Private Sub Descripcion_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Unidad_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Largo_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub KgsxMts_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub StockMinimo_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
