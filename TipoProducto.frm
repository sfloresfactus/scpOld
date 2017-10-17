VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form TipoProducto 
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
      Height          =   480
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Agregar"
            Object.Tag             =   "[Agregando]"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "[Modificando]"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "[Eliminando]"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Listar"
            Object.Tag             =   "[Listando]"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Deshacer"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      MouseIcon       =   "TipoProducto.frx":0000
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
      Picture         =   "TipoProducto.frx":001C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   300
   End
   Begin VB.Label lbl 
      Caption         =   "&Descripción"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&Código"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   975
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
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TipoProducto.frx":011E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TipoProducto.frx":0230
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TipoProducto.frx":0342
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TipoProducto.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TipoProducto.frx":0566
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TipoProducto.frx":0678
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "TipoProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Obj As String, Objs As String, nuevo As Boolean, Accion As String, old_accion As String
Private Db As Database, Rs As Recordset
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnListar As Button, btnGrabar As Button, btnDesHacer As Button
Private Sub Form_Load()

Inicializa

Me.Caption = "MANTENCIÓN DE " & Objs
Botones_Enabled True, True, True, True, False, False

' abre archivos
Set Db = OpenDatabase(data_file)
Set Rs = Db.OpenRecordset("Tipo Producto")
Rs.Index = "Codigo"

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

Obj = "TIPO DE PRODUCTO"
Objs = "TIPOS DE PRODUCTO"

Accion = ""
old_accion = ""

btnBuscar.Visible = False
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
    btnBuscar.Visible = True
Case "Eliminar"
    codigo.Enabled = True
    descripcion.Enabled = False
    codigo.SetFocus
'    nuevo = False
    btnBuscar.Visible = True
Case "Listar"
    MousePointer = vbHourglass
    Cr.WindowTitle = Objs
    Cr.WindowMaxButton = False
    Cr.WindowMinButton = False
    Cr.WindowState = crptMaximized
    Cr.DataFiles(0) = data_file & ".MDB"
    Cr.Formulas(0) = "RAZON SOCIAL=""" & Empresa.Razon & """"
    Cr.ReportSource = crptReport
    Cr.ReportFileName = Drive_Server & Path_Rpt & "TipoProducto.Rpt"
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
    btnBuscar.Visible = False
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
    Rs.Seek "=", codigo
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
    Rs.Seek "=", codigo
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        codigo.SetFocus
    Else
        Campos_Enabled True
        codigo.Enabled = False
        descripcion.SetFocus
        Registro_Leer
        btnGrabar.Enabled = True
        btnBuscar.Visible = False
    End If
Case "Eliminar"
    Rs.Seek "=", codigo
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        codigo.SetFocus
    Else
        Campos_Enabled False
        Registro_Leer
        If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo, "Atención") = vbYes Then
            Rs.Delete
        End If
        btnBuscar.Visible = True
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

Search.Muestra data_file, "Tipo Producto", "Codigo", "Descripcion", Obj, Objs

codigo = Search.codigo
descripcion = Search.descripcion

If codigo <> "" Then
    Rs.Seek "=", codigo
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
If IsObjBlanco(descripcion, "DESCRIPCIÓN", btnGrabar) Then Exit Function

'' valida que razón social no esté repetida
GoTo Sigue
If nuevo Then
    Rs.Index = "Descripcion"
    Rs.Seek "=", descripcion
    If Not Rs.NoMatch Then
        MsgBox "DESCRIPCIÓN YA EXISTE"
        btnGrabar.Value = tbrUnpressed
        Rs.Index = "Código"
        descripcion.SetFocus
        Exit Function
    End If
    btnGrabar.Value = tbrUnpressed
    Rs.Index = "Código"
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
