VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form sql_contratistas 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3240
   ClientLeft      =   3135
   ClientTop       =   2805
   ClientWidth     =   6195
   Icon            =   "sql_contratistas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6195
      _ExtentX        =   10927
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
   Begin VB.CheckBox Activo 
      Caption         =   "Activo"
      Height          =   195
      Left            =   1320
      TabIndex        =   14
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Comuna 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Direccion 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   1680
      Width           =   4575
   End
   Begin VB.TextBox Giro 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox Razon 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox codigo 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   5520
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.ComboBox Clasificacion 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2400
      Width           =   3375
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2880
      MousePointer    =   14  'Arrow and Question
      Picture         =   "sql_contratistas.frx":0442
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
            Picture         =   "sql_contratistas.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_contratistas.frx":0656
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_contratistas.frx":0768
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_contratistas.frx":087A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_contratistas.frx":098C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_contratistas.frx":0A9E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "C&lasificación"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&Comuna"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&Giro"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&Dirección"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Razón &Social"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&RUT"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "sql_contratistas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Obj As String, Objs As String, nuevo As Boolean, Accion As String, old_accion As String
'Dim Db As Database, Rs As Recordset
Private Db As Database
Private RsSc As New ADODB.Recordset
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnListar As Button, btnGrabar As Button, btnDesHacer As Button
Private lc(10) As Integer ' largo de los campos
Private m_Rut As String, RsCla As Recordset
Private m_TipoTabla As String
Private ac(9, 1) As String, av(9) As String, asn(9) As Boolean
Private Sub Campos_Enabled(Si As Boolean)
Codigo.Enabled = Si
Razon.Enabled = Si
Giro.Enabled = Si
Direccion.Enabled = Si
Comuna.Enabled = Si
Clasificacion.Enabled = Si
activo.Enabled = Si
End Sub
Private Sub btnBuscar_Click()

Dim cod_search As String, arreglo(1) As String
'arreglo(1) = "contratistas"
arreglo(1) = "razon_social"

sql_Search.Muestra "personas", "RUT", arreglo(), Obj, Objs, "contratista='S'"
cod_search = sql_Search.Codigo

If cod_search <> "" Then
    
    If Registro_Existe("personas", "contratista='S' AND rut='" & cod_search & "'") Then
    
        Codigo = cod_search
        After_Enter
        
    Else
    
        MsgBox Obj & " NO EXISTE"
        Codigo.SetFocus
        
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

'm_Rut = Rut_Formato(codigo)
m_Rut = Codigo ' 09/06/10

Select Case Accion
Case "Agregar"
    If Not Registro_Existe("personas", "contratista='S' AND rut='" & m_Rut & "'") Then
        Campos_Enabled True
        Codigo.Enabled = False
        Razon.SetFocus
        btnGrabar.Enabled = True
    Else
        Registro_Leer
        MsgBox Obj & " YA EXISTE"
        Campos_Inicializa
        Codigo.Enabled = True
        Codigo.SetFocus
    End If
Case "Modificar"
    If Not Registro_Existe("personas", "contratista='S' AND rut='" & m_Rut & "'") Then
        MsgBox Obj & " NO EXISTE"
        Codigo.SetFocus
    Else
        Campos_Enabled True
        Codigo.Enabled = False
        Razon.SetFocus
        Registro_Leer
        btnGrabar.Enabled = True
        btnBuscar.visible = False
    End If
Case "Eliminar"

    If Not Registro_Existe("personas", "contratista='S' AND rut='" & m_Rut & "'") Then
        MsgBox Obj & " NO EXISTE"
        Codigo.SetFocus
    Else
        Campos_Enabled False
        Registro_Leer
        If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo, "Atención") = vbYes Then
        
            av(1) = m_TipoTabla
            av(2) = Codigo.Text

            'Registro_Eliminar CnxSqlServer, "personas", numero.
            
        End If
        btnBuscar.visible = True
        Campos_Inicializa
        Codigo.Enabled = True
        Codigo.SetFocus
    End If
End Select
End Sub
Private Sub Registro_Leer()
Dim clas As String

Rs_Abrir RsSc, "SELECT * FROM personas WHERE rut='" & Codigo.Text & "'"

If Not RsSc.EOF Then

    Razon.Text = RsSc![razon_social]
    Giro.Text = NoNulo(RsSc!Giro)
    Direccion.Text = NoNulo(RsSc!Direccion)
    Comuna.Text = NoNulo(RsSc!Comuna)
    
    clas = NoNulo(RsSc!dato1)
    If clas = "" Then
        clas = " "
    Else
        RsCla.Seek "=", clas
        If RsCla.NoMatch Then
            clas = " "
        Else
            clas = RsCla!Codigo & " -> " & RsCla!Descripcion
        End If
    End If
    Clasificacion.Text = clas
    
    activo.Value = IIf(RsSc!activo = "S", 1, 0)

End If

End Sub
Private Sub Inicializa()
Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnListar = Toolbar.Buttons(4)
Set btnGrabar = Toolbar.Buttons(5)
Set btnDesHacer = Toolbar.Buttons(6)

Obj = "CONTRATISTA"
Objs = "CONTRATISTAS"

Accion = ""
old_accion = ""

btnBuscar.visible = False
btnBuscar.ToolTipText = "Busca " & StrConv(Obj, vbProperCase)

' largo del los campos
lc(1) = 10 'rut
lc(2) = 50 'razón social
lc(3) = 30 'giro
lc(4) = 50 'dirección
lc(5) = 30 'comuna
lc(6) = 10 'clasificación

ac(1, 0) = "rut"
ac(1, 1) = "'"
ac(2, 0) = "razon_social"
ac(2, 1) = "'"
ac(3, 0) = "giro"
ac(3, 1) = "'"
ac(4, 0) = "direccion"
ac(4, 1) = "'"
ac(5, 0) = "comuna"
ac(5, 1) = "'"
ac(6, 0) = "dato1"
ac(6, 1) = "'"
ac(7, 0) = "activo"
ac(7, 1) = "'"
ac(8, 0) = "contratista"
ac(8, 1) = "'"

' false es campo clave
asn(1) = False
asn(2) = True
asn(3) = True
asn(4) = True
asn(5) = True
asn(6) = True
asn(7) = True
asn(8) = True

With Codigo
'    .Mask = String(lc(1), "&") '"A" con este no deja espacios en blanco
    .MaxLength = lc(1)
'    .PromptInclude = False
End With

With Razon
'    .Mask = ">" & String(lc(2), "&")
    .MaxLength = lc(2)
'    .PromptInclude = False
End With

With Giro
'    .Mask = ">" & String(lc(3), "&")
    .MaxLength = lc(3)
'    .PromptInclude = False
End With

With Direccion
'    .Mask = ">" & String(lc(4), "&")
    .MaxLength = lc(4)
'    .PromptInclude = False
End With

With Comuna
'    .Mask = ">" & String(lc(5), "&")
    .MaxLength = lc(5)
'    .PromptInclude = False
End With

With Clasificacion
End With

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
If KeyAscii = vbKeyReturn Then Clasificacion.SetFocus
End Sub
Private Sub Clasificacion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Razon.SetFocus
End Sub
Private Sub Form_Load()

Inicializa

Me.Caption = "MANTENCIÓN DE " & Objs
Botones_Enabled True, True, True, True, False, False

' abre archivos
Set Db = OpenDatabase(data_file)
'Set Rs = Db.OpenRecordset("Contratistas")
'Rs.Index = "RUT"

Set RsCla = Db.OpenRecordset("Clasificacion de Contratistas")
RsCla.Index = "Codigo"

Clasi_Poblar

Campos_Inicializa

nuevo = False

m_TipoTabla = "CON"

End Sub
Private Sub Clasi_Poblar()
' puebla combo de clasificacion para contratista
Clasificacion.AddItem " "
If RsCla.RecordCount > 0 Then RsCla.MoveFirst
Do While Not RsCla.EOF
    Clasificacion.AddItem RsCla!Codigo & " -> " & RsCla!Descripcion
    RsCla.MoveNext
Loop
End Sub
Private Sub Form_Unload(Cancel As Integer)
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
    cr.ReportSource = crptReport
    cr.ReportFileName = Drive_Server & Path_Rpt & "Contratistas.Rpt"
    cr.Action = 1
    MousePointer = vbDefault
Case "Grabar"
    If Valida(nuevo) Then
        Registro_Grabar nuevo
    Else
        Exit Sub
    End If
Case "Deshacer"
    Campos_Inicializa
    btnBuscar.visible = False
End Select

Select Case Button.Index
Case 5  ' btnGrabar
    Campos_Inicializa
    
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
Private Function Valida(nuevo As Boolean) As Boolean
Valida = False
If IsObjBlanco(Razon, "RAZÓN SOCIAL", btnGrabar) Then Exit Function
'If IsObjBlanco(Giro, "GIRO", btnGrabar) Then Exit Function
'If IsObjBlanco(Direccion, "DIRECCIÓN", btnGrabar) Then Exit Function
'If IsObjBlanco(Comuna, "COMUNA", btnGrabar) Then Exit Function

'' valida que razón social no esté repetida
If nuevo Then
'    Rs.Index = "Razon Social"
'    Rs.Seek "=", Razon
'    If Not Rs.NoMatch Then
'        MsgBox "RAZÓN SOCIAL YA EXISTE"
'        btnGrabar.Value = tbrUnpressed
'        Rs.Index = "RUT"
'        Razon.SetFocus
'        Exit Function
'    End If
'    btnGrabar.Value = tbrUnpressed
'    Rs.Index = "RUT"
End If
''

Valida = True
End Function
Private Sub Registro_Grabar(nuevo As Boolean)

Dim clas As String, p As Integer

'av(1) = m_TipoTabla
av(1) = Codigo.Text
av(2) = Razon.Text
av(3) = Giro.Text
av(4) = Direccion.Text
av(5) = Comuna.Text
clas = Clasificacion.Text
If clas <> " " Then
    p = InStr(1, clas, "->")
    clas = Left(clas, p - 2)
End If
av(6) = clas
av(7) = IIf(activo.Value = 0, "N", "S")
av(8) = "S"

If nuevo Then

    Registro_Agregar CnxSqlServer_scp0, "personas", ac, av, 8
    
Else

    Registro_Modificar CnxSqlServer_scp0, "personas", ac, av, asn, 8
    
End If

If False Then

    RsSc![Razon Social] = Razon.Text
    RsSc!Giro = Giro.Text
    RsSc!Direccion = Direccion.Text
    RsSc!Comuna = Comuna.Text
    
    clas = Clasificacion.Text
    If clas <> " " Then
        p = InStr(1, clas, "->")
        clas = Left(clas, p - 2)
    End If
    RsSc!Clasificacion = clas
    
    RsSc!activo = activo.Value
    
    RsSc.Update

End If

Campos_Inicializa
Accion = old_accion

End Sub
Private Sub Campos_Inicializa()
Codigo.Text = ""
Razon.Text = ""
Giro.Text = ""
Direccion.Text = ""
Comuna.Text = ""
Clasificacion.Text = " "
activo.Value = False
Campos_Enabled False
End Sub
