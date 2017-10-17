VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form sql_TablasMaestras 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2805
   ClientLeft      =   3135
   ClientTop       =   2805
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox imputable 
      Caption         =   "&Imputable"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1920
      Width           =   1695
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   8
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
   Begin VB.CheckBox activo 
      Caption         =   "&Activo"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox observacion 
      Height          =   300
      Left            =   1440
      TabIndex        =   6
      Top             =   1440
      Width           =   3855
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
      Left            =   3000
      MousePointer    =   14  'Arrow and Question
      Picture         =   "sql_TablasMaestras.frx":0000
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
            Picture         =   "sql_TablasMaestras.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_TablasMaestras.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_TablasMaestras.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_TablasMaestras.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_TablasMaestras.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_TablasMaestras.frx":065C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "&Observación"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   975
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
End
Attribute VB_Name = "sql_TablasMaestras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Obj As String, Objs As String, nuevo As Boolean, Accion As String, old_accion As String
Private RsTabla As New ADODB.Recordset
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnListar As Button, btnGrabar As Button, btnDesHacer As Button
Private m_TipoTabla As String
Private ac(6, 1) As String, av(6) As String, asn(6) As Boolean
'Private aCodigoDescripcion(999, 3) As String ' arreglo de codigos y descripciones, en centrosdecosto->  0:codigo, 1:descripcion, 2:imputable, 3:orden
Public Property Let TipoTabla(ByVal New_Opcion As String)
m_TipoTabla = New_Opcion
End Property
Private Sub Form_Load()

Inicializa

Me.Caption = "MANTENCIÓN DE " & Objs
Botones_Enabled True, True, True, True, False, False

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

Case "GER" ' para noconformidad

    Obj = "GERENCIA"
    Objs = "GERENCIAS"
    
Case "CUCO"

    Obj = "CUENTA CONTABLE"
    Objs = "CUENTAS CONTABLES"
    
Case "CECO"

    Obj = "CENTRO DE COSTO"
    Objs = "CENTROS DE COSTO"
    
End Select

ac(1, 0) = "tipo"
ac(1, 1) = "'"
ac(2, 0) = "codigo"
ac(2, 1) = "'"
ac(3, 0) = "descripcion"
ac(3, 1) = "'"
ac(4, 0) = "observacion"
ac(4, 1) = "'"
ac(5, 0) = "dato1"
ac(5, 1) = "'"
ac(6, 0) = "activo"
ac(6, 1) = "'"

Accion = ""
old_accion = ""

btnBuscar.visible = False
btnBuscar.ToolTipText = "Busca " & StrConv(Obj, vbProperCase)

Codigo.MaxLength = 10
Descripcion.MaxLength = 50

End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Arch As String, i As Integer
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
    cr.DataFiles(0) = repo_file & ".MDB"
    cr.Formulas(0) = "RAZON SOCIAL=""" & Empresa.Razon & """"
    cr.ReportSource = crptReport
    
    Select Case m_TipoTabla
    
    Case "CUCO"
        
        cr.Formulas(1) = "TITULO=""" & "Cuentas Contables" & """"
    
        'i = cuentasContablesLeer(aCodigoDescripcion)
        sql2access cuentasContablesTotal
            
        cr.ReportFileName = Drive_Server & Path_Rpt & "centrosdecosto.rpt"
    
    Case "CECO"
    
        cr.Formulas(1) = "TITULO=""" & "Centros de Costo" & """"
        
'        i = centrosCostoLeer(aCodigoDescripcion)
        sql2access centrosCostoTotal
            
        cr.ReportFileName = Drive_Server & Path_Rpt & "centrosdecosto.rpt"
        
    End Select
    
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

'    RsTabla.Open "SELECT * FROM maestros WHERE tipo='" & m_TipoTabla & "' AND codigo='" & codigo.Text & "'", CnxSqlServer

    If Not Registro_Existe("maestros", "tipo='" & m_TipoTabla & "' AND codigo='" & Codigo.Text & "'") Then
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

'    RsTabla.Open "SELECT * FROM usuarios WHERE tipo='" & m_TipoTabla & "' AND codigo='" & codigo.Text & "'", CnxSqlServer
    If Not Registro_Existe("maestros", "tipo='" & m_TipoTabla & "' AND codigo='" & Codigo.Text & "'") Then
'    If RsTabla.EOF Then
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
    
    If Not Registro_Existe("maestros", "tipo='" & m_TipoTabla & "' AND codigo='" & Codigo.Text & "'") Then
'    If RsTabla.EOF Then
        MsgBox Obj & " NO EXISTE"
        Codigo.SetFocus
    Else
        Campos_Enabled False
        Registro_Leer
        If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo, "Atención") = vbYes Then
        
            av(1) = m_TipoTabla
            av(2) = UCase(Codigo.Text)
            av(3) = Descripcion.Text
            av(4) = Observacion.Text
            av(5) = IIf(imputable.Value = 0, "N", "S")
            av(6) = IIf(activo.Value = 0, "N", "S")
            
            asn(1) = False
            asn(2) = False
            asn(3) = True
            asn(4) = True
            asn(5) = True
            asn(6) = True
        
            'Registro_Eliminar CnxSqlServer, "maestros", ac, av, asn, 6
            
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
Descripcion.Enabled = Si
Observacion.Enabled = Si
imputable.Enabled = Si
activo.Enabled = Si
End Sub
Private Sub btnBuscar_Click()

Dim arreglo(1) As String
arreglo(1) = "descripcion"

sql_Search.Muestra "maestros", "Codigo", arreglo(), Obj, Objs, "Tipo='" & m_TipoTabla & "'"

Codigo = sql_Search.Codigo
Descripcion = sql_Search.Descripcion

If Codigo <> "" Then

'    RsTabla.Open "SELECT * FROM maestros WHERE tipo='" & m_TipoTabla & "' AND codigo='" & codigo.Text & "'", CnxSqlServer
    If Not Registro_Existe("maestros", "tipo='" & m_TipoTabla & "' AND codigo='" & Codigo.Text & "'") Then
'    If RsTabla.EOF Then
        MsgBox Obj & " NO EXISTE"
        Codigo.SetFocus
    Else
        After_Enter
    End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
'RsTabla.Close
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
Valida = True
End Function
Private Sub Registro_Grabar(nuevo As Boolean)

save:

av(1) = m_TipoTabla
av(2) = UCase(Codigo.Text)
av(3) = Descripcion.Text
av(4) = Observacion.Text
av(5) = IIf(imputable.Value = 0, "N", "S")
av(6) = IIf(activo.Value = 0, "N", "S")

asn(1) = False
asn(2) = False
asn(3) = True
asn(4) = True
asn(5) = True
asn(6) = True

If nuevo Then

    Registro_Agregar CnxSqlServer_scp0, "maestros", ac, av, 6
    
Else

    Registro_Modificar CnxSqlServer_scp0, "maestros", ac, av, asn, 6
    
End If

Campos_Limpiar

Accion = old_accion

End Sub
Private Sub Registro_Leer()

RsTabla.Open "SELECT * FROM maestros WHERE tipo='" & m_TipoTabla & "' AND codigo='" & Codigo.Text & "'", CnxSqlServer_scp0

If Not RsTabla.EOF Then

    Descripcion.Text = RsTabla!Descripcion
    Observacion.Text = NoNulo(RsTabla!Observacion)
    imputable.Value = IIf(RsTabla!dato1 = "S", 1, 0)
    activo.Value = IIf(RsTabla!activo = "S", 1, 0)
    
End If

RsTabla.Close

End Sub
Private Sub Campos_Limpiar()
Codigo.Text = ""
Descripcion.Text = ""
Observacion.Text = ""
imputable.Value = 0
activo.Value = 0
Campos_Enabled False
End Sub
Private Sub Descripcion_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Observacion_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub sql2access(nr As Integer)
' traspasa datos de tabla maestra a tabla access para reporte
Dim i As Integer
Dim DbR As Database, RsR As Recordset
Set DbR = OpenDatabase(repo_file)
DbR.Execute "DELETE * FROM maestros"
Set RsR = DbR.OpenRecordset("maestros")
With RsR

Select Case m_TipoTabla
Case "CUCO"
    For i = 1 To nr
        .AddNew
        !Codigo = aCuCo(i, 0) ' aCodigoDescripcion(i, 0)
        !Descripcion = aCuCo(i, 1) ' aCodigoDescripcion(i, 1)
        !dato1 = aCuCo(i, 2) ' aCodigoDescripcion(i, 2)
        !orden = aCuCo(i, 3) ' aCodigoDescripcion(i, 3)
        .Update
    Next
Case "CECO"
    For i = 1 To nr
        .AddNew
        !Codigo = aCeCo(i, 0) ' aCodigoDescripcion(i, 0)
        !Descripcion = aCeCo(i, 1) ' aCodigoDescripcion(i, 1)
        !dato1 = aCeCo(i, 2) ' aCodigoDescripcion(i, 2)
        !orden = aCeCo(i, 3) ' aCodigoDescripcion(i, 3)
        .Update
    Next
End Select

End With

End Sub
