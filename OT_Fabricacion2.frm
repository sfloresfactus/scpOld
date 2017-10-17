VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form OT_Fabricacion2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orden De Trabajo"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCopiar 
      Caption         =   "Copiar Fechas"
      Height          =   315
      Left            =   7200
      TabIndex        =   27
      Top             =   1200
      Width           =   1215
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ingresar Nueva OT"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar OT"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar OT"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir OT"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "DesHacer"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Grabar OT"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mantención de Subcontratistas"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox Razon 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6120
      TabIndex        =   23
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton btnSearch 
      Height          =   300
      Left            =   5640
      Picture         =   "OT_Fabricacion2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   720
      Width           =   300
   End
   Begin VB.TextBox Nv 
      Height          =   300
      Left            =   2280
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox ComboNV 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   3
      Left            =   120
      MaxLength       =   30
      TabIndex        =   17
      Top             =   5700
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   2
      Left            =   120
      MaxLength       =   30
      TabIndex        =   16
      Top             =   5400
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   1
      Left            =   120
      MaxLength       =   30
      TabIndex        =   15
      Top             =   5100
      Width           =   5000
   End
   Begin MSMask.MaskEdBox EditFecha 
      Height          =   300
      Left            =   8040
      TabIndex        =   19
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   327680
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox ComboMarca 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox ComboPlano 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtEditOT 
      Height          =   285
      Left            =   8040
      TabIndex        =   11
      Text            =   "txtEditOT"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   0
      Left            =   120
      MaxLength       =   30
      TabIndex        =   14
      Top             =   4800
      Width           =   5000
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   960
      TabIndex        =   6
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   327680
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Numero 
      Height          =   300
      Left            =   960
      TabIndex        =   4
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "##########"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid Detalle 
      Height          =   2925
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5159
      _Version        =   327680
      ScrollBars      =   2
   End
   Begin MSMask.MaskEdBox Rut 
      Height          =   300
      Left            =   4080
      TabIndex        =   24
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   327680
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   2760
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OT_Fabricacion2.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OT_Fabricacion2.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OT_Fabricacion2.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OT_Fabricacion2.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OT_Fabricacion2.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OT_Fabricacion2.frx":065C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OT_Fabricacion2.frx":076E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "Contratista"
      Height          =   255
      Index           =   4
      Left            =   6120
      TabIndex        =   26
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "RUT"
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   25
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "FABRICACIÓN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label TotalPrecio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6720
      TabIndex        =   21
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label TotalKilos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   5400
      TabIndex        =   20
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "OBRA"
      Height          =   255
      Index           =   7
      Left            =   3480
      TabIndex        =   8
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "FECHA"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "OT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lbl 
      Caption         =   "Nº"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
End
Attribute VB_Name = "OT_Fabricacion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button

'Private DbD As Database , RsCl As Recordset
'Private RsSc As Recordset
'Private SqlRsSc As New ADODB.Recordset
Private Dbm As Database, RsOTc As Recordset, RsOTd As Recordset, RsITOfd As Recordset

' por ahora solo estos dos tablas seran tratadas con sql_mdb
Private RsPc As Recordset, RsPd As Recordset

'Private RsNVc As Recordset

Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, d As Variant

Private n_filas As Integer, n_columnas As Integer
Private prt As Printer
Private Rev(2999) As String
Private n7 As Double, n8 As Double, n12 As Double
Private linea As String, m_Nv As Double
Private DbH As Database, RsOTcH As Recordset
Private n_marcas As Integer
' 0: numero,  1: nombre obra
Private a_Nv(2999, 1) As String, m_NvArea As Integer

' variables para impresion de etiq
Private m_obra As String, m_Plano As String, m_Rev As String, m_Marca As String, m_Peso As Double
Private m_ClienteRazon As String, AjusteX As Double, AjusteY As Double

'///////////////////////////////
' variable para sql
Private sql As String
'///////////////////////////////
Private Sub btnCopiar_Click()
Dim m_f1 As String, m_f2 As String, i As Integer
m_f1 = Detalle.TextMatrix(1, 10)
If m_f1 = "" Then
    MsgBox "Debe digitar fecha en linea 1"
    Exit Sub
End If
m_f2 = Detalle.TextMatrix(1, 11)
If m_f2 = "" Then
    MsgBox "Debe digitar fecha en linea 1"
    Exit Sub
End If

For i = 1 To n_filas
    If Detalle.TextMatrix(i, 1) <> "" Then
        Detalle.TextMatrix(i, 10) = m_f1
        Detalle.TextMatrix(i, 11) = m_f2
    End If
Next
    
End Sub
Private Sub Form_Load()

'Set DbD = OpenDatabase(data_file)

' abre archivos
'If Not Usando_SQL Then
'    Set RsSc = DbD.OpenRecordset("Contratistas")
'    RsSc.Index = "RUT"
'End If

'Set RsCl = DbD.OpenRecordset("Clientes")
'RsCl.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)

If Not Usuario.ObrasTerminadas Then

    ' si usuario esta en obras en proceso, abre movs historico
    Dim hist_file As String
    hist_file = Movs_Path(Empresa.Rut, True)
    Set DbH = OpenDatabase(hist_file)
    Set RsOTcH = DbH.OpenRecordset("OT Fab Cabecera")
    RsOTcH.Index = "Numero"
    
End If

Set RsOTc = Dbm.OpenRecordset("OT Fab Cabecera")
RsOTc.Index = "Numero"

Set RsOTd = Dbm.OpenRecordset("OT Fab Detalle")
RsOTd.Index = "Numero-Linea"

Set RsITOfd = Dbm.OpenRecordset("ITO Fab Detalle")
RsITOfd.Index = "NV-Plano-Marca"

'Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
'RsNVc.Index = Nv_Index ' "Número"

nvListar Usuario.Nv_Activas

ComboNV.AddItem " "
For i = 1 To nvTotal
    a_Nv(i, 0) = aNv(i).Numero
    a_Nv(i, 1) = aNv(i).obra
    ComboNV.AddItem Format(aNv(i).Numero, "0000") & " - " & aNv(i).obra
Next

'Set RsPc = Dbm.OpenRecordset("Planos Cabecera")
'RsPc.Index = "NV-Plano"

Set RsPc = Dbm.Recordsets(0) ' hubo que poner esta linea para setear rspc
Set RsPd = Dbm.Recordsets(0)

'Set RsPd = Dbm.OpenRecordset("Planos Detalle")
'RsPd.Index = "NV-Plano-Item"

Inicializa
Detalle_Config

Privilegios

AjusteX = 0
AjusteY = -0.3

m_NvArea = 0

End Sub
Private Sub Inicializa()

Obj = "ORDEN DE TRABAJO"
Objs = "ÓRDENES DE TRABAJO"

Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnDesHacer = Toolbar.Buttons(6)
Set btnGrabar = Toolbar.Buttons(7)

Accion = ""
'old_accion = ""

btnSearch.visible = False
btnSearch.ToolTipText = "Busca Contatista"
Campos_Enabled False

Nv.MaxLength = 4

End Sub
Private Sub Detalle_Config()
Dim i As Integer, ancho As Integer

n_filas = 25
n_columnas = 14 ' 13

Detalle.Left = 100
Detalle.WordWrap = True
Detalle.RowHeight(0) = 450
Detalle.Cols = n_columnas + 1
Detalle.Rows = n_filas + 1

Detalle.TextMatrix(0, 0) = ""
Detalle.TextMatrix(0, 1) = "Plano"
Detalle.TextMatrix(0, 2) = "Rev"                '*
Detalle.TextMatrix(0, 3) = "Marca"
Detalle.TextMatrix(0, 4) = "Descripción"        '*
Detalle.TextMatrix(0, 5) = "Cant Total"     '*
Detalle.TextMatrix(0, 6) = "Total Asig"  '*
Detalle.TextMatrix(0, 7) = "Cant a Asi"
Detalle.TextMatrix(0, 8) = "Peso Unitario"      '*
Detalle.TextMatrix(0, 9) = "Peso TOTAL"         '*
Detalle.TextMatrix(0, 10) = "Fecha Piezas" ' nuevo
Detalle.TextMatrix(0, 11) = "Fecha Entrega"
Detalle.TextMatrix(0, 12) = "Precio Unitario"
Detalle.TextMatrix(0, 13) = "Precio TOTAL"      '*
Detalle.TextMatrix(0, 14) = "superfice unitaria" ' campo oculto

Detalle.ColWidth(0) = 250
Detalle.ColWidth(1) = 2100 ' plano
Detalle.ColWidth(2) = 400  ' rev
Detalle.ColWidth(3) = 2300 ' marca
Detalle.ColWidth(4) = 1800 ' descripcion

Detalle.ColWidth(5) = 500 ' cant total
Detalle.ColWidth(6) = 500 ' total asig
Detalle.ColWidth(7) = 500 ' cant a asi

Detalle.ColWidth(8) = 700 ' peso uni
Detalle.ColWidth(9) = 800 ' peso total

Detalle.ColWidth(10) = 800 ' fecha piezas
Detalle.ColWidth(11) = 800 ' fecha entrega

Detalle.ColWidth(12) = 700 ' $ unitario
Detalle.ColWidth(13) = 800 ' $ total

Detalle.ColWidth(14) = 0 ' suni

'Detalle.ColAlignment(2) = 0

ancho = 350 ' con scroll vertical
'ancho = 100 ' sin scroll vertical

TotalKilos.Width = Detalle.ColWidth(9)
TotalPrecio.Width = Detalle.ColWidth(13) * 1.2 ' 12 ?
For i = 0 To n_columnas
    If i = 9 Then TotalKilos.Left = ancho + Detalle.Left - 300 ' peso total
    If i = 13 Then TotalPrecio.Left = ancho + Detalle.Left - 400 ' precio total '300
    ancho = ancho + Detalle.ColWidth(i)
Next

Detalle.Width = ancho
Me.Width = ancho + Detalle.Left * 2

' col y row fijas
'Detalle.BackColorFixed = vbCyan

' establece colores a columnas
' columnas    modificables : NEGRAS
' columnas no modificables : ROJAS
For i = 1 To n_filas
    Detalle.TextMatrix(i, 0) = i
    Detalle.Row = i
    Detalle.col = 2
    Detalle.CellForeColor = vbRed
    Detalle.CellAlignment = flexAlignLeftCenter
    Detalle.col = 3
    Detalle.CellAlignment = flexAlignLeftCenter
    Detalle.col = 4
    Detalle.CellForeColor = vbRed
    Detalle.col = 5
    Detalle.CellForeColor = vbRed
    Detalle.col = 6
    Detalle.CellForeColor = vbRed
    Detalle.col = 8
    Detalle.CellForeColor = vbRed
    Detalle.col = 9
    Detalle.CellForeColor = vbRed
    Detalle.col = 13
    Detalle.CellForeColor = vbRed
Next

txtEditOT.Text = ""

'Detalle.ScrollTrack = True 'no se nota
'Detalle.TextMatrix(1, 1) = "hola" 'ok

End Sub
Private Sub ComboNV_Click()

MousePointer = vbHourglass

Dim m_Plano As String

ComboPlano.visible = False
ComboMarca.visible = False

i = 0
m_Nv = Val(Left(ComboNV.Text, 6))
If m_Nv = 0 Then
    Nv.Text = ""
Else
    Nv.Text = m_Nv
End If

ComboPlano.Clear

ComboPlano.AddItem " " ' para "borrar" linea
Rev(i) = " "

'sql = "SELECT * FROM " & Tabla_PlanosCabecera
'sql = sql & " WHERE nv=" & m_Nv
'Rs_Abrir_MDB Dbm, RsPc, sql

sql = "SELECT * FROM " & Tabla_PlanosCabecera
sql = sql & " WHERE nv=" & m_Nv
Rs_Abrir_MDB Dbm, RsPc, sql

If False Then
    RsPc.Seek ">=", m_Nv, 0
    If Not RsPc.NoMatch Then
        Do While Not RsPc.EOF
        
            If RsPc!Nv = m_Nv Then
    
    '            sql = "SELECT * FROM [planos detalle] WHERE nv=" & m_Nv
    '            Rs_Abrir_MDB Dbm, RsPd, sql
    
                ComboPlano.AddItem RsPc![Plano]
                i = i + 1
                Rev(i) = RsPc![Rev]
                
            Else
                Exit Do
            End If
            RsPc.MoveNext
        Loop
    End If
End If

m_Plano = ""
Do While Not RsPc.EOF

'    If RsPc![Cantidad total] > RsPc![ot fab] Then
        ComboPlano.AddItem RsPc![Plano]
        i = i + 1
        Rev(i) = RsPc![Rev]
'    End If
        
    RsPc.MoveNext
    
Loop

' limpia combos de marca
ComboMarca.Clear
For i = 1 To n_filas
    Detalle.TextMatrix(i, 1) = ""
    Detalle.TextMatrix(i, 2) = ""
Next

MousePointer = vbDefault

End Sub
Private Sub ComboPlano_Click()
' el número del plano NO es único, puede repetirse en otras NV
Dim old_plano As String
Dim filaFlex As Integer
Dim np As String, indice_plano As Integer, Marca_Unica_enPlano As String

old_plano = Detalle

filaFlex = Detalle.Row

np = ComboPlano.Text

indice_plano = 0
ComboMarca.Clear

sql = "SELECT * FROM " & Tabla_PlanosDetalle
sql = sql & " WHERE nv=" & m_Nv & " AND plano='" & np & "'"
sql = sql & " ORDER BY marca"
Rs_Abrir_MDB Dbm, RsPd, sql

Do While Not RsPd.EOF
    indice_plano = indice_plano + 1
    Marca_Unica_enPlano = RsPd!Marca
    ComboMarca.AddItem Marca_Unica_enPlano
    RsPd.MoveNext
Loop

ComboPlano.visible = False
Detalle = ComboPlano.Text

If Detalle <> old_plano Then
    For i = 2 To n_columnas
        Detalle.TextMatrix(filaFlex, i) = ""
    Next
End If

' revision
If ComboPlano.ListIndex > 0 Then Detalle.TextMatrix(filaFlex, 2) = Rev(ComboPlano.ListIndex)

' recalcula total
Totales_Actualiza

If indice_plano = 1 Then
    ' plano tiene una sola marca
    ComboMarca.ListIndex = 0 ' para que pueble grid
    Detalle.TextMatrix(filaFlex, 3) = Marca_Unica_enPlano ' ComboMarca.Text
End If

End Sub
Private Sub ComboPlano_LostFocus()
ComboPlano.visible = False
End Sub
Private Sub ComboMarca_Click()
Dim m_Plano As String, m_Marca As String, fil As Integer, repetido As Boolean
Dim c_total As Integer, c_ot As Integer

fil = Detalle.Row
ComboMarca.visible = False
m_Plano = Detalle.TextMatrix(fil, 1)
m_Marca = ComboMarca.Text

'///
' verifica si Plano-Marca ya están en esta OT
For i = 1 To n_filas
    If fil <> i Then
        If m_Plano = Detalle.TextMatrix(i, 1) And m_Marca = Detalle.TextMatrix(i, 3) Then
            Beep
            MsgBox "MARCA YA EXISTE EN OT"
            Detalle.Row = i
            Detalle.col = 3
            Detalle.SetFocus
            Exit Sub
        End If
    End If
Next
'///

'Detalle = m_Marca ' bien hasta aqui
Detalle.TextMatrix(fil, 3) = m_Marca ' bien hasta aqui
c_total = 0
c_ot = 0

' busca marca en plano
sql = "SELECT * FROM " & Tabla_PlanosDetalle
sql = sql & " WHERE nv=" & m_Nv & " AND plano='" & m_Plano & "' AND marca='" & m_Marca & "'"
Rs_Abrir_MDB Dbm, RsPd, sql

Do While Not RsPd.EOF
    
    c_total = RsPd![Cantidad Total]
    c_ot = RsPd![OT fab]
    
    ' verifica que quede algo por asignar
    If c_total - c_ot <= 0 Then
        Beep
        MsgBox "No queda NADA por asignar" & vbCr & _
                "de la marca """ & m_Marca & """"
        Detalle.TextMatrix(fil, 3) = ""
        Detalle.SetFocus
        Exit Sub
    End If
    
    Detalle.TextMatrix(fil, 4) = RsPd!Descripcion
    Detalle.TextMatrix(fil, 5) = c_total
    Detalle.TextMatrix(fil, 6) = c_ot
    Detalle.TextMatrix(fil, 8) = Replace(RsPd![Peso], ",", ".")
    
    Detalle.TextMatrix(fil, 14) = Replace(RsPd![Superficie], ",", ".")
    
'            Linea_Actualiza
    RsPd.MoveNext
    
Loop

End Sub
Private Sub ComboMarca_LostFocus()
ComboMarca.visible = False
End Sub
Private Sub Fecha_GotFocus()
ComboPlano.visible = False
ComboMarca.visible = False
End Sub
Private Sub Fecha_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Fecha_LostFocus()
d = Fecha_Valida(Fecha, Now)
End Sub
Private Sub Form_Unload(Cancel As Integer)
DataBases_Cerrar
End Sub
Private Function Archivo_Abrir(Db As Database, Archivo As String) As Boolean
' intenta abrir archivo en forma compartida
Archivo_Abrir = True
On Error GoTo Error
Set Db = OpenDatabase(Archivo)
Exit Function
Error:
MsgBox "Archivo YA está en uso"
Archivo_Abrir = False
End Function
Private Sub Numero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then After_Enter
End Sub
Private Sub After_Enter()
If Val(Numero.Text) < 1 Then Beep: Exit Sub
Select Case Accion
Case "Agregando"
    
    If Not Usuario.ObrasTerminadas Then
        ' busca en historico
        RsOTcH.Seek "=", Numero.Text
        If Not RsOTcH.NoMatch Then
            MsgBox Obj & " YA EXISTE" & Chr(10) & "EN OBRAS TERMINADAS"
            Detalle.Enabled = False
            Campos_Limpiar
            Numero.Enabled = True
            Numero.SetFocus
            
            Exit Sub
            
        End If
    End If
    
    RsOTc.Seek "=", Numero.Text
    
    If RsOTc.NoMatch Then
        Campos_Enabled True
        Numero.Enabled = False
        Fecha.SetFocus
        btnGrabar.Enabled = True
        btnSearch.visible = True
    Else
        Doc_Leer
        MsgBox Obj & " YA EXISTE"
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
    End If
    
Case "Modificando"

    RsOTc.Seek "=", Numero.Text
    If RsOTc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Botones_Enabled 0, 0, 0, 0, 1, 1
        Doc_Leer
        Campos_Enabled True
        Numero.Enabled = False
        
        'busca si hay algo recibido de esta ot
'        If OT_ItoBuscar Then
'            MsgBox "ADVERTENCIA" & vbCr & "ITOf Nº" & RsITOfD!Número
'        End If

'        OT_ItoBuscar

        'muestra itos afectadas
'        aqui voy
'        ITOfaAnular.NV = m_NV
'        ITOfaAnular.PlanoNombre = "m_Plano"
'        ITOfaAnular.Rev = "m_Rev"
'        ITOfaAnular.NumerodeMarcas = n_marcas
'        ITOfaAnular.Show 1
        
    End If

Case "Eliminando"
    RsOTc.Seek "=", Numero.Text
    If RsOTc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Doc_Leer
'        Numero.Enabled = False
        If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo) = vbYes Then
        
            Doc_Eliminar
            
            PlanoDetalle_Actualizar Dbm, Nv.Text, m_NvArea, RsOTd, "nv-plano-marca", "ot fab"
            
        End If
        Campos_Limpiar
        Campos_Enabled False
        Numero.Enabled = True
        Numero.SetFocus
    End If
   
Case "Imprimiendo"
    
    RsOTc.Seek "=", Numero.Text
    If RsOTc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Doc_Leer
        Numero.Enabled = False
        Detalle.Enabled = True
    End If
    
End Select

End Sub
Private Sub Doc_Leer()

Dim obra As String
Dim m_resta As Integer
Dim m_PesoActual As Double ' peso actual en base a plano
' CABECERA

'On Error Resume Next

Fecha.Text = Format(RsOTc!Fecha, Fecha_Format)
m_Nv = RsOTc!Nv
Nv.Text = m_Nv
Rut.Text = RsOTc![Rut contratista]

If Usuario.Tipo = "C" Then
    ' contratista
    If Usuario.Rut <> Rut.Text Then
        MsgBox "OT es de Otro Contratista"
'        Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
        ' desHacer
        Toolbar_ButtonClick btnDesHacer
        Exit Sub
    End If
End If

obra = nv2Obra(m_Nv)

If obra <> "" Then

    ComboNV.Text = Format(m_Nv, "0000") & " - " & obra
    ComboNV_Click
    
    m_ClienteRazon = ""
    
End If

Obs(0).Text = NoNulo(RsOTc![Observacion 1])
Obs(1).Text = NoNulo(RsOTc![Observacion 2])
Obs(2).Text = NoNulo(RsOTc![Observacion 3])
Obs(3).Text = NoNulo(RsOTc![Observacion 4])

'DETALLE
'RsPd.Index = "NV-Plano-Marca"

RsOTd.Seek "=", Numero.Text, 1
If Not RsOTd.NoMatch Then
    Do While Not RsOTd.EOF
        If RsOTd!Numero = Numero.Text Then
        
            i = RsOTd!linea
            
            Detalle.TextMatrix(i, 1) = RsOTd!Plano
            Detalle.TextMatrix(i, 2) = RsOTd!Rev
            Detalle.TextMatrix(i, 3) = RsOTd!Marca
            
'            RsPd.Seek "=", m_Nv, m_NvArea, RsOTd!Plano, RsOTd!Marca
            sql = "SELECT * FROM " & Tabla_PlanosDetalle
            sql = sql & " WHERE nv=" & m_Nv & " AND plano='" & RsOTd!Plano & "' AND marca='" & RsOTd!Marca & "'"
            Rs_Abrir_MDB Dbm, RsPd, sql
            
            If RsPd.RecordCount > 0 Then
            
                If Not RsPd.NoMatch Then
                
                    Detalle.TextMatrix(i, 4) = RsPd!Descripcion
                    Detalle.TextMatrix(i, 5) = RsPd![Cantidad Total]
                    
                    m_resta = IIf(Accion = "Modificando", RsOTd!Cantidad, 0)
                    Detalle.TextMatrix(i, 6) = RsPd![OT fab] - m_resta
                    
                    m_PesoActual = RsPd![Peso]
                    
                    ' superficie unitaria, SUNI
                    Detalle.TextMatrix(i, 14) = RsPd![Superficie]
                    
                End If
            End If
            
            Detalle.TextMatrix(i, 7) = RsOTd!Cantidad
'            Detalle.TextMatrix(i, 8) = RsOTd![Peso Unitario] ' peso unitario en la ot
            Detalle.TextMatrix(i, 8) = m_PesoActual ' peso unitario en el plano
            Detalle.TextMatrix(i, 10) = NoNulo(RsOTd![fecha2]) ' fecha piezas
            Detalle.TextMatrix(i, 11) = RsOTd![Fecha Entrega]
            Detalle.TextMatrix(i, 12) = RsOTd![Precio Unitario]
            
            n7 = m_CDbl(Detalle.TextMatrix(i, 7)) ' cant
            n8 = m_CDbl(Detalle.TextMatrix(i, 8)) ' peso uni
            n12 = m_CDbl(Detalle.TextMatrix(i, 12)) ' $ uni
            
            Detalle.TextMatrix(i, 9) = Format(n7 * n8, num_Formato)
            Detalle.TextMatrix(i, 13) = Format(n7 * n8 * n12, num_fmtgrl)
            
            If m_PesoActual <> RsOTd![Peso Unitario] Then
                MsgBox "Peso Unitario en Plano es distinto al Peso Unitario en esta OT, linea " & i, , "ATENCION"
            End If
            
        Else
            Exit Do
        End If
        RsOTd.MoveNext
    Loop
End If

'RsPd.Index = "NV-Plano-Item"

Razon.Text = Contratista_Lee(SqlRsSc, Rut.Text)
Detalle.Row = 1 ' para q' actualice la primera fila del detalle

Actualiza

End Sub
Private Sub Contratista_Lee_OLD(ByVal Rut)
Rut = Trim(Rut)
SqlRsSc.Open "SELECT * FROM personas WHERE contratista='S' AND rut='" & Rut & "'", CnxSqlServer_scp0
If SqlRsSc.EOF Then
    Razon.Text = "NO Encontrado"
Else
    Razon.Text = SqlRsSc![razon_social]
End If
SqlRsSc.Close
End Sub
Private Function Doc_Validar() As Boolean
Dim porAsignar As Integer
Doc_Validar = False
If Rut.Text = "" Then
    MsgBox "DEBE ELEGIR CONTRATISTA"
    btnSearch.SetFocus
    Exit Function
End If

For i = 1 To n_filas

    ' plano
    If Trim(Detalle.TextMatrix(i, 1)) <> "" Then
    
        ' Revision          2
    
        ' marca
        If Not CampoReq_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
        
        ' cantidad total    4
        ' cantidad asignada 5
        
        ' cantidad a asignar
        If Not Numero_Valida(Detalle.TextMatrix(i, 7), i, 7) Then Exit Function
        ' [can total]-[can asignada]>=[can a asignar]
        porAsignar = Val(Detalle.TextMatrix(i, 5)) - Val(Detalle.TextMatrix(i, 6))
        If porAsignar < Detalle.TextMatrix(i, 7) Then
            MsgBox "Sólo quedan " & porAsignar & " por asignar", , "ATENCIÓN"
            Detalle.Row = i
            Detalle.col = 7
            Detalle.SetFocus
            Exit Function
        End If
        
        ' articulo
'        If Not LargoString_Valida(Detalle.Textmatrix((i, 6)), 30, i, 6) Then Exit Function
        
        ' peso unitario  8
        ' peso total     9
        
        ' fecha piezas 10
        If Not Fecha_Req(Detalle.TextMatrix(i, 10), i, 10) Then Exit Function
'        EditFecha.Text = Detalle.Textmatrix((i, 10))
        Detalle.Row = i
        Detalle.col = 10
        If Fecha_Valida(Detalle) = False Then Exit Function
        
        ' fecha entrega 11
        If Not Fecha_Req(Detalle.TextMatrix(i, 11), i, 11) Then Exit Function
        Detalle.Row = i
        Detalle.col = 11
        If Fecha_Valida(Detalle) = False Then Exit Function
        
        
        ' peso unitario 12
        If Not Numero_Valida(Detalle.TextMatrix(i, 12), i, 12) Then Exit Function
    
        ' peso total    13
        
    End If
    
Next

Doc_Validar = True

End Function
Private Function CampoReq_Valida(txt As String, fil As Integer, col As Integer) As Boolean
' valida si campo requerido
If Len(Trim(txt)) = 0 Then
    CampoReq_Valida = False
    Beep
    MsgBox "CAMPO OBLIGATORIO"
    Detalle.Row = fil
    Detalle.col = col
    Detalle.SetFocus
Else
    CampoReq_Valida = True
End If
End Function
Private Function LargoString_Valida(txt As String, max As Integer, fil As Integer, col As Integer) As Boolean
If Len(Trim(txt)) > max Then
    LargoString_Valida = False
    Beep
    MsgBox "Largo Máximo es " & max & " caracteres"
    Detalle.Row = fil
    Detalle.col = col
    Detalle.SetFocus
Else
    LargoString_Valida = True
End If
End Function
Private Function Numero_Valida(txt As String, fil As Integer, col As Integer) As Boolean
Dim num As String
Numero_Valida = False
num = txt
If Not IsNumeric(num) Then
    Beep
    MsgBox "Número no Válido"
    Detalle.Row = fil
    Detalle.col = col
    Detalle.SetFocus
    Exit Function
Else
    If Val(num) < 0 Then ' solo mayores que cero
        Beep
        MsgBox "Número no Válido"
        Detalle.Row = fil
        Detalle.col = col
        Detalle.SetFocus
        Exit Function
    End If
End If
Numero_Valida = True
End Function
Private Function Fecha_Req(Fecha As String, fil As Integer, col As Integer) As Boolean
Fecha_Req = False
Fecha = Replace(Fecha, "_")
If Fecha = "//" Or Fecha = "" Then
    Beep
    MsgBox "DEBE digitar Fecha"
    Detalle.Row = fil
    Detalle.col = col
    Detalle.SetFocus
    Exit Function
End If
Fecha_Req = True
End Function
Private Sub Doc_Grabar(Nueva As Boolean)

Dim m_Plano As String, m_Rev As String, m_Marca As String, m_cantidad As Integer
Dim planoeditable As Boolean
planoeditable = True

Dim RsPaso As Recordset

Dim precioUnitario As Double, PrecioTotal As Double
PrecioTotal = 0
Dim aITOF(9) As Integer, i As Integer
i = 0

Rut.Text = SqlRutPadL(Rut.Text)

save:

' DETALLE DE OT

Doc_Detalle_Eliminar

sql = "SELECT * FROM " & Tabla_PlanosDetalle
sql = sql & " WHERE nv=" & m_Nv
'Rs_Abrir_MDB Dbm, RsPc, sql

'RsPd.Index = "NV-Plano-Marca"
j = 0

For i = 1 To n_filas

    m_Plano = Trim(Detalle.TextMatrix(i, 1))
    
    If m_Plano <> "" Then
    
        j = j + 1
'        If j = 14 Then
'        MsgBox ""
'        End If
        m_Marca = Detalle.TextMatrix(i, 3)
        m_cantidad = Val(Detalle.TextMatrix(i, 7))
        RsOTd.AddNew
        RsOTd!Numero = Numero.Text
        RsOTd!linea = j
        
        RsOTd!Fecha = Fecha.Text
        RsOTd!Nv = m_Nv
        RsOTd![Rut contratista] = Rut.Text

        RsOTd!Plano = m_Plano
        m_Rev = Detalle.TextMatrix(i, 2)
'        RsPc.Seek "=", m_Nv, m_NvArea, m_Plano
        
        sql = "SELECT * FROM " & Tabla_PlanosCabecera
        sql = sql & " WHERE nv=" & m_Nv & " AND plano='" & m_Plano & "'"
        Rs_Abrir_MDB Dbm, RsPc, sql
        
        If Not RsPc.NoMatch Then
            m_Rev = RsPc![Rev]
        End If
        
        RsOTd!Rev = m_Rev
        
        RsOTd!Marca = m_Marca
        RsOTd!Cantidad = m_cantidad
        RsOTd![fecha2] = Detalle.TextMatrix(i, 10)
        RsOTd![Fecha Entrega] = Detalle.TextMatrix(i, 11)
        RsOTd![Peso Unitario] = m_CDbl(Detalle.TextMatrix(i, 8))
        precioUnitario = m_CDbl(Detalle.TextMatrix(i, 12))
        RsOTd![Precio Unitario] = precioUnitario
        RsOTd![Cantidad Recibida] = 0
        RsOTd.Update
        
If False Then
        RsPd.Seek "=", m_Nv, m_NvArea, m_Plano, m_Marca
        If RsPd.NoMatch Then
            ' no existe marca en el plano
        Else
            ' actualiza archivo detalle planos
'            RsPd.Edit
'            RsPd![OT fab] = RsPd![OT fab] + m_cantidad
'            RsPd.Update
            ' actualiza plano a no editable
            If planoeditable Then
                RsPc.Index = "NV-Plano"
                RsPc.Seek "=", CDbl(m_Nv), m_NvArea, m_Plano
                If Not RsPc.NoMatch Then
                    RsPc.Edit
                    RsPc![Editable] = False
                    RsPc.Update
                    planoeditable = False
                End If
            End If
            
        End If
End If

     ' actualiza precio unitario en ITO Fab detalle
     sql = "UPDATE [ito fab detalle] SET [precio unitario]=" & precioUnitario
     sql = sql & " WHERE [numero ot]=" & Numero.Text
     sql = sql & " AND [rut contratista]='" & Rut.Text & "'"
     sql = sql & " AND [nv]=" & m_Nv
     sql = sql & " AND [plano]='" & m_Plano & "'"
     sql = sql & " AND [marca]='" & m_Marca & "'"
     
     Debug.Print "OTF doc_Grabar|" & sql & "|"
     
     Dbm.Execute sql
          
     PrecioTotal = PrecioTotal + m_cantidad * m_CDbl(Detalle.TextMatrix(i, 8)) * precioUnitario
        
    End If
Next

' CABECERA DE OT
If Nueva Then
    RsOTc.AddNew
    RsOTc!Numero = Numero.Text
Else
    RsOTc.Edit
End If

RsOTc!Fecha = Fecha.Text
RsOTc!Nv = m_Nv
RsOTc![Rut contratista] = Rut.Text
RsOTc![Peso Total] = TotalKilos.Caption
RsOTc![Precio Total] = PrecioTotal
'If TotalPrecio.Caption = "" Then
'    RsOTc![Precio Total] = 0
'Else
'    RsOTc![Precio Total] = Replace(TotalPrecio.Caption, ".")
'End If
RsOTc![Observacion 1] = Obs(0).Text
RsOTc![Observacion 2] = Obs(1).Text
RsOTc![Observacion 3] = Obs(2).Text
RsOTc![Observacion 4] = Obs(3).Text
RsOTc.Update

'////////////////////////////////////////////////////////////////////////////////
If False Then
    ' actualiza precio total en ito fab cabecera
    sql = "SELECT numero, sum([Precio Unitario]*cantidad*[peso unitario]) AS precio"
    sql = sql & " FROM [ITO Fab Detalle]"
    sql = sql & " WHERE nv=" & Nv.Text
    sql = sql & " AND [rut contratista]='" & Rut.Text & "'"
    sql = sql & " GROUP BY numero"
    '////////////////////////////////////////////////////////////////////////////////
    
    Set RsPaso = Dbm.OpenRecordset(sql)
    With RsPaso
    Do While Not .EOF
        PrecioTotal = Int(![precio] + 0.5)
        sql = "UPDATE [ito fab cabecera] SET [precio total]=" & PrecioTotal
        sql = sql & " WHERE numero=" & ![Numero]
        Debug.Print "OTF ito fab cabecera actualiza|" & sql & "|"
        Dbm.Execute sql
        .MoveNext
    Loop
    End With
End If
'///////////////////////////////////////////////////////////////////////////////

'RsPd.Index = "NV-Plano-Item"

PlanoDetalle_Actualizar Dbm, Nv.Text, m_NvArea, RsOTd, "nv-plano-marca", "ot fab"

' Recalcula Ot
OT_Detalle_Recalcular Numero.Text
'OT_Detalle_Recalcular RsOTd, RsITOfd, m_NV, 0, Numero.Text

Select Case Accion
Case "Agregando"
    Track_Registrar "OTf", Numero.Text, "AGR"
Case "Modificando"
    Track_Registrar "OTf", Numero.Text, "MOD"
End Select

End Sub
Private Sub OT_Detalle_Recalcular(N_Ot As Double)
' recalcula piezas recibidas
Dim m_recibidas As Integer
RsOTd.Seek "=", N_Ot, 1
If Not RsOTd.NoMatch Then
    Do While Not RsOTd.EOF
        If RsOTd!Numero <> N_Ot Then Exit Do
        ' busca itos
        m_recibidas = 0
        RsITOfd.Seek "=", m_Nv, m_NvArea, RsOTd!Plano, RsOTd!Marca
        If Not RsITOfd.NoMatch Then
            Do While Not RsITOfd.EOF
                If RsITOfd!Nv <> m_Nv Or RsITOfd!Plano <> RsOTd!Plano Or RsITOfd!Marca <> RsOTd!Marca Then Exit Do
                    If RsITOfd![Numero OT] = N_Ot Then
                        m_recibidas = m_recibidas + RsITOfd!Cantidad
                    End If
                RsITOfd.MoveNext
            Loop
        End If
        
        RsOTd.Edit
        RsOTd![Cantidad Recibida] = m_recibidas
        RsOTd.Update
        '
        RsOTd.MoveNext
        
    Loop
End If
End Sub
Private Sub Doc_Eliminar()

' borra CABECERA DE OT
RsOTc.Seek "=", Numero.Text
If Not RsOTc.NoMatch Then

    RsOTc.Delete

End If

Doc_Detalle_Eliminar

Track_Registrar "OTf", Numero.Text, "ELI"

End Sub
Private Sub Doc_Detalle_Eliminar()
'Dim m_Plano As String, m_marca As String, m_cantidad As Integer
' elimina detalle OT

' al anular detalle OT debe actualizar detalle plano
'If False Then
'    RsPd.Index = "NV-Plano-Marca"
    RsOTd.Seek "=", Numero.Text, 1
    If Not RsOTd.NoMatch Then
        Do While Not RsOTd.EOF
        
            If RsOTd!Numero <> Numero.Text Then Exit Do
    '        RsPd.Seek "=", m_Nv, m_NvArea, RsOTd!Plano, RsOTd!Marca
    '        If Not RsPd.NoMatch Then
    '            RsPd.Edit
    '            RsPd![OT fab] = RsPd![OT fab] - RsOTd!Cantidad
    '            RsPd.Update
    '        End If
        
            ' borra detalle
            RsOTd.Delete
        
            RsOTd.MoveNext
        Loop
    End If
'    RsPd.Index = "NV-Plano-Item"
'End If

'Registro_Eliminar_MDB Dbm, Tabla_PlanosDetalle, ""

End Sub
Private Sub OT_ItoBuscar()
' busca en ITOs Fabricacion para ver si se puede modificar o borrar OT
n_marcas = 0
RsOTd.Seek "=", Numero.Text, 1
If Not RsOTd.NoMatch Then
    Do While Not RsOTd.EOF
        If RsOTd!Número <> Numero.Text Then Exit Do
        
        With RsITOfd
        .Seek "=", m_Nv, RsOTd!Plano, RsOTd!Marca
        If Not .NoMatch Then
            Do While Not .EOF
                If !Nv <> m_Nv Or !Plano <> RsOTd!Plano Or !Marca <> RsOTd!Marca Then Exit Do
                'encontrando itos
                n_marcas = n_marcas + 1
                Plano_Dig.Marcas_Agregar n_marcas, RsOTd!Marca, 0, 0
                .MoveNext
            Loop
        End If
        End With
        
        RsOTd.MoveNext
    Loop
End If

End Sub
Private Sub Campos_Limpiar()
Numero.Text = ""
Fecha.Text = Fecha_Vacia
Nv.Text = ""
ComboNV.Text = " "
Rut.Text = ""
Razon.Text = ""
'Direccion.Text = ""
'Comuna.Text = ""
Detalle_Limpiar
Obs(0).Text = ""
Obs(1).Text = ""
Obs(2).Text = ""
Obs(3).Text = ""
TotalKilos.Caption = "0"
TotalPrecio.Caption = "0"
End Sub
Private Sub Detalle_Limpiar()
Dim fi As Integer, co As Integer
For fi = 1 To n_filas
    For co = 1 To n_columnas
        Detalle.TextMatrix(fi, co) = ""
    Next
Next
End Sub
Private Sub Nv_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Nv_LostFocus()
m_Nv = Val(Nv.Text)
If m_Nv = 0 Then Exit Sub
' busca nv en combo
i = 1
Do Until a_Nv(i, 0) = ""
    If Val(a_Nv(i, 0)) = m_Nv Then
        ComboNV.ListIndex = i
        Exit Sub
    End If
    i = i + 1
Loop

MsgBox "NV no existe"
Nv.SetFocus

End Sub
Private Sub Obs_GotFocus(Index As Integer)
ComboPlano.visible = False
ComboMarca.visible = False
End Sub
Private Sub Obs_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If Index = 3 Then
        Obs(0).SetFocus
    Else
        Obs(Index + 1).SetFocus
    End If
End If
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim cambia_titulo As Boolean, n_Copias As Integer

ComboPlano.visible = False
ComboMarca.visible = False

cambia_titulo = True
'Accion = "" rem accion
Select Case Button.Index
Case 1 ' Agregar
    Accion = "Agregando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    
    Numero.Text = Documento_Numero_Nuevo(RsOTc, "Numero")
    
    Numero.Enabled = True
    Numero.SetFocus
Case 2 ' Modificar
    Accion = "Modificando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    Numero.Enabled = True
    Numero.SetFocus
Case 3 ' Eliminar
    Accion = "Eliminando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    Numero.Enabled = True
    Numero.SetFocus
Case 4 ' Imprimir
    Accion = "Imprimiendo"
    If Numero.Text = "" Then
        Botones_Enabled 0, 0, 0, 1, 1, 0
        Campos_Enabled False
        Numero.Enabled = True
        Numero.SetFocus
    Else
    
        n_Copias = 1
        PrinterNCopias.Numero_Copias = n_Copias
        PrinterNCopias.Show 1
        n_Copias = PrinterNCopias.Numero_Copias
        
        If n_Copias > 0 Then
        
            Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Docs")
            
            Doc_Imprimir n_Copias
            
            Piezas_Imprimir n_Copias
            
            Impresora_Predeterminada "default"
            
        End If
        
'        If MsgBox("¿ Imprime Etiquetas ?", vbYesNo) = vbYes Then
''            If MsgBox("Debe configurar Impresora ZEBRA como Prederminada", vbYesNo) = vbYes Then
''                Impresora_Predeterminada "zebra"
'                Etiquetas_Imprimir
''                Impresora_Predeterminada "default"
''            End If
'        End If

        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
        
    End If
Case 6 ' DesHacer
    If Numero.Text = "" Then
    
        Privilegios
        
        Campos_Limpiar
        Campos_Enabled False
        
    Else
        If Accion = "Imprimiendo" Then
            Privilegios
            Campos_Limpiar
            Campos_Enabled False
        Else
            If MsgBox("¿Seguro que quiere DesHacer?", vbYesNo) = vbYes Then
                Privilegios
                Campos_Limpiar
                Campos_Enabled False
            End If
        End If
    End If
    Accion = ""
Case 7 ' grabar

    If Doc_Validar Then
    
        If MsgBox("¿ GRABA " & Obj & " ?", vbYesNo) = vbYes Then
        
            If Accion = "Agregando" Then
                Doc_Grabar True
            Else
                Doc_Grabar False
            End If
            
            n_Copias = 1
            PrinterNCopias.Numero_Copias = n_Copias
            PrinterNCopias.Show 1
            n_Copias = PrinterNCopias.Numero_Copias
            
            If n_Copias > 0 Then
            
                Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Docs")
                Doc_Imprimir n_Copias
                Piezas_Imprimir n_Copias
                Impresora_Predeterminada "default"

            End If
            
'            If MsgBox("¿ Imprime Etiquetas ?", vbYesNo) = vbYes Then
''                If MsgBox("Debe configurar Impresora ZEBRA como Prederminada", vbYesNo) = vbYes Then
''                    Impresora_Predeterminada "zebra"
'                    Etiquetas_Imprimir
''                    Impresora_Predeterminada "default"
''                End If
'            End If
            
            Botones_Enabled 0, 0, 0, 0, 1, 0
            Campos_Limpiar
            Campos_Enabled False
            Numero.Enabled = True
            Numero.SetFocus
            
        End If
    End If
Case 8 ' Separador
Case 9 ' Contratistas
    MousePointer = 11
    Load sql_contratistas
    MousePointer = 0
    sql_contratistas.Show 1
    cambia_titulo = False
End Select

If cambia_titulo Then
    If Accion = "" Then
        Me.Caption = "Mantención de " & StrConv(Objs, vbProperCase)
    Else
        Me.Caption = "Mantención de " & StrConv(Objs, vbProperCase) & " [" & Accion & "]"
    End If
End If

End Sub
Private Sub Botones_Enabled(btn_Agregar As Boolean, btn_Modificar As Boolean, _
                            btn_Eliminar As Boolean, btn_Imprimir As Boolean, _
                            btn_DesHacer As Boolean, btn_Grabar As Boolean)
                            
btnAgregar.Enabled = btn_Agregar
btnModificar.Enabled = btn_Modificar
btnEliminar.Enabled = btn_Eliminar
btnImprimir.Enabled = btn_Imprimir
btnDesHacer.Enabled = btn_DesHacer
btnGrabar.Enabled = btn_Grabar

btnAgregar.Value = tbrUnpressed
btnModificar.Value = tbrUnpressed
btnEliminar.Value = tbrUnpressed
btnImprimir.Value = tbrUnpressed
btnDesHacer.Value = tbrUnpressed
btnGrabar.Value = tbrUnpressed

End Sub
Private Sub Campos_Enabled(Si As Boolean)
Numero.Enabled = Si
Fecha.Enabled = Si
btnSearch.Enabled = Si
Nv.Enabled = Si
ComboNV.Enabled = Si
btnCopiar.Enabled = Si
Detalle.Enabled = Si
Obs(0).Enabled = Si
Obs(1).Enabled = Si
Obs(2).Enabled = Si
Obs(3).Enabled = Si
End Sub
Private Sub btnSearch_Click()

Dim arreglo(1) As String
arreglo(1) = "razon_social"

ComboPlano.visible = False
ComboMarca.visible = False

sql_Search.Muestra "personas", "RUT", arreglo(), Obj, Objs, "contratista='S'"
Rut.Text = sql_Search.Codigo
Razon.Text = sql_Search.Descripcion
    
End Sub
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'
' RUTINAS PARA EL FLEXGRID
'
'
Private Sub Detalle_Click()
If Accion = "Imprimiendo" Then Exit Sub
After_Detalle_Click
End Sub
Private Sub After_Detalle_Click()
Dim fil As Integer
fil = Detalle.Row - 1
Select Case Detalle.col
    Case 1 ' plano
        On Error GoTo error1
        If Detalle <> "" Then ComboPlano.Text = Detalle
error1:
        On Error GoTo 0
        ComboPlano.Top = Detalle.CellTop + Detalle.Top
        ComboPlano.Left = Detalle.CellLeft + Detalle.Left
'        ComboPlano.Width = 2500 ' Int(Detalle.CellWidth * 1.5)
        ComboPlano.Width = Int(Detalle.CellWidth * 1.5)
        ComboPlano.visible = True
        ComboMarca.visible = False
    Case 3 ' marca
        ComboMarca_Poblar Detalle.TextMatrix(Detalle.Row, 1)
        On Error GoTo error3
        If Detalle <> "" Then ComboMarca.Text = Detalle
error3:
        On Error GoTo 0
'        ComboMarca.Text = ""
        ComboMarca.Top = Detalle.CellTop + Detalle.Top
        ComboMarca.Left = Detalle.CellLeft + Detalle.Left
'        ComboMarca.Width = 2800 ' Int(Detalle.CellWidth * 1.5)
        ComboMarca.Width = Int(Detalle.CellWidth * 1.5)
        ComboPlano.visible = False
        ComboMarca.visible = True
        
    Case 10 ' fecha de entrega
    Case Else
        ComboPlano.visible = False
        ComboMarca.visible = False
End Select
End Sub
Private Sub ComboMarca_Poblar(Plano As String)
' llena combo marcas
ComboMarca.Clear

'RsPd.Seek "=", Val(m_Nv), m_NvArea, Plano, 1

sql = "SELECT * FROM " & Tabla_PlanosDetalle
sql = sql & " WHERE nv=" & m_Nv & " AND plano='" & Plano & "'"
sql = sql & " ORDER BY marca"
Rs_Abrir_MDB Dbm, RsPd, sql

Do While Not RsPd.EOF
    If RsPd!Nv = m_Nv And RsPd!Plano = Plano Then
        ComboMarca.AddItem RsPd!Marca
    Else
        Exit Do
    End If
    RsPd.MoveNext
Loop

End Sub
Private Sub Detalle_DblClick()
If Accion = "Imprimiendo" Then Exit Sub
' simula un espacio
If Detalle.col = 10 Then
    MSFlexGridEdit Detalle, EditFecha, 32  'FECHA
Else
    MSFlexGridEdit Detalle, txtEditOT, 32
End If
End Sub
Private Sub Detalle_GotFocus()
If Accion = "Imprimiendo" Then Exit Sub
Select Case True
Case txtEditOT.visible
    Detalle = txtEditOT
    txtEditOT.visible = False
Case EditFecha.visible
    Detalle = EditFecha
    EditFecha.visible = False
End Select
End Sub
Private Sub Detalle_LeaveCell()
If Accion = "Imprimiendo" Then Exit Sub
Select Case True
Case txtEditOT.visible
    Detalle = txtEditOT
    txtEditOT.visible = False
Case EditFecha.visible
    Detalle = EditFecha
    EditFecha.visible = False
End Select
End Sub
Private Sub Detalle_KeyPress(KeyAscii As Integer)
If Accion = "Imprimiendo" Then Exit Sub
If Detalle.col = 10 Or Detalle.col = 11 Then
    MSFlexGridEdit Detalle, EditFecha, KeyAscii 'fecha
Else
    MSFlexGridEdit Detalle, txtEditOT, KeyAscii
End If
End Sub
Private Sub txtEditOT_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCodeP Detalle, txtEditOT, KeyCode, Shift
End Sub
Private Sub txtEditOT_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Private Sub txtEditOT_LostFocus()
'txtEditOT.Visible = False 07/03/98
'EditKeyCodeP Detalle, txtEditOT, vbkeyreturn, 0
' ó
'Detalle.SetFocus
'DoEvents
'Actualiza
End Sub
Private Sub EditFecha_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCodeP Detalle, EditFecha, KeyCode, Shift
End Sub
Sub EditKeyCodeP(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
' rutina que se ejecuta con los keydown de Edt
Dim m_col As Integer
m_col = MSFlexGrid.col

Dim dif As Integer
dif = Val(Detalle.TextMatrix(MSFlexGrid.Row, 5)) - Val(Detalle.TextMatrix(MSFlexGrid.Row, 6))
Select Case KeyCode
Case vbKeyEscape ' Esc
    Edt.visible = False
    MSFlexGrid.SetFocus
Case vbKeyReturn ' Enter
    Select Case m_col
    Case 7 ' Cantidad a Asignar
        If Asignada_Validar(MSFlexGrid.col, dif, Edt) Then
            MSFlexGrid.SetFocus
            DoEvents
            Linea_Actualiza
        End If
    Case 10, 11 ' Fecha
'        If Fecha_Valida(Edt) Then
'            MSFlexGrid.SetFocus
'            DoEvents
'            Linea_Actualiza
'        End If
        Detalle.SetFocus
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
        Linea_Actualiza
    End Select
    Cursor_Mueve MSFlexGrid
Case vbKeyUp ' Flecha Arriba
    Select Case m_col
    Case 7 ' Cantidad a Asignar
        If Asignada_Validar(MSFlexGrid.col, dif, Edt) Then
            MSFlexGrid.SetFocus
            DoEvents
            Linea_Actualiza
            If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
                MSFlexGrid.Row = MSFlexGrid.Row - 1
            End If
        End If
    Case 10, 11 ' Fecha
        If Fecha_Valida(Edt) Then
            MSFlexGrid.SetFocus
            DoEvents
            Linea_Actualiza
            If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
                MSFlexGrid.Row = MSFlexGrid.Row - 1
            End If
        End If
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
        Linea_Actualiza
        If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
            MSFlexGrid.Row = MSFlexGrid.Row - 1
        End If
    End Select
Case vbKeyDown ' Flecha Abajo
    Select Case m_col
    Case 7 ' Cantidad a Asignar
        If Asignada_Validar(MSFlexGrid.col, dif, Edt) Then
            MSFlexGrid.SetFocus
            DoEvents
            Linea_Actualiza
            If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
                MSFlexGrid.Row = MSFlexGrid.Row + 1
            End If
        End If
    Case 10, 11 ' Fecha
        If Fecha_Valida(Edt) Then
            MSFlexGrid.SetFocus
            DoEvents
            Linea_Actualiza
            If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
                MSFlexGrid.Row = MSFlexGrid.Row + 1
            End If
        End If
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
        Linea_Actualiza
        If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
            MSFlexGrid.Row = MSFlexGrid.Row + 1
        End If
    End Select
End Select
End Sub
Private Function Asignada_Validar(Colu As Integer, porAsignar As Integer, Edt As Control) As Boolean
' verifica que Ctotal-CAsignada >= CAAsignar
Asignada_Validar = True
If Colu <> 7 Then Exit Function
If porAsignar < Val(Edt) Then
    MsgBox "Sólo quedan " & porAsignar & " por asignar", , "ATENCIÓN"
    Asignada_Validar = False
End If
End Function
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)
Select Case MSFlexGrid.col
Case 1, 3
'    After_Detalle_Click
Case 2, 4, 5, 6, 8, 9, 13
    ' no editables
    Exit Sub
Case 10, 11 'fecha
    Select Case KeyAscii
    Case 0 To 32
        If MSFlexGrid = "" Then
            Edt = "__/__/__"
            Edt.SelStart = 0
        Else
            Edt = Format(MSFlexGrid, "dd/mm/yy")
            Edt.SelStart = 1000
        End If
    Case 48 To 51 ' "0" al "3"
        Edt = Chr(KeyAscii) & "_/__/__"
        Edt.SelStart = 1
    End Select
    Edt.Move MSFlexGrid.CellLeft + MSFlexGrid.Left, MSFlexGrid.CellTop + MSFlexGrid.Top, MSFlexGrid.CellWidth, MSFlexGrid.CellHeight + 50
    Edt.visible = True
    Edt.SetFocus
Case Else
    Select Case KeyAscii
    Case 0 To 32
        Edt = MSFlexGrid
        Edt.SelStart = 1000
    Case Else
        Edt = Chr(KeyAscii)
        Edt.SelStart = 1
    End Select
    Edt.Move MSFlexGrid.CellLeft + MSFlexGrid.Left, MSFlexGrid.CellTop + MSFlexGrid.Top, MSFlexGrid.CellWidth, MSFlexGrid.CellHeight
    Edt.visible = True
    Edt.SetFocus
    'opGrabar True
End Select
End Sub
Private Sub Detalle_KeyDown(KeyCode As Integer, Shift As Integer)
If Accion = "Imprimiendo" Then Exit Sub
If KeyCode = vbKeyF2 Then
    MSFlexGridEdit Detalle, txtEditOT, 32
End If
End Sub
Private Sub Linea_Actualiza()
' actualiza solo linea, y totales generales
Dim fi As Integer, co As Integer

fi = Detalle.Row
co = Detalle.col

n7 = m_CDbl(Detalle.TextMatrix(fi, 7))
n8 = m_CDbl(Detalle.TextMatrix(fi, 8))
n12 = m_CDbl(Detalle.TextMatrix(fi, 12))

' peso total
Detalle.TextMatrix(fi, 9) = Format(n7 * n8, num_Formato)
' precio total
Detalle.TextMatrix(fi, 13) = Format(n7 * n8 * n12, num_fmtgrl)

Totales_Actualiza

End Sub
Private Sub Actualiza()
' actualiza todo el detalle y totales generales
Dim fi As Integer
For fi = 1 To n_filas
    If Detalle.TextMatrix(fi, 1) <> "" Then
    
        n7 = m_CDbl(Detalle.TextMatrix(fi, 7))
        n8 = m_CDbl(Detalle.TextMatrix(fi, 8))
        n12 = m_CDbl(Detalle.TextMatrix(fi, 12))

        ' peso total
        Detalle.TextMatrix(fi, 9) = Format(n7 * n8, num_Formato)
        ' precio total
        Detalle.TextMatrix(fi, 13) = Format(n7 * n8 * n12, num_fmtgrl)
    End If
Next

Totales_Actualiza

End Sub
Private Sub Totales_Actualiza()
Dim Tot_Kilos As Double, Tot_Precio As Double
Tot_Kilos = 0
Tot_Precio = 0
For i = 1 To n_filas
    Tot_Kilos = Tot_Kilos + m_CDbl(Detalle.TextMatrix(i, 9))
    Tot_Precio = Tot_Precio + m_CDbl(Detalle.TextMatrix(i, 13))
Next

TotalKilos.Caption = Format(Tot_Kilos, num_Formato)
TotalPrecio.Caption = Format(Tot_Precio, num_Format0)

End Sub
Private Sub Cursor_Mueve(MSFlexGrid As Control)
'MIA
Select Case MSFlexGrid.col
Case 7
    MSFlexGrid.col = 10
Case 10
    MSFlexGrid.col = 11
Case 11
    MSFlexGrid.col = 12
Case 12
    If MSFlexGrid.Row + 1 < MSFlexGrid.Rows Then
        MSFlexGrid.col = 1
        MSFlexGrid.Row = MSFlexGrid.Row + 1
    End If
Case Else
    MSFlexGrid.col = MSFlexGrid.col + 1
End Select
End Sub
'
' FIN RUTINAS PARA FLEXGRID
'
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
Private Sub Doc_Imprimir(n_Copias As Integer)
MousePointer = vbHourglass
linea = String(100, "-")
' imprime OT
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer, tab4 As Integer
Dim tab5 As Integer, tab6 As Integer, tab7 As Integer, tab8 As Integer, tab9 As Integer
Dim tab10 As Integer, tab11 As Integer, tab40 As Integer
Dim fc As Integer, fn As Integer, fe As Integer, k As Integer
tab0 = 3 'margen izquierdo
tab1 = tab0 + 0
tab2 = tab1 + 15 ' plano
tab3 = tab2 + 4  ' rev
tab4 = tab3 + 15 ' marca
tab5 = tab4 + 5  ' cant
tab6 = tab5 + 8  ' descrip
tab7 = tab6 + 8  ' kg uni
tab8 = tab7 + 9  ' kg total
tab9 = tab8 + 10 ' fecha piezas
tab10 = tab9 + 10 ' f.entrega
tab11 = tab10 + 6 ' $ uni
tab40 = 40

' font.size
fc = 10 'comprimida
fn = 12 'normal
fe = 15 'expandida

Dim can_valor As String

'Printer_Set "Documentos"
Set prt = Printer

prt.Orientation = vbPRORLandscape ' orientacion horizontal

Font_Setear prt

For k = 1 To n_Copias

prt.Font.Size = fe
prt.Print Tab(tab0 - 1); m_TextoIso
prt.Print Tab(tab0 + 37); "OT FABRICACIÓN Nº";
prt.Font.Bold = True
'prt.Print Tab(tab0 + 18); Format(Numero.Text, "#####");
prt.Print Format(Numero.Text, "#####");
prt.Font.Bold = False
prt.Print Tab(tab0 + 70); Fecha.Text
prt.Font.Size = fc
prt.Print ""

' cabecera
prt.Font.Bold = True
prt.Print Tab(tab0); Empresa.Razon;
prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0 + tab40); "CONTRATISTA :";
prt.Font.Size = fe
prt.Font.Bold = True
prt.Print Tab(tab0 + 40); Left(Razon, 32)

prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0); "GIRO: " & Empresa.Giro;

prt.Font.Size = fc
prt.Print Tab(tab0); Empresa.Direccion;
prt.Print Tab(tab0 + tab40); "OBRA      : ";
prt.Font.Size = fe
prt.Font.Bold = True
prt.Print Tab(tab0 + 40); Left(Format(ComboNV.Text, ">"), 32)
prt.Font.Size = fc

prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0); "Teléfono: "; Empresa.Telefono1; " "; Empresa.Comuna;
prt.Font.Size = fn

'prt.Print ""
prt.Print ""

' detalle
prt.Font.Bold = True
prt.Print Tab(tab1); "PLANO";
prt.Print Tab(tab2); "REV";
prt.Print Tab(tab3); "MARCA";
prt.Print Tab(tab4); "CANT";
prt.Print Tab(tab5); "DESCRIP";
prt.Print Tab(tab6); " KG UNI";
prt.Print Tab(tab7); "  KG TOT";
prt.Print Tab(tab8); "F.PIEZAS";
prt.Print Tab(tab9); " ENTREGA";
prt.Print Tab(tab10); "$ UNI";
prt.Print Tab(tab11); "  $ TOT"
prt.Font.Bold = False

prt.Print Tab(tab1); linea
j = -1
For i = 1 To n_filas

    can_valor = Detalle.TextMatrix(i, 7)
    
    If Val(can_valor) = 0 Then
    
        j = j + 1
        prt.Print Tab(tab1 + j * 3); "  \"
        
    Else
    
        ' PLANO
        prt.Print Tab(tab1); Detalle.TextMatrix(i, 1);
        
        ' REVISION
        prt.Print Tab(tab2); Detalle.TextMatrix(i, 2);
        
        ' MARCA
        prt.Print Tab(tab3); Detalle.TextMatrix(i, 3);
        
        ' CANTIDAD
        prt.Print Tab(tab4); m_Format(can_valor, "####");
        
        ' DESCRIPCION
        prt.Print Tab(tab5); Left(Detalle.TextMatrix(i, 4), 7);
        
        ' KG UNITARIO
        prt.Print Tab(tab6); m_Format(m_CDbl(Detalle.TextMatrix(i, 8)), "#,###.0");
        
        ' KG TOTAL
        prt.Print Tab(tab7); m_Format(Detalle.TextMatrix(i, 9), "##,###.0");
        
        ' FECHA ENTREGA
        prt.Print Tab(tab8); Format(Detalle.TextMatrix(i, 10), Fecha_Format);
        
        ' FECHA ENTREGA
        prt.Print Tab(tab9); Format(Detalle.TextMatrix(i, 11), Fecha_Format);
        
        ' $ UNITARIO
        prt.Print Tab(tab10); m_Format(Detalle.TextMatrix(i, 12), "#,###");
        
        ' $ TOTAL
        prt.Print Tab(tab11); m_Format(Detalle.TextMatrix(i, 13), "###,###")
        
    End If
    
Next

prt.Print Tab(tab1); linea
prt.Font.Bold = True
prt.Print Tab(tab7 - 5); m_Format(TotalKilos.Caption, "###,###,###.0");
prt.Print Tab(tab11 - 5); m_Format(TotalPrecio, "$###,###,###")
prt.Font.Bold = False
'prt.Print ""

prt.Print Tab(tab0); "OBSERVACIONES :";
prt.Print Tab(tab0 + 16); Obs(0).Text
prt.Print Tab(tab0 + 16); Obs(1).Text
prt.Print Tab(tab0 + 16); Obs(2).Text
prt.Print Tab(tab0 + 16); Obs(3).Text

'For i = 1 To 2
    prt.Print ""
'Next

prt.Print Tab(30); "__________________"; Tab(70); "__________________"
prt.Print Tab(30); "       VºBº       "; Tab(70); "       VºBº       "

If k < n_Copias Then
    prt.NewPage
Else
    prt.EndDoc
End If

Next

Impresora_Predeterminada "default"

MousePointer = vbDefault

End Sub
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
Private Sub Piezas_Imprimir_OLD(n_Copias As Integer)
MousePointer = vbHourglass
linea = String(100, "-")
' imprime OT
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer, tab4 As Integer
Dim tab5 As Integer, tab6 As Integer, tab7 As Integer, tab8 As Integer, tab9 As Integer
Dim tab10 As Integer, tab11 As Integer, tab12 As Integer, tab13 As Integer, tab14 As Integer, tab40 As Integer
Dim fc As Integer, fn As Integer, fe As Integer, k As Integer

Dim m_Pla As String, m_Mar As String, can_valor As Integer, can_doble As Double

tab0 = 3 'margen izquierdo
tab1 = tab0 + 0
tab2 = tab1 + 10  ' plano
tab3 = tab2 + 4   ' rev
tab4 = tab3 + 7   ' marca
tab5 = tab4 + 5   ' pos
tab6 = tab5 + 5   ' cant
tab7 = tab6 + 10  ' descrip
tab8 = tab7 + 5   ' anc
tab9 = tab8 + 5   ' esp
tab10 = tab9 + 5  ' lar
tab11 = tab10 + 7 ' puni
tab12 = tab11 + 10 ' ptot
tab13 = tab12 + 7 ' suni
tab14 = tab13 + 8 ' stot
tab40 = 40

' font.size
fc = 10 'comprimida
fn = 12 'normal
fe = 15 'expandida


'Printer_Set "Documentos"
Set prt = Printer

prt.Orientation = vbPRORLandscape ' orientacion horizontal

Font_Setear prt

For k = 1 To n_Copias

prt.Font.Size = fe
prt.Print Tab(tab0 - 1); m_TextoIso
prt.Print Tab(tab0 + 37); "PIEZAS OT FABRICACIÓN Nº";
prt.Font.Bold = True
'prt.Print Tab(tab0 + 18); Format(Numero.Text, "#####");
prt.Print Format(Numero.Text, "#####");
prt.Font.Bold = False
prt.Print Tab(tab0 + 70); Fecha.Text
prt.Font.Size = fc
prt.Print ""

' cabecera
prt.Font.Bold = True
prt.Print Tab(tab0); Empresa.Razon;
prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0 + tab40); "CONTRATISTA :";
prt.Font.Size = fe
prt.Font.Bold = True
prt.Print Tab(tab0 + 40); Left(Razon, 32)

prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0); "GIRO: " & Empresa.Giro;

prt.Font.Size = fc
prt.Print Tab(tab0); Empresa.Direccion;
prt.Print Tab(tab0 + tab40); "OBRA      : ";
prt.Font.Size = fe
prt.Font.Bold = True
prt.Print Tab(tab0 + 40); Left(Format(ComboNV.Text, ">"), 32)
prt.Font.Size = fc

prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0); "Teléfono: "; Empresa.Telefono1; " "; Empresa.Comuna;
prt.Font.Size = fn

'prt.Print ""
prt.Print ""

' detalle
prt.Font.Bold = True
prt.Print Tab(tab1); "PLANO";
prt.Print Tab(tab2); "REV";
prt.Print Tab(tab3); "MARCA";
prt.Print Tab(tab4); "POS";
prt.Print Tab(tab5); "CANT";
prt.Print Tab(tab6); "DESCRIP";

prt.Print Tab(tab7); " ANC";
prt.Print Tab(tab8); " ESP";
prt.Print Tab(tab9); " LAR";

prt.Print Tab(tab10); "   PUNI";
prt.Print Tab(tab11); "    PTOT";
prt.Print Tab(tab12); " SUNI";
prt.Print Tab(tab13); "  STOT";
prt.Print Tab(tab14); "OBSERV"

prt.Font.Bold = False

prt.Print Tab(tab1); linea
j = -1
For i = 1 To n_filas

    can_valor = m_CDbl(Detalle.TextMatrix(i, 7))
    
    If Val(can_valor) = 0 Then
    
'        j = j + 1
'        prt.Print Tab(tab1 + j * 3); "  \"
        
    Else
        
        If i > 1 Then
            prt.Print Tab(tab1); "-"
        End If
        
        ' PLANO
        m_Pla = Detalle.TextMatrix(i, 1)
        prt.Print Tab(tab1); m_Pla;
        
        ' REVISION
        prt.Print Tab(tab2); Detalle.TextMatrix(i, 2);
        
        ' MARCA
        m_Mar = Detalle.TextMatrix(i, 3)
        prt.Print Tab(tab3); m_Mar;
                        
        ' CANTIDAD
        prt.Print Tab(tab5); m_Format(can_valor, "####");
        
        ' DESCRIPCION
        prt.Print Tab(tab6); Left(Detalle.TextMatrix(i, 4), 7);
        
        ' KG UNITARIO
        prt.Print Tab(tab10); m_Format(m_CDbl(Detalle.TextMatrix(i, 8)), "#,###.0");
        
        ' KG TOTAL
        prt.Print Tab(tab11); m_Format(Detalle.TextMatrix(i, 9), "##,###.0");
        
        ' FECHA ENTREGA
'        prt.Print Tab(tab12); Format(Detalle.TextMatrix(i, 10), Fecha_Format);
                
        ' SUNI
        prt.Print Tab(tab12); m_Format(Detalle.TextMatrix(i, 14), "#,###");
        
        ' STOT
'        prt.Print Tab(tab13); m_Format(Detalle.TextMatrix(i, 13), "###,###")
        
        ' busca e imprime piezas
        sql = "SELECT * FROM piezas"
        sql = sql & " WHERE nv=" & Nv.Text
        sql = sql & " AND plano='" & m_Pla & "'"
        sql = sql & " AND marca='" & m_Mar & "'"
        sql = sql & " ORDER BY pieza"
        
        Rs_Abrir SqlRsSc, sql
        
        Do While Not SqlRsSc.EOF
        
            prt.Print Tab(tab4); SqlRsSc![pieza]; ' pos

            ' CANTIDAD
            can_valor = SqlRsSc![cantidad_total]
            prt.Print Tab(tab5); m_Format(can_valor, "####");

            ' DESCRIPCION
            m_Mar = SqlRsSc![Descripcion]
            prt.Print Tab(tab6); Left(m_Mar, 8);

            ' ANC
            can_doble = SqlRsSc![ancho]
            prt.Print Tab(tab7); m_Format(can_doble, "####");
            ' ESPESOR
            can_doble = SqlRsSc![espesor]
            prt.Print Tab(tab8); m_Format(can_doble, "####");
            ' LARGO
            can_doble = SqlRsSc![largo]
            prt.Print Tab(tab9); m_Format(can_doble, "####");

            ' KG UNITARIO
            can_doble = SqlRsSc![Peso]
            prt.Print Tab(tab10); m_Format(can_doble, "#,###.0");

            ' KG TOT
            can_doble = can_valor * SqlRsSc![Peso]
            prt.Print Tab(tab11); m_Format(can_doble, "##,###.0");

            ' SUNI
            can_doble = SqlRsSc![Superficie]
            prt.Print Tab(tab12); m_Format(can_doble, "#.000");

            ' STOT
            can_doble = can_valor * can_doble
            prt.Print Tab(tab13); m_Format(can_doble, "##.000");

            ' OBSERVACION
            m_Mar = SqlRsSc![Observacion]
            prt.Print Tab(tab14); Left(m_Mar, 10)

            SqlRsSc.MoveNext

        Loop

    End If

Next

prt.Print Tab(tab1); linea
prt.Font.Bold = True
prt.Print Tab(tab7 - 5); m_Format(TotalKilos.Caption, "###,###,###.0");
prt.Print Tab(tab11 - 5); m_Format(TotalPrecio, "$###,###,###")
prt.Font.Bold = False
'prt.Print ""

prt.Print Tab(tab0); "OBSERVACIONES :";
prt.Print Tab(tab0 + 16); Obs(0).Text
prt.Print Tab(tab0 + 16); Obs(1).Text
prt.Print Tab(tab0 + 16); Obs(2).Text
prt.Print Tab(tab0 + 16); Obs(3).Text

'For i = 1 To 2
    prt.Print ""
'Next

prt.Print Tab(30); "__________________"; Tab(70); "__________________"
prt.Print Tab(30); "       VºBº       "; Tab(70); "       VºBº       "

If k < n_Copias Then
    prt.NewPage
Else
    prt.EndDoc
End If

Next

Impresora_Predeterminada "default"

MousePointer = vbDefault

End Sub
Private Sub Piezas_Imprimir(n_Copias As Integer)
MousePointer = vbHourglass
linea = String(100, "-")
' imprime OT
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer, tab4 As Integer
Dim tab5 As Integer, tab6 As Integer, tab7 As Integer, tab8 As Integer, tab9 As Integer
Dim tab10 As Integer, tab11 As Integer ', tab12 As Integer, tab13 As Integer, tab14 As Integer, tab40 As Integer
Dim tab40 As Integer
Dim PesoTotal As Double

Dim fc As Integer, fn As Integer, fe As Integer, k As Integer

Dim m_Pla As String, m_Mar As String, can_valor As Integer, can_doble As Double

Dim a_Piezas(25, 11) As String, numeroPiezas As Integer
Dim Posicion As Integer ' en arreglo de piezas (la fila)
Dim piezaEncontrada As Boolean

tab0 = 3 'margen izquierdo
tab1 = tab0 + 5
tab2 = tab1 + 5   ' ancho del pos
tab3 = tab2 + 5   ' cant
tab4 = tab3 + 20  ' descripcion
tab5 = tab4 + 5   ' anc
tab6 = tab5 + 5   ' esp
tab7 = tab6 + 5  ' lar
tab8 = tab7 + 8 ' puni
tab9 = tab8 + 10 ' ptot
tab10 = tab9 + 8 ' suni
tab11 = tab10 + 9 ' stot
tab40 = 40

' font.size
fc = 10 'comprimida
fn = 12 'normal
fe = 15 'expandida

'Printer_Set "Documentos"
Set prt = Printer

prt.Orientation = vbPRORLandscape ' orientacion horizontal

Font_Setear prt

For k = 1 To n_Copias

prt.Font.Size = fe
prt.Print Tab(tab0 - 1); m_TextoIso
prt.Print Tab(tab0 + 37); "PIEZAS OT FABRICACIÓN Nº";
prt.Font.Bold = True
'prt.Print Tab(tab0 + 18); Format(Numero.Text, "#####");
prt.Print Format(Numero.Text, "#####");
prt.Font.Bold = False
prt.Print Tab(tab0 + 70); Fecha.Text
prt.Font.Size = fc
prt.Print ""

' cabecera
prt.Font.Bold = True
prt.Print Tab(tab0); Empresa.Razon;
prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0 + tab40); "CONTRATISTA :";
prt.Font.Size = fe
prt.Font.Bold = True
prt.Print Tab(tab0 + 40); Left(Razon, 32)

prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0); "GIRO: " & Empresa.Giro;

prt.Font.Size = fc
prt.Print Tab(tab0); Empresa.Direccion;
prt.Print Tab(tab0 + tab40); "OBRA      : ";
prt.Font.Size = fe
prt.Font.Bold = True
prt.Print Tab(tab0 + 40); Left(Format(ComboNV.Text, ">"), 32)
prt.Font.Size = fc

prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0); "Teléfono: "; Empresa.Telefono1; " "; Empresa.Comuna;
prt.Font.Size = fn

prt.Print " "
prt.Print " "

' detalle
prt.Font.Bold = True
prt.Print Tab(tab1); "POS";
prt.Print Tab(tab2); "CANT";
prt.Print Tab(tab3); "DESCRIPCION";

prt.Print Tab(tab4); " ANC";
prt.Print Tab(tab5); " ESP";
prt.Print Tab(tab6); " LAR";

prt.Print Tab(tab7); "   PUNI";
prt.Print Tab(tab8); "    PTOT";
prt.Print Tab(tab9); " SUNI";
prt.Print Tab(tab10); "  STOT";
prt.Print Tab(tab11); "OBSERV"

prt.Font.Bold = False

prt.Print Tab(tab1); linea
j = -1

numeroPiezas = 0

For i = 1 To n_filas

    ' CANTIDAD
    can_valor = m_CDbl(Detalle.TextMatrix(i, 7))
    
    If Val(can_valor) = 0 Then
    
'        j = j + 1
'        prt.Print Tab(tab1 + j * 3); "  \"
        
    Else
                
        ' PLANO
        m_Pla = Detalle.TextMatrix(i, 1)
          
        ' MARCA
        m_Mar = Detalle.TextMatrix(i, 3)
                        
        ' busca e imprime piezas
        sql = "SELECT * FROM piezas"
        sql = sql & " WHERE nv=" & Nv.Text
        sql = sql & " AND plano='" & m_Pla & "'"
        sql = sql & " AND marca='" & m_Mar & "'"
        sql = sql & " ORDER BY pieza"
        
        Rs_Abrir SqlRsSc, sql
        
        Do While Not SqlRsSc.EOF
        
            ' busca en arreglo
            piezaEncontrada = False
            If numeroPiezas = 0 Then
                '
            Else
                For j = 1 To numeroPiezas
                    If a_Piezas(j, 1) = SqlRsSc![pieza] Then
                        Posicion = j
                        piezaEncontrada = True
                        Exit For
                    End If
                Next
            End If
            
            If piezaEncontrada Then
            
                a_Piezas(Posicion, 2) = a_Piezas(Posicion, 2) + can_valor * SqlRsSc![cantidad_total]
                
            Else
            
                numeroPiezas = numeroPiezas + 1
                Posicion = numeroPiezas
                
                a_Piezas(Posicion, 1) = SqlRsSc![pieza]
                a_Piezas(Posicion, 2) = can_valor * SqlRsSc![cantidad_total]
                a_Piezas(Posicion, 3) = SqlRsSc![Descripcion]
                a_Piezas(Posicion, 4) = SqlRsSc![ancho]
                a_Piezas(Posicion, 5) = SqlRsSc![espesor]
                a_Piezas(Posicion, 6) = SqlRsSc![largo]
                a_Piezas(Posicion, 7) = SqlRsSc![Peso]
                a_Piezas(Posicion, 9) = SqlRsSc![Superficie]
                a_Piezas(Posicion, 11) = SqlRsSc![Observacion]
                
            End If
            
            SqlRsSc.MoveNext

        Loop
        
        SqlRsSc.Close
        
    End If

Next

' imprime resumen piezas
PesoTotal = 0
For j = 1 To numeroPiezas

    prt.Print Tab(tab1); a_Piezas(j, 1); ' pos

    ' CANTIDAD
    can_valor = a_Piezas(j, 2)
    prt.Print Tab(tab2); m_Format(can_valor, "####");

    ' DESCRIPCION
    m_Mar = a_Piezas(j, 3)
    prt.Print Tab(tab3); Left(m_Mar, 20);

    ' ANC
    can_doble = a_Piezas(j, 4)
    prt.Print Tab(tab4); m_Format(can_doble, "####");
    ' ESPESOR
    can_doble = a_Piezas(j, 5)
    prt.Print Tab(tab5); m_Format(can_doble, "####");
    ' LARGO
    can_doble = a_Piezas(j, 6)
    prt.Print Tab(tab6); m_Format(can_doble, "####");

    ' KG UNITARIO
    can_doble = a_Piezas(j, 7)
    prt.Print Tab(tab7); m_Format(can_doble, "#,###.0");

    ' KG TOT
    can_doble = can_valor * can_doble
    prt.Print Tab(tab8); m_Format(can_doble, "##,###.0");

    PesoTotal = PesoTotal + can_doble

    ' SUNI
    can_doble = a_Piezas(j, 9)
    prt.Print Tab(tab9); m_Format(can_doble, "#.000");

    ' STOT
    can_doble = can_valor * can_doble
    prt.Print Tab(tab10); m_Format(can_doble, "##.000");

    ' OBSERVACION
    m_Mar = a_Piezas(j, 11)
    prt.Print Tab(tab11); Left(m_Mar, 10)



Next

prt.Print Tab(tab1); linea

prt.Print Tab(tab8); m_Format(PesoTotal, "##,###.0");


'For i = 1 To 2
    prt.Print ""
'Next

prt.Print Tab(30); "__________________"; Tab(70); "__________________"
prt.Print Tab(30); "       VºBº       "; Tab(70); "       VºBº       "

If k < n_Copias Then
    prt.NewPage
Else
    prt.EndDoc
End If

Next

Impresora_Predeterminada "default"

MousePointer = vbDefault

End Sub
Private Sub SetpYX(Y As Double, x As Double)
prt.CurrentY = AjusteY + Y
prt.CurrentX = AjusteX + x
End Sub
Private Sub Privilegios()

If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 0, 0
End If

End Sub
