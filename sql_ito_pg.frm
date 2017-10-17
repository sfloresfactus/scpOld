VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form sql_ito_pg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ITO PG"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ingresar Nueva ITO"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar ITO"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar ITO"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir ITO"
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
            Object.ToolTipText     =   "Grabar ITO"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mantención de Contratistas"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox defecto 
      Height          =   300
      Left            =   8880
      TabIndex        =   27
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton btnInspectorBuscar 
      Height          =   300
      Left            =   5400
      Picture         =   "sql_ito_pg.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   960
      Width           =   300
   End
   Begin VB.CommandButton btnContratistaBuscar 
      Height          =   300
      Left            =   5400
      Picture         =   "sql_ito_pg.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   600
      Width           =   300
   End
   Begin VB.ComboBox CbTipoGranalla 
      Height          =   315
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1920
      Width           =   975
   End
   Begin VB.ComboBox CbTrabajadores 
      Height          =   315
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   4440
      Width           =   1815
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   8280
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.TextBox Nv 
      Height          =   300
      Left            =   840
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox ComboOT 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox ComboNV 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   1
      Left            =   360
      MaxLength       =   50
      TabIndex        =   11
      Top             =   5280
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   0
      Left            =   360
      MaxLength       =   50
      TabIndex        =   10
      Top             =   4920
      Width           =   5000
   End
   Begin VB.ComboBox ComboMarca 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   8040
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox ComboPlano 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtEditGD 
      Height          =   285
      Left            =   8040
      TabIndex        =   12
      Text            =   "txtEditGD"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   3120
      TabIndex        =   4
      Top             =   600
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
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "##########"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid Detalle 
      Height          =   3045
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5371
      _Version        =   327680
      Enabled         =   0   'False
      ScrollBars      =   2
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   10080
      Top             =   2640
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
            Picture         =   "sql_ito_pg.frx":0204
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_ito_pg.frx":0316
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_ito_pg.frx":0428
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_ito_pg.frx":053A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_ito_pg.frx":064C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_ito_pg.frx":075E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_ito_pg.frx":0870
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDefecto 
      Caption         =   "Defecto"
      Height          =   255
      Left            =   8880
      TabIndex        =   28
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblInspectorRazon 
      Height          =   255
      Left            =   5760
      TabIndex        =   26
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label lblContratistaRazon 
      Height          =   255
      Left            =   5760
      TabIndex        =   25
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lblInspector 
      Caption         =   "Inspector"
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Top             =   1005
      Width           =   735
   End
   Begin VB.Label lblContratista 
      Caption         =   "Contratista"
      Height          =   255
      Left            =   4560
      TabIndex        =   22
      Top             =   600
      Width           =   855
   End
   Begin VB.Label TotalKg 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   5760
      TabIndex        =   20
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Observación"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label TotalPrecio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6840
      TabIndex        =   17
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Totalm2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   4440
      TabIndex        =   15
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "&OBRA"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "&FECHA"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "&Nº"
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
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "sql_ito_pg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' las piezas van:
'    en negro: es decir, no pintadas, ni galvanizadas
'    pintadas:
'    galvanizadas:
' estos tres son excluyentes
' y ahora granallado

' tipos de ITO:

' P : pintura
' G : galvanizado YA NO SE USA
' R : granallado, Erwin
' S : granallado especial
' T : produccion pintura , Erwin
' U : produccion pintura especial
' IRP : ito reproceso , va despues de ot, en cualquier proceso 26/08/10

Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button
Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, k As Integer, d As Variant

Private DbD As Database, RsCl As Recordset, RsTra As Recordset
'Private RsSc As Recordset
Private Dbm As Database, RsNVc As Recordset
Private RsNvPla As Recordset, RsPd As Recordset
'Private RsITOpgc As Recordset, RsITOpgd As Recordset

' contratistas, inspectores,
Private SqlRs As New ADODB.Recordset

Private m_TipoDoc As String ', m_Tipo As String

Private n_filas As Integer, n_columnas As Integer
Private Rev(2999) As String
'Private NvTipo(1999) As String ' P o G
Private prt As Printer
Private n1 As Double, n4 As Double, n6 As Double, n7 As Double, n8 As Double, n10 As Double, n12 As Double
Private n14 As Double, n15 As Double
Private a_Nv(2999, 1) As String, m_Nv As Double, m_NvArea As Integer
Private Doc_Num_Nuevo As Double

' variables para impresion de etiq
Private m_obra As String, m_Plano As String, m_Rev As String, m_Marca As String, m_Peso As Double
Private m_ClienteRazon As String, AjusteX As Double, AjusteY As Double
Private a_Trabajadores(1, 199) As String, m_Nombre As String
Private conContratista As Boolean
Private a_TipoGranalla(9) As String, m_TotalTiposGranalla As Integer

Private m_ContratistaRut As String
Private m_InspectorRut As String

Private ac(21, 1) As String, av(21) As String, asn(21) As Boolean

Public Property Let TipoDoc(ByVal New_Tipo As String)
m_TipoDoc = New_Tipo
End Property
Private Sub btnContratistaBuscar_Click()

Dim arreglo(1) As String
arreglo(1) = "razon_social"

ComboPlano.visible = False
ComboMarca.visible = False

sql_Search.Muestra "personas", "RUT", arreglo(), Obj, Objs, "contratista='S'"
m_ContratistaRut = sql_Search.Codigo
lblContratistaRazon.Caption = sql_Search.Descripcion
    
End Sub

Private Sub btnInspectorBuscar_Click()

Dim arreglo(2) As String
arreglo(1) = "paterno"
arreglo(2) = "nombres"

ComboPlano.visible = False
ComboMarca.visible = False

sql_Search.Muestra "trabajadores", "RUT", arreglo, "Inspector", "Inspectores", "tipo2='S'"
m_InspectorRut = sql_Search.Codigo
lblInspectorRazon.Caption = sql_Search.Descripcion

End Sub

Private Sub CbTipoGranalla_Click()
Detalle = CbTipoGranalla.Text
CbTipoGranalla.visible = False
End Sub
Private Sub CbTrabajadores_Click()
Detalle = CbTrabajadores.Text
CbTrabajadores.visible = False
End Sub
Private Sub Form_Load()

' abre archivos
Set DbD = OpenDatabase(data_file)
'Set RsSc = DbD.OpenRecordset("Contratistas")
'RsSc.Index = "RUT"

Set RsCl = DbD.OpenRecordset("Clientes")
RsCl.Index = "RUT"

'Set RsTra = DbD.OpenRecordset("Trabajadores")
'RsTra.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

'Set RsITOpgc = Dbm.OpenRecordset("ITO PG Cabecera")
'RsITOpgc.Index = "Numero"
'RsITOpgc.Index = "Tipo-Numero"

'Set RsITOpgd = Dbm.OpenRecordset("ITO PG Detalle")
'RsITOpgd.Index = "Número-Línea"
'RsITOpgd.Index = "Tipo-Numero-Linea"

' puebla nv
' Combo obra
ComboNV.AddItem " "
i = 0
Do While Not RsNVc.EOF

    If Usuario.Nv_Activas = False Then ' todas
        GoTo IncluirNV
    Else
        If Usuario.Nv_Activas And RsNVc!Activa Then
            GoTo IncluirNV
        End If
    End If
    If False Then
IncluirNV:

    If m_TipoDoc = "P" And RsNVc!pintura Then
    
        i = i + 1
        a_Nv(i, 0) = RsNVc!Numero
        a_Nv(i, 1) = RsNVc!Obra
    
        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra
        
    End If
    If m_TipoDoc = "G" And RsNVc!galvanizado Then
    
        i = i + 1
        a_Nv(i, 0) = RsNVc!Numero
        a_Nv(i, 1) = RsNVc!Obra
    
        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra
        
    End If
    
    If m_TipoDoc = "R" Then  ' granallado
    
        i = i + 1
        a_Nv(i, 0) = RsNVc!Numero
        a_Nv(i, 1) = RsNVc!Obra
    
        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra
        
    End If
    
    If m_TipoDoc = "T" Then  ' produccion pintura
    
        i = i + 1
        a_Nv(i, 0) = RsNVc!Numero
        a_Nv(i, 1) = RsNVc!Obra
    
        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra

    End If

    If m_TipoDoc = "IRP" Then  ' ito reproceso

        i = i + 1
        a_Nv(i, 0) = RsNVc!Numero
        a_Nv(i, 1) = RsNVc!Obra

        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra
        
    End If
    
'    If RsNVc!Galvanizado Then
'        NvTipo(i) = "G"
'    End If
    
'    If RsNVc!pintura Then
'        NvTipo(i) = "P"
'    End If

    End If
    
    RsNVc.MoveNext
    
Loop

Set RsNvPla = Dbm.OpenRecordset("Planos Cabecera")
RsNvPla.Index = "NV-Plano"

Set RsPd = Dbm.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Item"

Inicializa

'n_columnas = 15 + 1
n_columnas = 17

If m_TipoDoc = "P" Then ' pintura
    n_columnas = 17
    conContratista = True
    Detalle_Config
End If

If m_TipoDoc = "G" Then ' galvanizado
    conContratista = True
    Detalle_Config
    n_columnas = 17
End If

If m_TipoDoc = "R" Then ' granallado

    n_columnas = 17
    Detalle_Config
    
    a_TipoGranalla(0) = "SP2"
    a_TipoGranalla(1) = "SP3"
    a_TipoGranalla(2) = "SP5"
    a_TipoGranalla(3) = "SP6"
    a_TipoGranalla(4) = "SP7"
    a_TipoGranalla(5) = "SP10"
    
    CbTipoGranalla.AddItem "SP2"
    CbTipoGranalla.AddItem "SP3"
    CbTipoGranalla.AddItem "SP5"
    CbTipoGranalla.AddItem "SP6"
    CbTipoGranalla.AddItem "SP7"
    CbTipoGranalla.AddItem "SP10"
    
    m_TotalTiposGranalla = 6
    
End If

If m_TipoDoc = "T" Then ' produccion pintura
    conContratista = False
    Detalle_Config
    n_columnas = 17
End If

If m_TipoDoc = "IRP" Then ' reproceso
    n_columnas = 17
    conContratista = True
    Detalle_Config
End If

Contratista_Visible conContratista

Privilegios

m_NvArea = 0

CbTipoGranalla.visible = False
CbTipoGranalla.Width = 100

CbTrabajadores.visible = False
CbTrabajadores.Width = 1000

End Sub
Private Sub Contratista_Visible(ByVal visible As Boolean)
lblContratista.visible = visible
btnContratistaBuscar.visible = visible
lblContratistaRazon.visible = visible
End Sub
Private Sub Inicializa()
Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnDesHacer = Toolbar.Buttons(6)
Set btnGrabar = Toolbar.Buttons(7)

If m_TipoDoc = "P" Then
    Obj = "ITO PINTURA"
    Objs = "ITOS PINTURA"
    Me.Caption = Obj
'    Trabajadores_Poblar
    n_filas = 20
End If
If m_TipoDoc = "G" Then
    Obj = "ITO GALVANIZADO"
    Objs = "ITOS GALVANIZADO"
    Me.Caption = Obj
    n_filas = 20
End If
If m_TipoDoc = "R" Then
    Obj = "ITO GRANALLADO"
    Objs = "ITOS GRANALLADO"
    Me.Caption = Obj
    Trabajadores_Poblar
    n_filas = 30
End If
If m_TipoDoc = "T" Then
    Obj = "ITO PRODUCCION PINTURA"
    Objs = "ITOS PRODUCCION PINTURA"
    Me.Caption = Obj
    Trabajadores_Poblar
    n_filas = 30
End If
If m_TipoDoc = "IRP" Then
    Obj = "ITO REPROCESO"
    Objs = "ITOS REPROCESO"
    Me.Caption = Obj
    n_filas = 20
End If

Accion = ""

defecto.MaxLength = 5

btnContratistaBuscar.visible = False
btnContratistaBuscar.ToolTipText = "Busca Contratistas"
Campos_Enabled False

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
Private Sub ComboNV_Click()

If ComboNV.Text = " " Then
    m_Nv = 0
    Exit Sub
End If

'If m_TipoDoc = "P" Then
'    If NvTipo(ComboNV.ListIndex) <> "P" Then
'        MsgBox "NV no es Pintura"
'        If ComboNV.Enabled Then ComboNV.SetFocus
'        Exit Sub
'    End If
'End If

'If m_TipoDoc = "G" Then
'    If NvTipo(ComboNV.ListIndex) <> "G" Then
'        MsgBox "NV no es Galvanizada"
'        If ComboNV.Enabled Then ComboNV.SetFocus
'        Exit Sub
'    End If
'End If

MousePointer = vbHourglass

i = 0

m_Nv = Val(Left(ComboNV.Text, 6))
'MsgBox "m_nv" & m_Nv
Nv.Text = m_Nv

ComboPlano.Clear

ComboPlano.AddItem " "
Rev(i) = " "

' busca cliente
m_ClienteRazon = ""
RsNVc.Seek "=", m_Nv, m_NvArea
If Not RsNVc.NoMatch Then
    RsCl.Seek "=", RsNVc![RUT CLiente]
    m_ClienteRazon = "Cliente NO Existe"
    If Not RsCl.NoMatch Then
        m_ClienteRazon = RsCl![Razon Social]
    End If
End If

RsNvPla.Seek ">=", m_Nv, ""
If Not RsNvPla.NoMatch Then
    Do While Not RsNvPla.EOF
        If RsNvPla!Nv = m_Nv Then
            ComboPlano.AddItem RsNvPla!Plano
            i = i + 1
            Rev(i) = RsNvPla!Rev
            
        Else
            Exit Do
        End If
        RsNvPla.MoveNext
    Loop
End If

Detalle_Limpiar Detalle, n_columnas

ComboMarca.Clear

'If Rut.Text <> "" Then Detalle.Enabled = True
Detalle.Enabled = True

MousePointer = vbDefault

End Sub
Private Sub ComboPlano_Click()
' supuesto: el numero del plano es único para toda nv
Dim old_plano As String, filaFlex As Integer

old_plano = Detalle

filaFlex = Detalle.Row

If ComboPlano.ListIndex > 0 Then Detalle.TextMatrix(filaFlex, 2) = Rev(ComboPlano.ListIndex)

'ComboMarca_Poblar np

ComboPlano.visible = False
Detalle = ComboPlano.Text

If Detalle <> old_plano Then
    For i = 2 To n_columnas
        Detalle.TextMatrix(filaFlex, i) = ""
    Next
End If

' revision
If ComboPlano.ListIndex > 0 Then Detalle.TextMatrix(filaFlex, 2) = Rev(ComboPlano.ListIndex)

End Sub
Private Sub ComboMarca_Poblar(Plano As String)
' llena combo marcas
ComboMarca.Clear

RsPd.Seek "=", m_Nv, m_NvArea, Plano, 1
If Not RsPd.NoMatch Then
    Do While Not RsPd.EOF
        If RsPd!Nv = m_Nv And RsPd!Plano = Plano Then
            ComboMarca.AddItem RsPd!Marca
        Else
            Exit Do
        End If
        RsPd.MoveNext
    Loop
End If
End Sub
Private Sub ComboMarca_Click()
Dim m_Plano As String, m_Marca As String, fil As Integer
Dim c_tot As Integer, c_otf As Integer, c_itof As Integer, c_gdgal As Integer, c_itopg As Integer, c_itogr As Integer, c_itopp As Integer

fil = Detalle.Row
ComboMarca.visible = False
m_Plano = Detalle.TextMatrix(fil, 1)
m_Marca = ComboMarca.Text

If m_TipoDoc = "P" Or m_TipoDoc = "G" Then ' verifica en itos que sean pintura y galvanizado
'///
' verifica si Plano-Marca ya están en esta ITO
For i = 1 To n_filas
    If m_Plano = Detalle.TextMatrix(i, 1) And m_Marca = Detalle.TextMatrix(i, 3) Then
        Beep
        MsgBox "MARCA YA EXISTE EN ITO"
        Detalle.Row = i
        Detalle.col = 3
        Detalle.SetFocus
        Exit Sub
    End If
Next
'///
End If

Detalle = m_Marca

c_tot = 0
c_itof = 0
c_gdgal = 0
c_itopg = 0
c_itogr = 0
c_itopp = 0

' busca marca en plano
RsPd.Seek "=", m_Nv, m_NvArea, m_Plano, 1
If Not RsPd.NoMatch Then

    Do While Not RsPd.EOF
    
        If RsPd!Marca = m_Marca Then
        
            c_tot = RsPd![Cantidad Total]
            c_otf = RsPd![OT fab]
            c_itof = RsPd![ITO fab]
            c_gdgal = RsPd![GD gal]
            c_itopg = RsPd![ITO pyg]
            c_itogr = RsPd![ito gr]
            c_itopp = RsPd![ito pp]
            
            ' verifica si está asignada
            If c_otf = 0 Then
                Beep
                MsgBox "La marca """ & m_Marca & """" & vbCr _
                    & "No está Asignada"
                Detalle.TextMatrix(fil, 3) = ""
                Detalle.SetFocus
                Exit Sub
            End If
            
            If m_TipoDoc <> "IRP" Then ' si no es ito reproceso
                If c_itof = 0 Then
                    Beep
                    MsgBox "La marca """ & m_Marca & """" & vbCr _
                        & "No está Recibida"
                    Detalle.TextMatrix(fil, 3) = ""
                    Detalle.SetFocus
                    Exit Sub
                End If
            End If
            
            If m_TipoDoc = "G" Then
                If c_gdgal = 0 Then
                    Beep
                    MsgBox "La marca """ & m_Marca & """" & vbCr _
                        & "No se ha enviado a Galvanizar"
                    Detalle.TextMatrix(fil, 3) = ""
                    Detalle.SetFocus
                    Exit Sub
                End If
            End If
            
            ' verifica que quede algo por despachar
            Select Case m_TipoDoc
            Case "R"
                ' granalla
                If c_tot - c_itogr <= 0 Then
                    Beep
                    MsgBox "La marca """ & m_Marca & """" & vbCr & "Ya se Granalló"
                    Detalle.TextMatrix(fil, 3) = ""
                    Detalle.SetFocus
                    Exit Sub
                End If
            Case "T"
                ' produccion pintura
                If c_tot - c_itopp <= 0 Then
                    Beep
                    MsgBox "La marca """ & m_Marca & """" & vbCr & "Ya se Está en Produccion Pintura"
                    Detalle.TextMatrix(fil, 3) = ""
                    Detalle.SetFocus
                    Exit Sub
                End If
            Case "IRP"
                ' no verifica
            Case Else
                If c_itof - c_itopg <= 0 Then
                    Beep
                    MsgBox "La marca """ & m_Marca & """" & vbCr & "Ya se Recibió"
                    Detalle.TextMatrix(fil, 3) = ""
                    Detalle.SetFocus
                    Exit Sub
                End If
            
            End Select
        
            Detalle.TextMatrix(fil, 4) = RsPd!Descripcion
            
            Select Case m_TipoDoc
            Case "P"
                Detalle.TextMatrix(fil, 5) = c_itof
                Detalle.TextMatrix(fil, 6) = c_itopg
            Case "G"
                Detalle.TextMatrix(fil, 5) = c_gdgal
                Detalle.TextMatrix(fil, 6) = c_itopg
            Case "R"
                Detalle.TextMatrix(fil, 5) = c_tot
                Detalle.TextMatrix(fil, 6) = c_itogr - Val(Detalle.TextMatrix(fil, 7))
            Case "T"
                Detalle.TextMatrix(fil, 5) = c_tot
                Detalle.TextMatrix(fil, 6) = c_itopp - Val(Detalle.TextMatrix(fil, 7))
            Case "IRP"
                Detalle.TextMatrix(fil, 5) = c_otf
                Detalle.TextMatrix(fil, 6) = c_itof
            End Select
            
            
            Detalle.TextMatrix(fil, 8) = Replace(RsPd![Superficie], ",", ".")
            Detalle.TextMatrix(fil, 10) = Replace(RsPd![Peso], ",", ".")
            
            Detalle.TextMatrix(fil, n_columnas) = RsPd![Peso]
            
            Fila_Calcular_Normal fil, True
            
            Exit Do
            
        End If
        RsPd.MoveNext
    Loop
End If

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
Private Sub Detalle_Config()

Dim i As Integer, ancho As Integer

Detalle.Left = 100
Detalle.WordWrap = True
Detalle.RowHeight(0) = 450
Detalle.Rows = n_filas + 1
Detalle.Cols = n_columnas + 1

Detalle.TextMatrix(0, 0) = ""
Detalle.TextMatrix(0, 1) = "Plano"
Detalle.TextMatrix(0, 2) = "Rev"                   '*
Detalle.TextMatrix(0, 3) = "Marca"
Detalle.TextMatrix(0, 4) = "Descripción"           '*

If m_TipoDoc = "P" Then

    Detalle.TextMatrix(0, 5) = "Cant ITOF"            '*

    Detalle.TextMatrix(0, 6) = "Cant ITOP"            '*

    Detalle.ColWidth(14) = 0
    Detalle.ColWidth(15) = 0
    Detalle.ColWidth(16) = 0
    Detalle.ColWidth(17) = 0

End If

If m_TipoDoc = "G" Then

''Detalle.TextMatrix(0, 6) = "Cant ITOG"            '*
    Detalle.TextMatrix(0, 6) = "Cant GDgal"            '*
    Detalle.ColWidth(14) = 0
    Detalle.ColWidth(15) = 0
    Detalle.ColWidth(16) = 0
    Detalle.ColWidth(17) = 0
    
End If

If m_TipoDoc = "R" Then

    Detalle.TextMatrix(0, 6) = "Cant itoGr"            '*
    
    Detalle.TextMatrix(0, 14) = "Tipo Grana"
    Detalle.TextMatrix(0, 15) = "Maquina"
    
    Detalle.ColWidth(14) = 700
    Detalle.ColWidth(15) = 700
    Detalle.ColWidth(16) = 550
    Detalle.ColWidth(17) = 2000
    
End If

If m_TipoDoc = "T" Then

    Detalle.TextMatrix(0, 6) = "Cant ITOP"            '*
    
    Detalle.TextMatrix(0, 14) = "nManos Antic"
    Detalle.TextMatrix(0, 15) = "nManos Termin"
    
    Detalle.ColWidth(14) = 700
    Detalle.ColWidth(15) = 700
    Detalle.ColWidth(16) = 550
    Detalle.ColWidth(17) = 2000
    
End If

If m_TipoDoc = "IRP" Then

    Detalle.TextMatrix(0, 5) = "Cant OTf"   '*

    Detalle.TextMatrix(0, 6) = "Cant ITOf"  '*

    Detalle.ColWidth(12) = 0
    Detalle.ColWidth(13) = 0
    Detalle.ColWidth(14) = 0
    Detalle.ColWidth(15) = 0
    Detalle.ColWidth(16) = 0
    Detalle.ColWidth(17) = 0

End If

Detalle.TextMatrix(0, 7) = "a Recib"
Detalle.TextMatrix(0, 8) = "m2 Uni"         '*
Detalle.TextMatrix(0, 9) = "m2 Tot"            '*
Detalle.TextMatrix(0, 10) = "Peso Uni"
Detalle.TextMatrix(0, 11) = "Peso Total"         '*
Detalle.TextMatrix(0, 12) = "Precio Uni"
Detalle.TextMatrix(0, 13) = "Precio Total"         '*

Detalle.TextMatrix(0, 16) = "Turno"

Detalle.TextMatrix(0, 17) = "Operador"
Detalle.ColWidth(17) = 0

'Detalle.TextMatrix(0, 16) = "peso unitario" ' columna oculta

Detalle.ColWidth(0) = 300
Detalle.ColWidth(1) = 1800 '2000 ' plano
Detalle.ColWidth(2) = 400
Detalle.ColWidth(3) = 2000 '2200 ' marca
Detalle.ColWidth(4) = 1200
Detalle.ColWidth(5) = 500
Detalle.ColWidth(6) = 500
Detalle.ColWidth(7) = 500
Detalle.ColWidth(8) = 600
Detalle.ColWidth(9) = 600
Detalle.ColWidth(10) = 650
Detalle.ColWidth(11) = 800

If m_TipoDoc <> "IRP" Then
    Detalle.ColWidth(12) = 650
    Detalle.ColWidth(13) = 800
End If

'Detalle.ColWidth(16) = 0 ' peso unitario

ancho = 350 ' con scroll vertical
'ancho = 100 ' sin scroll vertical

Totalm2.Width = Detalle.ColWidth(11)
For i = 0 To n_columnas
    If i = 9 Then Totalm2.Left = ancho + Detalle.Left - 350
    If i = 11 Then TotalKg.Left = ancho + Detalle.Left - 350
    If i = 13 Then TotalPrecio.Left = ancho + Detalle.Left - 350
    ancho = ancho + Detalle.ColWidth(i)
Next

Detalle.Width = ancho
Me.Width = ancho + Detalle.Left * 2

For i = 1 To n_filas
    Detalle.TextMatrix(i, 0) = i
Next

For i = 1 To n_filas
    Detalle.Row = i
    Detalle.col = 2
    Detalle.CellAlignment = flexAlignLeftCenter
    Detalle.CellForeColor = vbRed
    Detalle.col = 3
    Detalle.CellAlignment = flexAlignLeftCenter
    Detalle.CellForeColor = vbBlue
    Detalle.col = 4
    Detalle.CellForeColor = vbBlue
    Detalle.col = 5
    Detalle.CellForeColor = vbBlue
    Detalle.col = 6
    Detalle.CellForeColor = vbBlue
    
    Detalle.col = 8
    Detalle.CellForeColor = vbRed
    Detalle.col = 9
    Detalle.CellForeColor = vbRed
    Detalle.col = 10
    Detalle.CellForeColor = vbRed
    Detalle.col = 11
    Detalle.CellForeColor = vbRed
    
Next

txtEditGD.Text = ""

Detalle.Enabled = False

End Sub
Private Sub Numero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    After_Enter
End If
End Sub
Private Sub After_Enter()

Dim n_Copias As Integer

Select Case Accion

Case "Agregando"

    Dim Campos As Variant, Valores As Variant

    Campos = Array("tipo|'" & m_TipoDoc & "", "numero|" & Numero.Text)
'    Valores = Array(m_TipoDoc, Numero.Text)

    If Registro_Existe("doc_cabecera", "tipo='" & m_TipoDoc & "' AND numero=" & Numero.Text) Then

        Doc_Leer
        
        If m_TipoDoc = "IRP" Then
            MsgBox Obj & " YA EXISTE"
        End If
        
        If m_TipoDoc = "P" Then
'            If RsITOpgc!Tipo = "P" Then
'                MsgBox "ITO PINTURA YA EXISTE"
'            Else
'                MsgBox "ITO REPROCESO YA EXISTE"
'            End If
        Else
'            If RsITOpgc!Tipo = "G" Then
'                MsgBox "ITO REPROCESO YA EXISTE"
'            Else
'                MsgBox "ITO PINTURA YA EXISTE"
'            End If
        End If
        
        Detalle.Enabled = False
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
        
    Else
    
        Campos_Enabled True
        Numero.Enabled = False
        Detalle.Enabled = False
        
        Fecha.Text = Format(Now, Fecha_Format)
        
'        GDespecial.SetFocus
        
        btnGrabar.Enabled = True
        btnContratistaBuscar.visible = True
        
    End If
    
Case "Modificando"



'    RsITOpgc.Seek "=", m_TipoDoc, Numero.Text
'    If RsITOpgc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
'    Else
        
        Doc_Leer
        
            Campos_Enabled True
            Numero.Enabled = False
            
'            GDespecial.SetFocus
            
            btnGrabar.Enabled = True
            btnContratistaBuscar.visible = True

'    End If

Case "Eliminando"
    
'    RsITOpgc.Seek "=", m_TipoDoc, Numero.Text
'    If RsITOpgc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
'    Else
    
        Doc_Leer
'        If rsitopgc!Tipo = "N" Then
            Numero.Enabled = False
            If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo) = vbYes Then
                Doc_Eliminar
'                PlanoDetalle_Actualizar Dbm, Nv.Text, m_NvArea, RsITOpgd, "nv-plano-marca", "ito pyg", m_TipoDoc
            End If
'        Else
'            MsgBox "DEBE ELIMINAR COMO GUIA ESPECIAL"
'        End If
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
'    End If

Case "Imprimiendo"

'    RsITOpgc.Seek "=", m_TipoDoc, Numero.Text
'    If RsITOpgc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
'    Else
    
        Doc_Leer
        
        Numero.Enabled = False
            
        Detalle.visible = True
        Detalle.Enabled = True

'    End If

End Select

End Sub
Private Sub Trabajadores_Poblar()

Dim sql As String

CbTrabajadores.Clear

If m_TipoDoc = "T" Then ' pintura
    sql = "SELECT * FROM trabajadores WHERE tipo4 ORDER BY appaterno"
End If
If m_TipoDoc = "R" Then ' granalla
    sql = "SELECT * FROM trabajadores WHERE tipo5 ORDER BY appaterno"
End If

Set RsTra = DbD.OpenRecordset(sql)

With RsTra
i = 0
Do While Not .EOF
    i = i + 1
    a_Trabajadores(0, i) = !Rut
    m_Nombre = !nombres & " " & !appaterno & " " & !apmaterno
    a_Trabajadores(1, i) = m_Nombre
    CbTrabajadores.AddItem m_Nombre
'        Debug.Print !nombres, !appaterno, !apmaterno
    .MoveNext
Loop
.Close
End With
End Sub
Private Sub Doc_Leer()
Dim m_resta As Integer

' CABECERA

SqlRs.Open "SELECT * FROM " & "doc_detalle" & " WHERE tipo='" & m_TipoDoc & "' AND numero=" & Numero.Text, CnxSqlServer_scp0

With SqlRs
Do While Not .EOF
    Fecha.Text = Format(!Fecha, Fecha_Format)
    m_Nv = !Nv
    Nv.Text = m_Nv
    m_ContratistaRut = !Rut
    m_InspectorRut = !Rut2
    Exit Do
'    .MoveNext
Loop
.Close
End With

'MsgBox "m_nv" & m_Nv

' busca nv
RsNVc.Index = "Numero"

RsNVc.Seek "=", m_Nv, m_NvArea

m_ClienteRazon = "NV NO Existe"

If Not RsNVc.NoMatch Then

    ' busca nombre de cliente
    RsCl.Seek "=", RsNVc![RUT CLiente]
    m_ClienteRazon = "Cliente NO Existe"
    If Not RsCl.NoMatch Then
        m_ClienteRazon = RsCl![Razon Social]
    End If

    If m_TipoDoc = "P" And RsNVc!pintura Then
        ComboNV.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra
    End If
    If m_TipoDoc = "G" And RsNVc!galvanizado Then
        ComboNV.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra
    End If
    If m_TipoDoc = "R" Then
        ComboNV.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra
    End If
    If m_TipoDoc = "T" Then
        ComboNV.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra
    End If
    
End If

RsNVc.Index = Nv_Index ' "Número"

'Obs(0).Text = NoNulo(RsITOpgc![Observacion 1])
'Obs(1).Text = NoNulo(RsITOpgc![Observacion 2])

'DETALLE
RsPd.Index = "NV-Plano-Marca"

SqlRs.Open "SELECT * FROM " & "doc_detalle" & " WHERE tipo='" & m_TipoDoc & "' AND numero=" & Numero.Text, CnxSqlServer_scp0

With SqlRs

Do While Not .EOF

'RsITOpgd.Seek "=", m_TipoDoc, Numero.Text, 1
'If Not RsITOpgd.NoMatch Then

'    Do While Not RsITOpgd.EOF
    
'        If m_TipoDoc = "P" Then
        If True Then
        
'            If RsITOpgd!Numero = Numero.Text Then
            
                i = !linea
                
                Detalle.TextMatrix(i, 1) = !Plano
                Detalle.TextMatrix(i, 2) = !Rev
                Detalle.TextMatrix(i, 3) = !Marca
'                Detalle.TextMatrix(i, 7) = RsITOpgd!Cantidad
                
'                RsPd.Seek "=", m_Nv, m_NvArea, RsITOpgd!Plano, RsITOpgd!Marca
                
                If Not RsPd.NoMatch Then
                
                    Detalle.TextMatrix(i, 4) = RsPd!Descripcion
                    Detalle.TextMatrix(i, 5) = RsPd![ITO fab]
                    Detalle.TextMatrix(i, 10) = RsPd![Peso]
                    
 '                   m_resta = IIf(Accion = "Modificando", RsITOpgd!Cantidad, 0)
                    Detalle.TextMatrix(i, 6) = RsPd![ITO pyg] - m_resta
                    
                    If Accion = "Modificando" And m_TipoDoc = "R" Then
                        Detalle.TextMatrix(i, 5) = RsPd![Cantidad Total]
                        Detalle.TextMatrix(i, 6) = RsPd![ito gr] - Val(Detalle.TextMatrix(i, 7))
                    End If
                                    
                End If

'                Detalle.TextMatrix(i, 8) = RsITOpgd![m2 Unitario]
'                Detalle.TextMatrix(i, 12) = RsITOpgd![Precio Unitario]
                                
                If m_TipoDoc = "R" Then
                
'                    Detalle.TextMatrix(i, 14) = NoNulo(RsITOpgd![tipo2]) '12
'                    Detalle.TextMatrix(i, 15) = NoNulo(RsITOpgd![Maquina]) '13
                    
'                    Detalle.TextMatrix(i, 16) = RsITOpgd![Turno] ' 14
                    
                    For j = 0 To 199
'                        If a_Trabajadores(0, j) = RsITOpgd![RUT Operador] Then
                            Detalle.TextMatrix(i, 17) = a_Trabajadores(1, j) '15
                            Exit For
'                        End If
                    Next
                
                End If
                
                If m_TipoDoc = "T" Then
                
'                    Detalle.TextMatrix(i, 14) = RsITOpgd![manos1]
'                    Detalle.TextMatrix(i, 15) = RsITOpgd![manos2]
                    
'                    Detalle.TextMatrix(i, 16) = RsITOpgd![Turno]
                    
                    For j = 0 To 199
'                        If a_Trabajadores(0, j) = RsITOpgd![RUT Operador] Then
                            Detalle.TextMatrix(i, 17) = a_Trabajadores(1, j)
                            Exit For
'                        End If
                    Next
                
                End If
                
                Fila_Calcular_Normal i, False
                
            Else
            
'                Exit Do
                
            End If
        
'        End If
        
'        RsITOpgd.MoveNext
        
        .MoveNext
        
    Loop
    
.Close
End With
'End If

Contratista_Leer m_ContratistaRut
Inspector_Leer m_InspectorRut

Detalle.Row = 1 ' para q' actualice la primera fila del detalle
Detalle_Sumar_Normal

End Sub
Private Sub Contratista_Leer(ByVal Rut)
Rut = Trim(Rut)
SqlRs.Open "SELECT * FROM personas WHERE contratista='S' AND rut='" & Rut & "'", CnxSqlServer_scp0
If SqlRs.EOF Then
    lblContratistaRazon.Caption = "NO Encontrado"
Else
    lblContratistaRazon.Caption = SqlRs![razon_social]
End If
SqlRs.Close
End Sub
Private Sub Inspector_Leer(ByVal Rut)
Rut = Trim(Rut)
SqlRs.Open "SELECT * FROM trabajadores WHERE tipo2='S' AND rut='" & Rut & "'", CnxSqlServer_scp0
If SqlRs.EOF Then
    lblInspectorRazon.Caption = "NO Encontrado"
Else
    lblInspectorRazon.Caption = SqlRs![paterno] & " " & SqlRs![materno] & " " & SqlRs![nombres]
End If
SqlRs.Close
End Sub
Private Function Doc_Validar() As Boolean
Dim porRecibir As Integer, m_Maquina As String
Doc_Validar = False

If conContratista Then

    If m_ContratistaRut = "" Then
        MsgBox "DEBE ELEGIR CONTRATISTA"
    '    Rut.SetFocus
        btnContratistaBuscar.SetFocus
        Exit Function
    End If

End If

For i = 1 To n_filas

    ' plano
    If Trim(Detalle.TextMatrix(i, 1)) <> "" Then
    
        ' marca                3
        If Not CampoReq_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
        
        ' descripcion          4
        
        ' tot cant recibida    5
        ' tot cant despach     6
        
        ' cantidad a despachar 7
        If Not Numero_Valida(Detalle.TextMatrix(i, 7), i, 7) Then Exit Function
        
        ' [can asignada]-[can recibida]>=[can a recibir]
        porRecibir = Detalle.TextMatrix(i, 5) - Val(Detalle.TextMatrix(i, 6))
        If porRecibir < Detalle.TextMatrix(i, 7) Then
            MsgBox "Sólo quedan " & porRecibir & " por Recibir", , "ATENCIÓN"
            Detalle.Row = i
            Detalle.col = 7
            Detalle.SetFocus
            Exit Function
        End If
        
        ' m2 unitario      8
        ' m2 total         9
        ' precio unitario 10
        ' precio total    11
        
        If m_TipoDoc = "T" Then
        
            porRecibir = Val(Detalle.TextMatrix(i, 14))
            If 0 <= porRecibir And porRecibir <= 3 Then
            Else
                MsgBox "Nº de Manos debe ser entre 0 y 3", , "ATENCIÓN"
                Detalle.Row = i
                Detalle.col = 14
                Detalle.SetFocus
                Exit Function
            End If
            
            porRecibir = Val(Detalle.TextMatrix(i, 15))
            If 0 <= porRecibir And porRecibir <= 3 Then
            Else
                MsgBox "Nº de Manos debe ser entre 0 y 3", , "ATENCIÓN"
                Detalle.Row = i
                Detalle.col = 15
                Detalle.SetFocus
                Exit Function
            End If
            
        End If
        
        If m_TipoDoc = "R" Then ' granalla
        
'            m_Maquina = UCase(Detalle.TextMatrix(i, 13))
            m_Maquina = UCase(Detalle.TextMatrix(i, 15))
            If m_Maquina <> "A" And m_Maquina <> "M" Then
                MsgBox "Maquina debe ser A ó M", , "ATENCIÓN"
                Detalle.Row = i
'                Detalle.col = 13
                Detalle.col = 15
                Detalle.SetFocus
                Exit Function
            End If
                        
        End If
        
        If m_TipoDoc = "R" Or m_TipoDoc = "T" Then
        
'            porRecibir = Val(Detalle.TextMatrix(i, 14))
            porRecibir = Val(Detalle.TextMatrix(i, 16))
            If 1 <= porRecibir And porRecibir <= 2 Then
            Else
                MsgBox "Turno debe ser 1 ó 2", , "ATENCIÓN"
                Detalle.Row = i
'                Detalle.col = 14
                Detalle.col = 16
                Detalle.SetFocus
                Exit Function
            End If
            
If False Then
            If Detalle.TextMatrix(i, 15) = "" Then
                MsgBox "Debe escoger Operador", , "ATENCIÓN"
                Detalle.Row = i
                Detalle.col = 15
                Detalle.SetFocus
                Exit Function
            End If
End If

        End If
            
        ' peso unitario   15 (columna oculta)
        
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
'    If num <> "" Then
        Beep
        MsgBox "Número no Válido"
        Detalle.Row = fil
        Detalle.col = col
        Detalle.SetFocus
        Exit Function
'    End If
End If
Numero_Valida = True
End Function
Private Sub Doc_Grabar(Nueva As Boolean)

MousePointer = vbHourglass

Dim m_Plano As String, m_Marca As String, m_cantidad As Integer
Dim m_PesoUnitario As Double, m_PesoTotal As Double, TotalCampos As Integer
Dim qry As String, jj As Integer

m_PesoTotal = 0

If Not Nueva Then Doc_Detalle_Eliminar

' DETALLE DE ITO

Campos_Detalle_Definir
' limpia campos "valores"
asn(1) = False
asn(2) = False
av(1) = m_TipoDoc
av(2) = Numero.Text
'av(3) linea
av(4) = Fecha.Text
av(5) = Nv.Text
av(6) = m_ContratistaRut
av(7) = m_InspectorRut

TotalCampos = 11
For i = 3 To jj
    av(i) = ""
    asn(i) = True
Next

For i = 1 To n_filas

    m_Plano = Trim(Detalle.TextMatrix(i, 1))

    If m_Plano <> "" Then

        av(3) = i ' linea
        av(8) = m_Plano

        av(9) = Detalle.TextMatrix(i, 2) ' rev
        av(10) = Val(Detalle.TextMatrix(i, 3)) ' marca

        av(11) = Val(Detalle.TextMatrix(i, 7)) ' cantidad

        If m_TipoDoc = "R" Then

'            ![tipo2] = Detalle.TextMatrix(i, 14) '12
'            ![Maquina] = UCase(Detalle.TextMatrix(i, 15)) '13
            
'            ![Turno] = m_CDbl(Detalle.TextMatrix(i, 16)) '14
            
'            For jj = 0 To 199
'                If a_Trabajadores(1, jj) = Detalle.TextMatrix(i, 17) Then ' 15
'                    ![RUT Operador] = a_Trabajadores(0, jj)
'                    Exit For
'                End If
'            Next
            
        End If

        If m_TipoDoc = "T" Then ' ito produccion pintura
        
'            ![manos1] = m_CDbl(Detalle.TextMatrix(i, 14))
'            ![manos2] = m_CDbl(Detalle.TextMatrix(i, 15))
            
'            ![Turno] = m_CDbl(Detalle.TextMatrix(i, 16))
            
'            For jj = 0 To 199
'                If a_Trabajadores(1, jj) = Detalle.TextMatrix(i, 17) Then
'                    ![RUT Operador] = a_Trabajadores(0, jj)
'                    Exit For
'                End If
'            Next
            
        End If

        Registro_Agregar CnxSqlServer_scp0, "doc_detalle", ac, av, TotalCampos
        
        m_PesoTotal = m_PesoTotal + (m_cantidad * m_PesoUnitario)
        
    End If
    
Next
    
save:
' CABECERA DE ITO
'With RsITOpgc
If Nueva Then
'    .AddNew
'    !Número = Numero.Text
'    !Tipo = m_TipoDoc
'    !Numero = Numero.Text
Else

'    Doc_Detalle_Eliminar
    
'    .Edit
    
End If

'!Fecha = Fecha.Text
'!Nv = Val(m_Nv)
'![RUT Contratista] = m_ContratistaRut
'![m2 Total] = m_CDbl(Totalm2.Caption)
'![Peso Total] = m_PesoTotal
'![Precio Total] = Val(TotalPrecio.Caption)
'![Observacion 1] = Obs(0).Text
'![Observacion 2] = Obs(1).Text
'.Update

'End With

'PlanoDetalle_Actualizar Dbm, Nv.Text, m_NvArea, RsITOpgd, "nv-plano-marca", "ito pyg", m_TipoDoc
    
MousePointer = vbDefault

End Sub
Private Sub Doc_Eliminar()

' elimina cabecera
'RsITOpgc.Seek "=", m_TipoDoc, Numero.Text
'If Not RsITOpgc.NoMatch Then
    
'    RsITOpgc.Delete

'End If

' elimina detalle
Doc_Detalle_Eliminar

End Sub
Private Sub Doc_Detalle_Eliminar()
' elimina detalle ITO
' al anular detalle ITO debe actualizar detalle plano

RsPd.Index = "NV-Plano-Marca"
'RsITOpgd.Seek "=", m_TipoDoc, Numero.Text, 1
'If Not RsITOpgd.NoMatch Then
'    Do While Not RsITOpgd.EOF
'        If RsITOpgd!Numero <> Numero.Text Then Exit Do
'        RsPd.Seek "=", m_Nv, m_NvArea, RsITOpgd!Plano, RsITOpgd!Marca
        If Not RsPd.NoMatch Then
            RsPd.Edit
            If m_TipoDoc = "P" Or m_TipoDoc = "G" Then
'                RsPd![ITO pyg] = RsPd![ITO pyg] - RsITOpgd!Cantidad
            End If
            If m_TipoDoc = "R" Then
'                RsPd![ITO gr] = RsPd![ITO gr] - RsITOpgd!Cantidad
            End If
            RsPd.Update
        End If
    
        ' borra detalle
'        RsITOpgd.Delete
    
'        RsITOpgd.MoveNext
'    Loop
'End If
RsPd.Index = "NV-Plano-Item"


End Sub
Private Sub Campos_Limpiar()
Numero.Text = ""
'GDespecial.Value = 0
Fecha.Text = Fecha_Vacia
'Fecha.Text = Format(Now, Fecha_Format)
ComboNV.Text = " "
m_Nv = 0
Nv.Text = m_Nv

m_ContratistaRut = ""
lblContratistaRazon.Caption = ""

m_InspectorRut = ""
lblInspectorRazon.Caption = ""

defecto.Text = ""

Detalle_Limpiar Detalle, n_columnas
Obs(0).Text = ""
Obs(1).Text = ""
Totalm2.Caption = ""
TotalPrecio.Caption = ""

End Sub
Private Sub Detalle_Limpiar(Detalle As Control, n_columnas As Integer)

Dim fi As Integer, co As Integer

For fi = 1 To n_filas
    For co = 1 To n_columnas
        Detalle.TextMatrix(fi, co) = ""
    Next
Next

'Detalle.Row = 1
If Detalle.Enabled Then
    Detalle.SetFocus
'    SendKeys "^{HOME}", True
End If

End Sub
Private Sub Obs_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If Index = 0 Then
        Obs(1).SetFocus
    Else
        Obs(0).SetFocus
    End If
End If
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim cambia_titulo As Boolean, n_Copias As Integer ', m_ImprimeITO As Boolean
cambia_titulo = True
'Accion = ""
Select Case Button.Index
Case 1 ' agregar
    Accion = "Agregando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    
'    Numero.Text = Documento_Numero_Nuevo_PG(m_TipoDoc, RsITOpgc)

    Numero.Enabled = True
    Numero.SetFocus
Case 2 ' modificar
    Accion = "Modificando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    Numero.Enabled = True
    Numero.SetFocus
Case 3 ' eliminar
    Accion = "Eliminando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    Numero.Enabled = True
    Numero.SetFocus
Case 4 ' imprimir

    Accion = "Imprimiendo"
    If Numero.Text = "" Then
    
        Botones_Enabled 0, 0, 0, 1, 1, 0
        Campos_Enabled False
        Numero.Enabled = True
        Numero.SetFocus
        
    Else

        If MsgBox("¿ IMPRIMIR " & Obj & " ?", vbYesNo) = vbYes Then
            Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Docs")
            Doc_Imprimir
        End If

        If m_TipoDoc = "P" Or m_TipoDoc = "G" Then
            If MsgBox("¿ IMPRIMIR ETIQUETAS ?", vbYesNo) = vbYes Then
                Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Etiquetas")
                Etiquetas_Imprimir
    '            Impresora_Predeterminada "default"
            End If
        End If

        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
        
    End If

Case 5 ' separador
Case 6 ' DesHacer
    If Numero.Text = "" Then
        GoTo DesHace
    Else
        If Accion = "Imprimiendo" Then
            GoTo DesHace
        Else
DesHace:
                Privilegios
                Campos_Limpiar
                Campos_Enabled False
                Accion = ""
        End If
    End If
Case 7 ' grabar

    If Doc_Validar Then
        
        If MsgBox("¿ GRABAR " & Obj & " ?", vbYesNo) = vbYes Then
        
            If Accion = "Agregando" Then
                Doc_Grabar True
            Else
                Doc_Grabar False
            End If
            
            If MsgBox("¿ IMPRIMIR " & Obj & " ?", vbYesNo) = vbYes Then
                Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Docs")
                Doc_Imprimir
'                Impresora_Predeterminada "default"
            End If
            
            If m_TipoDoc = "P" Or m_TipoDoc = "G" Then
                If MsgBox("¿ IMPRIMIR ETIQUETAS ?", vbYesNo) = vbYes Then
    '                If MsgBox("Debe configurar Impresora ZEBRA como Prederminada", vbYesNo) = vbYes Then
    '                    Impresora_Predeterminada "zebra"
                        Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Etiquetas")
                        Etiquetas_Imprimir
    '                    Impresora_Predeterminada "default"
    '                End If
                End If
            End If
            
            Botones_Enabled 0, 0, 0, 0, 1, 0
            Campos_Limpiar
            Campos_Enabled False
            Numero.Enabled = True
            Numero.SetFocus
            
'            If Accion = "Agregando" Then Numero.Text = Documento_Numero_Nuevo_PG(m_TipoDoc, RsITOpgc)
            
        End If
    End If
Case 8 ' separador
Case 9 ' contratistas
    MousePointer = vbHourglass
    Load sql_contratistas
    MousePointer = vbDefault
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

If Usuario.AccesoTotal Then
    Fecha.Enabled = Si
Else
    Fecha.Enabled = False
End If

btnContratistaBuscar.Enabled = Si
btnInspectorBuscar.Enabled = Si
Nv.Enabled = Si
ComboNV.Enabled = Si

defecto.Enabled = Si

Detalle.Enabled = Si

Obs(0).Enabled = Si
Obs(1).Enabled = Si

End Sub
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
' RUTINAS PARA EL FLEXGRID
Private Sub Detalle_Click()
If Accion = "Imprimiendo" Then Exit Sub
After_Detalle_Click
End Sub
Private Sub After_Detalle_Click()
ComboPlano.visible = False
ComboMarca.visible = False
Select Case Detalle.col
    Case 1 ' plano
        If Detalle <> "" Then ComboPlano.Text = Detalle
        ComboPlano.Top = Detalle.CellTop + Detalle.Top
        ComboPlano.Left = Detalle.CellLeft + Detalle.Left
        ComboPlano.Width = Int(Detalle.CellWidth * 1.5)
        ComboPlano.visible = True
        ComboMarca.visible = False
    Case 3 ' marca
        ComboMarca_Poblar Detalle.TextMatrix(Detalle.Row, 1)
        If Detalle <> "" Then ComboMarca.Text = Detalle
        ComboMarca.Top = Detalle.CellTop + Detalle.Top
        ComboMarca.Left = Detalle.CellLeft + Detalle.Left
        ComboMarca.Width = Int(Detalle.CellWidth * 1.5)
        ComboPlano.visible = False
        ComboMarca.visible = True
    Case Else
End Select
End Sub
Private Sub Detalle_DblClick()
If Accion = "Imprimiendo" Then Exit Sub
MSFlexGridEdit Detalle, txtEditGD, 32
End Sub
Private Sub Detalle_GotFocus()
If txtEditGD.visible Then
    Detalle = txtEditGD
    txtEditGD.visible = False
End If
End Sub
Private Sub Detalle_LeaveCell()
If txtEditGD.visible Then
    Detalle = txtEditGD
    txtEditGD.visible = False
End If
End Sub
Private Sub Detalle_KeyPress(KeyAscii As Integer)
If Accion = "Imprimiendo" Then Exit Sub
MSFlexGridEdit Detalle, txtEditGD, KeyAscii
End Sub
Private Sub txtEditGD_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode_N Detalle, txtEditGD, KeyCode, Shift
End Sub
Sub EditKeyCode_N(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
' rutina que se ejecuta con los keydown de Edt
Dim m_fil As Integer, m_col As Integer, dif As Integer
m_fil = MSFlexGrid.Row
m_col = MSFlexGrid.col
dif = Val(Detalle.TextMatrix(MSFlexGrid.Row, 5)) - Val(Detalle.TextMatrix(MSFlexGrid.Row, 6))
Select Case KeyCode
Case vbKeyEscape ' Esc
    Edt.visible = False
    MSFlexGrid.SetFocus
Case vbKeyReturn ' Enter
    Select Case m_col
    Case 7 ' Cantidad a Recibir
        If Recibida_Validar(MSFlexGrid.col, dif, Edt) Then
            MSFlexGrid.SetFocus
            DoEvents
            Fila_Calcular_Normal m_fil, True
        End If
    Case 14
        If m_TipoDoc = "T" Then ' manos anticorrisivo
            If Manos_Validar(MSFlexGrid.col, Edt) Then
                MSFlexGrid.SetFocus
                DoEvents
                Fila_Calcular_Normal m_fil, True
            End If
        End If
    Case 15
        If m_TipoDoc = "T" Then ' manos terminacion
'            dif = Val(Detalle.TextMatrix(MSFlexGrid.Row, 13))
            If Manos_Validar(MSFlexGrid.col, Edt) Then
                MSFlexGrid.SetFocus
                DoEvents
                Fila_Calcular_Normal m_fil, True
            End If
        End If
        If m_TipoDoc = "R" Then ' maquina
            If Maquina_Validar(MSFlexGrid.col, Edt) Then
                MSFlexGrid.SetFocus
                DoEvents
            End If
        End If
    Case 16 ' turno
        If m_TipoDoc = "R" Or m_TipoDoc = "T" Then
'            dif = Val(Detalle.TextMatrix(MSFlexGrid.Row, 14))
            If Turno_Validar(MSFlexGrid.col, Edt) Then
                MSFlexGrid.SetFocus
                DoEvents
            End If
        End If
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
        Fila_Calcular_Normal m_fil, True
    End Select
    Cursor_Mueve_N MSFlexGrid
Case vbKeyUp ' Flecha Arriba
    Select Case m_col
    Case 7 ' Cantidad a recibir
        If Recibida_Validar(MSFlexGrid.col, dif, Edt) Then
            MSFlexGrid.SetFocus
            DoEvents
            Fila_Calcular_Normal m_fil, True
            If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
                MSFlexGrid.Row = MSFlexGrid.Row - 1
            End If
        End If
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
        Fila_Calcular_Normal m_fil, True
        If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
            MSFlexGrid.Row = MSFlexGrid.Row - 1
        End If
    End Select
Case vbKeyDown ' Flecha Abajo
    Select Case m_col
    Case 7 ' Cantidad a recibir
        If Recibida_Validar(MSFlexGrid.col, dif, Edt) Then
            MSFlexGrid.SetFocus
            DoEvents
            Fila_Calcular_Normal m_fil, True
            If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
                MSFlexGrid.Row = MSFlexGrid.Row + 1
            End If
        End If
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
        Fila_Calcular_Normal m_fil, True
        If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
            MSFlexGrid.Row = MSFlexGrid.Row + 1
        End If
    End Select
End Select
End Sub
Private Function Recibida_Validar(Colu As Integer, porRecibir As Integer, Edt As Control) As Boolean
' verifica que CRecibida-CDespachada >= CADespachar
Recibida_Validar = True
If Colu <> 7 Then Exit Function
If porRecibir < Val(Edt) Then
    MsgBox "Sólo quedan " & porRecibir & " por Recibir", , "ATENCIÓN"
    Recibida_Validar = False
End If
End Function
Private Function Manos_Validar(Colu As Integer, Edt As Control) As Boolean
Manos_Validar = True
If 0 <= Val(Edt) And Val(Edt) <= 3 Then
Else
    If Detalle.col = 12 Then ' manos anticorrosivo
        MsgBox "Nº de manos de Anticorrisivo debe ser entre 0 y 3", , "ATENCIÓN"
        Detalle.col = 11
        Detalle.SetFocus
    End If
    If Detalle.col = 13 Then ' manos terminacion
        MsgBox "Nº de manos de Terminación debe ser entre 0 y 3", , "ATENCIÓN"
        Detalle.col = 12
        Detalle.SetFocus
    End If
    Manos_Validar = False
End If
End Function
Private Function Maquina_Validar(Colu As Integer, Edt As Control) As Boolean
Maquina_Validar = True
Edt.Text = UCase(Edt.Text)
If Edt.Text <> "M" And Edt <> "A" Then
'Else
    MsgBox "Maquina debe ser A ó M", , "ATENCIÓN"
    Detalle.col = 15 '12
    Detalle.SetFocus
    Maquina_Validar = False
End If
End Function
Private Function Turno_Validar(Colu As Integer, Edt As Control) As Boolean
Turno_Validar = True
If 1 <= Val(Edt) And Val(Edt) <= 2 Then
Else
    MsgBox "Turno debe ser 1 ó 2", , "ATENCIÓN"
    Detalle.col = 16 '13
    Detalle.SetFocus
    Turno_Validar = False
End If
End Function
Private Sub txtEditGD_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)
Select Case MSFlexGrid.col
Case 1, 2, 3
'    After_Detalle_Click
Case 4, 5, 6, 8, 9, 10, 11
    ' no editables
    Exit Sub
Case 14 ' tipo granalla
    If m_TipoDoc = "R" Then
        CbTipoGranalla.Left = MSFlexGrid.CellLeft + MSFlexGrid.Left
        CbTipoGranalla.Top = MSFlexGrid.CellTop + MSFlexGrid.Top
        CbTipoGranalla.Width = Detalle.ColWidth(12)
        CbTipoGranalla.visible = True
        CbTipoGranalla.SetFocus
    End If
    If m_TipoDoc = "T" Then GoTo Editar
Case 17 ' combo trabajador ex 15
    CbTrabajadores.Left = MSFlexGrid.CellLeft + MSFlexGrid.Left
    CbTrabajadores.Top = MSFlexGrid.CellTop + MSFlexGrid.Top
    CbTrabajadores.Width = Detalle.ColWidth(17) '15
    CbTrabajadores.visible = True
    CbTrabajadores.SetFocus
Case Else
Editar:
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
    MSFlexGridEdit Detalle, txtEditGD, 32
End If
End Sub
Private Sub Cursor_Mueve_N(MSFlexGrid As Control)
'MIA
Select Case MSFlexGrid.col
Case 7 ' cantidad a recibir
        MSFlexGrid.col = MSFlexGrid.col + 5
'Case 10 '
'        MSFlexGrid.col = MSFlexGrid.col + 4
'Case 10
'    If MSFlexGrid.Row + 1 < MSFlexGrid.Rows Then
'        MSFlexGrid.col = 1
'        MSFlexGrid.Row = MSFlexGrid.Row + 1
'    End If
Case Else
    MSFlexGrid.col = MSFlexGrid.col + 1
End Select
End Sub
Private Sub Fila_Calcular_Normal(Fila As Integer, Actualizar As Boolean)

' actualiza solo linea, y totales generales

n7 = m_CDbl(Detalle.TextMatrix(Fila, 7)) ' a recibir
n8 = m_CDbl(Detalle.TextMatrix(Fila, 8)) ' m2 uni
n10 = m_CDbl(Detalle.TextMatrix(Fila, 10)) ' peso uni
n12 = m_CDbl(Detalle.TextMatrix(Fila, 12)) ' precio uni

If m_TipoDoc = "T" Then ' produccion pintura


    n14 = m_CDbl(Detalle.TextMatrix(Fila, 14))
    n15 = m_CDbl(Detalle.TextMatrix(Fila, 15))
    ' m2 total
    Detalle.TextMatrix(Fila, 9) = Format(n7 * n8 * (n14 + n15), "#.00")
    
    ' peso total linea
    Detalle.TextMatrix(Fila, 11) = Format(n7 * n10, "#.00")
    
Else
    ' m2 total
    Detalle.TextMatrix(Fila, 9) = Format(n7 * n8, "#.00")
    ' peso total
    Detalle.TextMatrix(Fila, 11) = Format(n7 * n10, "#.00")
End If

' precio total
Detalle.TextMatrix(Fila, 13) = Format(n7 * n8 * n12, "#.00")

If Actualizar Then Detalle_Sumar_Normal

End Sub
Private Sub Detalle_Sumar_Normal()
Dim Tot_m2 As Double, Tot_Kg As Double, Tot_Precio As Double
Tot_m2 = 0
Tot_Kg = 0
Tot_Precio = 0
For i = 1 To n_filas
    Tot_m2 = Tot_m2 + m_CDbl(Detalle.TextMatrix(i, 9))
    Tot_Kg = Tot_Kg + m_CDbl(Detalle.TextMatrix(i, 11))
    Tot_Precio = Tot_Precio + m_CDbl(Detalle.TextMatrix(i, 13))
Next

Totalm2.Caption = Format(Tot_m2, "#,###.00")
TotalKg.Caption = Format(Tot_Kg, "#,###.00")
TotalPrecio.Caption = Format(Tot_Precio, num_fmtgrl)

End Sub
' FIN RUTINAS PARA FLEXGRID
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
Private Sub Doc_Imprimir()
' imprime ITOf
MousePointer = vbHourglass
Dim can_valor As String, can_col As Integer
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer, tab4 As Integer
Dim tab9 As Integer, tab10 As Integer
Dim tab5 As Integer, tab6 As Integer, tab7 As Integer, tab8 As Integer, tab40 As Integer
tab0 = 2 'margen izquierdo era 7
tab1 = tab0 + 0 ' plano
tab2 = tab1 + 10 ' rev
tab3 = tab2 + 4 ' marca
tab4 = tab3 + 10 ' descripcion
tab5 = tab4 + 20 ' cant
tab6 = tab5 + 7 ' m2 uni
tab7 = tab6 + 6 ' m2 tot
tab8 = tab7 + 10 ' kg uni
tab9 = tab8 + 9 ' kg tot
tab10 = tab9 + 10
tab40 = 50

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt

prt.Font.Size = 15
prt.Print Tab(tab0 + 14); "VALE " & Obj & " Nº" & Format(Numero.Text, "000")
prt.Font.Size = 10
prt.Print ""

' cabecera
prt.Font.Bold = True
prt.Print Tab(tab0); Empresa.Razon;
prt.Font.Bold = False
prt.Print Tab(tab0 + tab40); "FECHA     : " & Fecha.Text

prt.Print Tab(tab0); "GIRO: " & Empresa.Giro;
prt.Print Tab(tab0 + tab40); "SEÑOR(ES) : " & lblContratistaRazon.Caption,

prt.Print Tab(tab0); Empresa.Direccion;
prt.Print Tab(tab0 + tab40); "RUT       : " & m_ContratistaRut

prt.Print Tab(tab0); "Teléfono: " & Empresa.Telefono1 & " - " & Empresa.Comuna;
'prt.Print Tab(tab0 + tab40); "DIRECCIÓN : " & Direccion,

'prt.Print Tab(tab0 + tab40); "COMUNA    : " & Comuna

prt.Print Tab(tab0 + tab40); "OBRA      : ";
prt.Font.Bold = True
'prt.Print Format(Mid(ComboNV.Text, 8), ">")
prt.Print ComboNV.Text
prt.Font.Bold = False

prt.Print ""
' detalle
prt.Font.Bold = True
prt.Print Tab(tab1); "PLANO";
prt.Print Tab(tab2); "REV";
prt.Print Tab(tab3); "MARCA";
prt.Print Tab(tab4); "DESCRIPCIÓN";
'prt.Print Tab(tab5); "Nº OT";
prt.Print Tab(tab5); "";
prt.Print Tab(tab6); "CANT";
prt.Print Tab(tab7); "  m2 UNIT";
prt.Print Tab(tab8); " m2 TOTAL";
prt.Print Tab(tab9); "  Kg UNIT";
prt.Print Tab(tab10); " Kg TOTAL"
prt.Font.Bold = False

prt.Print Tab(tab1); String(110, "-")
j = -1
For i = 1 To n_filas

    can_valor = Detalle.TextMatrix(i, 7)
    
    If Val(can_valor) = 0 Then
    
        j = j + 1
        prt.Print Tab(tab1 + j * 5); "    \"
        
    Else
    
        ' PLANO
        prt.Print Tab(tab1); Detalle.TextMatrix(i, 1);
        
        ' REVISIÓN
        prt.Print Tab(tab2); Detalle.TextMatrix(i, 2);
        
        ' MARCA
        prt.Print Tab(tab3); Detalle.TextMatrix(i, 3);
        
        ' DESCRIPCIÓN
        prt.Print Tab(tab4); Left(Detalle.TextMatrix(i, 4), 18);
        
'        ' Nº OT
'        prt.Print Tab(tab5); Detalle.TextMatrix(i, 7);
        
        ' CANTIDAD
        can_valor = Trim(Format(can_valor, "####"))
        can_col = 4 - Len(can_valor)
        prt.Print Tab(tab6 + can_col); can_valor;
        
        ' m2 UNITARIO
        can_valor = Trim(Format(m_CDbl(Detalle.TextMatrix(i, 8)), "##,###.00"))
        can_col = 9 - Len(can_valor)
        prt.Print Tab(tab7 + can_col); can_valor;
        
        ' m2 TOTAL
        can_valor = Trim(Format(m_CDbl(Detalle.TextMatrix(i, 9)), "##,###.00"))
        can_col = 9 - Len(can_valor)
        prt.Print Tab(tab8 + can_col); can_valor;
        
        ' Kg UNITARIO
        can_valor = Trim(Format(m_CDbl(Detalle.TextMatrix(i, 10)), "#,###.00"))
        can_col = 9 - Len(can_valor)
        prt.Print Tab(tab9 + can_col); can_valor;
        
        ' KG TOTAL
        can_valor = Trim(Format(m_CDbl(Detalle.TextMatrix(i, 11)), "##,###.00"))
        can_col = 9 - Len(can_valor)
        prt.Print Tab(tab10 + can_col); can_valor
        
    End If
    
Next

prt.Print Tab(tab1); String(110, "-")
prt.Print Tab(tab0 + 40); "TOTAL m2 : "; Format(Totalm2.Caption, "#,###,###.00");
prt.Print "  TOTAL Kg : "; Format(TotalKg.Caption, "#,###,###.00");
'prt.Print "TOTAL $ : "; Format(TotalPrecio.Caption, "#,###,###.00")
'prt.Font.Bold = True
'prt.Print
'prt.Font.Bold = False
prt.Print ""
prt.Print Tab(tab0); "OBSERVACIONES :";
prt.Print Tab(tab0 + 16); Obs(0).Text
prt.Print Tab(tab0 + 16); Obs(1).Text
'prt.Print Tab(tab0 + 16); Obs(2).Text
'prt.Print Tab(tab0 + 16); Obs(3).Text

For i = 1 To 1
    prt.Print ""
Next

prt.Print Tab(tab0); Tab(14), "__________________", Tab(56), "__________________"
prt.Print Tab(tab0); Tab(14), "       VºBº       ", Tab(56), "       VºBº       "

prt.EndDoc

'Impresora_Predeterminada "default"

MousePointer = vbDefault

End Sub
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
Private Sub Etiquetas_Imprimir()

Dim li As Integer

Dim li1 As Double, li2 As Double, li3 As Double, li4 As Double, li5 As Double ', li6 As Double
Dim dif_linea As Double, Copia As Integer, n_Copias As Integer
Dim Margen_Izquierdo As Double, m_TamanoCodeBar As Integer

dif_linea = 0.57

'Margen_Izquierdo = 0.9
Margen_Izquierdo = 0.7

m_TamanoCodeBar = Val(ReadIniValue(Path_Local & "scp.ini", "Printer", "LabelBarCodeSize"))
If m_TamanoCodeBar = 0 Then m_TamanoCodeBar = 30

li1 = 1.2
li2 = li1 + dif_linea
li3 = li2 + dif_linea
li4 = li3 + dif_linea
'li6 = li5 + dif_linea * 1.5 ' para codigo de barras
'li5 = li4 + dif_linea * 1.5
'li5 = li4 + dif_linea * 1.3
li5 = li4 + dif_linea * 1.2

For li = 1 To n_filas

    n_Copias = Val(Trim(Detalle.TextMatrix(li, 7))) ' Cant a Asignar
'    If n_Copias > 0 Then n_Copias = 1
    
    If n_Copias > 0 Then
    
        For Copia = 1 To n_Copias
        
'            Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Etiquetas") ' puesta el 16/03/06
'            MsgBox Printer.DeviceName & vbLf & prt.DeviceName
            
            Prt_Ini
        
            'Debug.Print m_ClienteRazon
'            m_NV = Detalle.TextMatrix(li, 1)
            m_obra = Trim(Mid(ComboNV.Text, 7)) 'Detalle.TextMatrix(li, 2)
            m_Plano = Detalle.TextMatrix(li, 1)
            m_Rev = Detalle.TextMatrix(li, 2)
            m_Marca = Detalle.TextMatrix(li, 3)
            
            m_Peso = m_CDbl(Detalle.TextMatrix(li, 10))
        
            ' font para logo
            Printer.Font.Name = "delgado"
            Printer.Font.Size = 32
            Printer.Font.Bold = True
            
            SetpYX -0.15, Margen_Izquierdo
            Printer.Print "Delgado"
            '//////////////////
            
            Printer.Font.Name = "Arial Black" ' oficial
'            Printer.Font.Name = "Arial"
            Printer.Font.Bold = False
            Printer.Font.Size = 12 '14
            
            SetpYX li1, Margen_Izquierdo
            Printer.Print "Cliente: "; Left(m_ClienteRazon, 23)
            
    '        SetpYX 1.8, 0.5
    '        prt.Print "Obra: "; m_NV
    '        SetpYX 2.4, 0.9
    '        prt.Print m_Obra
            
            SetpYX li2, Margen_Izquierdo
            prt.Print "Obra: "; m_Nv; " "; Left(m_obra, 20)
            
            SetpYX li3, Margen_Izquierdo
            prt.Print "Plano: "; m_Plano; " rev "; m_Rev;
            
'            SetpYX li3, Margen_Izquierdo + 6.5
'            prt.Print "OT: "; Numero.Text
'            prt.Print "itopin: "; Numero.Text
            
            Printer.Font.Size = 11
            SetpYX li4, Margen_Izquierdo
            prt.Print "Marca: "; m_Marca;
'            SetpYX li4, 5.3
''''            SetpYX li5, Margen_Izquierdo
            prt.Print "  Peso(Kg): "; m_Peso
'            Debug.Print m_Peso
            
            '//////////////////////////////
'            prt.Font.Name = "barcod39"

'            prt.Font.Name = "IDAutomationHC39M"
'            prt.Font.Size = 10 ' 29

            prt.Font.Name = "code 128"
            prt.Font.Size = m_TamanoCodeBar ' 32 oficial
            prt.Font.Bold = False
            
            SetpYX li5, 0.4 'Margen_Izquierdo
'            prt.Print txt2code128(m_Nv & "/" & m_Plano & "/" & m_Rev & "/" & m_Marca)
            prt.Print txt2code128(m_Plano & m_Rev & "/" & m_Marca)
            
            prt.Font.Name = "Arial"
            prt.Font.Size = 8
'            SetpYX 6.15, 9
'            SetpYX 4.882, 9
'            SetpYX 4.93, 9
            SetpYX 4.91, 9
'            SetpYX 4.65, 9
            prt.Print "."
        
        Next
        
    End If
Next
    
prt.EndDoc
    
End Sub
Private Sub Prt_Ini()
Set prt = Printer
Printer.ScaleMode = 1 ' twips : 576 twips x cm
Printer.ScaleMode = 7 ' centimetros
End Sub
Private Sub Prt_Ini_Zebra_nousada()
Dim impr As Printer

'Debug.Print "0"; Printer.DeviceName
For Each impr In Printers
    If UCase(Left(impr.DeviceName, 5)) = "ZEBRA" Then
        ' deja la impresora como predeterminada para VB
        Set Printer = impr
'        Debug.Print "1"; impr.DeviceName
    End If
Next

'Debug.Print "2"; Printer.DeviceName
Set prt = Printer
'Debug.Print "3"; Printer.DeviceName
'Debug.Print "4"; prt.DeviceName

'prt.Height = 3744 ' 6.5 cms
prt.Height = 3014 ' 5.532 cms
'prt.Height = 2880 ' 5 cms
prt.ScaleMode = 1 ' twips : 576 twips x cm
'prt.Scale (0, 0)-(10, 6.5)
prt.Scale (0, 0)-(10, 5.232)
'prt.Scale (0, 0)-(10, 5)
prt.ScaleMode = 7 ' centimetros

End Sub
Private Sub SetpYX(Y As Double, x As Double)
Printer.CurrentY = AjusteY + Y
Printer.CurrentX = AjusteX + x
End Sub
Private Sub Doc_ImprimeRPT(Numero_Ito)

cr.WindowTitle = "ITO Nº " & Numero_Ito
cr.ReportSource = crptReport
cr.WindowState = crptMaximized
'Cr.WindowBorderStyle = crptFixedSingle
cr.WindowMaxButton = False
cr.WindowMinButton = False
cr.Formulas(0) = "RAZON SOCIAL=""" & EmpOC.Razon & """"
cr.Formulas(1) = "GIRO=""" & "GIRO: " & EmpOC.Giro & """"
cr.Formulas(2) = "DIRECCION=""" & EmpOC.Direccion & """"
cr.Formulas(3) = "TELEFONOS=""" & "Teléfono: " & EmpOC.Telefono1 & " " & EmpOC.Comuna & """"
cr.Formulas(4) = "RUT=""" & "RUT: " & EmpOC.Rut & """"

'MsgBox Certificado.Value

cr.DataFiles(0) = repo_file & ".MDB"

'If Tipo = "E" Then
'    CR.Formulas(5) = "COTIZACION=""" & "GUÍA DESP. Nº:" & """"
'    CR.ReportFileName = Drive_Server & Path_Server & EmpOC.Fantasia & "\Oc_Especial.Rpt"
'Else
'    CR.Formulas(5) = "COTIZACION=""" & "COTIZACIÓN Nº:" & """"
    cr.ReportFileName = Drive_Server & Path_Rpt & "itopg_legal.rpt"
'End If

cr.Action = 1

End Sub
Private Sub Privilegios()
If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 0, 0
End If
End Sub
Private Sub Campos_Cabecera_Definir()
ac(1, 0) = "tipo"
ac(1, 1) = "'"
ac(2, 0) = "numero"
ac(2, 1) = ""
ac(3, 0) = "fecha"
ac(3, 1) = "'"
ac(4, 0) = "nv"
ac(4, 1) = ""
ac(5, 0) = "rut" ' contratista
ac(5, 1) = "'"
ac(6, 0) = "rut2" ' inspector
ac(6, 1) = "'"
ac(7, 0) = "texto2" ' defecto
ac(7, 1) = "'"
End Sub
Private Sub Campos_Detalle_Definir()

ac(1, 0) = "tipo"
ac(1, 1) = "'"
ac(2, 0) = "numero"
ac(2, 1) = ""
ac(3, 0) = "linea"
ac(3, 1) = ""
ac(4, 0) = "fecha"
ac(4, 1) = "'"
ac(5, 0) = "nv"
ac(5, 1) = ""
ac(6, 0) = "rut" ' contratista
ac(6, 1) = "'"
ac(7, 0) = "rut2" ' inspector
ac(7, 1) = "'"
ac(8, 0) = "plano"
ac(8, 1) = "'"
ac(9, 0) = "rev"
ac(9, 1) = "'"
ac(10, 0) = "marca"
ac(10, 1) = "'"
ac(11, 0) = "cantidad"
ac(11, 1) = ""

End Sub
