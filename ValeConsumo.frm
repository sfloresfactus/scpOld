VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ValeConsumo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vale de Consumo"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ingresar Nuevo Vale"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar Vale"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar Vale"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Vale"
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
            Object.ToolTipText     =   "Grabar Vale"
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
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mantención de Productos"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox Nv 
      Height          =   300
      Left            =   2400
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox CbTrabajadores 
      Height          =   315
      Left            =   7320
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Numero 
      Height          =   300
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox ComboNV 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Frame Frame 
      Height          =   1455
      Left            =   3360
      TabIndex        =   9
      Top             =   480
      Width           =   5655
      Begin VB.ComboBox CbSeccion 
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Opcion 
         Caption         =   "N Trabajadores"
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   13
         Top             =   150
         Width           =   1575
      End
      Begin VB.OptionButton Opcion 
         Caption         =   "1 Trabajador"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   12
         Top             =   150
         Width           =   1335
      End
      Begin VB.OptionButton Opcion 
         Caption         =   "Contratista"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   11
         Top             =   150
         Width           =   1095
      End
      Begin VB.CommandButton btnDetalleTrabajador 
         Caption         =   "Detalle Trabajador"
         Height          =   495
         Left            =   2760
         TabIndex        =   17
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton btnSearch 
         Height          =   300
         Left            =   2400
         Picture         =   "ValeConsumo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   480
         Width           =   300
      End
      Begin VB.TextBox Razon 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   5415
      End
      Begin MSMask.MaskEdBox Rut 
         Height          =   300
         Left            =   960
         TabIndex        =   15
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label lblSeccion 
         Caption         =   "Sección"
         Height          =   255
         Left            =   3840
         TabIndex        =   18
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   10
         Top             =   150
         Width           =   615
      End
      Begin VB.Label lbl 
         Caption         =   "RUT"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lbl 
         Caption         =   "Señor(es)"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.TextBox txtEditOT 
      Height          =   285
      Left            =   8040
      TabIndex        =   22
      Text            =   "txtEditOT"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   960
      TabIndex        =   4
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   327680
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid Detalle 
      Height          =   2445
      Left            =   120
      TabIndex        =   21
      Top             =   2040
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4313
      _Version        =   327680
      ScrollBars      =   2
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
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ValeConsumo.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ValeConsumo.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ValeConsumo.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ValeConsumo.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ValeConsumo.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ValeConsumo.frx":065C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ValeConsumo.frx":076E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ValeConsumo.frx":0A88
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "NV"
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   5
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label TotalPrecio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6720
      TabIndex        =   25
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "OBRA"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "FECHA"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "Vale de Consumo"
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
      TabIndex        =   0
      Top             =   480
      Width           =   2175
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
      TabIndex        =   1
      Top             =   840
      Width           =   375
   End
   Begin VB.Menu MenuPop 
      Caption         =   "MenuPop"
      Visible         =   0   'False
      Begin VB.Menu FilaInsertar 
         Caption         =   "&Insertar Fila"
      End
      Begin VB.Menu FilaEliminar 
         Caption         =   "&Eliminar Fila"
      End
      Begin VB.Menu FilaBorrarContenido 
         Caption         =   "&Borrar Contenido"
      End
   End
End
Attribute VB_Name = "ValeConsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button

Private DbD As Database
', RsSc As Recordset
'Private SqlRsSc As New ADODB.Recordset
Private RsTra As Recordset, RsTraSec As Recordset, RsPrd As Recordset
Private Dbm As Database, RsVc As Recordset, RsNVc As Recordset
Private DbAdq As Database, RsOCd As Recordset

Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, d As Variant

Private n_filas As Integer, n_columnas As Integer
Private prt As Printer, TipoDoc As String
Private n3 As Double, n7 As Double
Private linea As String, m_Nv As Double
Private a_Seccion(1, 29) As String, a_Trabajadores(1, 299) As String, MaximoDeTrabajadores As Integer
Private Detalle_Ancho As Integer, Col9_Ancho As Integer
Private a_Nv(2999, 1) As String, m_NvArea As Integer
Private Sub btnDetalleTrabajador_Click()
TrabajadorMostrar.Rut = Rut.Text
TrabajadorMostrar.Show 1
End Sub
Private Sub CbTrabajadores_Click()
Detalle = CbTrabajadores.Text
CbTrabajadores.visible = False
End Sub
Private Sub Form_Load()

MaximoDeTrabajadores = 299

' abre archivos
Set DbD = OpenDatabase(data_file)
'Set RsSc = DbD.OpenRecordset("Contratistas")
'RsSc.Index = "RUT"
Set RsTra = DbD.OpenRecordset("Trabajadores")
RsTra.Index = "RUT"
Set RsPrd = DbD.OpenRecordset("Productos")
RsPrd.Index = "Codigo"

Set Dbm = OpenDatabase(mpro_file)

Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set DbAdq = OpenDatabase(Madq_file)

' ojo correccion para livio figueroa, porque rut lo esta grabando con sin espacio
'DbAdq.Execute "update documentos set rut=' 9379547-4' where tipo='VC' AND rut = '9379547-4'"


Set RsVc = DbAdq.OpenRecordset("Documentos")
RsVc.Index = "Tipo-Numero-Linea"

Set RsOCd = DbAdq.OpenRecordset("OC Detalle")
RsOCd.Index = "CodigoPrd-Fecha"

' Combo obra
i = 0
ComboNV.AddItem " "
Do While Not RsNVc.EOF
    If Usuario.Nv_Activas Then
        If RsNVc!Activa Then
            i = i + 1
            a_Nv(i, 0) = RsNVc!Numero
            a_Nv(i, 1) = RsNVc!Obra
            ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra
        End If
    Else
        i = i + 1
        a_Nv(i, 0) = RsNVc!Numero
        a_Nv(i, 1) = RsNVc!Obra
        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra
    End If
    RsNVc.MoveNext
Loop

a_Seccion(0, 1) = "ADQ"
a_Seccion(1, 1) = "Adquisicion"
a_Seccion(0, 2) = "ARS"
a_Seccion(1, 2) = "Arco Sumergido"
a_Seccion(0, 3) = "BOD"
a_Seccion(1, 3) = "Bodega"
a_Seccion(0, 4) = "CON"
a_Seccion(1, 4) = "Contabilidad"
a_Seccion(0, 5) = "CHO"
a_Seccion(1, 5) = "Chofer"
a_Seccion(0, 6) = "DES"
a_Seccion(1, 6) = "Despacho"
a_Seccion(0, 7) = "GER"
a_Seccion(1, 7) = "Gerencia"
a_Seccion(0, 8) = "GRA"
a_Seccion(1, 8) = "Granalla"
a_Seccion(0, 9) = "GRU"
a_Seccion(1, 9) = "Gruero"
a_Seccion(0, 10) = "GUI"
a_Seccion(1, 10) = "Guillotina"
a_Seccion(0, 11) = "INS"
a_Seccion(1, 11) = "Inspección"
a_Seccion(0, 12) = "MEL"
a_Seccion(1, 12) = "Mantencion Electrica"
a_Seccion(0, 13) = "MON"
a_Seccion(1, 13) = "Montaje"
a_Seccion(0, 14) = "OPE"
a_Seccion(1, 14) = "Operaciones"
a_Seccion(0, 15) = "OXI"
a_Seccion(1, 15) = "Oxicorte"
a_Seccion(0, 16) = "PPL"
a_Seccion(1, 16) = "Patio Plancha"
a_Seccion(0, 17) = "PIN"
a_Seccion(1, 17) = "Pintura"
a_Seccion(0, 18) = "PLA"
a_Seccion(1, 18) = "Plasma"
a_Seccion(0, 19) = "PMA"
a_Seccion(1, 19) = "Prep. Material"
a_Seccion(0, 20) = "PRE"
a_Seccion(1, 20) = "Prevensión"
a_Seccion(0, 21) = "PRO"
a_Seccion(1, 21) = "Producción"

'CbSeccion.AddItem " "
CbSeccion.AddItem "Todas"
For i = 1 To 21
    CbSeccion.AddItem a_Seccion(1, i)
Next

Inicializa

Detalle_Config

If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 0, 0
End If

Opcion(0).Value = True
btnDetalleTrabajador.Enabled = False

CbTrabajadores.visible = False

m_NvArea = 0

Exit Sub

' busca vales de consumo digitados desde la segunda linea
With RsVc
.Seek ">=", "VC", 0
Dim nvale As Double
nvale = 0
Do While Not .EOF
    If !Tipo <> "VC" Then Exit Do
    If nvale <> !Numero Then
        If !linea = 1 Then
            ' linea ok
        Else
            Debug.Print !Numero, !Fecha
        End If
        nvale = !Numero
    End If
    .MoveNext
Loop
End With

End Sub
Private Sub Inicializa()

Obj = "VALE DE CONSUMO"
Objs = "VALES DE CONSUMO"

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

TipoDoc = "VC"

End Sub
Private Sub Detalle_Config()
Dim i As Integer

n_filas = 25
' col9  trabajador
' col10 largo especial ?
n_columnas = 10
Col9_Ancho = 2000

Detalle.Left = 100
Detalle.WordWrap = True
'Detalle.RowHeight(0) = 450
Detalle.Cols = n_columnas + 1
Detalle.Rows = n_filas + 1

'Detalle.TextMatrix(0, 0) = ""
'Detalle.TextMatrix(0, 1) = "Código"
'Detalle.TextMatrix(0, 2) = "Descripción"
'Detalle.TextMatrix(0, 3) = "Cantidad"
'Detalle.TextMatrix(0, 4) = "$ Uni"
'Detalle.TextMatrix(0, 5) = "Total"

'Detalle.ColWidth(0) = 250
'Detalle.ColWidth(1) = 1500
'Detalle.ColWidth(2) = 3500
'Detalle.ColWidth(3) = 1000
'Detalle.ColWidth(4) = 1000
'Detalle.ColWidth(5) = 1000

Detalle.TextMatrix(0, 0) = ""
Detalle.TextMatrix(0, 1) = "Código"
Detalle.TextMatrix(0, 2) = "L"
Detalle.TextMatrix(0, 3) = "Cantidad"
Detalle.TextMatrix(0, 4) = "Uni"
Detalle.TextMatrix(0, 5) = "Descripción"
Detalle.TextMatrix(0, 6) = "Largo(mm)"
Detalle.TextMatrix(0, 7) = "$ Uni"
Detalle.TextMatrix(0, 8) = "Total"
Detalle.TextMatrix(0, 9) = "Trabajador"

Detalle.ColWidth(0) = 250
Detalle.ColWidth(1) = 1500
Detalle.ColWidth(2) = 300
Detalle.ColWidth(3) = 800
Detalle.ColWidth(4) = 500
Detalle.ColWidth(5) = 3000
Detalle.ColWidth(6) = 950
Detalle.ColWidth(7) = 950
Detalle.ColWidth(8) = 950
Detalle.ColWidth(9) = Col9_Ancho
Detalle.ColWidth(10) = 0

'Detalle.ColAlignment(2) = 0

Detalle_Ancho = 350 ' con scroll vertical
'ancho = 100 ' sin scroll vertical

TotalPrecio.Width = Detalle.ColWidth(5)
For i = 0 To n_columnas
    If i = 8 Then
        TotalPrecio.Left = Detalle_Ancho + Detalle.Left - 300
        TotalPrecio.Width = Detalle.ColWidth(8)
    End If
    Detalle_Ancho = Detalle_Ancho + Detalle.ColWidth(i)
Next

Detalle.Width = Detalle_Ancho
Me.Width = Detalle_Ancho + 50 + Detalle.Left * 2

' col y row fijas
'Detalle.BackColorFixed = vbCyan

' establece colores a columnas
' columnas    modificables : NEGRAS
' columnas no modificables : ROJAS

For i = 1 To n_filas

    Detalle.TextMatrix(i, 0) = i
    Detalle.Row = i
    Detalle.col = 1
    Detalle.CellAlignment = flexAlignLeftCenter
    Detalle.col = 2
    Detalle.CellForeColor = vbRed
    Detalle.CellAlignment = flexAlignLeftCenter
'    Detalle.col = 3
    Detalle.col = 4
    Detalle.CellForeColor = vbRed
    Detalle.col = 5
    Detalle.CellForeColor = vbRed
    Detalle.CellAlignment = flexAlignLeftCenter
    
Next

txtEditOT.Text = ""

Detalle.ColWidth(9) = 0

'Detalle.ScrollTrack = True 'no se nota
'Detalle.TextMatrix(1, 1) = "hola" 'ok

End Sub
Private Sub ComboNV_Click()

MousePointer = vbHourglass

i = 0
m_Nv = Val(Left(ComboNV.Text, 6))
If m_Nv > 0 Then Nv.Text = m_Nv

MousePointer = vbDefault

End Sub
Private Sub Fecha_GotFocus()
'
End Sub
Private Sub Fecha_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Fecha_LostFocus()
d = Fecha_Valida(Fecha, Now)
End Sub
Private Sub CbSeccion_Click()
' cambia el combobox de trabajadores de acuerdo a seccion
Dim i As Integer, li As Integer, m_Nombre As String
i = CbSeccion.ListIndex
'MsgBox i & a_Seccion(0, i)

Select Case i
Case -1
Case 0
    ' todas las secciones
    a_Trabajadores_Limpiar
    CbTrabajadores.Clear
    Set RsTraSec = DbD.OpenRecordset("SELECT * FROM trabajadores WHERE activo ORDER BY appaterno")
    With RsTraSec
    li = 0
    Do While Not .EOF
        li = li + 1
        a_Trabajadores(0, li) = !Rut
        m_Nombre = !nombres & " " & !appaterno & " " & !apmaterno
        a_Trabajadores(1, li) = m_Nombre
        CbTrabajadores.AddItem m_Nombre
'        Debug.Print !nombres, !appaterno, !apmaterno
        .MoveNext
    Loop
    RsTraSec.Close
    End With
Case Else
    a_Trabajadores_Limpiar
    CbTrabajadores.Clear
    Set RsTraSec = DbD.OpenRecordset("SELECT * FROM trabajadores WHERE clase1='" & a_Seccion(0, i) & "' AND activo ORDER BY appaterno")
    With RsTraSec
    li = 0
    Do While Not .EOF
        li = li + 1
        a_Trabajadores(0, li) = !Rut
        m_Nombre = !nombres & " " & !appaterno & " " & !apmaterno
        a_Trabajadores(1, li) = m_Nombre
        CbTrabajadores.AddItem m_Nombre
'        Debug.Print !nombres, !appaterno, !apmaterno
        .MoveNext
    Loop
    RsTraSec.Close
    End With
End Select

End Sub
Private Sub a_Trabajadores_Limpiar()
' limoia arreglo de trabajadores
For i = 0 To MaximoDeTrabajadores
    a_Trabajadores(0, i) = ""
    a_Trabajadores(1, i) = ""
Next
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
   
    RsVc.Seek ">=", TipoDoc, Numero.Text, 0
'    RsVc.Seek "=", TipoDoc, Numero.Text, 1
'    If RsVc.NoMatch Then
    If RsVc.EOF Then
        GoTo Agregar
    End If
    If RsVc.NoMatch Then
        GoTo Agregar
    End If
    If TipoDoc <> RsVc!Tipo Or Numero.Text <> RsVc!Numero Then
Agregar:
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
'    End If
    
Case "Modificando"

    RsVc.Seek ">=", TipoDoc, Numero.Text, 0
'    RsVc.Seek "=", TipoDoc, Numero.Text, 1
'    If RsVc.NoMatch Then
    If TipoDoc <> RsVc!Tipo Or Numero.Text <> RsVc!Numero Then
        MsgBox Obj & " NO EXISTE"
    Else
        Botones_Enabled 0, 0, 0, 0, 1, 1
        Doc_Leer
        Campos_Enabled True
        Numero.Enabled = False
        
    End If

Case "Eliminando"

    RsVc.Seek ">=", TipoDoc, Numero.Text, 0
'    RsVc.Seek "=", TipoDoc, Numero.Text, 1
'    If RsVc.NoMatch Then
    If TipoDoc <> RsVc!Tipo Or Numero.Text <> RsVc!Numero Then
        MsgBox Obj & " NO EXISTE"
    Else
        Doc_Leer
'        Numero.Enabled = False
        If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo) = vbYes Then
            Doc_Eliminar
        End If
        Campos_Limpiar
        Campos_Enabled False
        Numero.Enabled = True
        Numero.SetFocus
    End If
   
Case "Imprimiendo"
    
    RsVc.Seek ">=", TipoDoc, Numero.Text, 0
'    RsVc.Seek "=", TipoDoc, Numero.Text, 1
'    If RsVc.NoMatch Then
    If TipoDoc <> RsVc!Tipo Or Numero.Text <> RsVc!Numero Then
        MsgBox Obj & " NO EXISTE"
    Else
        Doc_Leer
        Numero.Enabled = False
        Detalle.Enabled = True
    End If
    
End Select

End Sub
Private Sub Doc_Leer()

Dim m_resta As Integer, m_LargoE As Boolean, m_Tipo As String, m_Rut_Guia As String, m_Ruts_Distintos As Boolean, primera As Boolean

m_Tipo = ""
' CABECERA
Fecha.Text = Format(RsVc!Fecha, Fecha_Format)
m_Nv = RsVc!Nv
Rut.Text = NoNulo(RsVc![Rut])

RsNVc.Seek "=", m_Nv, m_NvArea
If Not RsNVc.NoMatch Then
    On Error GoTo NoNv
    ComboNV.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra
    GoTo Sigue
NoNv:
    ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra
    ComboNV.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra
Sigue:
    On Error GoTo 0
'    ComboNV_Click
End If

With RsVc
.Seek ">=", TipoDoc, Numero.Text, 0

i = 0
primera = True
m_Ruts_Distintos = False

If Not .NoMatch Then

    If primera Then
        CbSeccion.ListIndex = 0
        m_Rut_Guia = NoNulo(!Rut)
        primera = False
    End If

    Do While Not .EOF
    
        If TipoDoc = !Tipo And Numero.Text = !Numero Then
        
            i = !linea
            
            m_Tipo = NoNulo(!TipoNE)
            
            m_LargoE = False
            Detalle.TextMatrix(i, 1) = ![codigo producto]
            RsPrd.Seek "=", ![codigo producto]
            If Not RsPrd.NoMatch Then
                Detalle.TextMatrix(i, 4) = RsPrd![unidad de medida]
                Detalle.TextMatrix(i, 5) = RsPrd!Descripcion
                m_LargoE = RsPrd![Largo Especial]
            End If
            
            Detalle.TextMatrix(i, 2) = NoNulo(![Largo Especial])
            Detalle.TextMatrix(i, 3) = !Cant_Sale
            Detalle.TextMatrix(i, 6) = ![largo]
            Detalle.TextMatrix(i, 7) = ![Precio Unitario]
            Detalle.TextMatrix(i, 10) = m_LargoE
            
            n3 = m_CDbl(Detalle.TextMatrix(i, 3))
            n7 = m_CDbl(Detalle.TextMatrix(i, 7))
            
            Detalle.TextMatrix(i, 8) = Format(n3 * n7, num_Formato)

            For j = 0 To MaximoDeTrabajadores
                If a_Trabajadores(0, j) = !Rut Then
                    Detalle.TextMatrix(i, 9) = a_Trabajadores(1, j)
                    Exit For
                End If
            Next

            If !Rut <> m_Rut_Guia Then
                m_Ruts_Distintos = True
            End If

        Else
        
            Exit Do
            
        End If
        
        .MoveNext
        
    Loop
    
End If

End With

If m_Tipo = "T" Then

    If m_Ruts_Distintos Then
        ' varias lineas distintos rut
        Opcion(2).Value = True
'        Detalle.Width = Detalle_Ancho + Col9_Ancho
'        Detalle.ColWidth(9) = Col9_Ancho
        
    Else
        Trabajador_Lee Rut.Text
    End If

Else
    Contratista_Lee Rut.Text
End If

Detalle.Row = 1 ' para q' actualice la primera fila del detalle

Actualiza

End Sub
Private Sub Contratista_Lee(Rut)

'RsSc.Seek "=", Rut
'If Not RsSc.NoMatch Then
'    Razon.Text = RsSc![Razon Social]
'End If

Rut = Trim(Rut)
SqlRsSc.Open "SELECT * FROM personas WHERE contratista='S' AND rut='" & Rut & "'", CnxSqlServer_scp0
If SqlRsSc.EOF Then
    Razon.Text = "NO Encontrado"
Else
    Razon.Text = SqlRsSc![razon_social]
End If
SqlRsSc.Close

End Sub
Private Sub Trabajador_Lee(pRut)
RsTra.Seek "=", pRut
If Not RsTra.NoMatch Then
    Opcion(1).Value = True
    Rut.Text = pRut
    Razon.Text = RsTra!appaterno & " " & RsTra!apmaterno & " " & RsTra!nombres
    btnDetalleTrabajador.visible = True
    btnDetalleTrabajador.Enabled = True
    btnSearch.visible = True
'    Direccion.Text = RsSc!Dirección
'    Comuna.Text = NoNulo(RsSc!Comuna)
End If
End Sub
Private Function Doc_Validar() As Boolean
Doc_Validar = False

If m_Nv = 0 Then
    MsgBox "DEBE ELEGIR OBRA"
    ComboNV.SetFocus
    Exit Function
End If

Select Case True
Case Opcion(0)
    If Rut.Text = "" Then
        MsgBox "DEBE ELEGIR CONTRATISTA"
        btnSearch.SetFocus
        Exit Function
    End If
Case Opcion(1)
    If Rut.Text = "" Then
        MsgBox "DEBE ELEGIR TRABAJADOR"
        btnSearch.SetFocus
        Exit Function
    End If
Case Opcion(2)
    ' n trabajadores
End Select

For i = 1 To n_filas

    ' codigo prod
    If Trim(Detalle.TextMatrix(i, 1)) <> "" Then
    
        ' cantidad 3
        If Not CampoReq_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
        
        Detalle.Row = i
        
        ' precio unitario
        If Not Numero_Valida(Detalle.TextMatrix(i, 7), i, 7) Then Exit Function
        
        If Opcion(2).Value = True Then
            ' trabajador x linea
            If Not CampoReq_Valida(Detalle.TextMatrix(i, 9), i, 9) Then Exit Function
        End If
    
        
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

save:

Detalle:
' DETALLE DE OT
Doc_Detalle_Eliminar

' graba detalle
With RsVc

For i = 1 To n_filas

    If Trim(Detalle.TextMatrix(i, 1)) <> "" Then
            
        .AddNew
        
        !Tipo = TipoDoc
        !Numero = Numero.Text
        !linea = i
        !Fecha = Fecha.Text
        !Nv = m_Nv
        
        If Rut.Text <> "" Then
        
            ![Rut] = SqlRutPadL(Rut.Text)
            
        Else
            ' n trabajadores, 1 x linea
            ' busca rut del trabajador
            For j = 0 To MaximoDeTrabajadores
                If a_Trabajadores(1, j) = Detalle.TextMatrix(i, 9) Then
                    ![Rut] = a_Trabajadores(0, j)
                    Exit For
                End If
            Next
            
        End If
        
        ![codigo producto] = Detalle.TextMatrix(i, 1)
        ![Cant_Sale] = Detalle.TextMatrix(i, 3)
        ![Largo Especial] = Trim(Detalle.TextMatrix(i, 2))
        ![largo] = Val(Detalle.TextMatrix(i, 6))
        ![Precio Unitario] = Detalle.TextMatrix(i, 7)
        
        If Opcion(0).Value = True Then
            !TipoNE = " "
        Else
            !TipoNE = "T" ' trabajador
        End If

        .Update
        
    End If
    
Next

End With

End Sub
Private Sub Doc_Eliminar()

' borra CABECERA DE OT
'RsVc.Seek "=", Numero.Text
'If Not RsVc.NoMatch Then

'    RsVc.Delete

'End If

Doc_Detalle_Eliminar

End Sub
Private Sub Doc_Detalle_Eliminar()

DbAdq.Execute "DELETE * FROM Documentos WHERE tipo='VC' AND numero=" & Numero.Text

'RsVc.Seek "=", Numero.Text, 1
'If Not RsVc.NoMatch Then
'    Do While Not RsVc.EOF
'
'        ' borra detalle
'        RsVc.Delete
'
'        RsVc.MoveNext
'
'    Loop
'End If

End Sub
Private Sub Campos_Limpiar()
Numero.Text = ""
Fecha.Text = Fecha_Vacia
Nv.Text = ""
ComboNV.Text = " "

Opcion(0).Value = True ' la mayoria de las veces va a ser contratista

Rut.Text = ""
Razon.Text = ""
Opcion(0).Value = True

CbSeccion.Text = "Todas"

btnDetalleTrabajador.visible = False
'Direccion.Text = ""
'Comuna.Text = ""
Detalle_Limpiar
'Obs(0).Text = ""
'Obs(1).Text = ""
'Obs(2).Text = ""
'Obs(3).Text = ""
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

Private Sub Opcion_Click(Index As Integer)
Select Case Index
Case 0
    ' contratista
    btnSearch.visible = True
    btnDetalleTrabajador.visible = False
    btnSearch.ToolTipText = "Busca Contratista"
'    CbSeccion.ListIndex = -1
'    CbSeccion.Enabled = False
    lblSeccion.visible = False
    CbSeccion.visible = False
    Detalle.Width = Detalle_Ancho - Col9_Ancho
    Detalle.ColWidth(9) = 0
Case 1
    ' 1 trabajador
    btnSearch.visible = True
    btnDetalleTrabajador.visible = False
    btnSearch.ToolTipText = "Busca Trabajador"
'    CbSeccion.ListIndex = -1
'    CbSeccion.Enabled = False
    
    lblSeccion.visible = False
    CbSeccion.visible = False
    
    Detalle.Width = Detalle_Ancho - Col9_Ancho
    Detalle.ColWidth(9) = 0
Case Else
    ' N trabajadores
    btnSearch.visible = False
    btnDetalleTrabajador.visible = False
    btnSearch.ToolTipText = ""
'    CbSeccion.Enabled = True
    lblSeccion.visible = True
    CbSeccion.visible = True
    Detalle.Width = Detalle_Ancho
    Detalle.ColWidth(9) = Col9_Ancho
End Select
Rut.Text = ""
Razon.Text = ""
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim cambia_titulo As Boolean, n_Copias As Integer

cambia_titulo = True
'Accion = "" rem accion
Select Case Button.Index
Case 1 ' Agregar
    Accion = "Agregando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    
    Numero.Text = Documento_Numero_Nuevo(RsVc, "Numero")
    
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
            Doc_Imprimir n_Copias
        End If

        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
        
    End If
Case 6 ' DesHacer
    If Numero.Text = "" Then
        Botones_Enabled 1, 1, 1, 1, 0, 0
        Campos_Limpiar
        Campos_Enabled False
    Else
        If Accion = "Imprimiendo" Then
            Botones_Enabled 1, 1, 1, 1, 0, 0
            Campos_Limpiar
            Campos_Enabled False
        Else
            If MsgBox("¿Seguro que quiere DesHacer?", vbYesNo) = vbYes Then
                Botones_Enabled 1, 1, 1, 1, 0, 0
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
                Doc_Imprimir n_Copias
            End If
            
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
Case 10 ' Productos
    MousePointer = 11
    Load Productos
    MousePointer = 0
    Productos.Show 1
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

Opcion(0).Enabled = Si
Opcion(1).Enabled = Si
Opcion(2).Enabled = Si

btnDetalleTrabajador.Enabled = Si

lblSeccion.Enabled = Si
CbSeccion.Enabled = Si

'If Si Then btnDetalleTrabajador.Visible = False
Nv.Enabled = Si
ComboNV.Enabled = Si
Detalle.Enabled = Si
'Obs(0).Enabled = Si
'Obs(1).Enabled = Si
'Obs(2).Enabled = Si
'Obs(3).Enabled = Si
End Sub
Private Sub btnSearch_Click()

If Opcion(0).Value = True Then

'    Search.Muestra data_file, "Contratistas", "RUT", "Razon Social", "Contratista", "Contratistas", "Activo"
'    Rut.Text = Search.codigo
'    If Rut.Text <> "" Then
'        RsSc.Seek "=", Rut
'        If RsSc.NoMatch Then
'            MsgBox "CONTRATISTA NO EXISTE"
'            Rut.SetFocus
'        Else
'            Razon.Text = Search.descripcion
'        End If
'    End If
        
    Dim arreglo(1) As String
    arreglo(1) = "razon_social"

    sql_Search.Muestra "personas", "RUT", arreglo(), Obj, Objs, "contratista='S'"
    Rut.Text = sql_Search.Codigo
    Razon.Text = sql_Search.Descripcion
    

Else

    Search.Muestra data_file, "Trabajadores", "RUT", "ApPaterno]+' '+[ApMaterno]+' '+[Nombres", "Trabajador", "Trabajadores" ', "Activo"
    
    Rut.Text = Search.Codigo
    If Rut.Text <> "" Then
        RsTra.Seek "=", Rut
        If RsTra.NoMatch Then
            MsgBox "TRABAJADOR NO EXISTE"
            Rut.SetFocus
        Else
            Razon.Text = Search.Descripcion
            btnDetalleTrabajador.Enabled = True
            btnDetalleTrabajador.visible = True
    '        Direccion.Text = RsSc!Dirección
    '        Comuna.Text = NoNulo(RsSc!Comuna)
        End If
    End If

End If

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
'        If Detalle <> "" Then ComboPlano.Text = Detalle
'        ComboPlano.Top = Detalle.CellTop + Detalle.Top
'        ComboPlano.Left = Detalle.CellLeft + Detalle.Left
'        ComboPlano.Width = Int(Detalle.CellWidth * 1.5)
'        ComboPlano.Visible = True
'        ComboMarca.Visible = False
    Case 3 ' marca
'        ComboMarca_Poblar Detalle.TextMatrix(Detalle.Row, 1)
        On Error GoTo Error
'        If Detalle <> "" Then ComboMarca.Text = Detalle
Error:
        On Error GoTo 0
'        ComboMarca.Text = ""
'        ComboMarca.Top = Detalle.CellTop + Detalle.Top
'        ComboMarca.Left = Detalle.CellLeft + Detalle.Left
'        ComboMarca.Width = Int(Detalle.CellWidth * 1.5)
'        ComboPlano.Visible = False
'        ComboMarca.Visible = True
        
    Case 10 ' fecha de entrega
    Case Else
'        ComboPlano.Visible = False
'        ComboMarca.Visible = False
End Select
End Sub
Private Sub Detalle_DblClick()
If Accion = "Imprimiendo" Then Exit Sub
' simula un espacio
'If Detalle.col = 10 Then
'    MSFlexGridEdit Detalle, EditFecha, 32  'FECHA
'Else
    MSFlexGridEdit Detalle, txtEditOT, 32
'End If
End Sub
Private Sub Detalle_GotFocus()
If Accion = "Imprimiendo" Then Exit Sub
Select Case True
Case txtEditOT.visible
    Detalle = txtEditOT
    txtEditOT.visible = False
'Case EditFecha.Visible
'    Detalle = EditFecha
'    EditFecha.Visible = False
End Select
End Sub
Private Sub Detalle_LeaveCell()
If Accion = "Imprimiendo" Then Exit Sub
Select Case True
Case txtEditOT.visible
    Detalle = txtEditOT
    txtEditOT.visible = False
'Case EditFecha.Visible
'    Detalle = EditFecha
'    EditFecha.Visible = False
End Select
End Sub
Private Sub Detalle_KeyPress(KeyAscii As Integer)
If Accion = "Imprimiendo" Then Exit Sub
If Detalle.col = 10 Then
'    MSFlexGridEdit Detalle, EditFecha, KeyAscii 'fecha
Else
    MSFlexGridEdit Detalle, txtEditOT, KeyAscii
End If
End Sub
Private Sub txtEditOT_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCodeP Detalle, txtEditOT, KeyCode, Shift
End Sub
Private Sub txtEditOT_LostFocus()
'txtEditOT.Visible = False 07/03/98
'EditKeyCodeP Detalle, txtEditOT, vbkeyreturn, 0
' ó
'Detalle.SetFocus
'DoEvents
'Actualiza
End Sub
Sub EditKeyCodeP(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
' rutina que se ejecuta con los keydown de Edt
Dim m_fil As Integer, m_col As Integer
m_fil = MSFlexGrid.Row
m_col = MSFlexGrid.col

Dim dif As Integer
'dif = Val(Detalle.TextMatrix(MSFlexGrid.Row, 5)) - Val(Detalle.TextMatrix(MSFlexGrid.Row, 6))
Select Case KeyCode
Case vbKeyEscape ' Esc
    Edt.visible = False
    MSFlexGrid.SetFocus
Case vbKeyReturn ' Enter
    Select Case m_col
    Case 1 ' Codigo del Producto
    
        ' busca codigo
        
        MSFlexGrid.SetFocus
        DoEvents
        Linea_Actualiza
        
        Codigo2Descripcion Detalle.TextMatrix(m_fil, 1)
        
    Case 10 ' Fecha
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
    Case 10 ' Fecha
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
Private Sub txtEditOT_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)
Dim m_Codigo As String

Select Case MSFlexGrid.col
Case 1 'codigo
    Edt.MaxLength = 15
Case 2 'largo especial
    Edt.MaxLength = 1
Case 4, 5, 8 'unidad, desc, total
'    no editables
Case Else
    Edt.MaxLength = 10
End Select

Select Case MSFlexGrid.col
Case 2 'largo especial
'    m_Codigo = MSFlexGrid.TextMatrix(MSFlexGrid.Row,10)
    If Not CBool(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 10)) Then Exit Sub
    GoTo Edita
Case 6
    'largo
    'si y no editable
    m_Codigo = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 2)
    If m_Codigo <> "E" Then Exit Sub
    GoTo Edita
Case 4, 5, 8
    ' no editables
    Exit Sub
Case 9 ' combo trabajador
    CbTrabajadores.Left = MSFlexGrid.CellLeft + MSFlexGrid.Left
    CbTrabajadores.Top = MSFlexGrid.CellTop + MSFlexGrid.Top
    CbTrabajadores.visible = True
    CbTrabajadores.SetFocus
Case Else
Edita:
    Select Case KeyAscii
    Case 0 To 32
        Edt = MSFlexGrid
        Edt.SelStart = 1000
    Case Else
        Edt = Chr(KeyAscii)
        Edt.SelStart = 1
    End Select
    Edt.Move MSFlexGrid.CellLeft + MSFlexGrid.Left, MSFlexGrid.CellTop + MSFlexGrid.Top, MSFlexGrid.CellWidth, MSFlexGrid.CellHeight * 1.2
    Edt.visible = True
    Edt.SetFocus
    'opGrabar True
End Select

Exit Sub

Select Case MSFlexGrid.col
Case 4, 5, 8
    ' no editables
    Exit Sub
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

Select Case KeyCode
Case vbKeyF1
    If Detalle.col = 1 Then CodigoProducto_Buscar
Case vbKeyF2
    MSFlexGridEdit Detalle, txtEditOT, 32
End Select
End Sub
Private Sub Linea_Actualiza()
' actualiza solo linea, y totales generales
Dim fi As Integer, co As Integer

fi = Detalle.Row
co = Detalle.col

n3 = m_CDbl(Detalle.TextMatrix(fi, 3))
n7 = m_CDbl(Detalle.TextMatrix(fi, 7))

' precio total
Detalle.TextMatrix(fi, 8) = Format(n3 * n7, num_fmtgrl)

Totales_Actualiza

End Sub
Private Sub Actualiza()
' actualiza todo el detalle y totales generales
Dim fi As Integer
For fi = 1 To n_filas
    If Detalle.TextMatrix(fi, 1) <> "" Then
    
        n3 = m_CDbl(Detalle.TextMatrix(fi, 3))
        n7 = m_CDbl(Detalle.TextMatrix(fi, 7))

        ' precio total
        Detalle.TextMatrix(fi, 8) = Format(n3 * n7, num_fmtgrl)
        
    End If
Next

Totales_Actualiza

End Sub
Private Sub Totales_Actualiza()
Dim Tot_Precio As Double
Tot_Precio = 0
For i = 1 To n_filas
    Tot_Precio = Tot_Precio + m_CDbl(Detalle.TextMatrix(i, 8))
Next

TotalPrecio.Caption = Format(Tot_Precio, num_Format0)

End Sub
Private Sub Cursor_Mueve(MSFlexGrid As Control)
'MIA
Select Case MSFlexGrid.col
Case 4
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
Private Sub Detalle_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = 2 Then
    If Detalle.ColSel = 1 And Detalle.col = 1 Then
        'como F1
        CodigoProducto_Buscar
    End If
    If Detalle.ColSel = n_columnas And Detalle.col = 1 Then
        PopupMenu MenuPop
    End If
End If

End Sub
Private Sub FilaInsertar_Click()
' inserta fila en FlexGrid
Dim fi As Integer, co As Integer, fi_ini As Integer
fi_ini = Detalle.Row
For fi = n_filas To fi_ini + 1 Step -1
    For co = 1 To n_columnas
        Detalle.TextMatrix(fi, co) = Detalle.TextMatrix(fi - 1, co)
    Next
Next
' fila nueva
For co = 1 To n_columnas
    Detalle.TextMatrix(fi_ini, co) = ""
Next
Detalle.col = 1
Detalle.Row = fi_ini
End Sub
Private Sub FilaEliminar_Click()
' elimina fila en FlexGrid
Dim fi As Integer, co As Integer, fi_ini As Integer
fi_ini = Detalle.Row
For fi = fi_ini To n_filas - 1
    For co = 1 To n_columnas
        Detalle.TextMatrix(fi, co) = Detalle.TextMatrix(fi + 1, co)
    Next
Next
' última fila
For co = 1 To n_columnas
    Detalle.TextMatrix(fi, co) = ""
Next
Detalle.col = 1
Detalle.Row = fi_ini
'Detalle_Sumar
End Sub
Private Sub FilaBorrarContenido_Click()
' borra contenido de la fila en flexgrid
Dim fi As Integer, co As Integer
fi = Detalle.Row
For co = 1 To n_columnas
    Detalle.TextMatrix(fi, co) = ""
Next
'Detalle_Sumar
End Sub
Private Sub Doc_Imprimir(n_Copias As Integer)
MousePointer = vbHourglass
linea = String(78, "-")
' imprime OT
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer, tab4 As Integer, tab5 As Integer, tab6 As Integer
', tab7 As Integer, tab8 As Integer, tab9 As Integer
Dim tab10 As Integer, tab40 As Integer
Dim fc As Integer, fn As Integer, fe As Integer, k As Integer
tab0 = 3 'margen izquierdo
tab1 = tab0 'cod
tab2 = tab1 + 17 ' desc
tab3 = tab2 + 30 ' cant 36
tab4 = tab3 + 6 ' cant
tab5 = tab4 + 5  ' $ uni
tab6 = tab5 + 7  ' $ tot

tab40 = 43

' font.size
fc = 10 'comprimida
fn = 12 'normal
fe = 15 'expandida

Dim can_valor As String
Dim LargoEspecial As Integer

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt

For k = 1 To n_Copias

prt.Font.Size = fe
prt.Print Tab(tab0 - 1); m_TextoIso
prt.Print Tab(tab0 + 25); "VALE DE CONSUMO Nº";
prt.Font.Bold = True
'prt.Print Tab(tab0 + 18); Format(Numero.Text, "#####");
prt.Print Format(Numero.Text, "#####");
prt.Font.Bold = False
prt.Print Tab(tab0 + 52); Fecha.Text
prt.Font.Size = fc
prt.Print ""

' cabecera
prt.Font.Bold = True
prt.Print Tab(tab0); Empresa.Razon;
prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0 + tab40); "CONTRATISTA :"
prt.Print Tab(tab0); "GIRO: " & Empresa.Giro;
prt.Font.Size = fe
prt.Font.Bold = True
prt.Print Tab(tab0 + 28); Left(Razon, 32)

prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0); Empresa.Direccion;
prt.Font.Size = fc
prt.Print Tab(tab0 + tab40); "OBRA      : "
prt.Print Tab(tab0); "Teléfono: "; Empresa.Telefono1;
prt.Font.Size = fe
prt.Font.Bold = True
prt.Print Tab(tab0 + 28); Left(Format(Mid(ComboNV.Text, 8), ">"), 32)

prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0); Empresa.Comuna;
prt.Font.Size = fn

prt.Print ""
prt.Print ""

' detalle
prt.Font.Bold = True
prt.Print Tab(tab1); "COD";
prt.Print Tab(tab2); "Descripción";
prt.Print Tab(tab3); " L.Esp";
prt.Print Tab(tab4); "CANT";
prt.Print Tab(tab5); " $ UNI";
prt.Print Tab(tab6); "   $ TOT"
prt.Font.Bold = False

prt.Print Tab(tab1); linea
j = -1

For i = 1 To n_filas

    can_valor = Detalle.TextMatrix(i, 3)
    
    If Val(can_valor) = 0 Then
    
        j = j + 1
        prt.Print Tab(tab1 + j * 3); "  \"
        
    Else
    
        ' COD PRO
        prt.Print Tab(tab1); Detalle.TextMatrix(i, 1);
        
        ' DESCRIPCION
        prt.Print Tab(tab2); Left(Detalle.TextMatrix(i, 5), 35);
        
        ' LARGO ESPECIAL
        LargoEspecial = Detalle.TextMatrix(i, 6)
        If LargoEspecial > 0 Then
            prt.Print Tab(tab3); m_Format(LargoEspecial, "##,###");
        End If
        
        ' CANTIDAD
        prt.Print Tab(tab4); m_Format(can_valor, "#,###");
        
        ' $ UNITARIO
        prt.Print Tab(tab5); m_Format(Detalle.TextMatrix(i, 7), "###,###");
        
        ' $ TOTAL
        prt.Print Tab(tab6); m_Format(Detalle.TextMatrix(i, 8), "#,###,###")
        
    End If
    
Next

prt.Print Tab(tab1); linea
prt.Font.Bold = True
prt.Print Tab(tab5 - 5); m_Format(TotalPrecio, "$###,###,###")
prt.Font.Bold = False
prt.Print ""

'prt.Print Tab(tab0); "OBSERVACIONES :";
'prt.Print Tab(tab0 + 16); Obs(0).Text
'prt.Print Tab(tab0 + 16); Obs(1).Text
'prt.Print Tab(tab0 + 16); Obs(2).Text
'prt.Print Tab(tab0 + 16); Obs(3).Text

For i = 1 To 2
    prt.Print ""
Next

prt.Print Tab(tab0); Tab(14), "__________________", Tab(56), "__________________"
prt.Print Tab(tab0); Tab(14), "       VºBº       ", Tab(56), "       VºBº       "

If k < n_Copias Then
    prt.NewPage
Else
    prt.EndDoc
End If

Next

Impresora_Predeterminada "default"

MousePointer = vbDefault
End Sub
Private Sub CodigoProducto_Buscar()
Dim m_cod As String

MousePointer = vbHourglass
Product_Search.Condicion = ""
Load Product_Search
MousePointer = vbDefault
Product_Search.Show 1
m_cod = Product_Search.CodigoP

Codigo2Descripcion m_cod

End Sub
Private Sub Codigo2Descripcion(CodigoP As String) ', fila As Integer)
' busca descripcion del codigo del producto

Dim m_Upc As Double
If CodigoP = "" Then

    Detalle.TextMatrix(Detalle.Row, 4) = ""
    Detalle.TextMatrix(Detalle.Row, 5) = ""
   
Else

    With RsPrd
    .Seek "=", CodigoP
    If Not .NoMatch Then
    
        Detalle.TextMatrix(Detalle.Row, 1) = CodigoP
        Detalle.TextMatrix(Detalle.Row, 4) = ![unidad de medida]
        Detalle.TextMatrix(Detalle.Row, 5) = !Descripcion
        Detalle.TextMatrix(Detalle.Row, 10) = ![Largo Especial]
        
        ' busca ultimo precio de compra
        m_Upc = 0
        RsOCd.Seek ">=", CodigoP
        If Not RsOCd.NoMatch Then
            Do While Not RsOCd.EOF
                If RsOCd![codigo producto] = CodigoP Then
                    m_Upc = RsOCd![Precio Unitario]
                Else
                    Exit Do
                End If
                RsOCd.MoveNext
            Loop
        End If
        Detalle.TextMatrix(Detalle.Row, 7) = m_Upc
        
    '    Detalle.TextMatrix(Detalle.Row, 4) = ![Unidad de Medida]
    '    Detalle.TextMatrix(Detalle.Row, 5) = !Descripción
    '    Detalle.TextMatrix(Detalle.Row, 6) = !Largo
    '    Detalle.TextMatrix(Detalle.Row,10) = ![Largo Especial]
        
    '    If Detalle.TextMatrix(Detalle.Row,10) Then
    '        Detalle.col = 2 'foco en largo especial
    '    Else
         Detalle.col = 2 '3
    '    End If

    Else
    
        Detalle.TextMatrix(Detalle.Row, 5) = "-- CODIGO NO EXISTE --"
        Detalle.TextMatrix(Detalle.Row, 10) = False
    
    End If
    
    End With

End If

End Sub
