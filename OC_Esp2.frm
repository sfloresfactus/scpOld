VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form OC_Esp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orden De Compra"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   14820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ComboCtaContable 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   3960
      Width           =   2535
   End
   Begin VB.ComboBox ComboCCosto 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   3600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ingresar Nueva Orden de Compra Especial"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar Orden de Compra Especial"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar Orden de Compra Especial"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Orden de Compra Especial"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anular Orden de Compra Especial"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Grabar Orden de Compra Especial"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Deshacer"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mantención de Proveedores"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Orden de Compra Especial C2"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CheckBox conIva 
      Alignment       =   1  'Right Justify
      Caption         =   "con IVA"
      Height          =   255
      Left            =   8160
      TabIndex        =   23
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox MEdit 
      Height          =   495
      Left            =   7920
      TabIndex        =   50
      Text            =   "MEdit"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox atencion 
      Height          =   300
      Left            =   960
      TabIndex        =   15
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox condiciones 
      Height          =   300
      Left            =   4920
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Nv 
      Height          =   300
      Left            =   960
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.CheckBox Certificado 
      Caption         =   "No se recepcionará material sin Certificado de Calidad adjunto"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   6120
      Width           =   4815
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   8640
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowBorderStyle=   1
      WindowControlBox=   -1  'True
   End
   Begin VB.ComboBox EntregarEn 
      Height          =   315
      Left            =   5040
      TabIndex        =   17
      Top             =   2040
      Width           =   2775
   End
   Begin MSMask.MaskEdBox Descuento 
      Height          =   300
      Left            =   6720
      TabIndex        =   22
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "##########"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox pDescuento 
      Height          =   300
      Left            =   5520
      TabIndex        =   21
      Top             =   4920
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   529
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   2
      Mask            =   "##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox FechaaRecibir 
      Height          =   300
      Left            =   7080
      TabIndex        =   13
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   327680
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox ComboNV 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   3
      Left            =   120
      MaxLength       =   50
      TabIndex        =   28
      Top             =   5700
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   2
      Left            =   120
      MaxLength       =   50
      TabIndex        =   27
      Top             =   5400
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   1
      Left            =   120
      MaxLength       =   50
      TabIndex        =   26
      Top             =   5100
      Width           =   5000
   End
   Begin VB.Frame Frame 
      Caption         =   "Proveedor"
      Height          =   1095
      Left            =   2400
      TabIndex        =   33
      Top             =   480
      Width           =   5655
      Begin VB.ComboBox Direccion 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton btnSearch 
         Height          =   300
         Left            =   2400
         Picture         =   "OC_Esp2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   300
      End
      Begin MSMask.MaskEdBox Rut 
         Height          =   300
         Left            =   720
         TabIndex        =   30
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label lbl 
         Caption         =   "COM"
         Height          =   255
         Index           =   12
         Left            =   3600
         TabIndex        =   49
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Razon 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2880
         TabIndex        =   48
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lbl 
         Caption         =   "DIR"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   47
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Comuna 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3960
         TabIndex        =   46
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lbl 
         Caption         =   "&RUT"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   0
      Left            =   120
      MaxLength       =   50
      TabIndex        =   25
      Top             =   4800
      Width           =   5000
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   327680
      AutoTab         =   -1  'True
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Numero 
      Height          =   300
      Left            =   960
      TabIndex        =   1
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
      Height          =   1725
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3043
      _Version        =   327680
      ScrollBars      =   2
   End
   Begin MSMask.MaskEdBox Cotizacion 
      Height          =   300
      Left            =   1800
      TabIndex        =   19
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "##########"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   8520
      Top             =   840
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
            Picture         =   "OC_Esp2.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OC_Esp2.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OC_Esp2.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OC_Esp2.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OC_Esp2.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OC_Esp2.frx":065C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OC_Esp2.frx":076E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OC_Esp2.frx":0880
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblccCodigo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   8760
      TabIndex        =   54
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblccDescripcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   10200
      TabIndex        =   53
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label lbl 
      Caption         =   "&Guía Despacho Nº"
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   18
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblpDescuento 
      Caption         =   "% Desc"
      Height          =   255
      Left            =   6000
      TabIndex        =   45
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label lblTotal 
      Caption         =   "TOTAL"
      Height          =   255
      Left            =   5520
      TabIndex        =   44
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lblIva 
      Caption         =   "IVA"
      Height          =   255
      Left            =   5520
      TabIndex        =   43
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label lblNeto 
      Caption         =   "NETO"
      Height          =   255
      Left            =   5520
      TabIndex        =   42
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lblDescuento 
      Caption         =   "DESCUENTO"
      Height          =   255
      Left            =   5520
      TabIndex        =   41
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lblSubTotal 
      Caption         =   "SUBTOTAL"
      Height          =   255
      Left            =   5520
      TabIndex        =   40
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6720
      TabIndex        =   39
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Iva 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6720
      TabIndex        =   38
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Neto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6720
      TabIndex        =   37
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label pDesc_a_Dinero 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6720
      TabIndex        =   36
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "&Entregar en"
      Height          =   255
      Index           =   11
      Left            =   4080
      TabIndex        =   16
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&At. Sr."
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   14
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "&Condic."
      Height          =   255
      Index           =   6
      Left            =   4320
      TabIndex        =   10
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "a Reci&bir"
      Height          =   255
      Index           =   5
      Left            =   6240
      TabIndex        =   12
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "ESPECIAL"
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
      TabIndex        =   35
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label SubTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6720
      TabIndex        =   34
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   24
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "&OBRA"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "&FECHA"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "OC"
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
      TabIndex        =   32
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
      TabIndex        =   0
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
Attribute VB_Name = "OC_Esp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button, btnImprimir As Button
Private btnAnular As Button, btnGrabar As Button, btnDesHacer As Button, btnImprimirC2 As Button

Private DbD As Database, RsPrv As Recordset, RsProDir As Recordset
Private Dbm As Database, RsNVc As Recordset
Private Dba As Database, RsOcc As Recordset, RsOCd, RsCorre As Recordset

Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, d As Variant

Private n_filas As Integer, n_columnas As Integer
Private prt As Printer
Private n1 As Double, n4 As Double
Private NV_Numero As Double, NV_Obra As String
Private linea As String
Private m_Certificado As Boolean
Private a_Nv(2999, 3) As String, m_NvArea As Integer

Private aLineas(99, 1) As String ' aqui van los codigos de las cuentas contables (0) y codigos de centros de costo(1), elegidos para cada linea

Private m_ImpresoraNombre As String
Private FormatoDoc As String
Private Sub ComboCCosto_LostFocus()
ComboCCosto.visible = False
End Sub

Private Sub ComboCtaContable_Click()

If ComboCtaContable.ListIndex = -1 Then Exit Sub

Dim fi As Integer, j As Integer, paso As String
fi = Detalle.Row
If aCuCo(ComboCtaContable.ListIndex + 1, 2) = "S" Then

    paso = Trim(ComboCtaContable.Text)
    
    'MsgBox "|" & paso & "|"
    
    'Detalle = paso
    paso = Mid(paso, 1, InStr(1, paso, " "))
    paso = Trim(paso)
    
    'MsgBox "|" & paso & "|"
        
    j = cuentaContableBuscarIndice(NoNulo(paso))
    If j > 0 Then
        Detalle.TextMatrix(fi, 6) = aCuCo(j, 0) & " " & aCuCo(j, 1)
        aLineas(fi, 0) = aCuCo(j, 0)
    End If
        
Else
    MsgBox "Cuenta Contable NO es Imputable"
End If

Detalle.SetFocus

ComboCtaContable.visible = False

End Sub
Private Sub ComboCtaContable_LostFocus()
ComboCtaContable.visible = False
End Sub
Private Sub ComboCCosto_Click()
If ComboCCosto.ListIndex = -1 Then Exit Sub
Dim fi As Integer
fi = Detalle.Row
If aCeCo(ComboCCosto.ListIndex, 2) = "S" Then
    Detalle = Trim(ComboCCosto.Text)
    aLineas(fi, 1) = aCeCo(ComboCCosto.ListIndex, 0)
Else
    MsgBox "Centro Costo NO es Imputable"
End If
ComboCCosto.visible = False
End Sub
Private Sub Form_Load()

Dim indice As Integer

' abre archivos
Set DbD = OpenDatabase(data_file)
Set RsPrv = DbD.OpenRecordset("Proveedores")
RsPrv.Index = "RUT"

Set RsProDir = DbD.OpenRecordset("Proveedores-Direcciones")
RsProDir.Index = "RUT-Codigo"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set Dba = OpenDatabase(Madq_file)
Set RsOcc = Dba.OpenRecordset("OC Cabecera")
RsOcc.Index = "Numero"

Set RsOCd = Dba.OpenRecordset("OC Detalle")
RsOCd.Index = "Numero-Linea"
Set RsCorre = Dba.OpenRecordset("Correlativo")

i = 0
' Combo obra
ComboNV.AddItem " "
If False Then
Do While Not RsNVc.EOF
    If Usuario.Nv_Activas Then
        If RsNVc!Activa Then
            i = i + 1
            a_Nv(i, 0) = RsNVc!Numero
            a_Nv(i, 1) = RsNVc!obra
            ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
        End If
    Else
        i = i + 1
        a_Nv(i, 0) = RsNVc!Numero
        a_Nv(i, 1) = RsNVc!obra
        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
    End If
    RsNVc.MoveNext
Loop
End If

Do While Not RsNVc.EOF
    If Usuario.Nv_Activas Then
        If RsNVc!Activa Then
            i = i + 1
            a_Nv(i, 0) = RsNVc!Numero
            a_Nv(i, 1) = RsNVc!obra
            indice = scpNew_Nv2ccCodigo(RsNVc!Numero)
            If indice > -1 Then
                a_Nv(i, 2) = scpNew_aNv(indice, 2) ' codigo centro de costo scp nuevo
                a_Nv(i, 3) = scpNew_aNv(indice, 3) ' descripcion centro de costo scp nuevo
            End If
            ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
        End If
    Else
        i = i + 1
        a_Nv(i, 0) = RsNVc!Numero
        a_Nv(i, 1) = RsNVc!obra
        indice = scpNew_Nv2ccCodigo(RsNVc!Numero)
        If indice > -1 Then
            a_Nv(i, 2) = scpNew_aNv(indice, 2) ' codigo centro de costo scp nuevo
            a_Nv(i, 3) = scpNew_aNv(indice, 3) ' descripcion centro de costo scp nuevo
        End If
        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
    End If
    RsNVc.MoveNext
Loop



Inicializa
Detalle_Config

If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 1, 0, 0, 1
End If

conIva.Value = 1

'StatusBar.Panels(1) = EmpOC.Razon
m_NvArea = 0

FormatoDoc = ReadIniValue(Path_Local & "scp.ini", "GD", "Formato")

cuentaContableComboPoblar
centroCostoComboPoblar

End Sub
Private Sub Form_Unload(Cancel As Integer)
DataBases_Cerrar
End Sub
Private Sub Inicializa()

Obj = "ORDEN DE COMPRA"
Objs = "ÓRDENES DE COMPRA"
'TipoDoc = "OTE"

Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnAnular = Toolbar.Buttons(5)
Set btnGrabar = Toolbar.Buttons(7)
Set btnDesHacer = Toolbar.Buttons(8)
Set btnImprimirC2 = Toolbar.Buttons(12)

Accion = ""
'old_accion = ""

'btnSearch.Visible = False
btnSearch.ToolTipText = "Busca Proveedor"
Campos_Enabled False

condiciones.MaxLength = 20
atencion.MaxLength = 50

EntregarEn.AddItem Empresa.Direccion
EntregarEn.AddItem "RETIRAMOS"
EntregarEn.AddItem "OBRA"
EntregarEn.AddItem "STA. ALEJANDRA 03521"

End Sub
Private Sub Detalle_Config()
Dim i As Integer, ancho As Integer, c_totales As Integer, ancho_totales As Integer

ComboCtaContable.visible = False

n_filas = 19
n_columnas = 7

ancho_totales = 1000

Detalle.Left = 100
Detalle.WordWrap = True
Detalle.RowHeight(0) = 450
Detalle.Cols = n_columnas + 1
Detalle.Rows = n_filas + 1

Detalle.TextMatrix(0, 0) = ""
Detalle.TextMatrix(0, 1) = "Cantidad"
Detalle.TextMatrix(0, 2) = "Unidad"
Detalle.TextMatrix(0, 3) = "Descripción"
Detalle.TextMatrix(0, 4) = "Precio Unitario"
Detalle.TextMatrix(0, 5) = "Precio TOTAL"      '*
Detalle.TextMatrix(0, 6) = "Cuenta Contable"
Detalle.TextMatrix(0, 7) = "Centro Costo"

Detalle.ColWidth(0) = 250
Detalle.ColWidth(1) = 1000
Detalle.ColWidth(2) = 600
Detalle.ColWidth(3) = 4500
Detalle.ColWidth(4) = 1000
Detalle.ColWidth(5) = ancho_totales
Detalle.ColWidth(6) = 2000
Detalle.ColWidth(7) = 2000

'Detalle.ColAlignment(2) = 0

ancho = 350 ' con scroll vertical

SubTotal.Width = ancho_totales
pDesc_a_Dinero.Width = ancho_totales
Descuento.Width = ancho_totales
Neto.Width = ancho_totales
Iva.Width = ancho_totales
Total.Width = ancho_totales

For i = 0 To n_columnas
    If i = 5 Then
    
        conIva.Left = ancho + Detalle.Left - 2500
    
        c_totales = ancho + Detalle.Left - 350
        lblSubTotal.Left = c_totales - 1000
        SubTotal.Left = c_totales
        
        pDescuento.Left = c_totales - 1000
        lblpDescuento.Left = c_totales - 600
        pDesc_a_Dinero.Left = c_totales
        
        lblDescuento.Left = c_totales - 1000
        Descuento.Left = c_totales
        
        lblNeto.Left = c_totales - 1000
        Neto.Left = c_totales
        lblIva.Left = c_totales - 1000
        Iva.Left = c_totales
        lblTotal.Left = c_totales - 1000
        Total.Left = c_totales
        
    End If
    ancho = ancho + Detalle.ColWidth(i)
Next

Detalle.Width = ancho
Me.Width = ancho + Detalle.Left * 2 + 1400

For i = 1 To n_filas
    Detalle.TextMatrix(i, 0) = i
    Detalle.Row = i
    Detalle.col = 2
    Detalle.CellAlignment = flexAlignLeftCenter
    Detalle.col = 3
    Detalle.CellAlignment = flexAlignLeftCenter
    Detalle.col = 5
    Detalle.CellForeColor = vbRed
Next

MEdit.Text = ""

End Sub
Private Sub Nv_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Nv_LostFocus()
Dim m_Nv As Integer
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
'SendKeys "{Home}+{End}"

End Sub
Private Sub ComboNV_Click()
Dim indice As Integer
NV_Numero = Val(Left(ComboNV.Text, 6))
Nv.Text = NV_Numero
NV_Obra = Mid(ComboNV.Text, 8)

indice = ComboNV.ListIndex

lblccCodigo.Caption = a_Nv(indice, 2)
lblccDescripcion.Caption = a_Nv(indice, 3)

End Sub
Private Sub Numero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then After_Enter
End Sub
Private Sub After_Enter()
If Val(Numero.Text) < 1 Then Beep: Exit Sub
Select Case Accion
Case "Agregando"
    RsOcc.Seek "=", Numero
    If RsOcc.NoMatch Then
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
    RsOcc.Seek "=", Numero
    If RsOcc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Botones_Enabled 0, 0, 0, 0, 0, 1, 1, 0
        
        Doc_Leer
        
        If RsOcc!Tipo = "E" Then
            Campos_Enabled True
            Numero.Enabled = False
            btnGrabar.Enabled = True
            btnSearch.visible = True
        Else
            MsgBox "DEBE MODIFICAR COMO OC NORMAL"
            Campos_Limpiar
            Numero.Enabled = True
            Numero.SetFocus
        End If

    End If

Case "Eliminando"
    RsOcc.Seek "=", Numero
    If RsOcc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
    
        Doc_Leer
        If RsOcc!Tipo = "E" Then
            Numero.Enabled = False
            If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo) = vbYes Then
                Doc_Eliminar
            End If
        Else
            MsgBox "DEBE ELIMINAR COMO OC NORMAL"
        End If
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus

    End If
   
Case "Imprimiendo"
    
    RsOcc.Seek "=", Numero
    If RsOcc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Doc_Leer
        If RsOcc!Tipo = "E" Then
            Numero.Enabled = False
            Detalle.Enabled = True
        Else
            MsgBox "DEBE IMPRIMIR COMO OC NORMAL"
            Campos_Limpiar
            Numero.Enabled = True
            Numero.SetFocus
        End If

    End If
    
Case "Anulando"
    RsOcc.Seek "=", Numero
    If RsOcc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
    
        Doc_Leer
        If RsOcc!Tipo = "E" Then
            Numero.Enabled = False
            If MsgBox("¿ ANULA " & Obj & " ?", vbYesNo) = vbYes Then
                Doc_Anular
            End If
        Else
            MsgBox "DEBE ANULAR COMO OC NORMAL"
        End If
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus

    End If

Case "ImprimiendoC2"
    
    RsOcc.Seek "=", Numero
    If RsOcc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Doc_Leer
        If RsOcc!Tipo = "E" Then
            Numero.Enabled = False
            Detalle.Enabled = True
        Else
            MsgBox "DEBE IMPRIMIR COMO OC NORMAL"
            Campos_Limpiar
            Numero.Enabled = True
            Numero.SetFocus
        End If

    End If
    
End Select
   
End Sub
Private Sub Doc_Leer()
Dim m_resta As Integer
Dim primera As Boolean

' CABECERA


With RsOcc
Fecha.Text = Format(!Fecha, Fecha_Format)
NV_Numero = !Nv
rut.Text = ![RUT Proveedor]

NV_Obra = ""
RsNVc.Seek "=", NV_Numero, m_NvArea
If Not RsNVc.NoMatch Then
    NV_Obra = RsNVc!obra
    ComboNV.Text = Format(RsNVc!Numero, "0000") & " - " & NV_Obra
    ComboNV_Click
End If

condiciones.Text = NoNulo(![Condiciones de Pago])
FechaaRecibir.Text = Format(![Fecha a Recibir], Fecha_Format)
atencion.Text = NoNulo(!atencion)
EntregarEn.Text = NoNulo(![Entregar en])
Cotizacion.Text = !Cotizacion

SubTotal.Caption = !SubTotal
pDescuento.Text = ![% Descuento]
Descuento.Text = !Descuento

If !Neto = !Total Then
    conIva.Value = 0
End If

Totales_Calcular

Obs(0).Text = NoNulo(![Observacion 1])
Obs(1).Text = NoNulo(![Observacion 2])
Obs(2).Text = NoNulo(![Observacion 3])
Obs(3).Text = NoNulo(![Observacion 4])

Certificado.Value = IIf(!Certificado, 1, 0)

If !Nula Then
    MsgBox "OC NULA !!!"
End If

End With

'DETALLE
primera = True
If RsOcc!Tipo = "E" Then
    With RsOCd
    .Seek ">=", Numero.Text, 0
    If Not .NoMatch Then
        Do While Not .EOF
            If !Numero = Numero.Text Then
            
                i = !linea
                               
                Detalle.TextMatrix(i, 1) = !Cantidad 'Format(!Cantidad, "#")
                Detalle.TextMatrix(i, 2) = !unidad
                Detalle.TextMatrix(i, 3) = !Descripcion
'                Detalle.TextMatrix(i, 4) = Format(![Precio Unitario], "#")
                Detalle.TextMatrix(i, 4) = ![Precio Unitario]
                
                j = cuentaContableBuscarIndice(NoNulo(!cuentacontable))
                If j > 0 Then
                    Detalle.TextMatrix(i, 6) = aCuCo(j, 0) & " " & aCuCo(j, 1)
                    aLineas(i, 0) = aCuCo(j, 0)
                End If
                 j = centroCostoBuscarIndice(NoNulo(!CentroCosto))
                If j > 0 Then
                    Detalle.TextMatrix(i, 7) = aCeCo(j, 0) & " " & aCeCo(j, 1)
                    aLineas(i, 1) = aCeCo(j, 0)
                End If
                
                Fila_Calcular i, False
                
            Else
                Exit Do
            End If
            .MoveNext
        Loop
    End If
    End With
Else

End If

Proveedor_Leer rut.Text
If Direccion.ListCount > 0 Then Direccion.ListIndex = RsOcc![Codigo Direccion]

Detalle.Row = 1 ' para q' actualice la primera fila del detalle

'Detalle_Sumar


End Sub
Private Sub Proveedor_Leer(rut)
With RsPrv
.Seek "=", rut
If Not .NoMatch Then

    Razon.Caption = RsPrv![Razon Social]
    
    '/////////
    Direccion.Clear
    Direccion.AddItem RsPrv!Direccion
    Direccion.Text = RsPrv!Direccion ' es como un click
    
    With RsProDir
    .Seek "=", rut, 1
    If Not .NoMatch Then
        Do While Not .EOF
            If rut = !rut Then Direccion.AddItem !Direccion
            .MoveNext
        Loop
    End If
    End With

End If

End With

End Sub
Private Sub Direccion_Click()
Dim d As Integer
d = Direccion.ListIndex
If d = 0 Then
    Comuna.Caption = NoNulo(RsPrv!Comuna)
    condiciones.Text = NoNulo(RsPrv![Condiciones de Pago])
    atencion.Text = NoNulo(RsPrv!Contacto)
Else
    RsProDir.Seek "=", rut, d
    If Not RsProDir.NoMatch Then
        Comuna.Caption = NoNulo(RsProDir!Comuna)
        atencion.Text = NoNulo(RsProDir!Contacto)
    End If
End If
End Sub
Private Function Doc_Valido() As Boolean
Dim porAsignar As Integer
Doc_Valido = False
If Trim(ComboNV.Text) = "" Then
    MsgBox "DEBE ELEGIR OBRA"
    ComboNV.SetFocus
    Exit Function
End If
If rut.Text = "" Then
    MsgBox "DEBE ELEGIR PROVEEDOR"
    btnSearch.SetFocus
    Exit Function
End If
If FechaaRecibir.Text = Fecha_Vacia Then
    MsgBox "DEBE DIGITAR FECHA DE RECEPCIÓN"
    FechaaRecibir.SetFocus
    Exit Function
End If

For i = 1 To n_filas

    ' cant
    If Val(Detalle.TextMatrix(i, 1)) <> 0 Then
    
        If Not CampoReq_Valida(Detalle.TextMatrix(i, 2), i, 2) Then Exit Function
        If Not CampoReq_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
        If Not Numero_Valido(Detalle.TextMatrix(i, 4), i, 4) Then Exit Function
        
        If Detalle.TextMatrix(i, 6) = "" Then
            MsgBox "Debe Escoger Cuenta Contable"
            Detalle.Row = i
            Detalle.col = 6
            Detalle.SetFocus
            Exit Function
        End If
        
    End If
    
Next

Doc_Valido = True

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
Private Function Numero_Valido(txt As String, fil As Integer, col As Integer) As Boolean
Dim num As String
Numero_Valido = False
num = txt
If Not IsNumeric(num) Then
    GoTo NoValido
Else
    If Val(num) < 0 Then ' solo mayores que cero
        GoTo NoValido
    End If
End If
Numero_Valido = True
Exit Function

NoValido:
Beep
MsgBox "Número no Válido"
Detalle.Row = fil
Detalle.col = col
Detalle.SetFocus
End Function
Private Function Doc_Grabar(Nueva As Boolean) As Double
Dim m_cantidad As String, m_pr As Double, m_Numero As Double
Doc_Grabar = 0

save:
' CABECERA
With RsOcc
If Nueva Then

    m_Numero = GetNumDoc("OC", RsOcc, RsCorre)
    If m_Numero = 0 Then
        MsgBox "OC NO SE GRABÓ!"
        Doc_Grabar = 0
        MousePointer = vbDefault
        Exit Function
    End If

    Numero.Text = m_Numero
    Doc_Grabar = m_Numero

    .AddNew
    !Numero = Numero.Text
    !Tipo = "E"
Else
    .Edit
    !fechaModificacion = Format(Now, Fecha_Format)
End If

!Fecha = Fecha.Text
!Nv = NV_Numero
![RUT Proveedor] = rut.Text
![Codigo Direccion] = Direccion.ListIndex
![Condiciones de Pago] = condiciones.Text
![Fecha a Recibir] = CDate(FechaaRecibir.Text)
!atencion = atencion.Text
![Entregar en] = EntregarEn.Text
!Cotizacion = Val(Cotizacion.Text)
![Observacion 1] = Obs(0).Text
![Observacion 2] = Obs(1).Text
![Observacion 3] = Obs(2).Text
![Observacion 4] = Obs(3).Text
!SubTotal = SubTotal.Caption
![% Descuento] = Val(pDescuento.Text)
!Descuento = Val(Descuento.Text)
!Neto = Neto.Caption
!Iva = Iva.Caption
!Total = Total.Caption
!Pendiente = True
!Nula = False

!Certificado = Certificado.Value

.Update
End With

' DETALLE
Doc_Detalle_Eliminar

With RsOCd
j = 0
For i = 1 To n_filas
    m_cantidad = Detalle.TextMatrix(i, 1)
'    If m_cantidad = 0 Then
'        ' puede haber texto
'    Else
'        j = j + 1

    'graba detalle de todas las lineas que tengan texto
    If Detalle.TextMatrix(i, 3) <> "" Then
        .AddNew
        !Numero = Numero.Text
        !Tipo = "E"
        !linea = i 'j
        !Fecha = Fecha.Text
        !Nv = NV_Numero
        ![RUT Proveedor] = rut.Text
        !Cantidad = m_CDbl(m_cantidad)
        !unidad = Detalle.TextMatrix(i, 2)
        !Descripcion = Detalle.TextMatrix(i, 3)
        ![Precio Unitario] = m_CDbl(Detalle.TextMatrix(i, 4))
        !Pendiente = True
        ![Fecha a Recibir] = CDate(FechaaRecibir.Text)
'       !Total = Format(m_cantidad * m_CDbl(Detalle.TextMatrix(i, 4)), "#########0") 'num_fmtgrl)
        !Total = Val(Detalle.TextMatrix(i, 5))
        
        !cuentacontable = aLineas(i, 0)
        !CentroCosto = aLineas(i, 1)
        
        .Update
        
    End If
Next
End With

DesBloqueo "OC", RsCorre

End Function
Private Sub Doc_Eliminar()

' borra CABECERA
RsOcc.Seek "=", Numero.Text
If Not RsOcc.NoMatch Then

    RsOcc.Delete

End If

Doc_Detalle_Eliminar

End Sub
Private Sub Doc_Detalle_Eliminar()

' elimina detalle
With RsOCd
.Seek "=", Numero.Text, 1
If Not .NoMatch Then
    Do While Not .EOF
        If !Numero <> Numero.Text Then Exit Do
    
        ' borra detalle
        .Delete
    
        .MoveNext
    Loop
End If
End With

End Sub
Private Sub Doc_Anular()

' elimina cabecera
With RsOcc
.Seek "=", Numero.Text
If Not RsOcc.NoMatch Then
    .Edit
    !Nula = True
    .Update
End If
End With

' elimina detalle
Doc_Detalle_Eliminar

End Sub
Private Sub Campos_Limpiar()

Numero.Text = ""
Fecha.Text = Fecha_Vacia
Nv.Text = ""
ComboNV.Text = " "

rut.Text = ""
Razon.Caption = ""
Direccion.Clear
Comuna.Caption = ""
condiciones.Text = ""
FechaaRecibir.Text = Fecha_Vacia
atencion.Text = ""
EntregarEn.Text = ""
Cotizacion.Text = ""

ComboCtaContable.ListIndex = -1
ComboCCosto.ListIndex = -1

lblccCodigo.Caption = ""
lblccDescripcion.Caption = ""

Detalle_Limpiar

Obs(0).Text = ""
Obs(1).Text = ""
Obs(2).Text = ""
Obs(3).Text = ""

m_Certificado = IIf(Certificado.Value = 1, True, False)
Certificado.Value = 0

SubTotal.Caption = ""
pDescuento.Text = "__"
pDesc_a_Dinero.Caption = ""
Descuento.Text = "__________"
Neto.Caption = ""
Iva.Caption = ""
Total.Caption = ""

conIva.Value = 1

End Sub
Private Sub Detalle_Limpiar()
Dim fi As Integer, co As Integer
For fi = 1 To n_filas
    For co = 1 To n_columnas
        Detalle.TextMatrix(fi, co) = ""
    Next
    aLineas(fi, 0) = "" ' cuenta contable
    aLineas(fi, 1) = "" ' centro costo
Next
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
Dim cambia_titulo As Boolean, m_Numero As String, m_obra As String

cambia_titulo = True
'Accion = "" rem accion
Select Case Button.Index
Case 1 ' Agregar
    Accion = "Agregando"
    Botones_Enabled 0, 0, 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    
'    Numero.Text = Documento_Numero_Nuevo(RsOcc, "Número")
'    Numero.Enabled = True
'    Numero.SetFocus
    Campos_Enabled True
    Numero.Enabled = False
    Fecha.SetFocus
    btnGrabar.Enabled = True
    btnSearch.visible = True

Case 2 ' Modificar
    Accion = "Modificando"
    Botones_Enabled 0, 0, 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    Numero.Enabled = True
    Numero.SetFocus
Case 3 ' Eliminar
    Accion = "Eliminando"
    Botones_Enabled 0, 0, 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    Numero.Enabled = True
    Numero.SetFocus
    
Case 4 ' Imprimir

    Accion = "Imprimiendo"
    If Numero.Text = "" Then
    
        Botones_Enabled 0, 0, 0, 1, 0, 0, 1, 0
        Campos_Enabled False
        Numero.Enabled = True
        Numero.SetFocus
        
    Else
    
        prt_escoger.ImpresoraNombre = ""
        prt_escoger.Show 1
        m_ImpresoraNombre = prt_escoger.ImpresoraNombre
        
'        If MsgBox("¿ IMPRIME " & Obj & " ?", vbYesNo) = vbYes Then
        MousePointer = vbHourglass
        m_Numero = Numero.Text
'        OC_Prepara Numero.Text, Nv.Text, NV_Obra, m_ImpresoraNombre
        OC_Prepara m_Numero, NV_Numero, NV_Obra, m_ImpresoraNombre
'            Doc_Imprimir
'        End If
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
        
        MousePointer = vbDefault
                
        Select Case FormatoDoc
        Case "EML"
            OC_PrintLegal_EML_Esp m_Numero
        Case "AYD"
            OC_PrintLegal_AYD_Esp m_Numero
        Case "DELSA" ' delsa no usa OC
'            OC_PrintLegal_Delsa_esp m_Numero
        Case "EIFFEL"
            OC_PrintLegal_EIFFEL_Esp m_Numero
        End Select
                
    End If
Case 5 ' Anular
    Accion = "Anulando"
    Botones_Enabled 0, 0, 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    Numero.Enabled = True
    Numero.SetFocus

Case 7 ' grabar
    If Doc_Valido Then
        If MsgBox("¿ GRABA " & Obj & " ?", vbYesNo) = vbYes Then
            
            m_Numero = Numero.Text
            m_obra = NV_Obra
            If Accion = "Agregando" Then
                m_Numero = Doc_Grabar(True)
            Else
                Doc_Grabar False
            End If
            
            Botones_Enabled 0, 0, 0, 0, 0, 0, 1, 0
            Campos_Limpiar
            
            If Accion = "Agregando" Then
                Campos_Enabled True
                Numero.Enabled = False
                Fecha.SetFocus
                btnGrabar.Enabled = True
                btnSearch.visible = True
            Else
                Campos_Enabled False
                Numero.Enabled = True
                Numero.SetFocus
            End If
            
            If m_Numero Then
            
                prt_escoger.ImpresoraNombre = ""
                prt_escoger.Show 1
                m_ImpresoraNombre = prt_escoger.ImpresoraNombre
            
'                If MsgBox("¿ IMPRIME " & Obj & " ?", vbYesNo) = vbYes Then
                OC_Prepara m_Numero, Nv.Text, m_obra, m_ImpresoraNombre
                
                Select Case FormatoDoc
                Case "EML"
                    OC_PrintLegal_EML_Esp m_Numero
                Case "AYD"
                    OC_PrintLegal_AYD_Esp m_Numero
                Case "DELSA"
'                    OC_PrintLegal_Delsa m_Numero
                Case "EIFFEL"
                    OC_PrintLegal_EIFFEL_Esp m_Numero
                End Select
                
'                End If

            End If

        End If
    End If
    
Case 8 ' DesHacer
    
    If Numero.Text = "" Then
    
        Botones_Enabled 1, 1, 1, 1, 1, 0, 0, 1
        Campos_Limpiar
        Campos_Enabled False
        
    Else
    
        If Accion = "Imprimiendo" Then
            Botones_Enabled 1, 1, 1, 1, 1, 0, 0, 1
            Campos_Limpiar
            Campos_Enabled False
        Else
            If MsgBox("¿Seguro que quiere DesHacer?", vbYesNo) = vbYes Then
                Botones_Enabled 1, 1, 1, 1, 1, 0, 0, 1
                Campos_Limpiar
                Campos_Enabled False
            End If
        End If
    End If
    Accion = ""
    
Case 9 'Separador
Case 10 'Proveedores
    MousePointer = 11
    Load Proveedores
    MousePointer = 0
    Proveedores.Show 1
    cambia_titulo = False

Case 12 ' Imprimir C2

    Accion = "ImprimiendoC2"
    
    If Numero.Text = "" Then
    
        Botones_Enabled 0, 0, 0, 0, 0, 0, 1, 1
        Campos_Enabled False
        Numero.Enabled = True
        Numero.SetFocus
        
    Else
    
        prt_escoger.ImpresoraNombre = ""
        prt_escoger.Show 1
        m_ImpresoraNombre = prt_escoger.ImpresoraNombre
        
'        If MsgBox("¿ IMPRIME " & Obj & " ?", vbYesNo) = vbYes Then
        MousePointer = vbHourglass
        m_Numero = Numero.Text
'        OC_Prepara Numero.Text, Nv.Text, NV_Obra, m_ImpresoraNombre
        OC_PreparaC2 m_Numero, NV_Numero, NV_Obra, m_ImpresoraNombre
'            Doc_Imprimir
'        End If
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
        
        MousePointer = vbDefault
                
        Select Case FormatoDoc
        Case "EML"
            OC_PrintLegal_C2_Esp m_Numero
        End Select
                
    End If
    
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
                            btn_Anular As Boolean, _
                            btn_Grabar As Boolean, btn_DesHacer As Boolean, btn_ImprimirC2 As Boolean)
                            
btnAgregar.Enabled = btn_Agregar
btnModificar.Enabled = btn_Modificar
btnEliminar.Enabled = btn_Eliminar
btnImprimir.Enabled = btn_Imprimir
btnAnular.Enabled = btn_Anular
btnGrabar.Enabled = btn_Grabar
btnDesHacer.Enabled = btn_DesHacer
btnImprimirC2.Enabled = btn_ImprimirC2

btnAgregar.Value = tbrUnpressed
btnModificar.Value = tbrUnpressed
btnEliminar.Value = tbrUnpressed
btnImprimir.Value = tbrUnpressed
btnAnular.Value = tbrUnpressed
btnGrabar.Value = tbrUnpressed
btnDesHacer.Value = tbrUnpressed
btnImprimirC2.Value = tbrUnpressed

End Sub
Private Sub Campos_Enabled(Si As Boolean)
Numero.Enabled = Si
Fecha.Enabled = Si
btnSearch.Enabled = Si
Direccion.Enabled = Si
Nv.Enabled = Si
ComboNV.Enabled = Si
condiciones.Enabled = Si
FechaaRecibir.Enabled = Si
atencion.Enabled = Si
EntregarEn.Enabled = Si
Cotizacion.Enabled = Si

ComboCtaContable.Enabled = Si
ComboCCosto.Enabled = Si

Detalle.Enabled = Si
pDescuento.Enabled = Si
Descuento.Enabled = Si
conIva.Enabled = Si
Obs(0).Enabled = Si
Obs(1).Enabled = Si
Obs(2).Enabled = Si
Obs(3).Enabled = Si
End Sub
Private Sub btnSearch_Click()

Search.Muestra data_file, "Proveedores", "RUT", "Razon Social", "Proveedor", "Proveedores"

With RsPrv
rut.Text = Search.Codigo
If rut.Text <> "" Then
    Proveedor_Leer (rut.Text)
End If
End With

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
Case 1 ' cantidad
Case 2 ' marca
Case 3 ' descripcion
Case 4 ' $ uni
Case 5 ' $ tot
    Case 6 ' cuenta contable
        On Error GoTo error1
        If Detalle <> "" Then ComboCtaContable.Text = Detalle
error1:
        On Error GoTo 0
        ComboCtaContable.Top = Detalle.CellTop + Detalle.Top
        ComboCtaContable.Left = Detalle.CellLeft + Detalle.Left
        ComboCtaContable.Width = Int(Detalle.CellWidth * 1.5)
        ComboCtaContable.visible = True
Case 7 ' centro costo
        On Error GoTo error2
        If Detalle <> "" Then ComboCCosto.Text = "   " & Detalle
error2:
        On Error GoTo 0
        ComboCCosto.Top = Detalle.CellTop + Detalle.Top
        ComboCCosto.Left = Detalle.CellLeft + Detalle.Left
        ComboCCosto.Width = Int(Detalle.CellWidth * 1.5)
        ComboCCosto.visible = True
Case Else
End Select
End Sub
Private Sub Detalle_DblClick()
If Accion = "Imprimiendo" Then Exit Sub
' simula un espacio
MSFlexGridEdit Detalle, MEdit, 32
End Sub
Private Sub Detalle_GotFocus()
If Accion = "Imprimiendo" Then Exit Sub
Select Case True
Case MEdit.visible
    Detalle = MEdit
    MEdit.visible = False
End Select
End Sub
Private Sub Detalle_LeaveCell()
If Accion = "Imprimiendo" Then Exit Sub
Select Case True
Case MEdit.visible
    Detalle = MEdit
    MEdit.visible = False
End Select
End Sub
Private Sub Detalle_KeyPress(KeyAscii As Integer)
If Accion = "Imprimiendo" Then Exit Sub
MSFlexGridEdit Detalle, MEdit, KeyAscii
End Sub
Private Sub MEdit_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCodeP Detalle, MEdit, KeyCode, Shift
End Sub
Sub EditKeyCodeP(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
' rutina que se ejecuta con los keydown de Edt
Dim m_fil As Integer, m_col As Integer
m_fil = MSFlexGrid.Row
m_col = MSFlexGrid.col

Select Case KeyCode
Case vbKeyEscape ' Esc
    Edt.visible = False
    MSFlexGrid.SetFocus
Case vbKeyReturn ' Enter
    MSFlexGrid.SetFocus
    DoEvents
    Select Case m_col
    Case 2, 3
    Case Else
        Fila_Calcular m_fil, True
    End Select
    Cursor_Mueve MSFlexGrid
Case vbKeyUp ' Flecha Arriba
    Select Case m_col
    Case 2, 3
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
        Fila_Calcular m_fil, True
        If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
            MSFlexGrid.Row = MSFlexGrid.Row - 1
        End If
    End Select
Case vbKeyDown ' Flecha Abajo
    Select Case m_col
    Case 2, 3
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
        Fila_Calcular m_fil, True
        If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
            MSFlexGrid.Row = MSFlexGrid.Row + 1
        End If
    End Select
End Select
End Sub
Private Sub MEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)

If MSFlexGrid.col = 5 Then Exit Sub

Select Case MSFlexGrid.col
Case 2
    Edt.MaxLength = 3
Case 3
    Edt.MaxLength = 50 '30
Case Else
    Edt.MaxLength = 10
End Select

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

End Sub
Private Sub Detalle_KeyDown(KeyCode As Integer, Shift As Integer)
If Accion = "Imprimiendo" Then Exit Sub
If KeyCode = vbKeyF2 Then
    MSFlexGridEdit Detalle, MEdit, 32
End If
End Sub
Private Sub Fila_Calcular(Fila As Integer, Actualiza As Boolean)
'actualiza solo totales de fila
'fila = Detalle.Row
n1 = m_CDbl(Detalle.TextMatrix(Fila, 1))
n4 = m_CDbl(Detalle.TextMatrix(Fila, 4))
' precio total
Detalle.TextMatrix(Fila, 5) = Format(n1 * n4, num_fmtgrl)

If Actualiza Then Detalle_Sumar

End Sub
Private Sub Detalle_Sumar()
Dim m_Total As Double
m_Total = 0
For i = 1 To n_filas
    m_Total = m_Total + m_CDbl(Detalle.TextMatrix(i, 5))
Next

Totales_Calcular m_Total

End Sub
Private Sub Totales_Calcular(Optional Sub_Total As Double)
If Sub_Total = 0 Then Sub_Total = Val(SubTotal.Caption)
SubTotal.Caption = Format(Sub_Total, num_fmtgrl)
pDesc_a_Dinero.Caption = Int(Val(SubTotal.Caption) * Val(pDescuento.Text) / 100 + 0.5)
Neto.Caption = Val(SubTotal.Caption) - Val(pDesc_a_Dinero.Caption) - Val(Descuento.Text)

If conIva.Value = 1 Then
    Iva.Caption = Int(m_CDbl(Neto.Caption) * Parametro.Iva / 100)
Else
    Iva.Caption = "0"
End If

Total.Caption = Val(Neto.Caption) + Val(Iva.Caption)
End Sub
Private Sub Cursor_Mueve(MSFlexGrid As Control)
'MIA
If MSFlexGrid.col = 5 Then
    If MSFlexGrid.Row + 1 < MSFlexGrid.Rows Then
        MSFlexGrid.col = 1
        MSFlexGrid.Row = MSFlexGrid.Row + 1
    End If
Else
    MSFlexGrid.col = MSFlexGrid.col + 1
End If
End Sub
'
' FIN RUTINAS PARA FLEXGRID
'
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
Private Sub pDescuento_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Descuento.SetFocus
End Sub
Private Sub pDescuento_LostFocus()
Totales_Calcular
End Sub
Private Sub Descuento_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Obs(0).SetFocus
End Sub
Private Sub Descuento_LostFocus()
If Val(Descuento.Text) > Val(SubTotal.Caption) Then
    MsgBox "Descuento MAYOR QUE SubTotal"
    Exit Sub
End If
Totales_Calcular
End Sub
Private Sub Fecha_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Fecha_LostFocus()
d = Fecha_Valida(Fecha, Now)
End Sub
Private Sub Condiciones_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub FechaaRecibir_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub FechaaRecibir_LostFocus()
d = Fecha_Valida(FechaaRecibir, Fecha.Text)
End Sub
Private Sub Atencion_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Entregar_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Cotizacion_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////
Private Sub Detalle_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
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
Detalle_Sumar
End Sub
Private Sub FilaBorrarContenido_Click()
' borra contenido de la fila en flexgrid
Dim fi As Integer, co As Integer
fi = Detalle.Row
For co = 1 To n_columnas
    Detalle.TextMatrix(fi, co) = ""
Next
Detalle_Sumar
End Sub
Private Sub OC_PrintLegal_EML_Esp(Numero_Oc)
cr.WindowTitle = "ORDEN DE COMPRA ESPECIAL Nº " & Numero_Oc
cr.ReportSource = crptReport
cr.WindowState = crptMaximized
'Cr.WindowBorderStyle = crptFixedSingle
cr.WindowMaxButton = False
cr.WindowMinButton = False
cr.Formulas(0) = "RAZON SOCIAL=""" & EmpOC.Razon & """"
cr.Formulas(1) = "GIRO=""" & "GIRO: " & EmpOC.Giro & """"
cr.Formulas(2) = "DIRECCION=""" & EmpOC.Direccion & """"
cr.Formulas(3) = "TELEFONOS=""" & "Teléfono: " & EmpOC.Telefono1 & " " & EmpOC.Comuna & """"
cr.Formulas(4) = "RUT=""" & "RUT: " & EmpOC.rut & """"
cr.Formulas(5) = "certificado=""" & IIf(m_Certificado, "NO SE RECEPCIONARA MATERIAL SIN CERTIFICADO DE CALIDAD ADJUNTO", "") & """" ' indica si envia ono certificado de calidad

'If Tipo = "E" Then
'    CR.Formulas(5) = "COTIZACION=""" & "GUÍA DESP. Nº:" & """"
'Else
'    CR.Formulas(5) = "COTIZACION=""" & "COTIZACIÓN Nº:" & """"
'End If

cr.DataFiles(0) = repo_file & ".MDB"
cr.ReportFileName = Drive_Server & Path_Rpt & "Oc_Especial.Rpt"
cr.Action = 1

End Sub
Private Sub OC_PrintLegal_C2_Esp(Numero_Oc)
cr.WindowTitle = "ORDEN DE COMPRA ESPECIAL Nº " & Numero_Oc
cr.ReportSource = crptReport
cr.WindowState = crptMaximized
'Cr.WindowBorderStyle = crptFixedSingle
cr.WindowMaxButton = False
cr.WindowMinButton = False
cr.Formulas(0) = "RAZON SOCIAL=""" & EmpOC.Razon & """"
cr.Formulas(1) = "GIRO=""" & "GIRO: " & EmpOC.Giro & """"
cr.Formulas(2) = "DIRECCION=""" & EmpOC.Direccion & """"
cr.Formulas(3) = "TELEFONOS=""" & "Teléfono: " & EmpOC.Telefono1 & " " & EmpOC.Comuna & """"
cr.Formulas(4) = "RUT=""" & "RUT: " & EmpOC.rut & """"
cr.Formulas(5) = "certificado=""" & IIf(m_Certificado, "NO SE RECEPCIONARA MATERIAL SIN CERTIFICADO DE CALIDAD ADJUNTO", "") & """" ' indica si envia ono certificado de calidad

'If Tipo = "E" Then
'    CR.Formulas(5) = "COTIZACION=""" & "GUÍA DESP. Nº:" & """"
'Else
'    CR.Formulas(5) = "COTIZACION=""" & "COTIZACIÓN Nº:" & """"
'End If

cr.DataFiles(0) = repo_file & ".MDB"
cr.ReportFileName = Drive_Server & Path_Rpt & "ocEspecialC2.Rpt"
cr.Action = 1

End Sub
Private Sub OC_PrintLegal_AYD_Esp(Numero_Oc)
cr.WindowTitle = "ORDEN DE COMPRA ESPECIAL Nº " & Numero_Oc
cr.ReportSource = crptReport
cr.WindowState = crptMaximized
'Cr.WindowBorderStyle = crptFixedSingle
cr.WindowMaxButton = False
cr.WindowMinButton = False
cr.Formulas(0) = "RAZON SOCIAL=""" & EmpOC.Razon & """"
cr.Formulas(1) = "GIRO=""" & "GIRO: " & EmpOC.Giro & """"
cr.Formulas(2) = "DIRECCION=""" & EmpOC.Direccion & """"
cr.Formulas(3) = "COMUNA=""" & EmpOC.Comuna & """"
cr.Formulas(4) = "TELEFONOS=""" & "Fono: " & EmpOC.Telefono1 & " Fax: " & EmpOC.Telefono2 & """"
cr.Formulas(5) = "RUT=""" & "RUT: " & EmpOC.rut & """"

cr.Formulas(6) = "certificado=""" & IIf(m_Certificado, "Enviar Certificado de calidad adjunto a factura", "") & """" ' indica si envia ono certificado de calidad
cr.Formulas(7) = "pagofactura=""" & "PAGO FACTURAS fono " & EmpOC.Telefono3 & """"

'Cr.Formulas(7) = ""
'Cr.Formulas(8) = ""

'If Tipo = "E" Then
    cr.Formulas(8) = "COTIZACION=""" & "GUÍA DESP. Nº:" & """"
'Else
'    Cr.Formulas(6) = "COTIZACION=""" & "COTIZACIÓN Nº:" & """"
'End If

cr.DataFiles(0) = repo_file & ".MDB"
cr.ReportFileName = Drive_Server & Path_Rpt & "Oc_Especial.Rpt"
cr.Action = 1

End Sub
Private Sub OC_PrintLegal_EIFFEL_Esp(Numero_Oc)
cr.WindowTitle = "ORDEN DE COMPRA ESPECIAL Nº " & Numero_Oc
cr.ReportSource = crptReport
cr.WindowState = crptMaximized
'Cr.WindowBorderStyle = crptFixedSingle
cr.WindowMaxButton = False
cr.WindowMinButton = False
cr.Formulas(0) = "RAZON SOCIAL=""" & EmpOC.Razon & """"
cr.Formulas(1) = "GIRO=""" & "GIRO: " & EmpOC.Giro & """"
cr.Formulas(2) = "DIRECCION=""" & EmpOC.Direccion & """"
cr.Formulas(3) = "TELEFONOS=""" & "Teléfono: " & EmpOC.Telefono1 & " " & EmpOC.Comuna & """"
cr.Formulas(4) = "RUT=""" & "RUT: " & EmpOC.rut & """"
cr.Formulas(5) = "certificado=""" & IIf(m_Certificado, "NO SE RECEPCIONARA MATERIAL SIN CERTIFICADO DE CALIDAD ADJUNTO", "") & """" ' indica si envia ono certificado de calidad

'If Tipo = "E" Then
'    CR.Formulas(5) = "COTIZACION=""" & "GUÍA DESP. Nº:" & """"
'Else
'    CR.Formulas(5) = "COTIZACION=""" & "COTIZACIÓN Nº:" & """"
'End If

cr.DataFiles(0) = repo_file & ".MDB"
cr.ReportFileName = Drive_Server & Path_Rpt & "Oc_Especial.Rpt"
cr.Action = 1

End Sub
Private Sub conIva_Click()
' deja iva en 0% o 19%
' y recalcula

Totales_Calcular

End Sub
Public Sub cuentaContableComboPoblar()
Dim i As Integer, Total As Integer
'Total = cuentasContablesLeer(aCuCo)
For i = 1 To cuentasContablesTotal
    If aCuCo(i, 2) = "S" Then
        ComboCtaContable.AddItem "   " & aCuCo(i, 0) & " " & aCuCo(i, 1)
    Else
        ComboCtaContable.AddItem aCuCo(i, 1)
    End If
Next
End Sub
Private Sub centroCostoComboPoblar()
Dim i As Integer, Total As Integer
'Total = centrosCostoLeer(aCeCo)
ComboCCosto.AddItem ""
For i = 1 To centrosCostoTotal
    If aCeCo(i, 2) = "S" Then
        ComboCCosto.AddItem "   " & aCeCo(i, 0) & " " & aCeCo(i, 1)
    Else
        ComboCCosto.AddItem aCeCo(i, 1)
    End If
Next
End Sub

