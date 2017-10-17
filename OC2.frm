VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form OC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Compra"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   12915
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ComboCCosto 
      Height          =   315
      Left            =   8400
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox ComboCtaContable 
      Height          =   315
      Left            =   8400
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   4320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ingresar Nueva Orden de Compra"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar Orden de Compra"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar Orden de Compra"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Orden de Compra"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anular Orden de Compra"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Grabar Orden de Compra"
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
            Object.ToolTipText     =   "Mantención de Productos"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir OC Copia2"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CheckBox conIva 
      Alignment       =   1  'Right Justify
      Caption         =   "Con IVA"
      Height          =   255
      Left            =   7680
      TabIndex        =   22
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox Nv 
      Height          =   300
      Left            =   840
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Condiciones 
      Height          =   285
      Left            =   4920
      TabIndex        =   10
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Atencion 
      Height          =   300
      Left            =   840
      TabIndex        =   14
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CheckBox Certificado 
      Caption         =   "No se recepcionará material sin Certificado de Calidad adjunto"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   6840
      Width           =   4815
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   9240
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.ComboBox EntregarEn 
      Height          =   315
      Left            =   5160
      TabIndex        =   16
      Top             =   1920
      Width           =   3375
   End
   Begin MSMask.MaskEdBox Cotizacion 
      Height          =   300
      Left            =   1320
      TabIndex        =   18
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "##########"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MEdit 
      Height          =   375
      Left            =   9000
      TabIndex        =   46
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   327680
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Descuento 
      Height          =   300
      Left            =   6720
      TabIndex        =   21
      Top             =   5865
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "##########"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox FechaaRecibir 
      Height          =   300
      Left            =   7680
      TabIndex        =   12
      Top             =   1560
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
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   3
      Left            =   240
      MaxLength       =   50
      TabIndex        =   27
      Top             =   6405
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   2
      Left            =   240
      MaxLength       =   50
      TabIndex        =   26
      Top             =   6105
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   1
      Left            =   240
      MaxLength       =   50
      TabIndex        =   25
      Top             =   5805
      Width           =   5000
   End
   Begin VB.Frame Frame 
      Caption         =   "PROVEEDOR"
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   2280
      TabIndex        =   33
      Top             =   400
      Width           =   6375
      Begin VB.ComboBox Direccion 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   600
         Width           =   3255
      End
      Begin VB.CommandButton btnSearch 
         Height          =   300
         Left            =   2040
         Picture         =   "OC2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Busca -Proveedor"
         Top             =   240
         Width           =   300
      End
      Begin MSMask.MaskEdBox Rut 
         Height          =   300
         Left            =   600
         TabIndex        =   31
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label lbl 
         Caption         =   "COM"
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   47
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Comuna 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4440
         TabIndex        =   43
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Razon 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2520
         TabIndex        =   42
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label lbl 
         Caption         =   "DIR"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lbl 
         Caption         =   "&RUT"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   0
      Left            =   240
      MaxLength       =   50
      TabIndex        =   24
      Top             =   5505
      Width           =   5000
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   840
      TabIndex        =   3
      Top             =   1080
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
      TabIndex        =   1
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
      Height          =   2565
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4524
      _Version        =   327680
      Enabled         =   -1  'True
      ScrollBars      =   2
   End
   Begin MSMask.MaskEdBox pDescuento 
      Height          =   300
      Left            =   5640
      TabIndex        =   20
      Top             =   5565
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   529
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   2
      Mask            =   "##"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   9120
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OC2.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OC2.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OC2.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OC2.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OC2.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OC2.frx":065C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OC2.frx":076E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OC2.frx":0880
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OC2.frx":0992
            Key             =   ""
         EndProperty
      EndProperty
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
      Left            =   10440
      TabIndex        =   52
      Top             =   2280
      Width           =   2055
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
      Left            =   9120
      TabIndex        =   51
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Cotización &Nº"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblpDescuento 
      Caption         =   "% Desc"
      Height          =   255
      Left            =   6000
      TabIndex        =   45
      Top             =   5565
      Width           =   615
   End
   Begin VB.Label pDesc_a_Dinero 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6720
      TabIndex        =   44
      Top             =   5565
      Width           =   975
   End
   Begin VB.Label lblTotal 
      Caption         =   "TOTAL"
      Height          =   255
      Left            =   5640
      TabIndex        =   41
      Top             =   6765
      Width           =   975
   End
   Begin VB.Label lblIva 
      Caption         =   "IVA"
      Height          =   255
      Left            =   5640
      TabIndex        =   40
      Top             =   6465
      Width           =   975
   End
   Begin VB.Label lblNeto 
      Caption         =   "NETO"
      Height          =   255
      Left            =   5640
      TabIndex        =   39
      Top             =   6165
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
      TabIndex        =   38
      Top             =   6765
      Width           =   975
   End
   Begin VB.Label Iva 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6720
      TabIndex        =   37
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Neto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6720
      TabIndex        =   36
      Top             =   6165
      Width           =   975
   End
   Begin VB.Label lblDescuento 
      Caption         =   "DESCUENTO"
      Height          =   255
      Left            =   5640
      TabIndex        =   35
      Top             =   5865
      Width           =   1095
   End
   Begin VB.Label lblSubTotal 
      Caption         =   "SUBTOTAL"
      Height          =   255
      Left            =   5640
      TabIndex        =   29
      Top             =   5265
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&Entregar en:"
      Height          =   255
      Index           =   11
      Left            =   4080
      TabIndex        =   15
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&At. Sr:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "a Reci&bir"
      Height          =   255
      Index           =   9
      Left            =   6960
      TabIndex        =   11
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "&Condic."
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label SubTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6720
      TabIndex        =   30
      Top             =   5265
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   23
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "&OBRA"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "&FECHA"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   615
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
      Left            =   120
      TabIndex        =   0
      Top             =   600
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
Attribute VB_Name = "OC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button, btnImprimir As Button
Private btnAnular As Button, btnGrabar As Button, btnDesHacer As Button, btnImprimirC2 As Button
Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, d As Variant

Private DbD As Database, RsPrv As Recordset, RsProDir As Recordset, RsPrd As Recordset
Private Dbm As Database, RsNVc As Recordset
Private DbAdq As Database, RsOcc As Recordset, RsOCdN As Recordset, RsCorre As Recordset
Private RsRMd As Recordset
'Private DbHAdq As Database, RsHOCc As Recordset, RsHOCdN As Recordset

Private n_filas As Integer, n_columnas As Integer
Private prt As Printer
Private n3 As Double, n6 As Double, n7 As Double ', n8 As Double
Private NV_Numero As Double, NV_Obra As String
Private m_OcNumero As Double, primera As Boolean
Private m_Certificado As Boolean
Private a_Nv(2999, 3) As String, m_NvArea As Integer
' 0: nv
' 1: obra
' 2: ccCodigo
' 3: ccDescripcion

Private aLineas(99, 1) As String ' aqui van los codigos de las cuentas contables (0) y codigos de centros de costo(1), elegidos para cada linea

Private m_ImpresoraNombre As String
Private FormatoDoc As String
Private Const columnaLargoEspecial As Integer = 11
Private miNv As NotaVenta
Public Property Let NumeroOC(ByVal New_Numero As Double)
m_OcNumero = New_Numero
End Property
Private Sub ComboCCosto_LostFocus()
ComboCCosto.visible = False
End Sub
Private Sub ComboCtaContable_Click()
If ComboCtaContable.ListIndex = -1 Then Exit Sub
Dim fi As Integer, paso As String, j As Integer
fi = Detalle.Row
If aCuCo(ComboCtaContable.ListIndex + 1, 2) = "S" Then

    paso = Trim(ComboCtaContable.Text)
    
    paso = Mid(paso, 1, InStr(1, paso, " "))
    paso = Trim(paso)
            
    j = cuentaContableBuscarIndice(NoNulo(paso))
    If j > 0 Then
        Detalle.TextMatrix(fi, 9) = aCuCo(j, 0) & " " & aCuCo(j, 1)
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
Set RsPrd = DbD.OpenRecordset("Productos")
RsPrd.Index = "Codigo"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set DbAdq = OpenDatabase(Madq_file)
Set RsOcc = DbAdq.OpenRecordset("OC Cabecera")
RsOcc.Index = "Numero"
Set RsOCdN = DbAdq.OpenRecordset("OC Detalle")
RsOCdN.Index = "Numero-Linea"

Set RsRMd = DbAdq.OpenRecordset("Documentos")
RsRMd.Index = "oc-linea-rm"

Set RsCorre = DbAdq.OpenRecordset("Correlativo")

i = 0
' Combo obra
ComboNV.AddItem " "
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

'Debug.Print "i=|"; i; "|"

Inicializa
Detalle_Config

Privilegios

'StatusBar.Panels(1) = EmpOC.Razon
'StatusBar.Panels(2) = IIf(Usuario.Adquis_Actual, "Archivo: ACTUAL", "Archivo: HISTÓRICO")

primera = True
conIva.Value = 1

m_NvArea = 0

FormatoDoc = ReadIniValue(Path_Local & "scp.ini", "GD", "Formato")

End Sub
Private Sub Form_Activate()
' 21/01/2000
If primera Then
    If m_OcNumero <> 0 Then
        Tulbar 4 'boton imprimir
        Numero.Text = m_OcNumero
        After_Enter
    End If
    primera = False
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
DataBases_Cerrar
End Sub
Private Sub Inicializa()
Dim Formato As String

Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnAnular = Toolbar.Buttons(5)
Set btnGrabar = Toolbar.Buttons(7)
Set btnDesHacer = Toolbar.Buttons(8)
Set btnImprimirC2 = Toolbar.Buttons(13)

Obj = "ORDEN DE COMPRA"
Objs = "ÓRDENES DE COMPRA"

Accion = ""

btnSearch.visible = False
btnSearch.ToolTipText = "Busca Proveedores"
Campos_Enabled False

pDescuento.Format = "0"
Descuento.Format = num_Format0

'm_AbrMetro = "MET"
MEdit.PromptInclude = False

Formato = ReadIniValue(Path_Local & "scp.ini", "GD", "Formato")

'If Formato = "EML" Then
'    EntregarEn.AddItem "LAS ACACIAS 02500"
'End If
'If Formato = "AYD" Then
'    EntregarEn.AddItem "STA. ALEJANDRA 03521"
'End If

EntregarEn.AddItem Empresa.Direccion

EntregarEn.AddItem "RETIRAMOS"
EntregarEn.AddItem "OBRA"
EntregarEn.AddItem "STA. ALEJANDRA 03521"

condiciones.MaxLength = 20
atencion.MaxLength = 50

centroCostoComboPoblar
cuentaContableComboPoblar

End Sub
Private Sub Detalle_Config()
Dim i As Integer, ancho As Integer, c_totales As Integer, ancho_totales

n_filas = 19
n_columnas = columnaLargoEspecial

ancho_totales = 1000

Detalle.Left = 100
Detalle.WordWrap = True
Detalle.RowHeight(0) = 450
Detalle.Rows = n_filas + 1
Detalle.Cols = n_columnas + 1

Detalle.TextMatrix(0, 0) = ""
Detalle.TextMatrix(0, 1) = "Código"
Detalle.TextMatrix(0, 2) = "L"
Detalle.TextMatrix(0, 3) = "Cantidad"
Detalle.TextMatrix(0, 4) = "Uni."              '*
Detalle.TextMatrix(0, 5) = "Descripción"       '*
Detalle.TextMatrix(0, 6) = "Largo (mm)"        '* digitable y no
'Detalle.TextMatrix(0, 7) = "Medida Total (m)" '*
Detalle.TextMatrix(0, 7) = "Precio Unitario"
Detalle.TextMatrix(0, 8) = "Precio TOTAL"      '*
Detalle.TextMatrix(0, 9) = "Cuenta Contable"
Detalle.TextMatrix(0, 10) = "Centro Costo"

' solo para largo especial
Detalle.TextMatrix(0, columnaLargoEspecial) = ""
Detalle.ColWidth(columnaLargoEspecial) = 0

Detalle.ColWidth(0) = 300
Detalle.ColWidth(1) = 1500
Detalle.ColWidth(2) = 300
Detalle.ColWidth(3) = 1000
Detalle.ColWidth(4) = 500
Detalle.ColWidth(5) = 3500 ' 2900
Detalle.ColWidth(6) = 1000
Detalle.ColWidth(7) = 1000
Detalle.ColWidth(8) = ancho_totales
Detalle.ColWidth(9) = 2000
Detalle.ColWidth(10) = 2000

ancho = 350 ' con scroll vertical

SubTotal.Width = ancho_totales
pDesc_a_Dinero.Width = ancho_totales
Descuento.Width = ancho_totales
Neto.Width = ancho_totales
Iva.Width = ancho_totales
Total.Width = ancho_totales

For i = 0 To n_columnas

    If i = 8 Then
    
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
Me.Width = ancho + Detalle.Left * 2.5

For i = 1 To n_filas
    Detalle.TextMatrix(i, 0) = i
Next

For i = 1 To n_filas
    Detalle.Row = i
    Detalle.col = 4
    Detalle.CellForeColor = vbRed
    Detalle.col = 5
    Detalle.CellForeColor = vbRed
    Detalle.col = 7
    Detalle.CellForeColor = vbRed
Next

MEdit.Text = ""

'Top = Form_CentraY(Me)
'Left = Form_CentraX(Me)

'Detalle.Enabled = False
Detalle.col = 1
Detalle.Row = 1

End Sub
Private Sub Nv_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Nv_LostFocus()
Dim m_Nv As Integer
m_Nv = Val(Nv.Text)
If m_Nv = 0 Then Exit Sub

' busca centro de costo
'miNv = nv_Buscar(Int(m_Nv))
'Debug.Print "numero|" & miNv.Numero & "|" & miNv.centroCostoCodigo & "|"

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

If ComboNV.Text = " " Then Nv.Text = "": Exit Sub
MousePointer = vbHourglass
'i = 0
NV_Numero = Val(Left(ComboNV.Text, 4))
NV_Obra = Mid(ComboNV.Text, 8)
Nv.Text = NV_Numero
Detalle.Enabled = True

indice = ComboNV.ListIndex

'Debug.Print indice; " "; a_Nv(indice, 0); " "; a_Nv(indice, 1); " "; a_Nv(indice, 2); " "; a_Nv(indice, 3)

lblccCodigo.Caption = a_Nv(indice, 2)
lblccDescripcion.Caption = a_Nv(indice, 3)

MousePointer = vbDefault

End Sub
Private Sub Numero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    After_Enter
End If
End Sub
Private Sub After_Enter()
Select Case Accion
Case "Agregando"
    If Val(Numero.Text) <= 0 Then
        MsgBox "Número NO Válido"
        Exit Sub
    End If
    RsOcc.Seek "=", Numero
    If RsOcc.NoMatch Then
    
        Campos_Enabled True
        Numero.Enabled = False
'        Detalle.Enabled = False
        
        Fecha.SetFocus
        btnGrabar.Enabled = True
        btnSearch.visible = True
        
    Else
    
        Doc_Leer
        MsgBox Obj & " YA EXISTE"
'        Detalle.Enabled = False
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
        
    End If
    
Case "Modificando"
    RsOcc.Seek "=", Numero.Text
    If RsOcc.NoMatch Then
'        RsHOCc.Seek "=", Numero.Text
'        If RsHOCc.NoMatch Then
            MsgBox Obj & " NO EXISTE"
'        Else
            ' lo encontró en histórico
'            Doc_Leer
'            GoTo Modificar_Leer
'        End If
    Else
        Doc_Leer
Modificar_Leer:
        If RsOcc!Tipo = "N" Then
            Campos_Enabled True
            Numero.Enabled = False
            btnGrabar.Enabled = True
            btnSearch.visible = True
        Else
            MsgBox "DEBE MODIFICAR COMO OC ESPECIAL"
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
        If RsOcc!Tipo = "N" Then
            Numero.Enabled = False
            If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo) = vbYes Then
                Doc_Eliminar
            End If
        Else
            MsgBox "DEBE ELIMINAR COMO OC ESPECIAL"
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
        If RsOcc!Tipo = "N" Then
            Numero.Enabled = False
            Detalle.Enabled = True
        Else
            MsgBox "DEBE IMPRIMIR COMO OC ESPECIAL"
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
        If RsOcc!Tipo = "N" Then
            Numero.Enabled = False
            If MsgBox("¿ ANULA " & Obj & " ?", vbYesNo) = vbYes Then
                Doc_Anular
            End If
        Else
            MsgBox "DEBE ANULAR COMO OC ESPECIAL"
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
        If RsOcc!Tipo = "N" Then
            Numero.Enabled = False
            Detalle.Enabled = True
        Else
            MsgBox "DEBE IMPRIMIR COMO OC ESPECIAL"
            Campos_Limpiar
            Numero.Enabled = True
            Numero.SetFocus
        End If
    End If
    
End Select

End Sub
Private Sub Doc_Leer()
Dim m_Unidad As String, m_descri As String, m_LargoE As Boolean
Dim j As Integer
' CABECERA

With RsOcc

Fecha.Text = Format(!Fecha, Fecha_Format)
NV_Numero = !Nv
rut.Text = ![RUT Proveedor]
condiciones.Text = NoNulo(![Condiciones de Pago])
FechaaRecibir.Text = Format(![Fecha a Recibir], Fecha_Format)
EntregarEn.Text = NoNulo(![Entregar en])
Cotizacion.Text = !Cotizacion

On Error Resume Next
NV_Obra = ""
RsNVc.Seek "=", NV_Numero, m_NvArea
If Not RsNVc.NoMatch Then
    NV_Obra = RsNVc!obra
    ComboNV.Text = Format(NV_Numero, "0000") & " - " & NV_Obra
End If
On Error GoTo 0

SubTotal.Caption = !SubTotal
pDescuento.Text = ![% Descuento]
Descuento.Text = !Descuento
'Neto.Caption = !Neto
'Iva.Caption = !Iva
'Total.Caption = !Total

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

'primera = True

If RsOcc!Tipo = "E" Then
Else
    
    With RsOCdN
    
    .Seek "=", Numero.Text, 1
    
    If Not .NoMatch Then
        
        Do While Not .EOF
            
            If !Numero = Numero.Text Then
                        
                i = !linea
                               
                Detalle.TextMatrix(i, 1) = ![codigo producto]
                Detalle.TextMatrix(i, 2) = NoNulo(![Largo Especial])
                Detalle.TextMatrix(i, 3) = !Cantidad
                
                m_Unidad = "": m_descri = "": m_LargoE = False
                RsPrd.Seek "=", ![codigo producto]
                If Not RsPrd.NoMatch Then
                    m_Unidad = RsPrd![unidad de medida]
                    m_descri = RsPrd!Descripcion
                    m_LargoE = RsPrd![Largo Especial]
                End If
                Detalle.TextMatrix(i, 4) = m_Unidad
                Detalle.TextMatrix(i, 5) = m_descri
                
                Detalle.TextMatrix(i, 6) = !largo
'                Detalle.TextMatrix(i, 7) = ![Medida Total]
                
                Detalle.TextMatrix(i, 7) = ![Precio Unitario]
                
                j = cuentaContableBuscarIndice(NoNulo(!cuentacontable))
                If j > 0 Then
                    Detalle.TextMatrix(i, 9) = aCuCo(j, 0) & " " & aCuCo(j, 1)
                    aLineas(i, 0) = aCuCo(j, 0)
                End If
                j = centroCostoBuscarIndice(NoNulo(!CentroCosto))
                If j > 0 Then
                    Detalle.TextMatrix(i, 10) = aCeCo(j, 0) & " " & aCeCo(j, 1)
                    aLineas(i, 1) = aCeCo(j, 0)
                End If
                                
                Detalle.TextMatrix(i, columnaLargoEspecial) = m_LargoE
                
                Fila_Calcular i, False
                
            Else
                Exit Do
            End If
            .MoveNext
        Loop
    End If
    End With
End If

Proveedor_Lee rut.Text
If Direccion.ListCount > 0 Then Direccion.ListIndex = RsOcc![Codigo Direccion]
' "Atencion" esta aquí porque esta ligado a proveedor
atencion.Text = NoNulo(RsOcc!atencion)

Detalle.Row = 1 ' para q' actualice la primera fila del detalle

End Sub
Private Sub Proveedor_Lee(rut)
RsPrv.Seek "=", rut
If RsPrv.NoMatch Then
    MsgBox "PROVEEDOR NO EXISTE"
Else
    Razon.Caption = RsPrv![Razon Social]
    
    '/////////
    Direccion.Clear
    Direccion.AddItem RsPrv!Direccion
    Direccion.Text = RsPrv!Direccion 'es como un click
    
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
d = Fecha_Valida(FechaaRecibir, Now)
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
Private Function Doc_Validar() As Boolean
Dim porDespachar As Integer, m_Can As Double, m_Lar As Double, m_Entero As Double
Doc_Validar = False

If Val(NV_Numero) = 0 Then
    MsgBox "DEBE ESCOGER OBRA"
    ComboNV.SetFocus
    Exit Function
End If

If rut.Text = "" Then
    MsgBox "DEBE ELEGIR PROVEEDOR"
    btnSearch.SetFocus
    Exit Function
End If

If FechaaRecibir.Text = "__/__/__" Then
    MsgBox "DEBE DIGITAR FECHA"
    FechaaRecibir.SetFocus
    Exit Function
End If

'If cc.Text = "" Then
'    MsgBox "Debe Escoger Centro de Costo"
'    cc.SetFocus
'    Exit Function
'End If

For i = 1 To n_filas

    If Trim(Detalle.TextMatrix(i, 1)) <> "" Then

        ' cantidad
        If Not Numero_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
        ' largo
        If Not Numero_Valida(Detalle.TextMatrix(i, 6), i, 6) Then Exit Function
        ' medida total
        If Not Numero_Valida(Detalle.TextMatrix(i, 7), i, 7) Then Exit Function
        ' precio uni
        If Not Numero_Valida(Detalle.TextMatrix(i, 8), i, 8) Then Exit Function

        If Detalle.TextMatrix(i, 9) = "" Then
            MsgBox "Debe Escoger Cuenta Contable"
            Detalle.Row = i
            Detalle.col = 9
            Detalle.SetFocus
            Exit Function
        End If
        
        ' valida que:
        ' si unidad es MTS o MET =>
        ' cantidad / (largo/1000) = debe ser entero
'        If Detalle.TextMatrix(i, 2) = "E" Then
        If Detalle.TextMatrix(i, 4) = "MTS" Or Detalle.TextMatrix(i, 4) = "MET" Then
        
            m_Can = CDbl(Detalle.TextMatrix(i, 3))
            m_Lar = CDbl(Detalle.TextMatrix(i, 6))
            
            If m_Lar = 0 Then
                MsgBox "El Producto " & Trim(Detalle.TextMatrix(i, 1)) & vbLf & "cuya unidad es MTS o MET, tiene largo=0"
                Detalle.Row = i
                Detalle.col = 3 ' foco en largo especial
                Detalle.SetFocus
                Exit Function
            Else
            
                m_Entero = m_Can * 1000 / m_Lar
            
                If m_Entero = Int(m_Can * 1000 / m_Lar) Then
                    ' cantidad y largo ok
                Else
                    MsgBox "Cantidad o Largo Erroneo(s)"
                    Detalle.Row = i
                    Detalle.col = 3 ' foco en largo especial
                    Detalle.SetFocus
                    Exit Function
                End If
            End If
            
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
Private Function Doc_Grabar(Nueva As Boolean) As Double
MousePointer = vbHourglass
Dim m_Numero As Double
Dim m_cantidad As String
Dim m_CanRec As Double, m_FecRec As Date
Dim qry As String
Doc_Grabar = 0 'Numero.Text
save:
' CABECERA DE OC
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
    !Tipo = "N"

Else

    Doc_Detalle_Eliminar
    .Edit
    !fechaModificacion = Format(Now, Fecha_Format)

End If

!Fecha = Fecha.Text
!Nv = NV_Numero
![RUT Proveedor] = rut.Text
![Codigo Direccion] = Direccion.ListIndex
![Condiciones de Pago] = condiciones.Text
![Fecha a Recibir] = FechaaRecibir.Text
!atencion = atencion.Text
![Entregar en] = EntregarEn.Text
!Cotizacion = Val(Cotizacion.Text)
![Observacion 1] = Obs(0).Text
![Observacion 2] = Obs(1).Text
![Observacion 3] = Obs(2).Text
![Observacion 4] = Obs(3).Text
!SubTotal = Val(SubTotal.Caption)
![% Descuento] = Val(pDescuento.Text)
!Descuento = Val(Descuento.Text)
!Neto = Val(Neto.Caption)
!Iva = Val(Iva.Caption)
!Total = Val(Total.Caption)
!Pendiente = True
!Nula = False

!Certificado = Certificado.Value

.Update

End With

' DETALLE DE OC
j = 0

With RsOCdN
For i = 1 To n_filas
    m_cantidad = Detalle.TextMatrix(i, 3)
    If Val(m_cantidad) <> 0 Then
        
        .AddNew
        !Numero = Numero.Text
        !Tipo = "N"
        j = j + 1
        !linea = j
        !Fecha = Fecha.Text
        !Nv = NV_Numero
        ![RUT Proveedor] = rut.Text
        ![Fecha a Recibir] = FechaaRecibir.Text
        ![codigo producto] = Detalle.TextMatrix(i, 1)
        ![Largo Especial] = Detalle.TextMatrix(i, 2)
        !Cantidad = m_CDbl(m_cantidad) 'esta es la forma de grabar?? 1,5 -> 1.5
        !largo = Val(Detalle.TextMatrix(i, 6))
'        ![Medida Total] = Val(Detalle.TextMatrix(i, 7))
        ![Precio Unitario] = m_CDbl(Detalle.TextMatrix(i, 7))
        !Total = Val(Detalle.TextMatrix(i, 8))

        !cuentacontable = aLineas(i, 0)
        !CentroCosto = aLineas(i, 1)
        
        ' actualiza cantidades recibidas desde recepcion de materiales
        ' 19/07/05
        m_CanRec = 0
        m_FecRec = Date
        RsRMd.Seek ">=", Numero.Text, i, 0
        If Not RsRMd.NoMatch Then ' agregado 26/08/05 para felix
            Do While Not RsRMd.EOF
                If RsRMd![Oc] <> Numero.Text Or RsRMd![Linea Oc] <> i Then Exit Do
                m_CanRec = m_CanRec + RsRMd![Cant_Entra]
                m_FecRec = RsRMd!Fecha
                RsRMd.MoveNext
            Loop
        End If
        If m_CanRec > 0 Then
            ![Cantidad Recibida] = m_CanRec
            ![Fecha Recepcion] = m_FecRec
        End If
        '/////////////////////////////////////////////////////////////
        
        !Pendiente = True

        .Update
        
        ' actualiza fecha, cantidad y precio de compra
        ' en maestro de productos
        
        ' ademas debe calcular precio promedio ponderado ultimos 6 meses
        'aqui voy
        
        RsPrd.Seek "=", Detalle.TextMatrix(i, 1)
        If Not RsPrd.NoMatch Then
            RsPrd.Edit
            RsPrd![Fecha Compra] = Fecha.Text
            RsPrd![Cantidad Compra] = m_CDbl(m_cantidad)
            RsPrd![Precio Compra] = m_CDbl(Detalle.TextMatrix(i, 7))
            RsPrd.Update
        End If
        
    End If
Next
End With

DesBloqueo "OC", RsCorre

MousePointer = vbDefault

End Function
Private Sub Doc_Eliminar()

' elimina cabecera
RsOcc.Seek "=", Numero.Text
If Not RsOcc.NoMatch Then

    RsOcc.Delete
   
End If

' elimina detalle
Doc_Detalle_Eliminar

End Sub
Private Sub Doc_Detalle_Eliminar()

DbAdq.Execute "DELETE * FROM [OC Detalle] WHERE Numero=" & Numero.Text

Exit Sub

With RsOCdN
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

lblccCodigo.Caption = ""
lblccDescripcion.Caption = ""

rut.Text = ""
Razon.Caption = ""
Direccion.Clear
Comuna.Caption = ""
condiciones.Text = ""
FechaaRecibir.Text = Fecha_Vacia
atencion.Text = ""
EntregarEn.Text = ""
Cotizacion.Text = ""
'cc.Text = ""
'ComboCCosto.Text = ""
Detalle_Limpiar

conIva.Value = 1

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
Detalle.col = 1
Detalle.Row = 1
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
Tulbar Button.Index
End Sub
Private Sub Tulbar(Boton_Indice As Integer)
Dim cambia_titulo As Boolean, m_Numero As String, m_obra As String
cambia_titulo = True
'Accion = ""
Select Case Boton_Indice
Case 1 ' agregar
    Accion = "Agregando"
    Botones_Enabled 0, 0, 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    
'    Numero.Text = Documento_Numero_Nuevo(RsOCc, "Número")
'    Numero.Enabled = True
'    Numero.SetFocus

'//////// agregado 08/09/99
    Campos_Enabled True
    Numero.Enabled = False
    Fecha.SetFocus
    btnGrabar.Enabled = True
    btnSearch.visible = True
'///////////////////////////////

Case 2 ' modificar
    Accion = "Modificando"
    Botones_Enabled 0, 0, 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    Numero.Enabled = True
    Numero.SetFocus
Case 3 ' eliminar
    Accion = "Eliminando"
    Botones_Enabled 0, 0, 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    Numero.Enabled = True
    Numero.SetFocus
Case 4 ' imprimir
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
            OC_Prepara m_Numero, NV_Numero, NV_Obra, m_ImpresoraNombre
            
            Campos_Limpiar
            Numero.Enabled = True
            Numero.SetFocus
            
            MousePointer = vbDefault
            
            If False Then
                If Empresa.rut = "89784800-7" Then
                    OC_PrintLegal_EML m_Numero
                End If
                If Empresa.rut = "76406180-2" Then
                    OC_PrintLegal_AYD m_Numero
                End If
                If Empresa.rut = "76008108-6" Then
                    OC_PrintLegal_Delsa m_Numero
                End If
            End If
            
            Select Case FormatoDoc
            Case "EML"
                OC_PrintLegal_EML m_Numero
            Case "AYD"
                OC_PrintLegal_AYD m_Numero
            Case "DELSA"
                OC_PrintLegal_Delsa m_Numero
            Case "EIFFEL"
                OC_PrintLegal_Eiffel m_Numero
            End Select
'        End If

'        Campos_Limpiar
'        Numero.Enabled = True
'        Numero.SetFocus
    End If
    
Case 5 ' Anular

    Accion = "Anulando"
    Botones_Enabled 0, 0, 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    Numero.Enabled = True
    Numero.SetFocus
    
Case 7 ' grabar

    If Doc_Validar Then
        
        If MsgBox("¿ GRABAR " & Obj & " ?", vbYesNo) = vbYes Then
        
            m_Numero = Numero.Text
            NV_Numero = Nv.Text
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
                OC_Prepara m_Numero, NV_Numero, m_obra, m_ImpresoraNombre
                
                
                If False Then
                    If Empresa.rut = "89784800-7" Then
                        OC_PrintLegal_EML m_Numero ', RsOcc!Tipo
                    End If
                    If Empresa.rut = "76406180-2" Then
                        OC_PrintLegal_AYD m_Numero
                    End If
                    If Empresa.rut = "76008108-6" Then
                        OC_PrintLegal_Delsa m_Numero
                    End If
                End If

                Select Case FormatoDoc
                Case "EML"
                    OC_PrintLegal_EML m_Numero
                Case "AYD"
                    OC_PrintLegal_AYD m_Numero
                Case "DELSA"
                    OC_PrintLegal_Delsa m_Numero
                Case "EIFFEL"
                    OC_PrintLegal_Eiffel m_Numero
                End Select

'                End If

            End If
            
        End If
        
    End If
    
Case 8 ' DesHacer

    If Numero.Text = "" Then
        Privilegios
        Campos_Limpiar
        Campos_Enabled False
        btnSearch.visible = False
    Else
        If Accion = "Imprimiendo" Then
            Privilegios
            Campos_Limpiar
            Campos_Enabled False
            btnSearch.visible = False
        Else
            If MsgBox("¿Seguro que quiere DesHacer?", vbYesNo) = vbYes Then
                Privilegios
                Campos_Limpiar
                Campos_Enabled False
                btnSearch.visible = False
            End If
        End If
    End If
    Accion = ""

Case 9  ' separador
Case 10 ' proveedores
    MousePointer = vbHourglass
    Load Proveedores
    MousePointer = vbDefault
    Proveedores.Show 1
    cambia_titulo = False
Case 11 ' productos
    MousePointer = vbHourglass
    Load Productos
    MousePointer = vbDefault
    Productos.Show 1
    cambia_titulo = False

Case 13 ' imprimir copia 2 (para contabilidad)

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

        MousePointer = vbHourglass
        m_Numero = Numero.Text
        
        OC_PreparaC2 m_Numero, NV_Numero, NV_Obra, m_ImpresoraNombre
        
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
        
        MousePointer = vbDefault
                
        Select Case FormatoDoc
        Case "EML"
            OC_PrintLegalC2 m_Numero '?
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
Direccion.Enabled = Si
Nv.Enabled = Si
ComboNV.Enabled = Si
condiciones.Enabled = Si
FechaaRecibir.Enabled = Si
atencion.Enabled = Si
EntregarEn.Enabled = Si
Cotizacion.Enabled = Si

ComboCCosto.Enabled = Si
ComboCtaContable.Enabled = Si

Detalle.Enabled = Si
pDescuento.Enabled = Si
Descuento.Enabled = Si
conIva.Enabled = Si
Obs(0).Enabled = Si
Obs(1).Enabled = Si
Obs(2).Enabled = Si
Obs(3).Enabled = Si
'Certificado.Enabled = Si
End Sub
Private Sub btnSearch_Click()

Search.Muestra data_file, "Proveedores", "RUT", "Razon Social", "Proveedor", "Proveedores"
rut.Text = Search.Codigo

If rut.Text <> "" Then
    Proveedor_Lee (rut.Text)
End If
End Sub
Private Sub Producto_Buscar()
Dim m_Codigo As String, fi As Integer
fi = Detalle.Row
m_Codigo = Detalle.TextMatrix(fi, 1)

With RsPrd
.Seek "=", m_Codigo
If .NoMatch Then
    MsgBox "PRODUCTO NO EXISTE"
    For i = 1 To n_columnas
        Detalle.TextMatrix(fi, i) = ""
    Next
    Detalle.TextMatrix(fi, 11) = False
Else
    Detalle.TextMatrix(fi, 2) = ""
    Detalle.TextMatrix(fi, 3) = 0
    Detalle.TextMatrix(fi, 4) = ![unidad de medida]
    Detalle.TextMatrix(fi, 5) = !Descripcion
    Detalle.TextMatrix(fi, 6) = !largo
    Detalle.TextMatrix(fi, 7) = 0
    Detalle.TextMatrix(fi, 8) = 0
    Detalle.TextMatrix(fi, 9) = ""
    Detalle.TextMatrix(fi, columnaLargoEspecial) = ![Largo Especial]
End If
End With

End Sub
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
Select Case Detalle.col
    Case 1 ' codigo
    Case 9 ' cuenta contable
        On Error GoTo error1
        If Detalle <> "" Then ComboCtaContable.Text = "   " & Detalle
error1:
        On Error GoTo 0
        ComboCtaContable.Top = Detalle.CellTop + Detalle.Top
        ComboCtaContable.Left = Detalle.CellLeft + Detalle.Left
        ComboCtaContable.Width = Int(Detalle.CellWidth * 1.5)
        ComboCtaContable.visible = True
    Case 10 ' centro costo
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
MSFlexGridEdit Detalle, MEdit, 32
End Sub
Private Sub Detalle_GotFocus()
If MEdit.visible Then
    Detalle = MEdit
    MEdit.visible = False
End If
End Sub
Private Sub Detalle_LeaveCell()
If MEdit.visible Then
    Detalle = MEdit
    MEdit.visible = False
End If
End Sub
Private Sub Detalle_KeyPress(KeyAscii As Integer)
If Accion = "Imprimiendo" Then Exit Sub
'If Right(Trim(Detalle.TextMatrix(Detalle.Row, 1)), 1) = "E" And Detalle.col = 5 Then Exit Sub
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
    Select Case m_col
    Case 1 ' Código producto
        MSFlexGrid.SetFocus
        DoEvents
        Producto_Buscar
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
        
        'aqui
'        If m_col = 3 Then
'            Celda_Calcular m_fil
'        End If
        
        Fila_Calcular m_fil, True
    End Select
    Cursor_Mueve MSFlexGrid
Case vbKeyUp ' Flecha Arriba
    Select Case m_col
    Case 1 ' Codigo
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
        'aqui
        Fila_Calcular m_fil, True
        If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
            MSFlexGrid.Row = MSFlexGrid.Row - 1
        End If
    End Select
Case vbKeyDown ' Flecha Abajo
    Select Case m_col
    Case 1 ' Codigo
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
        'aqui
        Fila_Calcular m_fil, True
        If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
            MSFlexGrid.Row = MSFlexGrid.Row + 1
        End If
    End Select
End Select
End Sub

Private Sub Celda_Calcular(Fila As Integer)
' calcula formula en celda (tipo excel)
Dim s As String
s = Detalle.TextMatrix(Fila, 3) 'can
On Error GoTo ErrorFormula
n3 = m_CDbl(s)
Detalle.TextMatrix(Fila, 3) = n3
Exit Sub
ErrorFormula:
MsgBox "Error en Formula"
End Sub

Private Sub MEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)
Dim m_Codigo As String

Select Case MSFlexGrid.col
Case 1 'codigo
    Edt.Mask = ">&&&&&&&&&&&&&&&&&&&&"
Case 2 'largo especial
    Edt.Mask = ">&"
Case 4 'unidad
    Edt.Mask = ">&&&"
Case Else
    Edt.Mask = "&&&&&&&&&&"
End Select

Select Case MSFlexGrid.col
Case 2 'largo especial
'    m_Codigo = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 9)
    If Not CBool(MSFlexGrid.TextMatrix(MSFlexGrid.Row, columnaLargoEspecial)) Then Exit Sub
    GoTo Edita
Case 6
    'largo
    'si y no editable
    m_Codigo = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 2)
    If m_Codigo <> "E" Then Exit Sub
    GoTo Edita
'Case 7
'    'medida total
'    'si y no editable
'    m_Codigo = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 1)
'    If Left(m_Codigo, 2) <> "PL" Then Exit Sub
'    GoTo Edita
Case 4, 5, columnaLargoEspecial
    ' no editables
    Exit Sub
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
End Sub
Private Sub Detalle_KeyDown(KeyCode As Integer, Shift As Integer)
If Accion = "Imprimiendo" Then Exit Sub
Select Case KeyCode
Case vbKeyF1
    If Detalle.col = 1 Then CodigoProducto_Buscar
Case vbKeyF2
    MSFlexGridEdit Detalle, MEdit, 32
End Select
End Sub
Private Sub CodigoProducto_Buscar()
Dim m_cod As String
MousePointer = vbHourglass
Product_Search.Condicion = ""
Load Product_Search
MousePointer = vbDefault
Product_Search.Show 1
m_cod = Product_Search.CodigoP

With RsPrd
.Seek "=", m_cod
If Not .NoMatch Then

    Detalle.TextMatrix(Detalle.Row, 1) = m_cod
    Detalle.TextMatrix(Detalle.Row, 2) = ""
    Detalle.TextMatrix(Detalle.Row, 4) = ![unidad de medida]
    Detalle.TextMatrix(Detalle.Row, 5) = !Descripcion
    Detalle.TextMatrix(Detalle.Row, 6) = !largo
    Detalle.TextMatrix(Detalle.Row, columnaLargoEspecial) = ![Largo Especial]
    
    If Detalle.TextMatrix(Detalle.Row, columnaLargoEspecial) Then
        Detalle.col = 2 'foco en largo especial
    Else
        Detalle.col = 3
    End If
End If
End With

End Sub
Private Sub Cursor_Mueve(MSFlexGrid As Control)
'MIA
Dim m_fil As Integer
m_fil = MSFlexGrid.Row
Select Case MSFlexGrid.col
Case 1 'codigo
    'si unidad es metro
    If MSFlexGrid.TextMatrix(m_fil, columnaLargoEspecial) Then
        MSFlexGrid.col = 2  'foco en largo especial
    Else
        MSFlexGrid.col = 3
    End If
Case 2 'largo especial E
    MSFlexGrid.col = 3
Case 3 'Cantidad
    'si largo especial = "E"
    If MSFlexGrid.TextMatrix(m_fil, 2) = "E" Then
        MSFlexGrid.col = 6
    Else
        MSFlexGrid.col = 7
    End If
'Case 6 'Largo en mm
'    MSFlexGrid.col = 8
Case 7 'Precio Unitario
    If MSFlexGrid.Row + 1 < MSFlexGrid.Rows Then
        MSFlexGrid.col = 1
        MSFlexGrid.Row = MSFlexGrid.Row + 1
    End If
Case Else
    MSFlexGrid.col = MSFlexGrid.col + 1
End Select
End Sub
Private Sub Fila_Calcular(Fila As Integer, Actualiza As Boolean)
' actualiza solo linea, y totales generales
Dim co As Integer

co = Detalle.col

n3 = m_CDbl(Detalle.TextMatrix(Fila, 3)) 'can
n6 = m_CDbl(Detalle.TextMatrix(Fila, 6)) 'largo mm
'n7 = n3 * n6 / 1000 ' medida total en metros
n7 = m_CDbl(Detalle.TextMatrix(Fila, 7)) 'pu

' precio total
Detalle.TextMatrix(Fila, 7) = Format(n7, "########.00") 'num_Formato)
'If Detalle.TextMatrix(Fila, 4) = m_AbrMetro Then
'    ' Medida Total * P.U.
'    Detalle.TextMatrix(Fila, 9) = Format(n7 * n8, num_fmtgrl)
'Else
    ' Cantidad * P.U.
    Detalle.TextMatrix(Fila, 8) = Format(n3 * n7, num_fmtgrl)
'End If

If Actualiza Then Detalle_Sumar

End Sub
Private Sub Detalle_Sumar()
Dim SubTotal As Double
SubTotal = 0
For i = 1 To n_filas
    SubTotal = SubTotal + m_CDbl(Detalle.TextMatrix(i, 8))
Next

Totales_Calcular SubTotal

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
'
' FIN RUTINAS PARA FLEXGRID
'
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
Private Sub Detalle_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If Imprimiendo Then Exit Sub
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
Private Sub OC_PrintLegal_EML(Numero_Oc)
cr.WindowTitle = "ORDEN DE COMPRA Nº " & Numero_Oc
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
cr.Formulas(5) = "certificado=""" & IIf(m_Certificado, "NO SE RECEPCIONARA MATERIAL SIN CERTIFICADO DE CALIDAD ADJUNTO", "") & """" ' indica si envia o no certificado de calidad

'MsgBox Certificado.Value

cr.DataFiles(0) = repo_file & ".MDB"

'If Tipo = "E" Then
'    CR.Formulas(5) = "COTIZACION=""" & "GUÍA DESP. Nº:" & """"
'    CR.ReportFileName = Drive_Server & Path_Server & EmpOC.Fantasia & "\Oc_Especial.Rpt"
'Else
'    CR.Formulas(5) = "COTIZACION=""" & "COTIZACIÓN Nº:" & """"
    cr.ReportFileName = Drive_Server & Path_Rpt & "Oc_Legal.Rpt"
'End If

cr.Action = 1

End Sub
Private Sub OC_PrintLegalC2(Numero_Oc)
cr.WindowTitle = "ORDEN DE COMPRA Nº " & Numero_Oc
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
'Cr.Formulas(5) = "certificado=""" & IIf(m_Certificado, "NO SE RECEPCIONARA MATERIAL SIN CERTIFICADO DE CALIDAD ADJUNTO", "") & """" ' indica si envia o no certificado de calidad
cr.Formulas(5) = ""

'MsgBox Certificado.Value

cr.DataFiles(0) = repo_file & ".MDB"

cr.ReportFileName = Drive_Server & Path_Rpt & "ocLegalC2.rpt"

cr.Action = 1

End Sub
Private Sub OC_PrintLegal_AYD(Numero_Oc)
cr.WindowTitle = "ORDEN DE COMPRA Nº " & Numero_Oc
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
cr.Formulas(6) = "certificado=""" & IIf(m_Certificado, "Enviar Certificado de calidad a mail ayanez@aydsa.cl", "") & """" ' indica si envia o no certificado de calidad
cr.Formulas(7) = "pagofactura=""" & "PAGO FACTURAS fono " & EmpOC.Telefono3 & """" ' telefono para pago de facturas

'MsgBox Certificado.Value

cr.DataFiles(0) = repo_file & ".MDB"

'If Tipo = "E" Then
'    CR.Formulas(5) = "COTIZACION=""" & "GUÍA DESP. Nº:" & """"
'    CR.ReportFileName = Drive_Server & Path_Server & EmpOC.Fantasia & "\Oc_Especial.Rpt"
'Else
'    CR.Formulas(5) = "COTIZACION=""" & "COTIZACIÓN Nº:" & """"
    cr.ReportFileName = Drive_Server & Path_Rpt & "Oc_Legal.Rpt"
'End If

cr.Action = 1

End Sub
Private Sub OC_PrintLegal_Delsa(Numero_Oc)
' 09/05/08
cr.WindowTitle = "ORDEN DE COMPRA Nº " & Numero_Oc
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
cr.Formulas(5) = "certificado=""" & IIf(m_Certificado, "Enviar Certificado de calidad adjunto a factura", "") & """" ' indica si envia ono certificado de calidad

'MsgBox Certificado.Value

cr.DataFiles(0) = repo_file & ".MDB"

'If Tipo = "E" Then
'    CR.Formulas(5) = "COTIZACION=""" & "GUÍA DESP. Nº:" & """"
'    CR.ReportFileName = Drive_Server & Path_Server & EmpOC.Fantasia & "\Oc_Especial.Rpt"
'Else
'    CR.Formulas(5) = "COTIZACION=""" & "COTIZACIÓN Nº:" & """"
    cr.ReportFileName = Drive_Server & Path_Rpt & "Oc_Legal.Rpt"
'End If

cr.Action = 1

End Sub
Private Sub OC_PrintLegal_Eiffel(Numero_Oc)
cr.WindowTitle = "ORDEN DE COMPRA Nº " & Numero_Oc
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
cr.Formulas(5) = "certificado=""" & IIf(m_Certificado, "NO SE RECEPCIONARA MATERIAL SIN CERTIFICADO DE CALIDAD ADJUNTO", "") & """" ' indica si envia o no certificado de calidad

cr.DataFiles(0) = repo_file & ".MDB"

cr.ReportFileName = Drive_Server & Path_Rpt & "Oc_Legal.Rpt"

cr.Action = 1

End Sub
Private Sub conIva_Click()
' deja iva en 0% o 19%
' y recalcula

Totales_Calcular

End Sub
Private Sub Privilegios()
If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0, 0, 1
Else
    Botones_Enabled 1, 1, 1, 1, 1, 0, 0, 1
End If
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
