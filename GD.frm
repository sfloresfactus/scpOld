VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form GD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guía de Despacho"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   13500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   9240
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnBarrick 
      Caption         =   "Barrick"
      Height          =   375
      Left            =   7680
      TabIndex        =   53
      Top             =   1920
      Visible         =   0   'False
      Width           =   700
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ingresar Nueva Guía"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar Guía"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar Guía"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Guía"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Autorizar Guia"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "DesHacer"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Grabar Guía"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mantención de Clientes"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame_Densidad 
      Caption         =   "Densidad"
      Height          =   975
      Left            =   7680
      TabIndex        =   43
      Top             =   5400
      Width           =   3375
      Begin VB.Label lbld 
         Caption         =   "7 Super Heavy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   2640
         TabIndex        =   50
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lbld 
         Caption         =   "6 Medium"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1920
         TabIndex        =   49
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lbld 
         Caption         =   "5 Heavy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1920
         TabIndex        =   48
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lbld 
         Caption         =   "4 Light"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1920
         TabIndex        =   47
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lbld 
         Caption         =   "3 Grating ARS 6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lbld 
         Caption         =   "2 Handrails"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lbld 
         Caption         =   "1 Stair Treads ARS 6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox LugarRecepcion 
      Height          =   300
      Left            =   3360
      TabIndex        =   42
      Top             =   1920
      Width           =   4095
   End
   Begin VB.TextBox oc 
      Height          =   300
      Left            =   840
      TabIndex        =   39
      Top             =   1920
      Width           =   735
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   10080
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton btnBultos 
      Caption         =   "Bul"
      Height          =   375
      Left            =   2520
      TabIndex        =   37
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton btnCapturadorLeer 
      Caption         =   "Cap"
      Height          =   375
      Left            =   1920
      TabIndex        =   36
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   2
      Left            =   4680
      MaxLength       =   50
      TabIndex        =   30
      Top             =   5760
      Width           =   3555
   End
   Begin VB.TextBox Numero 
      Height          =   300
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox CbTipo 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Nv 
      Height          =   300
      Left            =   840
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.ComboBox CbPatente 
      Height          =   315
      Left            =   1080
      TabIndex        =   27
      Top             =   5280
      Width           =   1935
   End
   Begin VB.ComboBox CbChofer 
      Height          =   315
      Left            =   1080
      TabIndex        =   25
      Top             =   4920
      Width           =   3615
   End
   Begin VB.ComboBox ComboOT 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox ComboNV 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   1
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   32
      Top             =   6120
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   0
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   29
      Top             =   5760
      Width           =   3555
   End
   Begin VB.ComboBox ComboMarca 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   8040
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox ComboPlano 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame 
      Caption         =   "CLIENTE"
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   3240
      TabIndex        =   9
      Top             =   480
      Width           =   5535
      Begin VB.CommandButton btnSearch 
         Height          =   300
         Left            =   1920
         Picture         =   "GD.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   300
      End
      Begin VB.TextBox Direccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Comuna 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   18
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox Razon 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   14
         Top             =   360
         Width           =   3015
      End
      Begin MSMask.MaskEdBox Rut 
         Height          =   300
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label lbl 
         Caption         =   "DIRECCIÓN"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "COMUNA"
         Height          =   255
         Index           =   6
         Left            =   2880
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "RUT"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbl 
         Caption         =   "SEÑOR(ES)"
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   13
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox txtEditGD 
      Height          =   285
      Left            =   8040
      TabIndex        =   23
      Text            =   "txtEditGD"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   840
      TabIndex        =   5
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   327680
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid Detalle_Normal 
      Height          =   2565
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4524
      _Version        =   393216
      Enabled         =   0   'False
      ScrollBars      =   2
   End
   Begin MSFlexGridLib.MSFlexGrid Detalle_Especial 
      Height          =   2445
      Left            =   120
      TabIndex        =   35
      Top             =   2400
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4313
      _Version        =   393216
      Enabled         =   0   'False
      ScrollBars      =   2
   End
   Begin MSFlexGridLib.MSFlexGrid Detalle_Pernos 
      Height          =   2445
      Left            =   120
      TabIndex        =   38
      Top             =   2400
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4313
      _Version        =   393216
      Enabled         =   0   'False
      ScrollBars      =   2
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
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GD.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GD.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GD.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GD.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GD.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GD.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GD.frx":0836
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GD.frx":0948
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "PRECIO"
      Height          =   255
      Left            =   7200
      TabIndex        =   52
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "PESO"
      Height          =   255
      Left            =   5640
      TabIndex        =   51
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label lblLR 
      Caption         =   "Lugar de Recepción"
      Height          =   255
      Left            =   1800
      TabIndex        =   41
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblOC 
      Caption         =   "OC"
      Height          =   255
      Left            =   360
      TabIndex        =   40
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblDirObra 
      Caption         =   "&Dirección Obra"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblContenido 
      Caption         =   "&Contenido"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label lblPatente 
      Caption         =   "&Patente"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label TotalPrecio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   8160
      TabIndex        =   34
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label TotalPeso 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   5880
      TabIndex        =   33
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label lblChofer 
      Caption         =   "C&hofer"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4920
      Width           =   735
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
      TabIndex        =   4
      Top             =   1080
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
Attribute VB_Name = "GD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnAnular As Button, btnDesHacer As Button, btnGrabar As Button
Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, k As Integer, d As Variant

Private DbD As Database, RsCl As Recordset, RsTb As Recordset, RsPrd As Recordset
Private Dbm As Database, RsNVc As Recordset
Private RsNvPla As Recordset, RsPd As Recordset
Private RsGDc As Recordset, RsGDdN As Recordset, RsGDdE As Recordset
Private RsBulDet As Recordset

Private DbAdq As Database, RsDoc As Recordset

Private m_Tipo As String

Private n_filas As Integer, n_columnas_N As Integer, n_columnas_E As Integer, n_columnas_P As Integer
Private RutClientes(2999) As String, Rev(2999) As String
Private prt As Printer
Private n1 As Double, n4 As Double, n6 As Double, n7 As Double, n8 As Double, n11 As Double
Private a_Nv(2999, 3) As String, m_Nv As Double, m_NvArea As Integer
Private Depende_ITOPyG As Boolean, Advertencia_ITOPyG As Boolean, Advertencia_ITOPyG_msg As String
'private Depende

Private Ret As Integer, Buffer As Variant
Private Capturando As Boolean, Reg As String, Arr_DatoCapturado(99, 1) As String, contador_registros As Integer
Private Arr_Bultos(9) As Double ' arreglo con bultos que van en la guia
Private FormatoGuia As String
Private guiaEspecial As Boolean ' indica si usuario puede hacer guia especial

Private Const NumeroCampos As Integer = 11
Private aCampos(NumeroCampos) As String ' para funcion split
Private Sub btnBarrick_Click()

' lee marcas desde archivo que genera el capturador

End Sub
Private Sub btnBultos_Click()
' busca bultos pendientes de despachar
Dim li As Integer, nBulto As Integer

If Nv.Text = "" Then
    MsgBox "Debe Escoger NV"
    Nv.SetFocus
    Exit Sub
End If

BultosPendientes.NV_Numero = Nv.Text
'BultosPendientes.Proveedor_Razon = Razon.Caption
BultosPendientes.Show 1

nBulto = BultosPendientes.nBulto

If nBulto <> 0 Then
    ' trae bulto pendiente (con todas sus lineas)
'    MsgBox "aqui traer bulto"
    ' busca bulto
    RsBulDet.Seek "=", nBulto, 1
    li = 0
    Do While Not RsBulDet.EOF
    
        If RsBulDet!Numero = nBulto Then
        
            li = li + 1
            Do While Detalle_Normal.TextMatrix(li, 1) <> ""
                li = li + 1
                If li > 20 Then
                    MsgBox "Más de 20 líneas"
                    Exit Sub
                End If
            Loop
            
            ' llena grid
'            Detalle_Normal.TextMatrix(li, 0) = ""
            Detalle_Normal.TextMatrix(li, 1) = RsBulDet!Plano
            Detalle_Normal.TextMatrix(li, 2) = RsBulDet!Rev
            Detalle_Normal.TextMatrix(li, 3) = RsBulDet!Marca
            Detalle_Normal.TextMatrix(li, 7) = RsBulDet!Cantidad
            
            Marca_MiClick RsBulDet!Marca, li, True
            
        Else
        
            Exit Do
            
        End If
        
'        Detalle_Normal.TextMatrix(li, 4) = RsBulDet!Descripcion
'        Detalle_Normal.TextMatrix(li, 5) = "Cant Reci."            '*
'        Detalle_Normal.TextMatrix(li, 6) = "Cant Desp."            '*
'        Detalle_Normal.TextMatrix(li, 7) = "a Desp."
'        Detalle_Normal.TextMatrix(li, 8) = "Peso Unitario"         '*
'        Detalle_Normal.TextMatrix(li, 9) = "Peso TOTAL"            '*
'        Detalle_Normal.TextMatrix(li, 10) = "Precio Unitario"
'        Detalle_Normal.TextMatrix(li, 11) = "Precio TOTAL"         '*
        
        RsBulDet.MoveNext
        
    Loop
    
    ' incorpora bulto a arreglo de bultos
    For i = 0 To 9
        If Arr_Bultos(i) = 0 Then
            Arr_Bultos(i) = nBulto
            Exit For
        End If
    Next
    
End If

End Sub
Private Sub btnCapturadorLeer_Click()
ArchivoAbrir
End Sub
Private Sub CbTipo_Click()

m_Tipo = Left(CbTipo.Text, 1)

ComboPlano.visible = False
ComboMarca.visible = False

guiaEspecial = False

Select Case m_Tipo

Case "N"

    ' NORMAL
    btnCapturadorLeer.Enabled = True
    
    Detalle_Normal.visible = True
    Detalle_Especial.visible = False
    Detalle_Pernos.visible = False
    
'    m_Tipo = "N"

    lblContenido.Caption = "Contenido"
    Obs(2).visible = False
    
Case "E"

'    If Not guiaEspecial Then
'        MsgBox "Usuario No Autorizado para emitir" & vbLf & "Guias de Despacho Especial"
'    Else
    guiaEspecial = True

    ' ESPECIAL
    btnCapturadorLeer.Enabled = False

    Detalle_Normal.visible = False
    Detalle_Especial.visible = True
    Detalle_Pernos.visible = False
    
    m_Tipo = "E"
    
    Detalle_Especial.TextMatrix(0, 4) = "Peso Unitario"
    Detalle_Especial.TextMatrix(0, 5) = "Peso TOTAL"
    
    lblContenido.Caption = "Contenido"
    Obs(2).visible = False
    
'   End If

Case "Pintura"
    
    ' PINTURA
    btnCapturadorLeer.Enabled = False
    
    Detalle_Normal.visible = False
    Detalle_Especial.visible = True
'    m_Tipo = "P"
    
    Detalle_Especial.TextMatrix(0, 4) = "m2 Unitario"
    Detalle_Especial.TextMatrix(0, 5) = "m2 TOTAL"
    
    lblContenido.Caption = "Esquema"
    Obs(2).visible = True
    
Case "G"
    
    ' GALVANIZADO
    btnCapturadorLeer.Enabled = False
    
    Detalle_Normal.visible = True
    Detalle_Especial.visible = False
    Detalle_Pernos.visible = False
    
    m_Tipo = "G"
    
    Detalle_Especial.TextMatrix(0, 4) = "m2 Unitario"
    Detalle_Especial.TextMatrix(0, 5) = "m2 TOTAL"
    
    Obs(2).visible = False

Case "P"
    
    ' PERNOS
    btnCapturadorLeer.Enabled = False
    
    btnBultos.Enabled = False
    btnBarrick.Enabled = False
    
    Detalle_Normal.visible = False
    Detalle_Especial.visible = False
    Detalle_Pernos.visible = True
    Detalle_Pernos.Enabled = True
    
'    m_Tipo = "P"
    
'    Detalle_Especial.TextMatrix(0, 4) = "m2 Unitario"
'    Detalle_Especial.TextMatrix(0, 5) = "m2 TOTAL"
'    lblContenido.Caption = "Esquema"
    Obs(2).visible = True

End Select

End Sub
Private Sub Detalle_Pernos_DblClick()
If Accion = "Imprimiendo" Then Exit Sub
MSFlexGridEdit_P Detalle_Pernos, txtEditGD, 32
End Sub
Private Sub Detalle_Pernos_GotFocus()
If txtEditGD.visible Then
    Detalle_Pernos = txtEditGD.Text
    txtEditGD.visible = False
End If
End Sub
Private Sub Detalle_Pernos_KeyDown(KeyCode As Integer, Shift As Integer)
If Accion = "Imprimiendo" Then Exit Sub
Select Case KeyCode
Case vbKeyF1
    If Detalle_Pernos.col = 1 Then CodigoProducto_Buscar
Case vbKeyF2
    MSFlexGridEdit_P Detalle_Pernos, txtEditGD, 32
End Select
End Sub
Private Sub CodigoProducto_Buscar()
Dim m_cod As String
MousePointer = vbHourglass
'Product_Search.Condicion = " AND [tipo producto]='PRN'"
Product_Search.Condicion = ""
Load Product_Search
MousePointer = vbDefault
Product_Search.Show 1
m_cod = Product_Search.CodigoP

With RsPrd
.Seek "=", m_cod
If Not .NoMatch Then
    Detalle_Pernos.TextMatrix(Detalle_Pernos.Row, 1) = m_cod
    Detalle_Pernos.TextMatrix(Detalle_Pernos.Row, 2) = !Descripcion
End If
End With

End Sub

Private Sub Detalle_Pernos_KeyPress(KeyAscii As Integer)
If Accion = "Imprimiendo" Then Exit Sub
MSFlexGridEdit_P Detalle_Pernos, txtEditGD, KeyAscii
End Sub
Private Sub Detalle_Pernos_LeaveCell()
If txtEditGD.visible Then
    Detalle_Pernos = txtEditGD.Text
    txtEditGD.visible = False
End If
End Sub
' procedimientos nuevos
' 25/07/06
' una guia se puede modificar y eliminar n veces antes de imprimir
' una vez impresa no se puede modificar, para hacerlo debe autorizar gerente (erwin)
Private Sub Form_Load()

guiaEspecial = False ' ojo

lblOC.visible = False
Oc.visible = False
lblLR.visible = False
LugarRecepcion.visible = False

Capturando = False  ' indica acaba de ser presionado boton capturar
btnCapturadorLeer.visible = True ' Capturador

'Depende_ITOPyG = False
Depende_ITOPyG = True ' siempre a partir de 29/11/05

Advertencia_ITOPyG = False
Advertencia_ITOPyG_msg = ""

' solo cliente sedgman
Frame_Densidad.visible = False ' a partir de 18/02/10

' abre archivos
Set DbD = OpenDatabase(data_file)
Set RsCl = DbD.OpenRecordset("Clientes")
RsCl.Index = "RUT"

Set RsPrd = DbD.OpenRecordset("Productos")
RsPrd.Index = "codigo"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

' para calculo de cantidad de piezas fabricadas x mes
'Dim qry As String
'qry = "SELECT "
'qry = qry & "SUM([Cantidad Total]) AS can"
'qry = qry & " FROM [Planos Detalle]"
'Set RsGDc = Dbm.OpenRecordset(qry)
'Debug.Print RsGDc!can

Set RsGDc = Dbm.OpenRecordset("GD Cabecera")
RsGDc.Index = "Numero"

Set RsGDdN = Dbm.OpenRecordset("GD Detalle")
RsGDdN.Index = "Numero-Linea"

Set RsGDdE = Dbm.OpenRecordset("GD Especial Detalle")
RsGDdE.Index = "Numero-Linea"

Set RsBulDet = Dbm.OpenRecordset("Bultos")
RsBulDet.Index = "Numero-Linea"

CbTipo.AddItem "Normal"
CbTipo.AddItem "Especial"
'CbTipo.AddItem "Pintura"
CbTipo.AddItem "Galvanizado"
CbTipo.AddItem "Pernos"

' Combo obra
ComboNV.AddItem " "
i = 0
RutClientes(i) = " "

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
        i = i + 1
        a_Nv(i, 0) = RsNVc!Numero
        a_Nv(i, 1) = RsNVc!obra
        a_Nv(i, 2) = RsNVc!galvanizado Or RsNVc!pintura
        a_Nv(i, 3) = RsNVc!galvanizado
        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
        RutClientes(i) = RsNVc![RUT CLiente]
    End If
    
    RsNVc.MoveNext
    
Loop

Set RsNvPla = Dbm.OpenRecordset("Planos Cabecera")
RsNvPla.Index = "NV-Plano"

Set RsPd = Dbm.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Item"

Set DbAdq = OpenDatabase(Madq_file)
Set RsDoc = DbAdq.OpenRecordset("Documentos")
RsDoc.Index = "Tipo-Numero-Linea"

Inicializa

Detalle_Config_Normal
Detalle_Config_Especial
Detalle_Config_Pernos
'Detalle_Config_Pintura

m_Tipo = "N"
Obs(2).visible = False

FormatoGuia = ReadIniValue(Path_Local & "scp.ini", "GD", "Formato")

m_NvArea = 0

Privilegios

'Gd_Anular 88501, 88550, #11/30/2010#

End Sub

Private Sub Nv_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Nv_LostFocus()
'Dim m_Nv As Integer
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

MousePointer = vbHourglass

Select Case m_Tipo

Case "G" ' guia galvanizado
'If m_Tipo = "G" Then ' guia galvanizado
'////////////////////////////////////////////////
    m_Nv = Val(Left(ComboNV.Text, 6))
    Nv.Text = m_Nv

    If CBool(a_Nv(ComboNV.ListIndex, 3)) Then ' si nv va galvanizada
    
        ComboPlano.Clear
        
        ComboPlano.AddItem " "
        i = 0
        Rev(i) = " "
        
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
        
''''''        Detalle_Limpiar Detalle_Normal, n_columnas_N

        Detalle_Normal.Enabled = True

    Else
        MsgBox "NV " & Nv.Text & " NO va Galvanizada"
        Nv.SetFocus
    End If

'////////////////////////////////////////////////
Case "N", "E", "P"
'Else ' normal y especial

    Depende_ITOPyG = True
    
    i = 0
    m_Nv = Val(Left(ComboNV.Text, 6))
    Nv.Text = m_Nv
    
    'Debug.Print ComboNV.ListIndex, ComboNV.Text, a_Nv(ComboNV.ListIndex, 2)
    '''Depende_ITOPyG = CBool(a_Nv(ComboNV.ListIndex, 2)) ' lo comente el 17/04/2014
    
    ComboPlano.Clear
    
    ComboPlano.AddItem " "
    Rev(i) = " "
    
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
    
    Detalle_Limpiar Detalle_Normal, n_columnas_N
    'no debe limpiar detalle de guia especial
    'Detalle_Limpiar Detalle_Especial, n_columnas_E
    
    ComboMarca.Clear
    
    Select Case m_Tipo
    Case "G"
    
        ' galvanizado
        Depende_ITOPyG = False
        
'    Case "P" ' pernos
'        ' no asocia nv con cliente
    Case Else
        
        ' datos del cliente
        rut.Text = RutClientes(ComboNV.ListIndex)
        RsCl.Seek "=", rut.Text
        If Not RsCl.NoMatch Then
            Razon.Text = RsCl![Razon Social]
            Direccion.Text = RsCl!Direccion
            Comuna.Text = RsCl!Comuna

            If rut.Text = "76300170-9" Then ' solo cliente SEDGMAN
                Frame_Densidad.visible = True
            Else
                Frame_Densidad.visible = False
            End If

        End If
        
    End Select
    
    'If Rut.Text <> "" Then Detalle.Enabled = True
    Detalle_Normal.Enabled = True

End Select
'End If ' guia normal

MousePointer = vbDefault

End Sub
Private Sub ComboPlano_Click()
' supuesto: el numero del plano es único para toda nv
Dim old_plano As String, filaFlex As Integer

old_plano = Detalle_Normal

filaFlex = Detalle_Normal.Row

If ComboPlano.ListIndex > 0 Then Detalle_Normal.TextMatrix(filaFlex, 2) = Rev(ComboPlano.ListIndex)

'ComboMarca_Poblar np

ComboPlano.visible = False
Detalle_Normal = ComboPlano.Text

If Detalle_Normal <> old_plano Then
    For i = 2 To n_columnas_N
        Detalle_Normal.TextMatrix(filaFlex, i) = ""
    Next
End If

' revision
If ComboPlano.ListIndex > 0 Then Detalle_Normal.TextMatrix(filaFlex, 2) = Rev(ComboPlano.ListIndex)

End Sub
Private Sub ComboMarca_Poblar(Plano As String)
' llena combo marcas
ComboMarca.Clear

RsPd.Seek ">=", m_Nv, m_NvArea, Plano, 0
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

Marca_MiClick "", 0, False

End Sub
Private Sub Marca_MiClick(Marca As String, Fila As Integer, bulto As Boolean)
Dim m_Plano As String, m_Marca As String, fil As Integer
Dim c_otf As Integer, c_itof As Integer, c_desp As Integer
'Dim c_itopg As Integer, c_gdgal As Integer, c_total As Integer
Dim c_itopp As Integer, c_total As Integer

If Fila <> 0 Then Detalle_Normal.Row = Fila

fil = Detalle_Normal.Row

ComboMarca.visible = False
m_Plano = Detalle_Normal.TextMatrix(fil, 1)

If Marca = "" Then
    m_Marca = ComboMarca.Text
Else
    m_Marca = Marca
End If

'If Not Capturador Then
    '///
If Not bulto Then
' verifica si Plano-Marca ya están en esta GD
If Not Capturando Then
For i = 1 To n_filas
    If m_Plano = Detalle_Normal.TextMatrix(i, 1) And m_Marca = Detalle_Normal.TextMatrix(i, 3) Then
        Beep
        MsgBox "MARCA YA EXISTE EN GD"
        Detalle_Normal.Row = i
        Detalle_Normal.col = 3
        Detalle_Normal.SetFocus
        Exit Sub
    End If
Next
End If
'///
End If

'End If

If Not bulto Then
Detalle_Normal = m_Marca
End If

c_itof = c_desp = 0
' busca marca en plano
RsPd.Seek ">=", m_Nv, m_NvArea, m_Plano, 0
If Not RsPd.NoMatch Then
    Do While Not RsPd.EOF
    
        If RsPd!Marca = m_Marca Then
        
            c_total = RsPd![Cantidad Total]
            c_otf = RsPd![OT fab]
            c_itof = RsPd![ITO fab]
            'c_gdgal = RsPd![GD gal]
            'c_itopg = RsPd![ITO pyg]
            c_itopp = RsPd![ITO pp]
            c_desp = RsPd![GD]
            
            ' verifica si está asignada
            If c_otf = 0 Then
                Beep
                MsgBox "La marca """ & m_Marca & """" & vbCr _
                    & "No está Asignada"
                Detalle_Normal.TextMatrix(fil, 3) = ""
                Detalle_Normal.SetFocus
                Exit Sub
            End If
            
            If c_itof = 0 Then
                Beep
                MsgBox "La marca """ & m_Marca & """" & vbCr _
                    & "No está Recibida de Fabricación"
                Detalle_Normal.TextMatrix(fil, 3) = ""
                Detalle_Normal.SetFocus
                Exit Sub
            End If
            
            If m_Tipo <> "G" Then
                If Depende_ITOPyG Then
                    If c_itopp = 0 Then
                        Beep
                        MsgBox "La marca """ & m_Marca & """" & vbCr _
                            & "No está Recibida de Produccion Pintura"
                            
                        If Advertencia_ITOPyG Then
                        Else
                            Detalle_Normal.TextMatrix(fil, 3) = ""
                            Exit Sub
                        End If
                        Detalle_Normal.SetFocus
                    End If
                End If
            End If
            
            ' si es "gd gal" verifica que quede algo por enviar al galvanizado
            'If m_Tipo = "G" Then
            '    If c_total > c_gdgal Then
            '        ' ok
            '    Else
            '        Beep
            '        MsgBox "La marca """ & m_Marca & """" & vbCr & "Ya se galvanizó"
            '        Detalle_Normal.TextMatrix(fil, 3) = ""
            '        Detalle_Normal.SetFocus
            '        Exit Sub
            '    End If
            'End If
        
            ' verifica que quede algo por despachar
            If Depende_ITOPyG Then
                'If c_itopg - c_desp <= 0 Then
                '    Beep
                '    MsgBox "La marca """ & m_Marca & """" & vbCr _
                '        & "Ya se despachó"
                '    Detalle_Normal.TextMatrix(fil, 3) = ""
                '    Detalle_Normal.SetFocus
                '    If Not Advertencia_ITOPyG Then
                '        Exit Sub
                '    End If
                'End If
            Else
                If c_itof - c_desp <= 0 Then
                    Beep
                    MsgBox "La marca """ & m_Marca & """" & vbCr & "Ya se despachó"
                    Detalle_Normal.TextMatrix(fil, 3) = ""
                    Detalle_Normal.SetFocus
                    Exit Sub
                End If
            End If
            
            
            Detalle_Normal.TextMatrix(fil, 4) = RsPd!Descripcion
            If m_Tipo = "G" Then
                'Detalle_Normal.TextMatrix(fil, 5) = c_itof
                'Detalle_Normal.TextMatrix(fil, 6) = c_gdgal '- m_cantGD ?
            Else
                If Depende_ITOPyG Then
                    Detalle_Normal.TextMatrix(fil, 5) = c_itopp
                Else
                    Detalle_Normal.TextMatrix(fil, 5) = c_itof
                End If
                
                Detalle_Normal.TextMatrix(fil, 6) = c_desp '- m_cantGD ?
            End If
            
            Detalle_Normal.TextMatrix(fil, 8) = Replace(RsPd![Peso], ",", ".")
            
            Detalle_Normal.TextMatrix(fil, 9) = Val(Detalle_Normal.TextMatrix(fil, 7)) * Detalle_Normal.TextMatrix(fil, 8)
            
            ' densidad
            Detalle_Normal.TextMatrix(fil, 10) = RsPd!densidad
            
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
Private Sub Inicializa()

Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnAnular = Toolbar.Buttons(5)
Set btnDesHacer = Toolbar.Buttons(7)
Set btnGrabar = Toolbar.Buttons(8)

Obj = "GUÍA DE DESPACHO"
Objs = "GUÍAS DE DESPACHO"

Accion = ""

btnSearch.visible = False
btnSearch.ToolTipText = "Busca Clientes"
Campos_Enabled False

n_filas = 20 ' 18
Select Case FormatoGuia
Case "EML"
    n_filas = 20 ' 18
Case "AYD"
    n_filas = 20
Case "DELSA"
    n_filas = 14
Case "EIFFEL"
    n_filas = 20
End Select

n_columnas_N = 13 ' 12 '11
n_columnas_E = 7
n_columnas_P = 5

Set DbD = OpenDatabase(data_file)
Set RsTb = DbD.OpenRecordset("Tablas")
RsTb.Index = "Tipo-Descripcion"

Do While Not RsTb.EOF
    If RsTb!Tipo = "CHOFER" Then
        CbChofer.AddItem RsTb!Descripcion
    End If
    If RsTb!Tipo = "PATENTE" Then
        CbPatente.AddItem RsTb!Descripcion
    End If
    RsTb.MoveNext
Loop

RsTb.Close

Oc.MaxLength = 10
LugarRecepcion.MaxLength = 30

Obs(0).MaxLength = 50
Obs(1).MaxLength = 50
Obs(2).MaxLength = 50
'Obs(3).MaxLength = 50
'Obs(4).MaxLength = 50


End Sub
Private Sub Numero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    After_Enter
End If
End Sub
Private Sub After_Enter()
Dim n_Copias As Integer

'If EmpOC.Rut <> Rut_Eml Then
'    'si empresa es PyP => guia despacho debe ser especial 23/08/1999
'    GDespecial.Value = 1
'End If

Select Case Accion
Case "Agregando"
    RsGDc.Seek "=", Numero.Text
    If RsGDc.NoMatch Then
    
        Campos_Enabled True
        Numero.Enabled = False
        
        Detalle_Normal.Enabled = False
        
        Fecha.Text = Format(Now, Fecha_Format)
        
'        If Usuario.AccesoTotal Then
'            Fecha.SetFocus
'        Else
'            ComboNV.SetFocus
'        End If
        CbTipo.Text = "Normal"
        CbTipo.SetFocus
'        GDespecial.SetFocus
        
        btnGrabar.Enabled = True
        btnSearch.visible = True
        
    Else
    
        Doc_Leer
        MsgBox Obj & " YA EXISTE"
        Detalle_Normal.Enabled = False
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
        
    End If
    
Case "Modificando"

    RsGDc.Seek "=", Numero.Text
    If RsGDc.NoMatch Then
    
        MsgBox Obj & " NO EXISTE"
        
    Else
    
        Doc_Leer
        
        If RsGDc!impresa Then
        
            MsgBox Obj & " ya esta impresa," & vbLf & "NO se puede Modificar"
            Detalle_Normal.Enabled = False
            Campos_Limpiar
            Numero.Enabled = True
            Numero.SetFocus
        
        Else
        
'        If RsGDc!Tipo = "N" Then
            Campos_Enabled True
            Numero.Enabled = False
            
'            If Usuario.AccesoTotal Then
'                Fecha.SetFocus
'            Else
'                ComboNV.SetFocus
'            End If
            CbTipo.SetFocus
'            GDespecial.SetFocus
            
            btnGrabar.Enabled = True
            btnSearch.visible = True
            
'        Else
'            MsgBox "DEBE MODIFICAR COMO GUIA ESPECIAL"
'            Campos_Limpiar
'            Numero.Enabled = True
'            Numero.SetFocus
'        End If

        End If
    End If
    
Case "Eliminando"

    RsGDc.Seek "=", Numero.Text
    
    If RsGDc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
    
        Doc_Leer
        
        If RsGDc!impresa Then
        
            MsgBox Obj & " ya esta impresa," & vbLf & "NO se puede Eliminar"
            Detalle_Normal.Enabled = False
            Campos_Limpiar
            Numero.Enabled = True
            Numero.SetFocus
        
        Else
        
'        If RsGDc!Tipo = "N" Then
            Numero.Enabled = False
            If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo) = vbYes Then
                Doc_Eliminar
            End If
'        Else
'            MsgBox "DEBE ELIMINAR COMO GUIA ESPECIAL"
'        End If
            Campos_Limpiar
            Numero.Enabled = True
            Numero.SetFocus
        
        End If
        
    End If
Case "Imprimiendo"

    RsGDc.Seek "=", Numero
    If RsGDc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
    
        Doc_Leer
        
'        Select Case rsdc!Tipo
'        Case "N", "G"
'
'            Numero.Enabled = False
'
'            Detalle_Normal.Visible = True
'            Detalle_Normal.Enabled = True
'            Detalle_Especial.Visible = False
'
'        Case "E"
'
'            Numero.Enabled = False
'
'           Detalle_Especial.Visible = True
'           Detalle_Especial.Enabled = True
'           Detalle_Normal.Visible = False
'
'        End Select
        
    End If
    
Case "Autorizando"

    RsGDc.Seek "=", Numero.Text
    
    If RsGDc.NoMatch Then
    
        MsgBox Obj & " NO EXISTE"
        
    Else
    
        Doc_Leer
        
        Select Case RsGDc!Tipo
        Case "N", "G"
        'If RsGDc!Tipo = "N" Then
            Numero.Enabled = False
            
            Detalle_Normal.visible = True
            Detalle_Normal.Enabled = True
            Detalle_Especial.visible = False
            Detalle_Pernos.visible = False
            
        Case "E"
        'Else
        
            Numero.Enabled = False
            
            Detalle_Normal.visible = False
            Detalle_Especial.visible = True
            Detalle_Especial.Enabled = True
            Detalle_Pernos.visible = False
            
        Case "P" ' pernos
        
            Numero.Enabled = False
            
            Detalle_Normal.visible = False
            Detalle_Especial.visible = False
            Detalle_Pernos.visible = True
            Detalle_Pernos.Enabled = True
        
        End Select
        'End If
        
        If RsGDc!impresa Then
            ' ok se puede autorizar
        Else
            MsgBox Obj & " NO está impresa"
        End If
        
    End If

End Select

End Sub
Private Sub Doc_Leer()
Dim m_resta As Integer, Cuenta_Lineas As Integer
' CABECERA
Fecha.Text = Format(RsGDc!Fecha, Fecha_Format)
m_Nv = RsGDc!Nv
rut.Text = RsGDc![RUT CLiente]

RsNVc.Seek "=", m_Nv, m_NvArea
If Not RsNVc.NoMatch Then
    On Error Resume Next
    ComboNV.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
    On Error GoTo 0
End If

rut.Text = RsGDc![RUT CLiente]

'Oc.Text = RsGDc!Oc
'LugarRecepcion.Text = NoNulo(RsGDc![Lugar Recepcion])

CbChofer.Text = NoNulo(RsGDc![Observacion 1])
CbPatente.Text = NoNulo(RsGDc![Observacion 2])
Obs(0).Text = NoNulo(RsGDc![Observacion 3])
Obs(1).Text = NoNulo(RsGDc![Observacion 4])
Obs(2).Text = NoNulo(RsGDc![Observacion 5])

'DETALLE
RsPd.Index = "NV-Plano-Marca"

m_Tipo = RsGDc!Tipo

Cuenta_Lineas = 0

Select Case m_Tipo

Case "N", "G"

    ' GUIA NORMAL
'    CbTipo.Text = "Normal"
'    GDespecial.Value = 0
'    m_Tipo = "N"
    RsGDdN.Seek "=", Numero.Text, 1
    If Not RsGDdN.NoMatch Then
        Do While Not RsGDdN.EOF
            If RsGDdN!Numero = Numero.Text Then
            
                i = RsGDdN!linea
                
                Cuenta_Lineas = Cuenta_Lineas + 1
                
                If Cuenta_Lineas > n_filas Then
                    GoTo Sigue
                End If
                
                Detalle_Normal.TextMatrix(i, 1) = RsGDdN!Plano
                Detalle_Normal.TextMatrix(i, 2) = RsGDdN!Rev
                Detalle_Normal.TextMatrix(i, 3) = RsGDdN!Marca
                
                RsPd.Seek "=", m_Nv, m_NvArea, RsGDdN!Plano, RsGDdN!Marca
                If Not RsPd.NoMatch Then
                    Detalle_Normal.TextMatrix(i, 4) = RsPd!Descripcion
'                    Detalle_Normal.TextMatrix(i, 5) = RsPd![ito fab]
                    Detalle_Normal.TextMatrix(i, 5) = RsPd![ITO pp]
                    m_resta = IIf(Accion = "Modificando", RsGDdN!Cantidad, 0)
                    Detalle_Normal.TextMatrix(i, 6) = RsPd![GD] - m_resta
                    'Detalle_Normal.TextMatrix(i, 10) = RsPd![densidad]
                End If
                
                Detalle_Normal.TextMatrix(i, 7) = RsGDdN!Cantidad
                Detalle_Normal.TextMatrix(i, 8) = RsGDdN![Peso Unitario]
                Detalle_Normal.TextMatrix(i, 11) = RsGDdN![Precio Unitario]
                
                Fila_Calcular_Normal i, False
                
            Else
                Exit Do
            End If
            RsGDdN.MoveNext
        Loop
    End If

Case "E" '"P"

    ' GUIA ESPECIAL
    CbTipo.Text = "Especial"
'    GDespecial.Value = 1
    m_Tipo = "E"
    
    RsGDdE.Seek "=", Numero.Text, 1
    If Not RsGDdE.NoMatch Then
        Do While Not RsGDdE.EOF
            If RsGDdE!Numero = Numero.Text Then
            
                i = RsGDdE!linea
                
                Cuenta_Lineas = Cuenta_Lineas + 1
                
                Detalle_Especial.TextMatrix(i, 1) = RsGDdE!Cantidad
                Detalle_Especial.TextMatrix(i, 2) = RsGDdE!unidad
                Detalle_Especial.TextMatrix(i, 3) = RsGDdE!Detalle
                Detalle_Especial.TextMatrix(i, 4) = RsGDdE![Peso Unitario]
                Detalle_Especial.TextMatrix(i, 6) = RsGDdE![Precio Unitario]
                
                Fila_Calcular_Especial i, False
                            
            Else
                Exit Do
            End If
            RsGDdE.MoveNext
        Loop
    End If
    
Case "P" ' pernos

    ' GUIA pernos
    CbTipo.Text = "Pernos"
'    GDespecial.Value = 1
    m_Tipo = "P"
    
    With RsDoc
    .Seek "=", "GP", Numero.Text, 1
    If Not .NoMatch Then
        Do While Not .EOF
        
            If !Tipo = "GP" And !Numero = Numero.Text Then
            
                i = !linea
                
                Cuenta_Lineas = Cuenta_Lineas + 1
                
                If i > 17 Then
                    ' 20/01/15
                    MsgBox "GD tiene mas de 17 lineas"
                    Exit Do
                End If
                
                Detalle_Pernos.TextMatrix(i, 1) = ![codigo producto]
                RsPrd.Seek "=", ![codigo producto]
                If Not RsPrd.NoMatch Then
                    Detalle_Pernos.TextMatrix(i, 2) = RsPrd![Descripcion]
                End If
                Detalle_Pernos.TextMatrix(i, 3) = !Cant_Sale
                Detalle_Pernos.TextMatrix(i, 4) = ![Precio Unitario]
                
                Fila_Calcular_Pernos i, False
                            
            Else
                Exit Do
            End If
            
            .MoveNext
            
        Loop
    End If
    End With

End Select

Sigue:

RsPd.Index = "NV-Plano-Item"

Cliente_Lee rut.Text

Select Case m_Tipo
Case "N", "G"

    If m_Tipo = "N" Then
        CbTipo.Text = "Normal"
    Else
        CbTipo.Text = "Galvanizado"
    End If
    
    Detalle_Normal.Row = 1 ' para q' actualice la primera fila del detalle
    Detalle_Sumar_Normal
    
Case "E"

    CbTipo.Text = "Especial"

    Detalle_Especial.Row = 1 ' para q' actualice la primera fila del detalle
    Detalle_Sumar_Especial

Case "P" ' pernos

    CbTipo.Text = "Pernos"
        
'    Detalle_Especial.TextMatrix(0, 4) = "m2 Unitario"
'    Detalle_Especial.TextMatrix(0, 5) = "m2 TOTAL"
        
    Detalle_Pernos.Row = 1 ' para q' actualice la primera fila del detalle
    Detalle_Sumar_Pernos

End Select

If Cuenta_Lineas > n_filas Then
    MsgBox "Numero de Lineas Mayor a " & n_filas, vbCritical, "ATENCION"
End If

End Sub
Private Sub Cliente_Lee(rut)
RsCl.Seek "=", rut
If Not RsCl.NoMatch Then
    Razon.Text = RsCl![Razon Social]
    Direccion.Text = RsCl!Direccion
    Comuna.Text = NoNulo(RsCl!Comuna)
End If
End Sub
Private Function Doc_Validar() As Boolean
Dim porDespachar As Integer

Doc_Validar = False

If rut.Text = "" Then
    MsgBox "DEBE ELEGIR CLIENTE"
'    Rut.SetFocus
    btnSearch.SetFocus
    Exit Function
End If

If m_Nv = 0 Then
    MsgBox "DEBE ELEGIR NV=0"
    ComboNV.SetFocus
    Exit Function
End If
If Nv.Text = "" Then
    MsgBox "DEBE ELEGIR NV"
    ComboNV.SetFocus
    Exit Function
End If

Select Case m_Tipo
Case "N", "G"

For i = 1 To n_filas

    ' plano
    If Trim(Detalle_Normal.TextMatrix(i, 1)) <> "" Then
    
        ' marca                3
        If Not CampoReq_Valida(Detalle_Normal.TextMatrix(i, 3), i, 3) Then Exit Function
        
        ' descripcion          4
        
        ' tot cant recibida    5
        ' tot cant despach     6
        
        ' cantidad a despachar 7
        If Not Numero_Valida(Detalle_Normal.TextMatrix(i, 7), i, 7) Then Exit Function
        
        ' [can asignada]-[can recibida]>=[can a recibir]
        porDespachar = Detalle_Normal.TextMatrix(i, 5) - Val(Detalle_Normal.TextMatrix(i, 6))
        If porDespachar < Detalle_Normal.TextMatrix(i, 7) Then
            MsgBox "Sólo quedan " & porDespachar & " por Despachar", , "ATENCIÓN"
            Detalle_Normal.Row = i
            Detalle_Normal.col = 7
            Detalle_Normal.SetFocus
            Exit Function
        End If
        
        ' peso unitario    8
        ' peso total       9
        ' precio unitario 10
        ' precio total    11
        
    End If
    
Next

End Select

Doc_Validar = True

End Function
Private Function CampoReq_Valida(txt As String, fil As Integer, col As Integer) As Boolean
' valida si campo requerido
If Len(Trim(txt)) = 0 Then
    CampoReq_Valida = False
    Beep
    MsgBox "CAMPO OBLIGATORIO"
    Detalle_Normal.Row = fil
    Detalle_Normal.col = col
    Detalle_Normal.SetFocus
Else
    CampoReq_Valida = True
End If
End Function
Private Function LargoString_Valida(txt As String, max As Integer, fil As Integer, col As Integer) As Boolean
If Len(Trim(txt)) > max Then
    LargoString_Valida = False
    Beep
    MsgBox "Largo Máximo es " & max & " caracteres"
    Detalle_Normal.Row = fil
    Detalle_Normal.col = col
    Detalle_Normal.SetFocus
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
        Detalle_Normal.Row = fil
        Detalle_Normal.col = col
        Detalle_Normal.SetFocus
        Exit Function
'    End If
End If
Numero_Valida = True
End Function
Private Sub Doc_Grabar(Nueva As Boolean)

Dim m_Plano As String, m_Marca As String, m_cantidad As Long
Dim qry As String

MousePointer = vbHourglass

save:

' CABECERA DE GD
With RsGDc
If Nueva Then
    .AddNew
    !Numero = Numero.Text
    !Tipo = m_Tipo '"N"
Else

    Doc_Detalle_Eliminar
    
    .Edit
    
End If

!Fecha = Fecha.Text
!Nv = Val(m_Nv)
![RUT CLiente] = rut.Text
'![Peso Total] = Val(TotalPeso.Caption)

'!Oc = Val(Oc.Text)
'![Lugar Recepcion] = LugarRecepcion.Text

![Peso Total] = m_CDbl(TotalPeso.Caption) ' modif 01/01/05
![Precio Total] = Val(TotalPrecio.Caption)
![Observacion 1] = Left(CbChofer.Text, 50)
![Observacion 2] = Left(CbPatente.Text, 50)
![Observacion 3] = Left(Obs(0).Text, 50)
![Observacion 4] = Left(Obs(1).Text, 50)
![Observacion 5] = Left(Obs(2).Text, 50)

.Update

End With

' DETALLE DE GD
Select Case m_Tipo
Case "N", "G"

    ' NORMAL
    With RsGDdN
    j = 0
    RsPd.Index = "NV-Plano-Marca"
    For i = 1 To n_filas
        m_Plano = Trim(Detalle_Normal.TextMatrix(i, 1))
        If m_Plano <> "" Then
            m_Marca = Detalle_Normal.TextMatrix(i, 3)
            m_cantidad = Val(Detalle_Normal.TextMatrix(i, 7))
            
            .AddNew
            !Numero = Numero.Text
            j = j + 1
            !linea = j
            
            !Nv = m_Nv
            !Fecha = Fecha.Text
            ![RUT CLiente] = rut.Text
            
            !Plano = m_Plano
            !Rev = Detalle_Normal.TextMatrix(i, 2)
            !Marca = m_Marca
            !Cantidad = m_cantidad
    '        RsGDd("Fecha Despacho") = Fecha.Text
            ![Peso Unitario] = m_CDbl(Detalle_Normal.TextMatrix(i, 8))
            ![Precio Unitario] = m_CDbl(Detalle_Normal.TextMatrix(i, 11)) '?
    
            .Update
            
            RsPd.Seek "=", m_Nv, m_NvArea, m_Plano, m_Marca
            If RsPd.NoMatch Then
                ' no existe marca en el plano
            Else
                ' actualiza archivo detalle planos
                
                RsPd.Edit
                If m_Tipo = "G" Then
                    RsPd![GD gal] = RsPd![GD gal] + m_cantidad
                Else
                    RsPd![GD] = RsPd![GD] + m_cantidad
                End If
                RsPd.Update
                
            End If
            
        End If
    Next
    RsPd.Index = "NV-Plano-Item"
    End With
    
Case "E"

    ' ESPECIAL
    With RsGDdE
    j = 0
    For i = 1 To n_filas
    
        m_cantidad = Val(Detalle_Especial.TextMatrix(i, 1))
        
        If m_cantidad <> 0 Then
            
            .AddNew
            !Numero = Numero.Text
            j = j + 1
            !linea = j
            
            !Nv = Val(m_Nv)
            !Fecha = Fecha.Text
            ![RUT CLiente] = rut.Text
            
            !Cantidad = m_cantidad
            !unidad = Detalle_Especial.TextMatrix(i, 2)
            !Detalle = Detalle_Especial.TextMatrix(i, 3)
            ![Peso Unitario] = m_CDbl(Detalle_Especial.TextMatrix(i, 4))
            ![Precio Unitario] = m_CDbl(Detalle_Especial.TextMatrix(i, 6))
    
            .Update
            
        End If
    Next
    End With
    
Case "P" ' pernos

    With RsDoc
    j = 0
    For i = 1 To n_filas
        m_cantidad = Val(Detalle_Pernos.TextMatrix(i, 3))
        If m_cantidad <> 0 Then
            .AddNew
            
            !Tipo = "GP" ' guia de despacho pernos
            !Numero = Numero.Text
            j = j + 1
            !linea = j
            !Fecha = Fecha.Text
            !Nv = Val(m_Nv)
            ![rut] = rut.Text
            
            ![codigo producto] = Detalle_Pernos.TextMatrix(i, 1)
            ![Precio Unitario] = m_CDbl(Detalle_Pernos.TextMatrix(i, 4))
            !Cant_Sale = m_cantidad
            
            .Update
            
        End If
        
    Next
    End With

End Select

' graba bultos incluidos en esta guia como NO pendientes
For i = 0 To 9
    If Arr_Bultos(i) > 0 Then
        Dbm.Execute "UPDATE bultos SET despachado = true WHERE numero=" & Arr_Bultos(i)
    End If
Next

Select Case Accion
Case "Agregando"
    Track_Registrar "GD" & m_Tipo, Numero.Text, "AGR"
Case "Modificando"
    Track_Registrar "GD" & m_Tipo, Numero.Text, "MOD"
End Select

MousePointer = vbDefault

End Sub
Private Sub Doc_Eliminar()

' elimina cabecera
RsGDc.Seek "=", Numero.Text
If Not RsGDc.NoMatch Then

    RsGDc.Delete
   
End If

' elimina detalle
Doc_Detalle_Eliminar

Track_Registrar "GD" & m_Tipo, Numero.Text, "ELI"

End Sub
Private Sub Doc_Detalle_Eliminar()
' elimina detalle GD
' al anular detalle GD debe actualizar detalle plano

Select Case m_Tipo
Case "N"

    RsPd.Index = "NV-Plano-Marca"
    RsGDdN.Seek "=", Numero.Text, 1
    If Not RsGDdN.NoMatch Then
        Do While Not RsGDdN.EOF
            If RsGDdN!Numero <> Numero.Text Then Exit Do
            RsPd.Seek "=", m_Nv, m_NvArea, RsGDdN!Plano, RsGDdN!Marca
            If Not RsPd.NoMatch Then
                RsPd.Edit
                RsPd![GD] = RsPd![GD] - RsGDdN!Cantidad
                RsPd.Update
            End If
        
            ' borra detalle
            RsGDdN.Delete
        
            RsGDdN.MoveNext
        Loop
    End If
    RsPd.Index = "NV-Plano-Item"

Case "G"

    RsPd.Index = "NV-Plano-Marca"
    RsGDdN.Seek "=", Numero.Text, 1
    If Not RsGDdN.NoMatch Then
        Do While Not RsGDdN.EOF
            If RsGDdN!Numero <> Numero.Text Then Exit Do
            RsPd.Seek "=", m_Nv, m_NvArea, RsGDdN!Plano, RsGDdN!Marca
            If Not RsPd.NoMatch Then
                RsPd.Edit
                RsPd![GD gal] = RsPd![GD gal] - RsGDdN!Cantidad
                RsPd.Update
            End If
        
            ' borra detalle
            RsGDdN.Delete
        
            RsGDdN.MoveNext
        Loop
    End If
    RsPd.Index = "NV-Plano-Item"

Case "E"
    
    RsGDdE.Seek "=", Numero.Text, 1
    If Not RsGDdE.NoMatch Then
        Do While Not RsGDdE.EOF
            If RsGDdE!Numero <> Numero.Text Then Exit Do
            ' borra detalle
            RsGDdE.Delete
    
            RsGDdE.MoveNext
        Loop
    End If
    
Case "P" ' pernos
    
    With RsDoc
    .Seek "=", "GP", Numero.Text, 1
    If Not .NoMatch Then
    
        Do While Not .EOF
            If !Tipo <> "GP" Or !Numero <> Numero.Text Then Exit Do
            ' borra detalle
            .Delete
            .MoveNext
            
        Loop
    End If
    End With
    
End Select

End Sub
Private Sub Campos_Limpiar()
Numero.Text = ""
m_Tipo = "N"
CbTipo.Text = "Normal"
'GDespecial.Value = 0
Fecha.Text = Fecha_Vacia
'Fecha.Text = Format(Now, Fecha_Format)
m_Nv = 0
Nv.Text = ""
ComboNV.Text = " "
m_Nv = 0
rut.Text = ""
Razon.Text = ""
Direccion.Text = ""
Comuna.Text = ""

Oc.Text = ""
LugarRecepcion.Text = ""

ComboMarca.Clear
ComboPlano.Clear

guiaEspecial = False

Detalle_Limpiar Detalle_Normal, n_columnas_N
Detalle_Limpiar Detalle_Especial, n_columnas_E
Detalle_Limpiar Detalle_Pernos, n_columnas_P

CbChofer.Text = ""
CbPatente.Text = ""
Obs(0).Text = ""
Obs(1).Text = ""
Obs(2).Text = ""
Obs(2).visible = False
TotalPeso.Caption = ""
TotalPrecio.Caption = ""

'Capturador = False
contador_registros = 0

For i = 0 To 9
Arr_Bultos(i) = 0
Next

' solo cliente sedgman
Frame_Densidad.visible = False ' a partir de 18/02/10

End Sub
Private Sub Detalle_Limpiar(Detalle As Control, n_columnas As Integer)
Dim fi As Integer, co As Integer
For fi = 1 To n_filas
    For co = 1 To n_columnas
        Detalle.TextMatrix(fi, co) = ""
    Next
Next
Detalle.Row = 1
'Detalle_Normal

End Sub
Private Sub Obs_KeyPress(Index As Integer, KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim cambia_titulo As Boolean, n_Copias As Integer, m_NVprt As Double
cambia_titulo = True
'Accion = ""
Select Case Button.Index
Case 1 ' agregar

    Accion = "Agregando"
    Botones_Enabled 0, 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    
    Numero.Text = Documento_Numero_Nuevo(RsGDc, "Numero")

    Numero.Enabled = True
    Numero.SetFocus
    
Case 2 ' modificar
    Accion = "Modificando"
    Botones_Enabled 0, 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    Numero.Enabled = True
    Numero.SetFocus
Case 3 ' eliminar
    Accion = "Eliminando"
    Botones_Enabled 0, 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    Numero.Enabled = True
    Numero.SetFocus
    
Case 4 ' imprimir

    Accion = "Imprimiendo"
    If Numero.Text = "" Then
        Botones_Enabled 0, 0, 0, 1, 0, 1, 0
        Campos_Enabled False
        Numero.Enabled = True
        Numero.SetFocus
    Else
        
'        n_Copias = 1
'        PrinterNCopias.Numero_Copias = n_Copias
'        PrinterNCopias.Show 1
'        n_Copias = PrinterNCopias.Numero_Copias
    
'        If n_Copias > 0 Then
        Dim m_Numero As String, m_obra As String, m_Tipo As String

        m_Numero = Numero.Text
        m_NVprt = Val(Nv.Text)
        m_obra = ComboNV.Text
        m_Tipo = Left(CbTipo.Text, 1)

        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus

        If MsgBox("¿ IMPRIMIR GUIA DE DESPACHO ?", vbYesNo) = vbYes Then
        
            Select Case FormatoGuia
            Case "EML"
                GD_PrintLegal str(m_Numero), m_obra ' Mid(m_obra, 8)
            Case "AYD"
                ' AyD
                GD_PrintLegal_AyD str(m_Numero), Mid(m_obra, 8)
            Case "DELSA"
                ' Delsa
                GD_PrintLegal_Delsa str(m_Numero), Mid(m_obra, 8)
            Case "EIFFEL"
                ' eiffel
                GD_PrintLegal_Eiffel str(m_Numero), m_obra
            End Select
            
            ' grabar "impresa" en GD Cabecera, para que no se pueda modificacar una vez que este impresa
            RsGDc.Edit
            RsGDc!impresa = True
            RsGDc.Update
            
        End If
        
        If MsgBox("¿ IMPRIMIR PACKING LIST ?", vbYesNo) = vbYes Then
        
            'Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Docs")

            Dim m_ImpresoraNombre As String

            prt_escoger.ImpresoraNombre = ""
            prt_escoger.Show 1
            m_ImpresoraNombre = prt_escoger.ImpresoraNombre

            GD_Prepara m_Numero, m_NVprt, m_obra ', m_ImpresoraNombre
            
'            Campos_Limpiar
'            Numero.Enabled = True
'            Numero.SetFocus
            
            GD_PrintLegal_EML m_Numero, m_Tipo
            
        End If

        'End If
        
    End If

Case 5 ' Anular

    Accion = "Autorizando"
    If Numero.Text = "" Then
        Botones_Enabled 0, 0, 0, 0, 1, 1, 0
        Campos_Enabled False
        Numero.Enabled = True
        Numero.SetFocus
    Else
        
        If MsgBox("¿ AUTORIZAR ?", vbYesNo) = vbYes Then
            
            ' autorizar gd para que sea modificada
            ' e impresa
            RsGDc.Edit
            RsGDc!impresa = False
'            RsGDc!autorizada = True
            RsGDc.Update
            
        End If

        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus

    End If
    
Case 6 ' separador
Case 7 ' DesHacer
    If Numero.Text = "" Then
        GoTo DesHace
    Else
        If Accion = "Imprimiendo" Then
            GoTo DesHace
        Else
            If CbTipo.Text = "Especial" Then
'            If GDespecial.Enabled Then
                If MsgBox("¿Seguro que quiere DesHacer?", vbYesNo) = vbYes Then
                    GoTo DesHace
                End If
            Else
DesHace:
                Privilegios
                
                Campos_Limpiar
                Campos_Enabled False
                Accion = ""
            End If
        End If
    End If

Case 8 ' grabar

    If Doc_Validar Then
        
        If MsgBox("¿ GRABAR " & Obj & " ?", vbYesNo) = vbYes Then
            
            If guiaEspecial Then
            
                'Dim MiValor As String, Codigo As String
                
                'Codigo = Acceso_Lee(SqlRsSc, "GDE")
                
                'MiValor = InputBox("Debe digitar codigo de autorizacion", "Atencion", "")
                'If MiValor = Codigo Then
                    GoTo Sigue
                'Else
                '    MsgBox "Codigo No Valido !!!"
                'End If
                
            Else
                GoTo Sigue
            End If
    
            If False Then
Sigue:
                If Accion = "Agregando" Then
                    Doc_Grabar True
                Else
                    Doc_Grabar False
                End If
                            
                m_Numero = Numero.Text
                m_NVprt = Val(Nv.Text)
                m_obra = ComboNV.Text
                m_Tipo = Left(CbTipo.Text, 1)
    
                Campos_Limpiar
                Numero.Enabled = True
                Numero.SetFocus
                
                If MsgBox("¿ IMPRIMIR " & Obj & " ?", vbYesNo) = vbYes Then
                
    '                If Empresa.Rut = Rut_Eml Then
                    Select Case FormatoGuia
                    Case "EML"
                        'Eml
                        GD_PrintLegal str(m_Numero), Mid(m_obra, 8)
                    Case "AYD"
                        ' AyD
                        GD_PrintLegal_AyD str(m_Numero), Mid(m_obra, 8)
                    Case "DELSA"
                        ' Delsa
                        GD_PrintLegal_Delsa str(m_Numero), Mid(m_obra, 8)
                    Case "EIFFEL"
                        ' Eiffel
                        GD_PrintLegal_Eiffel str(m_Numero), m_obra
                    End Select
                    
                    RsGDc.Seek "=", Numero.Text
                    If Not RsGDc.NoMatch Then
                        RsGDc.Edit
                        RsGDc!impresa = True
                        RsGDc.Update
                    End If
                
                End If
                
                If MsgBox("¿ IMPRIMIR PACKING LIST ?", vbYesNo) = vbYes Then
            
                    Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Docs")
    
                    GD_Prepara m_Numero, m_NVprt, m_obra ', m_ImpresoraNombre
                            
                    GD_PrintLegal_EML m_Numero, m_Tipo
                
                End If
                
            End If
            
            Botones_Enabled 0, 0, 0, 0, 0, 1, 0
            
            Campos_Limpiar
            Campos_Enabled False
            Numero.Enabled = True
            Numero.SetFocus
            
            If Accion = "Agregando" Then Numero.Text = Documento_Numero_Nuevo(RsGDc, "Numero")
            
        End If
        
    End If
    
Case 9 ' separador
Case 10 ' clientes
    MousePointer = vbHourglass
    Load Clientes
    MousePointer = vbDefault
    Clientes.Show 1
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
                            btn_Anular As Boolean, _
                            btn_DesHacer As Boolean, btn_Grabar As Boolean)
                            
btnAgregar.Enabled = btn_Agregar

btnModificar.Enabled = btn_Modificar
btnEliminar.Enabled = btn_Eliminar
btnImprimir.Enabled = btn_Imprimir

If Usuario.AccesoTotal Then
    btnAnular.Enabled = btn_Anular
Else
    btnAnular.Enabled = False
End If

btnDesHacer.Enabled = btn_DesHacer

btnGrabar.Enabled = btn_Grabar

btnAgregar.Value = tbrUnpressed
btnModificar.Value = tbrUnpressed
btnEliminar.Value = tbrUnpressed
btnImprimir.Value = tbrUnpressed
btnAnular.Value = tbrUnpressed
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

btnCapturadorLeer.Enabled = Si
btnBultos.Enabled = Si
btnBarrick.Enabled = Si

btnSearch.Enabled = Si
CbTipo.Enabled = Si
'GDespecial.Enabled = Si
Nv.Enabled = Si
ComboNV.Enabled = Si

Oc.Enabled = Si
LugarRecepcion.Enabled = Si

Detalle_Normal.Enabled = Si
Detalle_Especial.Enabled = Si
Detalle_Pernos.Enabled = Si
CbChofer.Enabled = Si
CbPatente.Enabled = Si
Obs(0).Enabled = Si
Obs(1).Enabled = Si
Obs(2).Enabled = Si

'Depende_ITOPyG = True ' ???? 06/04/08 por delsa

End Sub
Private Sub btnSearch_Click()

Search.Muestra data_file, "Clientes", "RUT", "Razon Social", "Cliente", "Clientes"
rut.Text = Search.Codigo

If rut.Text <> "" Then
    RsCl.Seek "=", rut.Text
    If RsCl.NoMatch Then
        MsgBox "CLIENTE NO EXISTE"
        rut.SetFocus
    Else
'        Rut.Text = ""
        Razon.Text = Search.Descripcion
        Direccion.Text = RsCl!Direccion
        Comuna.Text = NoNulo(RsCl!Comuna)
        If ComboNV.Text <> "" Then Detalle_Normal.Enabled = True
    End If
End If
End Sub
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
' RUTINAS PARA EL FLEXGRID
Private Sub Detalle_Normal_Click()
If Accion = "Imprimiendo" Then Exit Sub
After_Detalle_Click
End Sub
Private Sub After_Detalle_Click()
ComboPlano.visible = False
ComboMarca.visible = False
Select Case Detalle_Normal.col
    Case 1 ' plano
        If Detalle_Normal <> "" Then ComboPlano.Text = Detalle_Normal
        ComboPlano.Top = Detalle_Normal.CellTop + Detalle_Normal.Top
        ComboPlano.Left = Detalle_Normal.CellLeft + Detalle_Normal.Left
        ComboPlano.Width = Int(Detalle_Normal.CellWidth * 1.5)
        ComboPlano.visible = True
        ComboMarca.visible = False
    Case 3 ' marca
        ComboMarca_Poblar Detalle_Normal.TextMatrix(Detalle_Normal.Row, 1)
        If Detalle_Normal <> "" Then ComboMarca.Text = Detalle_Normal
        ComboMarca.Top = Detalle_Normal.CellTop + Detalle_Normal.Top
        ComboMarca.Left = Detalle_Normal.CellLeft + Detalle_Normal.Left
        ComboMarca.Width = Int(Detalle_Normal.CellWidth * 1.5)
        ComboPlano.visible = False
        ComboMarca.visible = True
    Case Else
End Select
End Sub
Private Sub Detalle_Normal_DblClick()
If Accion = "Imprimiendo" Then Exit Sub
MSFlexGridEdit_N Detalle_Normal, txtEditGD, 32
End Sub
Private Sub Detalle_Normal_GotFocus()
If txtEditGD.visible Then
    Detalle_Normal = txtEditGD
    txtEditGD.visible = False
End If
End Sub
Private Sub Detalle_Normal_LeaveCell()
If txtEditGD.visible Then
    Detalle_Normal = txtEditGD
    txtEditGD.visible = False
End If
End Sub
Private Sub Detalle_Normal_KeyPress(KeyAscii As Integer)
If Accion = "Imprimiendo" Then Exit Sub
MSFlexGridEdit_N Detalle_Normal, txtEditGD, KeyAscii
End Sub
Private Sub txtEditGD_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case m_Tipo
Case "N", "G"
    EditKeyCode_N Detalle_Normal, txtEditGD, KeyCode, Shift
Case "E"
    EditKeyCode_E Detalle_Especial, txtEditGD, KeyCode, Shift
Case "P"
    EditKeyCode_P Detalle_Pernos, txtEditGD, KeyCode, Shift
End Select
End Sub
Sub EditKeyCode_N(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
' rutina que se ejecuta con los keydown de Edt
Dim m_fil As Integer, m_col As Integer, dif As Integer
m_fil = MSFlexGrid.Row
m_col = MSFlexGrid.col
dif = Val(Detalle_Normal.TextMatrix(MSFlexGrid.Row, 5)) - Val(Detalle_Normal.TextMatrix(MSFlexGrid.Row, 6))
Select Case KeyCode
Case vbKeyEscape ' Esc
    Edt.visible = False
    MSFlexGrid.SetFocus
Case vbKeyReturn ' Enter
    Select Case m_col
    Case 7 ' Cantidad a Despachar
        If Despachada_Validar(MSFlexGrid.col, dif, Edt) Then
            MSFlexGrid.SetFocus
            DoEvents
            Fila_Calcular_Normal m_fil, True
        End If
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
        Fila_Calcular_Normal m_fil, True
    End Select
    Cursor_Mueve_N MSFlexGrid
Case vbKeyUp ' Flecha Arriba
    Select Case m_col
    Case 7 ' Cantidad a Despachar
        If Despachada_Validar(MSFlexGrid.col, dif, Edt) Then
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
    Case 7 ' Cantidad a Despachar
        If Despachada_Validar(MSFlexGrid.col, dif, Edt) Then
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
Private Function Despachada_Validar(Colu As Integer, porDespachar As Integer, Edt As Control) As Boolean
' verifica que CRecibida-CDespachada >= CADespachar
Despachada_Validar = True
If Colu <> 7 Then Exit Function
If porDespachar < Val(Edt) Then
    MsgBox "Sólo quedan " & porDespachar & " por Despachar", , "ATENCIÓN"
    Despachada_Validar = False
End If
End Function
Private Sub txtEditGD_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Sub MSFlexGridEdit_N(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)
Select Case MSFlexGrid.col
Case 1, 2, 3
'    After_Detalle_Click
Case 4, 5, 6, 8, 9, 10, 12
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
Private Sub Detalle_Normal_KeyDown(KeyCode As Integer, Shift As Integer)
If Accion = "Imprimiendo" Then Exit Sub
If KeyCode = vbKeyF2 Then
    MSFlexGridEdit_N Detalle_Normal, txtEditGD, 32
End If
End Sub
Private Sub Cursor_Mueve_N(MSFlexGrid As Control)
Select Case MSFlexGrid.col
Case 7
    MSFlexGrid.col = MSFlexGrid.col + 3
Case 10, 13 ' 12
    If MSFlexGrid.Row + 1 < MSFlexGrid.Rows Then
        MSFlexGrid.col = 1
        MSFlexGrid.Row = MSFlexGrid.Row + 1
    End If
Case Else
    MSFlexGrid.col = MSFlexGrid.col + 1
End Select
End Sub
Private Sub Cursor_Mueve_E(MSFlexGrid As Control)
Select Case MSFlexGrid.col
Case 4
    MSFlexGrid.col = MSFlexGrid.col + 2
Case 6 Or 7
    If MSFlexGrid.Row + 1 < MSFlexGrid.Rows Then
        MSFlexGrid.col = 1
        MSFlexGrid.Row = MSFlexGrid.Row + 1
    End If
Case Else
    MSFlexGrid.col = MSFlexGrid.col + 1
End Select
End Sub
Private Sub Cursor_Mueve_P(MSFlexGrid As Control)
Select Case MSFlexGrid.col
Case 1
    MSFlexGrid.col = MSFlexGrid.col + 1
Case 3
    MSFlexGrid.col = MSFlexGrid.col + 1
Case 4
    If MSFlexGrid.Row + 1 < MSFlexGrid.Rows Then
        MSFlexGrid.col = 1
        MSFlexGrid.Row = MSFlexGrid.Row + 1
    End If
Case Else
    MSFlexGrid.col = MSFlexGrid.col + 1
End Select
End Sub
Private Sub Fila_Calcular_Normal(Fila As Integer, Actualizar As Boolean)
' actualiza solo linea, y totales generales

n7 = m_CDbl(Detalle_Normal.TextMatrix(Fila, 7))
n8 = m_CDbl(Detalle_Normal.TextMatrix(Fila, 8))
n11 = m_CDbl(Detalle_Normal.TextMatrix(Fila, 11))

' peso total
Detalle_Normal.TextMatrix(Fila, 9) = Format(n7 * n8, num_Formato)
' precio total
Detalle_Normal.TextMatrix(Fila, 12) = Format(n7 * n8 * n11, num_fmtgrl)

If Actualizar Then Detalle_Sumar_Normal

End Sub
Private Sub Detalle_Sumar_Normal()
Dim Tot_Kilos As Double, Tot_Precio As Double
Tot_Kilos = 0
Tot_Precio = 0
For i = 1 To n_filas
    Tot_Kilos = Tot_Kilos + m_CDbl(Detalle_Normal.TextMatrix(i, 9))
    Tot_Precio = Tot_Precio + m_CDbl(Detalle_Normal.TextMatrix(i, 12))
Next

TotalPeso.Caption = Format(Tot_Kilos, num_Formato)
TotalPrecio.Caption = Format(Tot_Precio, num_fmtgrl)

End Sub
' FIN RUTINAS PARA FLEXGRID
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
Private Sub Detalle_Config_Normal()

Dim i As Integer, ancho As Integer

Detalle_Normal.Left = 100
Detalle_Normal.WordWrap = True
Detalle_Normal.RowHeight(0) = 450
Detalle_Normal.Rows = n_filas + 1
Detalle_Normal.Cols = n_columnas_N + 1

Detalle_Normal.TextMatrix(0, 0) = ""
Detalle_Normal.TextMatrix(0, 1) = "Plano"
Detalle_Normal.TextMatrix(0, 2) = "Rev"                   '*
Detalle_Normal.TextMatrix(0, 3) = "Marca"
Detalle_Normal.TextMatrix(0, 4) = "Descripción"           '*
'Detalle_Normal.TextMatrix(0, 5) = "Cant Reci."            '*
Detalle_Normal.TextMatrix(0, 5) = "Cant itopp"            '*
Detalle_Normal.TextMatrix(0, 6) = "Cant Desp."            '*
Detalle_Normal.TextMatrix(0, 7) = "a Desp."
Detalle_Normal.TextMatrix(0, 8) = "Peso Unitario"         '*
Detalle_Normal.TextMatrix(0, 9) = "Peso TOTAL"            '*
Detalle_Normal.TextMatrix(0, 10) = "D"           '* nueva linea : densidad
Detalle_Normal.TextMatrix(0, 11) = "Precio Unitario"
Detalle_Normal.TextMatrix(0, 12) = "Precio TOTAL"         '*
Detalle_Normal.TextMatrix(0, 13) = "Tipo Embalaje"

Detalle_Normal.ColWidth(0) = 300
Detalle_Normal.ColWidth(1) = 2000 ' plano
Detalle_Normal.ColWidth(2) = 390
Detalle_Normal.ColWidth(3) = 2200 ' marca
Detalle_Normal.ColWidth(4) = 1150
Detalle_Normal.ColWidth(5) = 500
Detalle_Normal.ColWidth(6) = 500
Detalle_Normal.ColWidth(7) = 500
Detalle_Normal.ColWidth(8) = 800
Detalle_Normal.ColWidth(9) = 800
Detalle_Normal.ColWidth(10) = 200
Detalle_Normal.ColWidth(11) = 700
Detalle_Normal.ColWidth(12) = 750
Detalle_Normal.ColWidth(13) = 750

ancho = 350 ' con scroll vertical
'ancho = 100 ' sin scroll vertical

TotalPeso.Width = Detalle_Normal.ColWidth(12)
For i = 0 To n_columnas_N
    If i = 9 Then TotalPeso.Left = 0 ' ancho + Detalle_Normal.Left - 350
    If i = 12 Then TotalPrecio.Left = 0 ' ancho + Detalle_Normal.Left - 350
    ancho = ancho + Detalle_Normal.ColWidth(i)
Next

Detalle_Normal.Width = ancho
Me.Width = ancho + Detalle_Normal.Left * 2

For i = 1 To n_filas
    Detalle_Normal.TextMatrix(i, 0) = i
Next

For i = 1 To n_filas

    Detalle_Normal.Row = i
    Detalle_Normal.col = 2
    Detalle_Normal.CellAlignment = flexAlignLeftCenter
    Detalle_Normal.CellForeColor = vbRed
    Detalle_Normal.col = 3
    Detalle_Normal.CellAlignment = flexAlignLeftCenter
    Detalle_Normal.CellForeColor = vbBlue
    Detalle_Normal.col = 4
    Detalle_Normal.CellForeColor = vbBlue
    Detalle_Normal.col = 5
    Detalle_Normal.CellForeColor = vbBlue
    Detalle_Normal.col = 6
    Detalle_Normal.CellForeColor = vbBlue
    
    Detalle_Normal.col = 8
    Detalle_Normal.CellForeColor = vbRed
    Detalle_Normal.col = 9
    Detalle_Normal.CellForeColor = vbRed
    Detalle_Normal.col = 10
    Detalle_Normal.CellForeColor = vbRed
    Detalle_Normal.col = 11
    Detalle_Normal.CellForeColor = vbRed
    
Next

txtEditGD.Text = ""

'Top = Form_CentraY(Me)
'Left = Form_CentraX(Me)

Detalle_Normal.Enabled = False

End Sub
Private Sub Detalle_Config_Especial()

Dim i As Integer, ancho As Integer

Detalle_Especial.Left = 100
Detalle_Especial.WordWrap = True
Detalle_Especial.RowHeight(0) = 450
Detalle_Especial.Rows = n_filas + 1
Detalle_Especial.Cols = n_columnas_E + 1

Detalle_Especial.TextMatrix(0, 0) = ""
Detalle_Especial.TextMatrix(0, 1) = "Cantidad"
Detalle_Especial.TextMatrix(0, 2) = "Unidad"
Detalle_Especial.TextMatrix(0, 3) = "Descripción"
Detalle_Especial.TextMatrix(0, 4) = "Peso Unitario"
Detalle_Especial.TextMatrix(0, 5) = "Peso TOTAL"           '*
Detalle_Especial.TextMatrix(0, 6) = "Precio Unitario"
Detalle_Especial.TextMatrix(0, 7) = "Precio TOTAL"         '*

Detalle_Especial.ColWidth(0) = 300
Detalle_Especial.ColWidth(1) = 800
Detalle_Especial.ColWidth(2) = 700
Detalle_Especial.ColWidth(3) = 3700
Detalle_Especial.ColWidth(4) = 700
Detalle_Especial.ColWidth(5) = 800
Detalle_Especial.ColWidth(6) = 700
Detalle_Especial.ColWidth(7) = 800

ancho = 350 ' con scroll vertical
'ancho = 100 ' sin scroll vertical

TotalPeso.Width = Detalle_Especial.ColWidth(5)
TotalPrecio.Width = Detalle_Especial.ColWidth(7)
For i = 0 To n_columnas_E
    If i = 5 Then TotalPeso.Left = ancho + Detalle_Especial.Left - 350
    If i = 7 Then TotalPrecio.Left = ancho + Detalle_Especial.Left - 350
    ancho = ancho + Detalle_Especial.ColWidth(i)
Next

Detalle_Especial.Width = ancho
'Me.Width = ancho + Detalle_Especial.Left * 2

For i = 1 To n_filas
    Detalle_Especial.TextMatrix(i, 0) = i
Next

For i = 1 To n_filas
    Detalle_Especial.Row = i
    Detalle_Especial.col = 2
    Detalle_Especial.CellAlignment = flexAlignLeftCenter
    Detalle_Especial.col = 3
    Detalle_Especial.CellAlignment = flexAlignLeftCenter
    Detalle_Especial.col = 5
    Detalle_Especial.CellForeColor = vbRed
    Detalle_Especial.col = 7
    Detalle_Especial.CellForeColor = vbRed
Next

txtEditGD.Text = ""

'Top = Form_CentraY(Me)
'Left = Form_CentraX(Me)

Detalle_Especial.Enabled = False

End Sub
Private Sub Detalle_Config_Pernos()

Dim i As Integer, ancho As Integer

n_filas = 17 ' 20/01/16
n_filas = 18 ' 04/03/16 solo por manuel

Detalle_Pernos.Left = 100
Detalle_Pernos.WordWrap = True
Detalle_Pernos.RowHeight(0) = 450
Detalle_Pernos.Rows = n_filas + 1
Detalle_Pernos.Cols = n_columnas_P + 1

Detalle_Pernos.TextMatrix(0, 0) = ""
Detalle_Pernos.TextMatrix(0, 1) = "Codigo"
Detalle_Pernos.TextMatrix(0, 2) = "Descripción"
Detalle_Pernos.TextMatrix(0, 3) = "Cantidad"
Detalle_Pernos.TextMatrix(0, 4) = "Precio Unitario"
Detalle_Pernos.TextMatrix(0, 5) = "Precio TOTAL"         '*

Detalle_Pernos.ColWidth(0) = 300
Detalle_Pernos.ColWidth(1) = 1500
Detalle_Pernos.ColWidth(2) = 3700
Detalle_Pernos.ColWidth(3) = 900
Detalle_Pernos.ColWidth(4) = 900
Detalle_Pernos.ColWidth(5) = 1200

ancho = 350 ' con scroll vertical
'ancho = 100 ' sin scroll vertical

'TotalPeso.Width = DETALLE_PERNOs.ColWidth(5)
'TotalPrecio.Width = DETALLE_PERNOs.ColWidth(7)
For i = 0 To n_columnas_P
'    If i = 5 Then TotalPeso.Left = ancho + DETALLE_PERNOs.Left - 350
'    If i = 7 Then TotalPrecio.Left = ancho + DETALLE_PERNOs.Left - 350
    ancho = ancho + Detalle_Pernos.ColWidth(i)
Next

Detalle_Pernos.Width = ancho
'Me.Width = ancho + Detalle_Pernos.Left * 2

For i = 1 To n_filas
    Detalle_Pernos.TextMatrix(i, 0) = i
Next

For i = 1 To n_filas
    Detalle_Pernos.Row = i
'    Detalle_Pernos.col = 2
'    Detalle_Pernos.CellAlignment = flexAlignLeftCenter
'    Detalle_Pernos.col = 3
'    Detalle_Pernos.CellAlignment = flexAlignLeftCenter
    Detalle_Pernos.col = 5
    Detalle_Pernos.CellForeColor = vbRed
Next

txtEditGD.Text = ""

'Top = Form_CentraY(Me)
'Left = Form_CentraX(Me)

'Detalle_Pernos.Enabled = False
Detalle_Pernos.Enabled = False

End Sub
Private Sub Detalle_Config_Pintura()

Dim i As Integer, ancho As Integer

Detalle_Especial.Left = 100
Detalle_Especial.WordWrap = True
Detalle_Especial.RowHeight(0) = 450
Detalle_Especial.Rows = n_filas + 1
Detalle_Especial.Cols = n_columnas_E + 1

Detalle_Especial.TextMatrix(0, 0) = ""
Detalle_Especial.TextMatrix(0, 1) = "Cantidad"
Detalle_Especial.TextMatrix(0, 2) = "Unidad"
Detalle_Especial.TextMatrix(0, 3) = "Descripción"
Detalle_Especial.TextMatrix(0, 4) = "m2 Unitario"
Detalle_Especial.TextMatrix(0, 5) = "m2 TOTAL"           '*
Detalle_Especial.TextMatrix(0, 6) = "Precio Unitario"
Detalle_Especial.TextMatrix(0, 7) = "Precio TOTAL"         '*

Detalle_Especial.ColWidth(0) = 300
Detalle_Especial.ColWidth(1) = 800
Detalle_Especial.ColWidth(2) = 700
Detalle_Especial.ColWidth(3) = 3700
Detalle_Especial.ColWidth(4) = 700
Detalle_Especial.ColWidth(5) = 800
Detalle_Especial.ColWidth(6) = 700
Detalle_Especial.ColWidth(7) = 800

ancho = 350 ' con scroll vertical
'ancho = 100 ' sin scroll vertical

TotalPeso.Width = Detalle_Especial.ColWidth(5)
TotalPrecio.Width = Detalle_Especial.ColWidth(7)
For i = 0 To n_columnas_E
    If i = 5 Then TotalPeso.Left = ancho + Detalle_Especial.Left - 350
    If i = 7 Then TotalPrecio.Left = ancho + Detalle_Especial.Left - 350
    ancho = ancho + Detalle_Especial.ColWidth(i)
Next

Detalle_Especial.Width = ancho
'Me.Width = ancho + Detalle_Especial.Left * 2

For i = 1 To n_filas
    Detalle_Especial.TextMatrix(i, 0) = i
Next

For i = 1 To n_filas
    Detalle_Especial.Row = i
    Detalle_Especial.col = 2
    Detalle_Especial.CellAlignment = flexAlignLeftCenter
    Detalle_Especial.col = 3
    Detalle_Especial.CellAlignment = flexAlignLeftCenter
    Detalle_Especial.col = 5
    Detalle_Especial.CellForeColor = vbRed
    Detalle_Especial.col = 7
    Detalle_Especial.CellForeColor = vbRed
Next

txtEditGD.Text = ""

Detalle_Especial.Enabled = False

End Sub
Private Sub Fila_Calcular_Especial(Fila As Integer, Actualizar As Boolean)
' actualiza solo linea, y totales generales

n1 = m_CDbl(Detalle_Especial.TextMatrix(Fila, 1))
n4 = m_CDbl(Detalle_Especial.TextMatrix(Fila, 4))
n6 = m_CDbl(Detalle_Especial.TextMatrix(Fila, 6))

' peso total
Detalle_Especial.TextMatrix(Fila, 5) = Format(n1 * n4, num_Formato)
' precio total
Detalle_Especial.TextMatrix(Fila, 7) = Format(n1 * n4 * n6, num_fmtgrl)

If Actualizar Then Detalle_Sumar_Especial

End Sub
Private Sub Fila_Calcular_Pernos(Fila As Integer, Actualizar As Boolean)
' actualiza solo linea, y totales generales

n4 = m_CDbl(Detalle_Pernos.TextMatrix(Fila, 3))
n6 = m_CDbl(Detalle_Pernos.TextMatrix(Fila, 4))

' precio total
Detalle_Pernos.TextMatrix(Fila, 5) = Format(n4 * n6, num_fmtgrl)

If Actualizar Then Detalle_Sumar_Pernos

End Sub
Private Sub Detalle_Sumar_Especial()
If m_Tipo = "N" Then Exit Sub
Dim Tot_Kilos As Double, Tot_Precio As Double
Tot_Kilos = 0
Tot_Precio = 0
For i = 1 To n_filas
    Tot_Kilos = Tot_Kilos + m_CDbl(Detalle_Especial.TextMatrix(i, 5))
    Tot_Precio = Tot_Precio + m_CDbl(Detalle_Especial.TextMatrix(i, 7))
Next

TotalPeso.Caption = Format(Tot_Kilos, num_Formato)
TotalPrecio.Caption = Format(Tot_Precio, num_fmtgrl)

If Tot_Precio = 0 Then
    Obs(0).Text = "NO CONSTITUYE VENTA" ' 22/06/98
Else
    If Obs(0).Text = "NO CONSTITUYE VENTA" Then
        Obs(0).Text = ""
    End If
End If
End Sub
Private Sub Detalle_Sumar_Pernos()

Dim Tot_Precio As Double
Tot_Precio = 0
For i = 1 To n_filas
    Tot_Precio = Tot_Precio + m_CDbl(Detalle_Pernos.TextMatrix(i, 5))
Next

TotalPrecio.Caption = Format(Tot_Precio, num_fmtgrl)

If Tot_Precio = 0 Then
    Obs(0).Text = "NO CONSTITUYE VENTA" ' 22/06/98
Else
    If Obs(0).Text = "NO CONSTITUYE VENTA" Then
        Obs(0).Text = ""
    End If
End If
End Sub
'Private Sub Detalle_Especial_Click()
'If Accion = "Imprimiendo" Then Exit Sub
'After_Detalle_e_Click
'End Sub
Private Sub Detalle_Especial_DblClick()
If Accion = "Imprimiendo" Then Exit Sub
MSFlexGridEdit_E Detalle_Especial, txtEditGD, 32
End Sub
Private Sub Detalle_Especial_GotFocus()
If txtEditGD.visible Then
    Detalle_Especial = txtEditGD
    txtEditGD.visible = False
End If
End Sub
Private Sub Detalle_Especial_KeyDown(KeyCode As Integer, Shift As Integer)
If Accion = "Imprimiendo" Then Exit Sub
If KeyCode = vbKeyF2 Then
    MSFlexGridEdit_E Detalle_Especial, txtEditGD, 32
End If
End Sub
Private Sub Detalle_Especial_KeyPress(KeyAscii As Integer)
If Accion = "Imprimiendo" Then Exit Sub
MSFlexGridEdit_E Detalle_Especial, txtEditGD, KeyAscii
End Sub
Private Sub Detalle_Especial_LeaveCell()
If txtEditGD.visible Then
    Detalle_Especial = txtEditGD
    txtEditGD.visible = False
End If
End Sub
Sub MSFlexGridEdit_E(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)

Select Case MSFlexGrid.col
Case 2
    Edt.MaxLength = 3
Case 3
'    Edt.MaxLength = 21 '50
'    Edt.MaxLength = 40
    Edt.MaxLength = 50
Case Else
    Edt.MaxLength = 10
End Select

Select Case MSFlexGrid.col
Case 5, 7
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
Sub EditKeyCode_E(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
' rutina que se ejecuta con los keydown de Edt
Dim m_col As Integer, m_fil As Integer
m_fil = MSFlexGrid.Row
m_col = MSFlexGrid.col
Select Case KeyCode
Case vbKeyEscape ' Esc
    Edt.visible = False
    MSFlexGrid.SetFocus
Case vbKeyReturn ' Enter
    MSFlexGrid.SetFocus
    DoEvents
    Fila_Calcular_Especial m_fil, True
    Cursor_Mueve_E MSFlexGrid
Case vbKeyUp ' Flecha Arriba
    MSFlexGrid.SetFocus
    DoEvents
    Fila_Calcular_Especial m_fil, True
Case vbKeyDown ' Flecha Abajo
    MSFlexGrid.SetFocus
    DoEvents
    Fila_Calcular_Especial m_fil, True
    If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
        MSFlexGrid.Row = MSFlexGrid.Row + 1
    End If
End Select
End Sub
Sub MSFlexGridEdit_P(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)

Select Case MSFlexGrid.col
Case 1 'codigo
'    Edt.Mask = ">&&&&&&&&&&&&&&&&&&&&"
    Edt.MaxLength = 15
Case 3
    Edt.MaxLength = 5 '4
Case 4 ' precio unitario
    Edt.MaxLength = 7
Case Else
    Edt.MaxLength = 10
End Select

Select Case MSFlexGrid.col
Case 2, 5
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
Sub EditKeyCode_P(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
' rutina que se ejecuta con los keydown de Edt
Dim m_col As Integer, m_fil As Integer
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
        Fila_Calcular_Pernos m_fil, True
    End Select
    Cursor_Mueve_P MSFlexGrid
Case vbKeyUp ' Flecha Arriba
    MSFlexGrid.SetFocus
    DoEvents
    Fila_Calcular_Pernos m_fil, True
Case vbKeyDown ' Flecha Abajo
    MSFlexGrid.SetFocus
    DoEvents
    Fila_Calcular_Pernos m_fil, True
    If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
        MSFlexGrid.Row = MSFlexGrid.Row + 1
    End If
End Select
End Sub
Private Sub Producto_Buscar()
Dim m_Codigo As String, fi As Integer
fi = Detalle_Pernos.Row
m_Codigo = Detalle_Pernos.TextMatrix(fi, 1)

With RsPrd
.Seek "=", m_Codigo
If .NoMatch Then
    MsgBox "PRODUCTO NO EXISTE"
    For i = 1 To n_columnas_P
        Detalle_Pernos.TextMatrix(fi, i) = ""
    Next
Else
    Detalle_Pernos.TextMatrix(fi, 2) = !Descripcion
    Detalle_Pernos.col = 2
End If
End With

End Sub
Private Sub Privilegios()

If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 1, 0, 0
End If

End Sub
Private Sub GD_PrintLegal_EML(Numero_GD, TipoGuia As String)
Dim repo As String
cr.WindowTitle = "GUIA DE DESPACHO Nº " & Numero_GD
cr.ReportSource = crptReport
cr.WindowState = crptMaximized
'Cr.WindowBorderStyle = crptFixedSingle
cr.WindowMaxButton = False
cr.WindowMinButton = False
'cr.Formulas(0) = "RAZON SOCIAL=""" & Empresa.Razon & """"
'cr.Formulas(1) = "GIRO=""" & "GIRO: " & Empresa.Giro & """"
'cr.Formulas(2) = "DIRECCION=""" & Empresa.Direccion & """"
'cr.Formulas(3) = "TELEFONOS=""" & "Teléfono: " & Empresa.Telefono1 & " " & Empresa.Comuna & """"
'cr.Formulas(4) = "RUT=""" & "RUT: " & Empresa.rut & """"

'MsgBox Certificado.Value

cr.DataFiles(0) = repo_file & ".MDB"

'If Tipo = "E" Then
'    CR.Formulas(5) = "COTIZACION=""" & "GUÍA DESP. Nº:" & """"
'    CR.ReportFileName = Drive_Server & Path_Server & EmpOC.Fantasia & "\Oc_Especial.Rpt"
'Else
'    CR.Formulas(5) = "COTIZACION=""" & "COTIZACIÓN Nº:" & """"
Select Case TipoGuia
Case "N"
    repo = "gdn_packinglist.rpt"
Case "E"
    repo = "gde_packinglist.rpt"
Case "P"
    repo = "gdp_packinglist.rpt"
End Select
cr.ReportFileName = Drive_Server & Path_Rpt & repo

cr.WindowTitle = "GUIA DE DESPACHO Nº " & Numero_GD & " " & repo

cr.Action = 1

End Sub
Private Sub Gd_Anular(Ini As Double, Fin As Double, Fecha As Date)
' anula guias de despacho dentro del rango
Dim n As Double
With RsGDc
For n = Ini To Fin
    .AddNew
    !Numero = n
    !Fecha = Fecha
    !Tipo = "E"
    !Nv = 138
    ![RUT CLiente] = "89784800-7"
    ![Observacion 1] = "NULA"
    .Update
Next
End With

With RsGDdE
For n = Ini To Fin
    .AddNew
    !Numero = n
    !linea = 1
    !Fecha = Fecha
    !Nv = 138
    ![RUT CLiente] = "89784800-7"
    !Cantidad = 1
    !unidad = "UNI"
    !Detalle = "NULA"
    .Update
Next
End With

End Sub
Private Sub ArchivoAbrir()

Dim mPath As String, mPathArchivo As String, mArchivo As String, p As Integer

cd.DialogTitle = "Buscar Carpeta"
cd.Filter = "Texto (*.txt)|*.txt|Todos los Archivos (*.*)|*.*"

' busca ultima ruta
mPath = GetSetting("scp", "gd", "ruta")
If mPath = "" Then mPath = "C:"

mPath = Directorio(mPath)

cd.InitDir = mPath
cd.ShowOpen

mPathArchivo = cd.filename

If mPathArchivo = "" Then
    Exit Sub
End If

mPathArchivo = cd.filename

' separa path y archivo
p = InStrLast(mPathArchivo, "\")
If p > 0 Then

    ' guarda ultima ruta usada
    SaveSetting "scp", "gd", "ruta", mPathArchivo

    mPath = Left(mPathArchivo, p)
    mArchivo = Mid(mPathArchivo, p + 1)
    
'    lblCarpeta.Caption = m_Path
    
    ArchivoLeer mPath, mArchivo
    
End If

End Sub
Private Sub ArchivoLeer(Path As String, Archivo As String)
' lee archivo del honey
' abre archivo
Dim li As Integer, mPlano As String, mMarca As String
Dim RsPaso As Recordset, arreglo(99, 9) As String
Dim i As Integer, j As Integer, k As Integer

li = 0

' lee archivo y lo deja en arreglo
Open Path & Archivo For Input As #1
Do While Not EOF(1)
    Line Input #1, Reg
'    Debug.Print Reg
    split Reg, "/", aCampos, NumeroCampos
    
    li = li + 1
    
    arreglo(li, 1) = aCampos(4) ' marca
    arreglo(li, 2) = 1 ' cantidad
    arreglo(li, 3) = aCampos(7) ' nv
    
Loop
Close #1
'//////////////////////////////////
' buscas marcas repetidas
For i = 1 To li - 1

    For j = i + 1 To li
    
        If arreglo(i, 1) = arreglo(j, 1) Then
            
            ' suma cantidad
            arreglo(i, 2) = arreglo(i, 2) + 1
            
            ' elimina fila j y desplaza filas hacia arriba
            For k = j To li
                arreglo(k, 1) = arreglo(k + 1, 1)
                arreglo(k, 2) = arreglo(k + 1, 2)
                arreglo(k, 3) = arreglo(k + 1, 3)
            Next
            li = li - 1
            
            j = j - 1
            
        End If
    Next
Next
'For i = 1 To li
'    Debug.Print i & "|" & arreglo(i, 1) & "|" & arreglo(i, 2)
'Next
'//////////////////////////////////

If li > n_columnas_N Then
    MsgBox "Hay mas de " & n_columnas_N & " lecturas, solo se muestran las primeras " & n_columnas_N
    li = n_columnas_N
End If

' lleva arreglo a pantalla (grilla)
For i = 1 To li
  
    If i = 1 Then
        ' pone nv
        Nv.Text = arreglo(i, 3)
        Nv_LostFocus
    End If
    
    'mPlano = aCampos(4)
    mMarca = arreglo(i, 1)
        
    Set RsPaso = Dbm.OpenRecordset("SELECT * FROM [planos detalle] WHERE nv=" & Nv.Text & " AND marca='" & mMarca & "'")
    With RsPaso
    If .RecordCount > 1 Then
        Detalle_Normal.TextMatrix(i, 1) = !Plano
        Detalle_Normal.TextMatrix(i, 2) = !Rev
        Detalle_Normal.TextMatrix(i, 3) = !Marca
        Detalle_Normal.TextMatrix(i, 4) = !Descripcion
        Detalle_Normal.TextMatrix(i, 5) = ![ITO pp]
        Detalle_Normal.TextMatrix(i, 6) = ![GD]
        Detalle_Normal.TextMatrix(i, 7) = arreglo(i, 2) ' cantidad
        Detalle_Normal.TextMatrix(i, 8) = ![Peso]
        Detalle_Normal.TextMatrix(i, 9) = ![Peso] * arreglo(i, 2) ' peso total
    End If
    End With

Next

End Sub
