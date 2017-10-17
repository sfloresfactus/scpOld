VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form NV 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8040
      _ExtentX        =   14182
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
            Object.ToolTipText     =   "Nueva NV"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar NV"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir NV"
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
            Object.ToolTipText     =   "Grabar NV"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clientes"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5055
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   8916
      _Version        =   393216
      TabHeight       =   529
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "NV.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(10)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPintura"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl(9)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl(8)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblCentroCosto"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Fecha_Inicial"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Fecha_Final"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CbTipo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CbPintura"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Activa"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame_Pernos"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Observaciones"
      TabPicture(1)   =   "NV.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Obs(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Obs(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Obs(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Obs(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Especificaciones"
      TabPicture(2)   =   "NV.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cd"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame_Planos"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame_Especifiaciones"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "FrameCarta"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin MSComDlg.CommonDialog cd 
         Left            =   -68160
         Top             =   1680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame_Planos 
         Caption         =   "LISTADO DE PLANOS"
         Height          =   1095
         Left            =   -74640
         TabIndex        =   40
         Top             =   2760
         Width           =   6255
         Begin VB.CommandButton btnPlanosLeer 
            Caption         =   "Leer"
            Height          =   375
            Left            =   3720
            TabIndex        =   42
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton btnPlanosSubir 
            Caption         =   "Subir"
            Height          =   375
            Left            =   360
            TabIndex        =   41
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame_Especifiaciones 
         Caption         =   "ESPECIFICACIONES TECNICAS"
         Height          =   975
         Left            =   -74640
         TabIndex        =   37
         Top             =   1680
         Width           =   6255
         Begin VB.CommandButton btnEspTecLeer 
            Caption         =   "Leer"
            Height          =   375
            Left            =   3720
            TabIndex        =   39
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton btnEspecificacionesSubir 
            Caption         =   "Subir"
            Height          =   375
            Left            =   360
            TabIndex        =   38
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame FrameCarta 
         Caption         =   "CARTA COTIZACION"
         Height          =   975
         Left            =   -74640
         TabIndex        =   34
         Top             =   600
         Width           =   6255
         Begin VB.CommandButton btnCartaLeer 
            Caption         =   "Leer"
            Height          =   375
            Left            =   3720
            TabIndex        =   36
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton btnCartaSubir 
            Caption         =   "Subir"
            Height          =   375
            Left            =   360
            TabIndex        =   35
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame_Pernos 
         Caption         =   "PERNOS"
         Height          =   975
         Left            =   4440
         TabIndex        =   22
         Top             =   2280
         Width           =   2895
         Begin VB.CheckBox ListaPernosIncluida 
            Caption         =   "Lleva Pernos"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox ListaPernosRecibida 
            Caption         =   "Listado de Pernos Recibido"
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   600
            Width           =   2295
         End
      End
      Begin VB.CheckBox Activa 
         Caption         =   "&Activa (Si está activa, aparece en todos los listados)"
         Height          =   495
         Left            =   4320
         TabIndex        =   29
         Top             =   3600
         Width           =   3015
      End
      Begin VB.ComboBox CbPintura 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1680
         Width           =   2415
      End
      Begin VB.ComboBox CbTipo 
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   840
         Width           =   2295
      End
      Begin VB.Frame Frame 
         Caption         =   "CLIENTE"
         Height          =   2535
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   3735
         Begin VB.CommandButton btnSearch 
            Height          =   300
            Left            =   2880
            Picture         =   "NV.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   300
         End
         Begin VB.TextBox Direccion 
            Enabled         =   0   'False
            Height          =   300
            Left            =   240
            TabIndex        =   15
            Top             =   1440
            Width           =   3135
         End
         Begin VB.TextBox Comuna 
            Enabled         =   0   'False
            Height          =   300
            Left            =   240
            TabIndex        =   17
            Top             =   2040
            Width           =   3135
         End
         Begin VB.TextBox Razon 
            Enabled         =   0   'False
            Height          =   300
            Left            =   240
            TabIndex        =   13
            Top             =   840
            Width           =   3135
         End
         Begin MSMask.MaskEdBox Rut 
            Height          =   300
            Left            =   1200
            TabIndex        =   10
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            _Version        =   327680
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.Label lbl 
            Caption         =   "DIRECCIÓN"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   14
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lbl 
            Caption         =   "COMUNA"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   16
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "RUT"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lbl 
            Caption         =   "SEÑOR(ES)"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   12
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.TextBox Obs 
         Height          =   660
         Index           =   0
         Left            =   -74760
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   720
         Width           =   7260
      End
      Begin VB.TextBox Obs 
         Height          =   660
         Index           =   1
         Left            =   -74760
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   1440
         Width           =   7260
      End
      Begin VB.TextBox Obs 
         Height          =   660
         Index           =   2
         Left            =   -74760
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   2160
         Width           =   7260
      End
      Begin VB.TextBox Obs 
         Height          =   660
         Index           =   3
         Left            =   -74760
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   2880
         Width           =   7260
      End
      Begin MSMask.MaskEdBox Fecha_Final 
         Height          =   300
         Left            =   2280
         TabIndex        =   28
         Top             =   3840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   327680
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Fecha_Inicial 
         Height          =   300
         Left            =   2280
         TabIndex        =   26
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   327680
         PromptChar      =   "_"
      End
      Begin VB.Label lblCentroCosto 
         Caption         =   "Centro de Costo ( de SCP nuevo )"
         Height          =   255
         Left            =   600
         TabIndex        =   45
         Top             =   4440
         Width           =   2535
      End
      Begin VB.Label lbl 
         Caption         =   "Fecha Inicio Obras"
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   25
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label lbl 
         Caption         =   "Fecha Término Obras"
         Height          =   255
         Index           =   9
         Left            =   600
         TabIndex        =   27
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lblPintura 
         Caption         =   "Recubrimiento"
         Height          =   255
         Left            =   4440
         TabIndex        =   20
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   10
         Left            =   4440
         TabIndex        =   18
         Top             =   840
         Width           =   375
      End
   End
   Begin VB.TextBox Numero 
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   44
      Top             =   6480
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox ComboNV 
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   3650
   End
   Begin VB.TextBox Obra 
      Height          =   300
      Left            =   3240
      MaxLength       =   30
      TabIndex        =   43
      Top             =   720
      Width           =   3660
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   327680
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   7080
      Top             =   600
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
            Picture         =   "NV.frx":0156
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NV.frx":0268
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NV.frx":037A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NV.frx":048C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NV.frx":059E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NV.frx":06B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NV.frx":07C2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "OBRA"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "NÚMERO"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "FECHA"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "NV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Obj As String, Objs As String
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button
Private DbD As Database, RsCl As Recordset
Private DbH As Database, RsNvcH As Recordset
Private Dbm As Database
'Private RsNVc As Recordset
Private RsNVp As Recordset
Private Accion As String, d As Variant, prt As Printer, i As Integer
' 0: numero,  1: nombre obra
'Private aNv(2999, 1) As String
Private m_Nv As Double, m_NvArea As Integer
Private sql As String
Private mvarURL As String ' variable para link a pagina intanet
Private Const nv_files As String = "nv_files\"  ' carpeta del servidor de destino para guardar PDF de nv
Private m_PathDestino As String
Private mNv As NotaVenta
Private RsNvSql As New ADODB.Recordset
Private aNotaVenta(999, 1) As String
Private Sub Form_Load()

Set DbD = OpenDatabase(data_file)
Set RsCl = DbD.OpenRecordset("Clientes")
RsCl.Index = "RUT"

Set DbH = OpenDatabase(Drive_Server & Path_Mdb & "ScpHist")
'sql = "SELECT * FROM [nv cabecera]"
sql = "nv cabecera"
Set RsNvcH = DbH.OpenRecordset(sql)
RsNvcH.Index = "Numero"

Set Dbm = OpenDatabase(mpro_file)
'Set RsNVc = Dbm.OpenRecordset(sql)
'RsNVc.Index = Nv_Index ' "Número"

Set RsNVp = Dbm.OpenRecordset("Planos Cabecera")
RsNVp.Index = "NV-Plano"

' ojo
'Debug.Print Dbm.Name
'Dbm.Execute "DELETE * FROM [nv cabecera] WHERE numero=3741"

Inicializa

Me.Caption = Obj

If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 0, 0
End If

btnSearch.visible = False

CbTipo.AddItem "Producción"
CbTipo.AddItem "Servicio"

CbPintura.AddItem "Pintura"
CbPintura.AddItem "Galvanizado"
CbPintura.AddItem "Negro"

obrasLeer

'RsNVc.Index = "Numero"  ' debe ser ´numero

StatusBar.Panels(1) = EmpOC.Razon

m_NvArea = 0

' copia archivo a carpeta de servidor \\acr3006-dualpro\scp\mdb\nv_files
m_PathDestino = Drive_Server & Path_Mdb & nv_files

'Dim l As Integer
'l = nv_Leer(aNotaVenta)
'For i = 1 To l
'    Debug.Print aNotaVenta(i, 0); aNotaVenta(i, 1)
'Next

' lista centros de costo
'For i = 1 To centrosCostoTotal
'    Debug.Print aCeCo(i, 0); aCeCo(i, 1)
'Next

End Sub
Private Sub Form_Unload(Cancel As Integer)
DataBases_Cerrar
End Sub
Private Sub Inicializa()

Obj = "Nota de Venta"
Objs = "Notas de Venta"

Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnDesHacer = Toolbar.Buttons(6)
Set btnGrabar = Toolbar.Buttons(7)

Accion = ""

'btnSearch.Visible = False
'btnSearch.ToolTipText = "Busca Cliente"

'Numero.Mask = "##########"
'Numero.PromptInclude = False
Numero.MaxLength = 10

Fecha.Mask = Fecha_Mask
Fecha_Inicial.Mask = Fecha_Mask
Fecha_Final.Mask = Fecha_Mask

Campos_Enabled False

'SSTab.Tab = 0

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
obra.Enabled = Si

'SSTab.Enabled = Si

Fecha_Inicial.Enabled = Si
Fecha_Final.Enabled = Si

CbTipo.Enabled = Si

CbPintura.Enabled = Si
ListaPernosIncluida.Enabled = Si

If ListaPernosIncluida.Value = 1 Then
    ListaPernosRecibida.Enabled = True
Else
    ListaPernosRecibida.Enabled = False
End If

Obs(0).Enabled = Si
Obs(1).Enabled = Si
Obs(2).Enabled = Si
Obs(3).Enabled = Si
Activa.Enabled = Si

btnCartaSubir.Enabled = Si
btnEspecificacionesSubir.Enabled = Si
btnPlanosSubir.Enabled = Si

btnCartaLeer.Enabled = Si
btnEspTecLeer.Enabled = Si
btnPlanosLeer.Enabled = Si

End Sub
Private Sub Campos_Limpiar()

SSTab.Tab = 0

Numero.Text = ""
Fecha.Text = Fecha_Vacia

obra.Text = ""

Fecha_Inicial.Text = Fecha_Vacia
Fecha_Final.Text = Fecha_Vacia

rut.Text = ""
Razon.Text = ""
Direccion.Text = ""
Comuna.Text = ""

CbTipo.ListIndex = 0

CbPintura.ListIndex = 0

ListaPernosIncluida.Value = False
ListaPernosRecibida.Value = False

Obs(0).Text = ""
Obs(1).Text = ""
Obs(2).Text = ""
Obs(3).Text = ""

'Activa.Value = True

End Sub
Private Sub btnCartaLeer_Click()
If Val(Numero.Text) > 0 Then
    PDF_Leer Numero.Text, "carta"
End If
End Sub
Private Sub btnEspTecLeer_Click()
If Val(Numero.Text) > 0 Then
    PDF_Leer Numero.Text, "esptec"
End If
End Sub
Private Sub btnPlanosLeer_Click()
If Val(Numero.Text) > 0 Then
    PDF_Leer Numero.Text, "planos"
End If
End Sub
Private Sub PDF_Leer(ByVal Nv As Double, ByVal Archivo As String)
' lee archivo pdf
'http://acr3006-dualpro/nv_files/nv_136_carta.pdf ok
    'mvarURL = "http://acr3006-dualpro/nv_files/nv_" & Nv & "_" & Archivo & ".pdf"
    Dim intranet As String
    intranet = ReadIniValue(Path_Local & "scp.ini", "Path", "intranet_server")
    mvarURL = intranet & "nv_files/nv_" & Nv & "_" & Archivo & ".pdf"
    Select Case Archivo
    Case "carta"
'        lblCartaLeer.ForeColor = &H40C0&
    Case "esptec"
'        lblEspecificacionesLeer.ForeColor = &H40C0&
    Case "planos"
'        lblPlanosLeer.ForeColor = &H40C0&
    End Select
    GoURL (mvarURL)
End Sub
Private Sub ListaPernosIncluida_Click()
If ListaPernosIncluida.Value = 0 Then
    ListaPernosRecibida.Value = 0
    ListaPernosRecibida.Enabled = False
Else
    ListaPernosRecibida.Value = 0
    ListaPernosRecibida.Enabled = True
End If
End Sub
Private Sub Numero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then After_Enter
End Sub
Private Sub After_Enter()
If Val(Numero.Text) < 1 Then Beep: Exit Sub
Select Case Accion
Case "Agregando"

''    RsNVc.Seek "=", Numero.Text, m_NvArea
    ' para historico
   sql = "SELECT * FROM [nv cabecera] WHERE numero=" & Numero.Text
'   Set RsNVc = Dbm.OpenRecordset(sql)
''    If RsNVc.NoMatch Then
    mNv = nvLeer(Numero.Text)
    If mNv.Numero = 0 Then
'    If nv2Obra(Numero.Text) = "" Then

'   If RsNVc.RecordCount = 0 Then
    
      '  busca nv en archivo historico
'      RsNvcH.Seek "=", Numero.Text, m_NvArea
      Set RsNvcH = DbH.OpenRecordset(sql)
'      If RsNvcH.NoMatch Then
      If RsNvcH.RecordCount = 0 Then
    
         Campos_Enabled True
         Numero.Enabled = False
         Fecha.SetFocus
         btnGrabar.Enabled = True
         btnSearch.visible = True
         Activa.Value = 1
            
      Else
        
         Doc_Leer Numero.Text
            
         MsgBox Obj & " YA EXISTE EN ARCHIVO HISTÓRICO"
         Campos_Limpiar
         Numero.Enabled = True
         Numero.SetFocus
        
      End If
        
   Else
    
      Doc_Leer Numero.Text
        
      MsgBox Obj & " YA EXISTE"
      Campos_Limpiar
      Numero.Enabled = True
      Numero.SetFocus
        
   End If
    
Case "Modificando"

'   sql = "SELECT * FROM [nv cabecera] WHERE numero=" & Numero.Text
'   Set RsNVc = Dbm.OpenRecordset(sql)
    mNv = nvLeer(Numero.Text)
    If mNv.Numero = 0 Then
'    If nv2Obra(Numero.Text) = "" Then
'   If RsNVc.RecordCount = 0 Then
     
        MsgBox Obj & " NO EXISTE"
        
    Else
    
      If Usuario.ReadOnly Then
          Botones_Enabled 0, 0, 0, 0, 1, 0
      Else
          Botones_Enabled 0, 0, 0, 0, 1, 1
      End If
      Doc_Leer Numero.Text
      Campos_Enabled True
      
      btnCartaSubir.Enabled = Not Usuario.ReadOnly
      btnEspecificacionesSubir.Enabled = Not Usuario.ReadOnly
      btnPlanosSubir.Enabled = Not Usuario.ReadOnly
      
      Numero.Enabled = False
      btnSearch.visible = True
        
      Nv_Combo
        
   End If

Case "Eliminando"

'    sql = "SELECT * FROM [nv cabecera] WHERE numero=" & Numero.Text
'    Set RsNVc = Dbm.OpenRecordset(sql)
    mNv = nvLeer(Numero.Text)
    If mNv.Numero = 0 Then
'    If RsNVc.RecordCount = 0 Then
        MsgBox Obj & " NO EXISTE"
    Else
        Doc_Leer Numero.Text
        If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo) = vbYes Then
            Doc_Eliminar
            obrasLeer
        End If
        Campos_Limpiar
        Campos_Enabled False
        Numero.Enabled = True
        Numero.SetFocus
    End If
   
Case "Imprimiendo"

'    sql = "SELECT * FROM [nv cabecera] WHERE numero=" & Numero.Text
'    Set RsNVc = Dbm.OpenRecordset(sql)
    mNv = nvLeer(Numero.Text)
    If mNv.Numero = 0 Then
'    If RsNVc.RecordCount = 0 Then
        MsgBox Obj & " NO EXISTE"
    Else
    
        Doc_Leer Numero.Text
        Numero.Enabled = False
        
        btnCartaLeer.Enabled = True
        btnEspTecLeer.Enabled = True
        btnPlanosLeer.Enabled = True
        
'        If MsgBox("¿ Imprime ?", vbYesNo) = vbYes Then
'            Doc_Imprimir
'        End If
'        Campos_Limpiar
'        Numero.Enabled = True
'        Numero.SetFocus

    End If
End Select

End Sub
Private Sub Fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then btnSearch.SetFocus
End Sub
Private Sub Fecha_LostFocus()
d = Fecha_Valida(Fecha, Now)
End Sub
Private Sub Fecha_Inicial_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Fecha_Final.SetFocus
End Sub
Private Sub Fecha_Inicial_LostFocus()
d = Fecha_Valida(Fecha_Inicial, Fecha.Text)
End Sub
Private Sub Fecha_Final_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Fecha_Final.SetFocus
End Sub
Private Sub Fecha_Final_LostFocus()
d = Fecha_Valida(Fecha_Final, Fecha_Inicial.Text)
End Sub
Private Sub btnSearch_Click()
Dim cod_search As String
Search.Muestra data_file, "Clientes", "RUT", "Razon Social", "Cliente", "Clientes"
cod_search = Search.Codigo
If Search.Descripcion <> "" Then
    RsCl.Seek "=", cod_search
    If RsCl.NoMatch Then
        MsgBox "CLIENTE NO EXISTE"
        btnSearch.SetFocus
    Else
        rut.Text = cod_search
        Razon.Text = Search.Descripcion
        Direccion.Text = RsCl!Direccion
        Comuna.Text = NoNulo(RsCl!Comuna)
        obra.SetFocus
    End If
End If
End Sub
Private Sub Obra_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Obs(0).SetFocus
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
Select Case Button.Index
Case 1 ' agregar
    Accion = "Agregando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    
    'sql = "SELECT * FROM [nv cabecera] ORDER BY numero"
    'Set RsNVc = Dbm.OpenRecordset(sql)
    'Numero.Text = Documento_Numero_Nuevo(RsNVc, "Numero")
    Numero.Text = nvNueva()
    
    Numero.Enabled = True
    Numero.SetFocus
Case 2 ' modificar
    Accion = "Modificando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Campos_Limpiar
    Campos_Enabled False
    Numero.Enabled = True ''
    Numero.SetFocus ''
'    Obras_Leer
    ComboNV.visible = True
Case 3 ' Eliminar
    Accion = "Eliminando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Campos_Enabled False
'    Numero.Enabled = True
'    Numero.SetFocus
'    Obras_Leer
    ComboNV.visible = True
    
Case 4 ' Imprimir

    Accion = "Imprimiendo"
    
    Botones_Enabled 0, 0, 0, 1, 1, 0
    
    Campos_Enabled False
    
'    btnCartaLeer.Enabled = True
'    btnEspTecLeer.Enabled = True
'    btnPlanosLeer.Enabled = True
    
'    Numero.Enabled = True
'    Numero.SetFocus
'    Obras_Leer

    ComboNV.visible = True
    
'    If Not ComboNV.Enabled Then
'        MsgBox "imprimir"
'    End If
    If Numero.Text <> "" Then
        If MsgBox("¿ Imprime ?", vbYesNo) = vbYes Then
            Doc_Imprimir
        End If
    End If
    
    
Case 5 ' Separador
Case 6 ' DesHacer

    If Numero.Text = "" Then
    
        If Usuario.ReadOnly Then '01/06/98
            Botones_Enabled 0, 0, 0, 1, 0, 0
        Else
            Botones_Enabled 1, 1, 1, 1, 0, 0
        End If
        
        Campos_Limpiar
        Campos_Enabled False
        Accion = ""
        btnSearch.visible = False
        
    Else
    
        If MsgBox("¿Seguro que quiere DesHacer?", vbYesNo) = vbYes Then
        
            If Usuario.ReadOnly Then '01/06/98
                Botones_Enabled 0, 0, 0, 1, 0, 0
            Else
                Botones_Enabled 1, 1, 1, 1, 0, 0
            End If

            Campos_Limpiar
            
            Campos_Enabled False
            Accion = ""
            btnSearch.visible = False
            
        End If
        
    End If
    ComboNV.visible = False
Case 7 ' grabar
    If Doc_Valido Then
        If MsgBox("¿ GRABA " & Obj & " ?", vbYesNo) = vbYes Then
        
            Doc_Grabar
            If Accion = "Agregando" Then obrasLeer
            
            If MsgBox("¿ IMPRIME " & Obj & " ?", vbYesNo) = vbYes Then
                Doc_Imprimir
            End If
            
            If Usuario.ReadOnly Then
                Botones_Enabled 0, 0, 0, 1, 0, 0
            Else
                Botones_Enabled 1, 1, 1, 1, 0, 0
            End If
            Campos_Limpiar
            Campos_Enabled False
            Accion = ""
            btnSearch.visible = False

'            btnDesHacer.Value = tbrPressed
            
'            Numero.Enabled = True
'            Numero.SetFocus
            
        End If
    End If
Case 8 'separador
Case 9
    MousePointer = vbHourglass
    Load Clientes
    MousePointer = vbDefault
    Clientes.Show 1
End Select

If Accion = "" Then
    Me.Caption = StrConv(Objs, vbProperCase)
Else
    Me.Caption = StrConv(Objs, vbProperCase) & " [" & Accion & "]"
End If

End Sub
Private Sub obrasLeer()
i = 0

nvListar False ' lee todas

ComboNV.Clear

For i = 1 To nvTotal
    ComboNV.AddItem Format(aNv(i).Numero, "0000") & " - " & aNv(i).obra
Next

End Sub
Private Sub ComboNV_Click()
ComboNV.visible = False
Numero.Text = Val(Left(ComboNV.Text, 6))
After_Enter
End Sub
Private Sub Doc_Leer(Nv As Double)

mNv = nvLeer(Nv)

Fecha.Text = Format(mNv.Fecha, Fecha_Format)
rut.Text = mNv.rutCliente
obra.Text = mNv.obra
Fecha_Inicial.Text = Format(mNv.fechaInicio, Fecha_Format)
Fecha_Final.Text = Format(mNv.fechaTermino, Fecha_Format)

If mNv.Tipo = "SV" Then
    CbTipo.ListIndex = 1
Else
    CbTipo.ListIndex = 0
End If

If mNv.pintura Then
    CbPintura.Text = "Pintura"
Else
    If mNv.galvanizado Then
        CbPintura.Text = "Galvanizado"
    Else
        CbPintura.Text = "Negro"
    End If
End If

ListaPernosIncluida.Value = IIf(mNv.ListaPernosIncluida, 1, 0)
ListaPernosRecibida.Value = IIf(mNv.ListaPernosRecibida, 1, 0)

Obs(0).Text = NoNulo(mNv.observacion1)
Obs(1).Text = NoNulo(mNv.observacion2)
Obs(2).Text = NoNulo(mNv.observacion3)
Obs(3).Text = NoNulo(mNv.observacion4)
Activa.Value = IIf(mNv.Activa, 1, 0)

Cliente_Lee rut.Text

' busca si hay archivos pdf
Archivos_Subidos_Buscar

End Sub
Private Sub Cliente_Lee(rut)
RsCl.Seek "=", rut
If Not RsCl.NoMatch Then
    Razon.Text = RsCl![Razon Social]
    Direccion.Text = RsCl!Direccion
    Comuna.Text = NoNulo(RsCl!Comuna)
End If
End Sub
Private Function Doc_Valido() As Boolean
Doc_Valido = False
If Trim(rut.Text) = "" Then
    Beep
    MsgBox "DEBE ELEGIR CLIENTE"
    btnSearch.SetFocus
    Exit Function
End If
If IsObjBlanco(obra, "OBRA", btnGrabar) Then Exit Function

If Fecha_Inicial.Text = "__/__/__" Then
    MsgBox "DEBE DIGITAR FECHA INICIO"
    Fecha_Inicial.SetFocus
    Exit Function
End If

If Fecha_Final.Text = "__/__/__" Then
    MsgBox "DEBE DIGITAR FECHA TÉRMINO"
    Fecha_Final.SetFocus
    Exit Function
End If

Doc_Valido = True
End Function
Private Sub Doc_Grabar()

mNv.Numero = Numero.Text
'!NvArea = m_NvArea
mNv.Fecha = Format(Fecha.Text, "yyyy-mm-dd")
mNv.rutCliente = rut.Text
mNv.obra = Left(Trim(obra.Text), 30)
mNv.fechaInicio = Fecha_Inicial.Text
mNv.fechaTermino = Fecha_Final.Text

If CbTipo.Text = "Servicio" Then
    mNv.Tipo = "SV"  ' servicio
Else
    mNv.Tipo = "PR"  ' produccion
End If

If CbPintura.Text = "Pintura" Then
    mNv.pintura = True
    mNv.galvanizado = False
Else
    If CbPintura.Text = "Galvanizado" Then
        mNv.pintura = False
        mNv.galvanizado = True
    Else
        mNv.pintura = False
        mNv.galvanizado = False
    End If
End If

mNv.ListaPernosIncluida = ListaPernosIncluida.Value
mNv.ListaPernosRecibida = ListaPernosRecibida.Value
mNv.observacion1 = Obs(0).Text
mNv.observacion2 = Obs(1).Text
mNv.observacion3 = Obs(2).Text
mNv.observacion4 = Obs(3).Text
mNv.Activa = Activa.Value

' OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO
If True Then
    If Accion = "Agregando" Then
        nvGrabar mNv, True
    Else
        nvGrabar mNv, False
    End If
End If
' OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO OJO

' ojo con eiffel
'NvSqlGrabar CnxSqlServerScc
'NvSqlGrabar CnxSqlServerScpNew

Botones_Enabled 0, 0, 0, 0, 1, 0

obrasLeer

If Accion = "Agregando" Then
    Track_Registrar "NV", Numero.Text, "AGR"
Else
    Track_Registrar "NV", Numero.Text, "MOD"
End If

End Sub
Private Sub Doc_Eliminar()
' borra cabecera
RsNVp.Seek ">=", Numero.Text, ""
If Not RsNVp.NoMatch Then
    If RsNVp!Nv = Numero.Text Then
        Do While Not RsNVp.EOF
            If RsNVp!Nv = Numero.Text Then
                MsgBox "NV tiene planos," & vbCr & "no se puede eliminar"
            End If
            Exit Sub
        Loop
    End If
End If

' borra cabecera
nvEliminar Numero.Text

Track_Registrar "NV", Numero.Text, "ELI"

End Sub
Private Sub Doc_Imprimir()
MousePointer = vbHourglass
' imprime nv
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer, tab4 As Integer
Dim tab5 As Integer, tab6 As Integer, tab7 As Integer, tab8 As Integer, tab9 As Integer
tab0 = 7 'margen izquierdo
tab1 = tab0 + 0
tab2 = tab0 + 10
tab3 = tab0 + 20
tab4 = tab0 + 30
tab5 = tab0 + 43
tab6 = tab0 + 54
tab7 = tab0 + 65
tab8 = tab0 + 83
tab9 = tab0 + 93

Dim can_valor As String, can_col As Integer

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt

prt.Font.Size = 15
prt.Print Tab(tab0 + 10); "NOTA DE VENTA Nº" & Numero.Text
prt.Font.Size = 10
prt.Print ""
prt.Print ""

' cabecera
prt.Print Tab(tab0); Empresa.Razon
prt.Print Tab(tab0); "GIRO: " & Empresa.Giro
prt.Print Tab(tab0); Empresa.Direccion
prt.Print Tab(tab0); "Teléfono: " & Empresa.Telefono1 & " - " & Empresa.Comuna
prt.Print ""
prt.Print Tab(tab0); "FECHA     : " & Fecha.Text

prt.Print Tab(tab0); "SEÑOR(ES) : " & Razon,
prt.Print Tab(tab0 + 50); "RUT    : " & rut
prt.Print Tab(tab0); "DIRECCIÓN : " & Direccion,
prt.Print Tab(tab0 + 50); "COMUNA : " & Comuna
prt.Print ""
prt.Font.Size = 15
prt.Print Tab(tab0); "OBRA : " & Numero.Text & " " & obra.Text
prt.Font.Size = 10
prt.Print ""
prt.Print Tab(tab0); "PINTURA : " & CbPintura.Text
'prt.Print Tab(tab0); "PINTURA : " & IIf(Pintura_Op(0), "SI", "NO")
'prt.Print ""
'prt.Print Tab(tab0); "GALVANIZADO : " & IIf(Galva_Op(0), "SI", "NO")
prt.Print ""
prt.Print Tab(tab0); "FECHA DE INICIO OBRAS  : " & Fecha_Inicial.Text
prt.Print Tab(tab0); "FECHA DE TÉRMINO OBRAS : " & Fecha_Final.Text
prt.Print ""
prt.Print Tab(tab0); "OBSERVACIONES :";
prt.Print Tab(tab0 + 16); Obs(0).Text
prt.Print Tab(tab0 + 16); Obs(1).Text
prt.Print Tab(tab0 + 16); Obs(2).Text
prt.Print Tab(tab0 + 16); Obs(3).Text

For i = 1 To 28
    prt.Print ""
Next

prt.Print Tab(tab0); Tab(14), "__________________", Tab(56), "__________________"
prt.Print Tab(tab0); Tab(14), "       VºBº       ", Tab(56), "       VºBº       "

prt.EndDoc

Impresora_Predeterminada "default"

MousePointer = vbDefault
End Sub
Private Sub Nv_Combo()

m_Nv = Val(Numero.Text)

If m_Nv = 0 Then Exit Sub
' busca nv en combo
i = 1
'Do Until aNv(i).Numero <> 0
Do While aNv(i).Numero <> 0
    If Val(aNv(i).Numero) = m_Nv Then
        ComboNV.ListIndex = i - 1
        Exit Sub
    End If
    i = i + 1
Loop

'MsgBox "NV no existe"
'Nv.SetFocus

End Sub
Private Sub btnCartaSubir_Click()
ArchivoSubir Val(Numero.Text), "carta"
End Sub
Private Sub btnEspecificacionesSubir_Click()
ArchivoSubir Val(Numero.Text), "esptec"
End Sub
Private Sub btnPlanosSubir_Click()
ArchivoSubir Val(Numero.Text), "planos"
End Sub
Private Sub ArchivoSubir(ByVal Nv As Double, ByVal Tipo As String)
' sube archivo pdf a la carpeta \\acr3006-dualpro\scp\mdb\nv_files
' parametros:
' Nv: numero delanota de venta
' Tipo: tipo o mejor dicho nombre del documento
'      carta: carta cotizacion (la que le envia estructuras al cliente)
'      esptec: especificaciones tecnicas pedidas por el cliente
'      planos: lista de planos
Dim p As Integer
Dim m_PathyArchivoOrigen As String
Dim m_ArchivoOrigenNombre As String
Dim m_NombreArchivoDestino As String
Dim extension As String

' adjunta archivo

cd.DialogTitle = "Buscar Archivo PDF"
'cd.Filter = "Documento PDF (*.pdf)|*.pdf|Todos los Archivos (*.*)|*.*"
cd.Filter = "Documento PDF (*.pdf)|*.pdf"

cd.ShowOpen

m_PathyArchivoOrigen = cd.filename

If m_PathyArchivoOrigen = "" Then
'    MsgBox "Debe escoger Archivo"
    Exit Sub
End If

'MsgBox Cd.filename ' viene con ruta clompleta
m_PathyArchivoOrigen = cd.filename

' separa path y archivo
p = InStrLast(m_PathyArchivoOrigen, ".")
extension = Mid(m_PathyArchivoOrigen, p)

p = InStrLast(m_PathyArchivoOrigen, "\")

If p > 0 Then

'    m_PathArchivo = Left(m_PathArchivo, p)
    m_ArchivoOrigenNombre = Mid(m_PathyArchivoOrigen, p + 1)
      
'   formato archivo de destino:  NV_nnnn_carta.pdf, o nv_nnnn_esptec.pdf o nv_nnnn_planos.pdf
    m_NombreArchivoDestino = "nv_" & Nv & "_" & Tipo & extension
    If Archivo_Existe(m_PathDestino, m_NombreArchivoDestino) Then
    Else
    
        FileCopy m_PathyArchivoOrigen, m_PathDestino & m_NombreArchivoDestino
    
    End If
        
End If

End Sub
Private Sub Archivos_Subidos_Buscar()
' busca nombres de archivos subidos
Dim ArchivoNombre As String
'lblAdjuntos.Caption = ""

ArchivoNombre = Dir(m_PathDestino & "nv_" & Numero.Text & "_carta.pdf", vbArchive)
If ArchivoNombre = "" Then
    btnCartaLeer.Enabled = False
Else
    btnCartaLeer.Enabled = True
End If

ArchivoNombre = Dir(m_PathDestino & "nv_" & Numero.Text & "_esptec.pdf", vbArchive)
If ArchivoNombre = "" Then
    btnEspTecLeer.Enabled = False
Else
    btnEspTecLeer.Enabled = True
End If

ArchivoNombre = Dir(m_PathDestino & "nv_" & Numero.Text & "_planos.pdf", vbArchive)
If ArchivoNombre = "" Then
    btnPlanosLeer.Enabled = False
Else
    btnPlanosLeer.Enabled = True
End If

'D:\acr3006-dualpro\scp\mdb\nv_files

'If ArchivoNombre <> "" Then
'    Do
'        ArchivoNombre = Dir()
'    Loop Until ArchivoNombre = ""
'End If

End Sub
Private Sub NvSqlGrabar(Cnx)
' graba esta nv en Sql Server

Exit Sub ' no grava en SqlServer

Dim recub As String, lpi As String, lpr As String, mActiva As String
recub = "N" ' negro
If mNv.galvanizado Then recub = "G"
If mNv.pintura Then recub = "P"

lpi = IIf(ListaPernosIncluida.Value, "S", "N")
lpr = IIf(ListaPernosRecibida.Value, "S", "N")
mActiva = IIf(Activa.Value, "S", "N")

SP_nvUpdate Numero.Text, fecha2aaaammdd(Fecha.Text), rut.Text, "SV", obra.Text, recub, fecha2aaaammdd(Fecha_Inicial.Text), fecha2aaaammdd(Fecha_Final.Text), lpi, lpr, mActiva, Obs(0), Obs(1), Obs(2), Obs(3)

Exit Sub

Dim Tabla As String

' primero verifica si existe cliente
Tabla = "clientes" ' scpold
Tabla = "tb_clientes" ' scpnew
' verifica si existe
sql = "SELECT * FROM " & Tabla
sql = sql & " WHERE rut='" & Trim(rut.Text) & "'"

With RsNvSql

    .Open sql, Cnx
    
    If .EOF Then
        
        ' nuevo cliente en sqlServer
    
        sql = "INSERT INTO [clientes] ("
        sql = sql & "[rut]" ' 1
        sql = sql & ",[razonsocial]" ' 2
        sql = sql & ",[direccion]" ' 3
        sql = sql & ",[comuna]" ' 4
        sql = sql & ") Values ("
        sql = sql & "'" & Trim(rut.Text) & "'"  ' 1
        sql = sql & ",'" & Trim(Razon.Text) & "'"   ' 2
        sql = sql & ",'" & Trim(Direccion.Text) & "'" ' 3
        sql = sql & ",'" & Trim(Comuna.Text) & "'"   ' 4
        sql = sql & ")"
        Cnx.Execute sql
        
    End If
    
    .Close


'//////////////////////////////////////////////////////////

Tabla = "tb_nv" ' sqlserver.delgado1303

' verifica si existe
sql = "SELECT * FROM " & Tabla
sql = sql & " WHERE nv=" & Numero.Text ' sql

.Open sql, Cnx

If .EOF Then
    
    ' nueva en sqlServer

    sql = "INSERT INTO [" & Tabla & "] ("
    sql = sql & "[nv]," ' 1
    sql = sql & "[area]," ' 2"
    sql = sql & "[fecha]," ' 3
    sql = sql & "[rut_cliente]," ' 4
    sql = sql & "[tipo]," ' 5
    sql = sql & "[obra]," ' 6
    sql = sql & "[galvanizado]," ' 7
    sql = sql & "[pintura]," ' 8
    sql = sql & "[fecha_inicio]," ' 9
    sql = sql & "[fecha_termino]," ' 10
    sql = sql & "[lista_pernos_incluida]," ' 11
    sql = sql & "[lista_pernos_recibida]," ' 12
    sql = sql & "[activa]," ' 13
    sql = sql & "[observacion1]," ' 14
    sql = sql & "[observacion2]," ' 15
    sql = sql & "[observacion3]," ' 16
    sql = sql & "[observacion4]" ' 17
    sql = sql & ") Values ("
    sql = sql & Numero.Text & "," ' 1
    sql = sql & "0," ' 2
    sql = sql & "'" & fecha2aaaammdd(Fecha.Text) & "'," ' 3
    sql = sql & "'" & rut.Text & "'," ' 4
    sql = sql & "'" & mNv.Tipo & " '," ' 5
    sql = sql & "'" & obra.Text & "'," ' 6
    sql = sql & "'" & IIf(mNv.galvanizado, "S", "N") & "'," ' 7
    sql = sql & "'" & IIf(mNv.pintura, "S", "N") & "'," ' 8
    sql = sql & "'" & fecha2aaaammdd(Fecha_Inicial.Text) & "'," ' 9
    sql = sql & "'" & fecha2aaaammdd(Fecha_Final.Text) & "'," ' 10
    sql = sql & "'" & IIf(ListaPernosIncluida.Value, "S", "N") & "'," ' 11
    sql = sql & "'" & IIf(ListaPernosRecibida.Value, "S", "N") & "'," ' 12
    sql = sql & "'" & IIf(Activa.Value, "S", "N") & "'," ' 13
    sql = sql & "'" & Obs(0) & "'," ' 14
    sql = sql & "'" & Obs(1) & "'," ' 15
    sql = sql & "'" & Obs(2) & "'," ' 16
    sql = sql & "'" & Obs(3) & "'" ' 17
    sql = sql & ")"
           
Else
    
    ' ya existe en sqlServer
    
    sql = "UPDATE " & Tabla & " SET "
    
    sql = sql & "[fecha]='" & fecha2aaaammdd(Fecha.Text) & "',"  ' formato aaaammdd
    sql = sql & "[rut_cliente]='" & rut.Text & "',"
    sql = sql & "[tipo]='" & mNv.Tipo & "',"
    sql = sql & "[obra]='" & obra.Text & "',"
    sql = sql & "[galvanizado]='" & IIf(mNv.galvanizado, "S", "N") & "',"
    sql = sql & "[pintura]='" & IIf(mNv.pintura, "S", "N") & "',"
    sql = sql & "[fecha_inicio]='" & fecha2aaaammdd(Fecha_Inicial.Text) & "',"
    sql = sql & "[fecha_termino]='" & fecha2aaaammdd(Fecha_Final.Text) & "',"
    sql = sql & "[lista_pernos_incluida]='" & IIf(ListaPernosIncluida.Value, "S", "N") & "',"
    sql = sql & "[lista_pernos_recibida]='" & IIf(ListaPernosRecibida.Value, "S", "N") & "',"
    sql = sql & "[activa] = '" & IIf(Activa.Value, "S", "N") & "',"
    sql = sql & "[observacion1]='" & Obs(0) & "',"
    sql = sql & "[observacion2]='" & Obs(1) & "',"
    sql = sql & "[observacion3]='" & Obs(2) & "',"
    sql = sql & "[observacion4]='" & Obs(3) & "'"
      
    sql = sql & " WHERE nv=" & Numero.Text
    
End If

Cnx.Execute sql

.Close

End With

End Sub
