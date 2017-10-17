VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form bultos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bultos"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport Cr 
      Left            =   8520
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11220
      _ExtentX        =   19791
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
            Object.ToolTipText     =   "Ingresar Nuevo Bulto"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar Bulto"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar Bulto"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Bulto"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Autorizar Bulto"
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
            Object.ToolTipText     =   "Grabar Bulto"
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
   Begin MSCommLib.MSComm MSComm 
      Left            =   9360
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton btnCapturadorLeer 
      Caption         =   "Capturar Datos"
      Height          =   375
      Left            =   1920
      TabIndex        =   19
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Numero 
      Height          =   300
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Nv 
      Height          =   300
      Left            =   840
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.ComboBox ComboOT 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox ComboNV 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
   Begin VB.ComboBox ComboMarca 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   8040
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox ComboPlano 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame 
      Caption         =   "CLIENTE"
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   3240
      TabIndex        =   8
      Top             =   480
      Width           =   3735
      Begin VB.CommandButton btnSearch 
         Height          =   300
         Left            =   1920
         Picture         =   "bultos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   300
      End
      Begin VB.TextBox Razon 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   3015
      End
      Begin MSMask.MaskEdBox Rut 
         Height          =   300
         Left            =   600
         TabIndex        =   10
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label lbl 
         Caption         =   "RUT"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbl 
         Caption         =   "SEÑOR(ES)"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.TextBox txtEditGD 
      Height          =   285
      Left            =   8040
      TabIndex        =   18
      Text            =   "txtEditGD"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   840
      TabIndex        =   4
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   327680
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid Detalle 
      Height          =   4005
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7064
      _Version        =   327680
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
            Picture         =   "bultos.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bultos.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bultos.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bultos.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bultos.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bultos.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bultos.frx":0836
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bultos.frx":0948
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "&OBRA"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "&FECHA"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
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
Attribute VB_Name = "bultos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnAnular As Button, btnDesHacer As Button, btnGrabar As Button
Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, k As Integer, d As Variant

Private DbD As Database, RsCl As Recordset, RsTb As Recordset
Private DbM As Database, RsNVc As Recordset
Private RsNvPla As Recordset, RsPd As Recordset
'Private rsbuldet As Recordset, rsbuldet As Recordset, RsGDdE As Recordset
Private RsBulDet As Recordset

'Private m_Tipo As String

Private n_filas As Integer, n_columnas_N As Integer, n_columnas_E As Integer
Private RutClientes(1999) As String, Rev(2999) As String
Private prt As Printer
Private n1 As Double, n4 As Double, n6 As Double, n7 As Double, n8 As Double, n10 As Double
Private a_Nv(2999, 2) As String, m_Nv As Double
Private Depende_ITOPyG As Boolean, Advertencia_ITOPyG As Boolean, Advertencia_ITOPyG_msg As String

Private Ret As Integer, Buffer As Variant
Private Capturando As Boolean, Reg As String, Arr_DatoCapturado(99, 1) As String, contador_registros As Integer
'Private Capturador As Boolean
Private AjusteX As Double, AjusteY As Double
' Arr_DatoCapturado(i,j)
' i=conjtador
' j: 0:dato , 1:cantidad de piezas(digitadas en el scanpal2)
Private m_NvArea As Integer

' procedimientos nuevos
' 25/07/06
' una guia se puede modificar y eliminar n veces antes de imprimir
' una vez impresa no se puede modificar, para hacerlo debe autorizar gerente (erwin)
Private Sub Form_Load()

AjusteX = 0
AjusteY = 0

'Capturador = True ' False

Capturando = False  ' indica acaba de ser presionado boton capturar
btnCapturadorLeer.visible = True 'Capturador

'Depende_ITOPyG = False
Depende_ITOPyG = True ' siempre a partir de 29/11/05

Advertencia_ITOPyG = False
Advertencia_ITOPyG_msg = ""

' abre archivos
Set DbD = OpenDatabase(data_file)
Set RsCl = DbD.OpenRecordset("Clientes")
RsCl.Index = "RUT"

Set DbM = OpenDatabase(mpro_file)
Set RsNVc = DbM.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

' para calculo de cantidad de piezas fabricadas x mes
'Dim qry As String
'qry = "SELECT "
'qry = qry & "SUM([Cantidad Total]) AS can"
'qry = qry & " FROM [Planos Detalle]"
'Set rsbuldet = Dbm.OpenRecordset(qry)
'Debug.Print rsbuldet!can

'Set rsbuldet = Dbm.OpenRecordset("GD Cabecera")
'rsbuldet.Index = "Número"

Set RsBulDet = DbM.OpenRecordset("bultos")
RsBulDet.Index = "Numero-Linea"

'Set RsGDdE = Dbm.OpenRecordset("GD Especial Detalle")
'RsGDdE.Index = "Número-Línea"

'CbTipo.AddItem "Normal"
'CbTipo.AddItem "Especial"
'CbTipo.AddItem "Pintura"

' Combo obra
ComboNV.AddItem " "
i = 0
RutClientes(i) = " "

Do While Not RsNVc.EOF
    i = i + 1
    a_Nv(i, 0) = RsNVc!Numero
    a_Nv(i, 1) = RsNVc!obra
    a_Nv(i, 2) = RsNVc!galvanizado Or RsNVc!pintura
    ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
'    i = i + 1
    RutClientes(i) = RsNVc![RUT CLiente]
    RsNVc.MoveNext
Loop

Set RsNvPla = DbM.OpenRecordset("Planos Cabecera")
RsNvPla.Index = "NV-Plano"

Set RsPd = DbM.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Item"

Inicializa
Detalle_Config

Privilegios
'm_Tipo = "N"
'Obs(2).Visible = False
m_NvArea = 0

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

End Sub
Private Sub ComboNV_Click()
If ComboNV.Text = " " Then
    m_Nv = 0
    Exit Sub
End If

MousePointer = vbHourglass

Depende_ITOPyG = True

i = 0
m_Nv = Val(Left(ComboNV.Text, 6))
Nv.Text = m_Nv

'Debug.Print ComboNV.ListIndex, ComboNV.Text, a_Nv(ComboNV.ListIndex, 2)
Depende_ITOPyG = CBool(a_Nv(ComboNV.ListIndex, 2))

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

Detalle_Limpiar Detalle, n_columnas_N
'no debe limpiar detalle de guia especial
'Detalle_Limpiar Detalle_Especial, n_columnas_E

ComboMarca.Clear

' datos del cliente
Rut.Text = RutClientes(ComboNV.ListIndex)
RsCl.Seek "=", Rut.Text
If Not RsCl.NoMatch Then
    Razon.Text = RsCl![Razon Social]
'    Direccion.Text = RsCl!Dirección
'    Comuna.Text = RsCl!Comuna
End If

'If Rut.Text <> "" Then Detalle.Enabled = True
Detalle.Enabled = True

MousePointer = vbDefault

End Sub
Private Sub btnCapturadorLeer_Click()
Dim commset As String, commport As Integer
'/////////////////////////////////////////////
' lee datos desde el capturador de datos

' Establecer y abrir el puerto
If MSComm.PortOpen Then
    MSComm.PortOpen = False
End If

' settings para comm port
commset = ReadIniValue(Path_Local & "scp.ini", "ScanPal2", "commsettings")
commport = ReadIniValue(Path_Local & "scp.ini", "ScanPal2", "commport")
MSComm.Settings = commset ' "115200,n,8,1"
MSComm.Handshaking = comNone
MSComm.InputMode = comInputModeBinary
MSComm.RThreshold = 1
MSComm.commport = commport

On Error GoTo NoPuerta
MSComm.PortOpen = True
On Error GoTo 0

Capturando = True

' inicia protocolo con ScanPal2
Buffer = "READ" & vbCr
MSComm.Output = Buffer

contador_registros = 0

Exit Sub

NoPuerta:
MsgBox "Error En Puerto COM" & commport & vbCr & commset

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
    For i = 2 To n_columnas_N
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
Dim c_otf As Integer, c_itof As Integer, c_desp As Integer
Dim c_itopg As Integer

fil = Detalle.Row
ComboMarca.visible = False
m_Plano = Detalle.TextMatrix(fil, 1)
m_Marca = ComboMarca.Text

'If Not Capturador Then
If Not Capturando Then
    '///
    ' verifica si Plano-Marca ya están en esta GD
    For i = 1 To n_filas
        If m_Plano = Detalle.TextMatrix(i, 1) And m_Marca = Detalle.TextMatrix(i, 3) Then
            Beep
            MsgBox "MARCA YA EXISTE EN GD"
            Detalle.Row = i
            Detalle.col = 3
            Detalle.SetFocus
            Exit Sub
        End If
    Next
    '///
End If
'End If

Detalle = m_Marca

c_itof = c_desp = 0
' busca marca en plano
RsPd.Seek ">=", m_Nv, m_NvArea, m_Plano, 0
If Not RsPd.NoMatch Then
    Do While Not RsPd.EOF
        If RsPd!Marca = m_Marca Then
        
            c_otf = RsPd![OT fab]
            c_itof = RsPd![ITO fab]
            c_itopg = RsPd![ITO pyg]
            c_desp = RsPd![GD]
            
            ' verifica si está asignada
            If c_otf = 0 Then
                Beep
                MsgBox "La marca """ & m_Marca & """" & vbCr _
                    & "No está Asignada"
                Detalle.TextMatrix(fil, 3) = ""
                Detalle.SetFocus
                Exit Sub
            End If
            
            If c_itof = 0 Then
                Beep
                MsgBox "La marca """ & m_Marca & """" & vbCr _
                    & "No está Recibida de Fabricación"
                Detalle.TextMatrix(fil, 3) = ""
                Detalle.SetFocus
                Exit Sub
            End If
            
            If Depende_ITOPyG Then
                If c_itopg = 0 Then
                    Beep
                    MsgBox "La marca """ & m_Marca & """" & vbCr _
                        & "No está Recibida de Pintura o Galvanizado"
                        
                    If Advertencia_ITOPyG Then
                    Else
                        Detalle.TextMatrix(fil, 3) = ""
                        Exit Sub
                    End If
                    Detalle.SetFocus
                End If
            End If
            
            ' verifica que quede algo por despachar
            If Depende_ITOPyG Then
                'If c_itopg - c_desp <= 0 Then
                '    Beep
                '    MsgBox "La marca """ & m_Marca & """" & vbCr _
                '        & "Ya se despachó"
                '    Detalle.TextMatrix(fil, 3) = ""
                '    Detalle.SetFocus
                '    If Not Advertencia_ITOPyG Then
                '        Exit Sub
                '    End If
                'End If
            Else
                If c_itof - c_desp <= 0 Then
                    Beep
                    MsgBox "La marca """ & m_Marca & """" & vbCr _
                        & "Ya se despachó"
                    Detalle.TextMatrix(fil, 3) = ""
                    Detalle.SetFocus
                    Exit Sub
                End If
            End If
            
            Detalle.TextMatrix(fil, 4) = RsPd!Descripcion
            If Depende_ITOPyG Then
                Detalle.TextMatrix(fil, 5) = c_itopg
            Else
                Detalle.TextMatrix(fil, 5) = c_itof
            End If
            Detalle.TextMatrix(fil, 6) = c_desp '- m_cantGD ?
            Detalle.TextMatrix(fil, 8) = Replace(RsPd![Peso], ",", ".")
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

Obj = "BULTO"
Objs = "BULTOS"

Accion = ""

btnSearch.visible = False
btnSearch.ToolTipText = "Busca Clientes"
Campos_Enabled False

n_filas = 20
n_columnas_N = 9 ' 11
n_columnas_E = 7

Set DbD = OpenDatabase(data_file)
'Set RsTb = DbD.OpenRecordset("Tablas")
'RsTb.Index = "Tipo-Descripcion"

'Do While Not RsTb.EOF
'    If RsTb!Tipo = "CHOFER" Then
'        CbChofer.AddItem RsTb!Descripcion
'    End If
'    If RsTb!Tipo = "PATENTE" Then
'        CbPatente.AddItem RsTb!Descripcion
'    End If
'    RsTb.MoveNext
'Loop

'RsTb.Close

'Obs(0).MaxLength = 50
'Obs(1).MaxLength = 50
'Obs(2).MaxLength = 50

End Sub
Private Sub Detalle_Config()

Dim i As Integer, ancho As Integer

Detalle.Left = 100
Detalle.WordWrap = True
Detalle.RowHeight(0) = 450
Detalle.Rows = n_filas + 1
Detalle.Cols = n_columnas_N + 1

Detalle.TextMatrix(0, 0) = ""
Detalle.TextMatrix(0, 1) = "Plano"
Detalle.TextMatrix(0, 2) = "Rev"                   '*
Detalle.TextMatrix(0, 3) = "Marca"
Detalle.TextMatrix(0, 4) = "Descripción"           '*
Detalle.TextMatrix(0, 5) = "Cant Reci."            '*
Detalle.TextMatrix(0, 6) = "Cant Desp."            '*
Detalle.TextMatrix(0, 7) = "a Desp."
Detalle.TextMatrix(0, 8) = "Peso Unitario"         '*
Detalle.TextMatrix(0, 9) = "Peso TOTAL"            '*
'Detalle.TextMatrix(0, 10) = "Precio Unitario"
'Detalle.TextMatrix(0, 11) = "Precio TOTAL"         '*

Detalle.ColWidth(0) = 300
Detalle.ColWidth(1) = 2100 '1200 ' plano
Detalle.ColWidth(2) = 400
Detalle.ColWidth(3) = 1700 ' marca
Detalle.ColWidth(4) = 1200
Detalle.ColWidth(5) = 500
Detalle.ColWidth(6) = 500
Detalle.ColWidth(7) = 500
Detalle.ColWidth(8) = 800
Detalle.ColWidth(9) = 800
'Detalle.ColWidth(10) = 700
'Detalle.ColWidth(11) = 800

ancho = 350 ' con scroll vertical
'ancho = 100 ' sin scroll vertical

'TotalPeso.Width = Detalle.ColWidth(11)
For i = 0 To n_columnas_N
'    If i = 9 Then TotalPeso.Left = ancho + Detalle.Left - 350
'    If i = 11 Then TotalPrecio.Left = ancho + Detalle.Left - 350
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
'    Detalle.col = 10
'    Detalle.CellForeColor = vbRed
'    Detalle.col = 11
'    Detalle.CellForeColor = vbRed
    
Next

txtEditGD.Text = ""

'Top = Form_CentraY(Me)
'Left = Form_CentraX(Me)

Detalle.Enabled = False

End Sub
Private Sub CbTipo_Click()

'm_Tipo = Left(CbTipo.Text, 1)

ComboPlano.visible = False
ComboMarca.visible = False

'Select Case m_Tipo
Select Case "N"

Case "N"

    ' NORMAL
    btnCapturadorLeer.Enabled = True
    Detalle.visible = True
'    Detalle_Especial.Visible = False
'    m_Tipo = "N"
'    lblContenido.Caption = "Contenido"
'    Obs(2).Visible = False
    
Case "E"

    ' ESPECIAL
    btnCapturadorLeer.Enabled = False
    Detalle.visible = False
'    Detalle_Especial.Visible = True
'    m_Tipo = "E"
    
'    Detalle_Especial.TextMatrix(0, 4) = "Peso Unitario"
'    Detalle_Especial.TextMatrix(0, 5) = "Peso TOTAL"
    
'    lblContenido.Caption = "Contenido"
'    Obs(2).Visible = False
    
Case "P"
    
    ' PINTURA
    btnCapturadorLeer.Enabled = False
    Detalle.visible = False
'    Detalle_Especial.Visible = True
'    m_Tipo = "P"
    
'    Detalle_Especial.TextMatrix(0, 4) = "m2 Unitario"
'    Detalle_Especial.TextMatrix(0, 5) = "m2 TOTAL"
    
'    lblContenido.Caption = "Esquema"
'    Obs(2).Visible = True

End Select

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

    RsBulDet.Seek "=", Numero.Text, 1
    If RsBulDet.NoMatch Then
    
        Campos_Enabled True
        Numero.Enabled = False
        Detalle.Enabled = False
        
        Fecha.Text = Format(Now, Fecha_Format)
        
'        If Usuario.AccesoTotal Then
'            Fecha.SetFocus
'        Else
'            ComboNV.SetFocus
'        End If
'        CbTipo.Text = "Normal"
'        CbTipo.SetFocus
'        GDespecial.SetFocus
        
        btnGrabar.Enabled = True
        btnSearch.visible = True
        
    Else
    
        Doc_Leer
        MsgBox Obj & " YA EXISTE"
        Detalle.Enabled = False
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
        
    End If
    
Case "Modificando"

    RsBulDet.Seek "=", Numero.Text, 1
    If RsBulDet.NoMatch Then
    
        MsgBox Obj & " NO EXISTE"
        
    Else
    
        Doc_Leer
        
'        If RsBulDet!impresa Then
        If False Then
        
            MsgBox Obj & " ya esta impresa," & vbLf & "NO se puede Modificar"
            Detalle.Enabled = False
            Campos_Limpiar
            Numero.Enabled = True
            Numero.SetFocus
        
        Else
        
'        If rsbuldet!Tipo = "N" Then
            Campos_Enabled True
            Numero.Enabled = False
            
'            If Usuario.AccesoTotal Then
'                Fecha.SetFocus
'            Else
'                ComboNV.SetFocus
'            End If
'            CbTipo.SetFocus
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

    RsBulDet.Seek "=", Numero.Text, 1
    
    If RsBulDet.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
    
        Doc_Leer
        
'        If RsBulDet!impresa Then
            If False Then
        
            MsgBox Obj & " ya esta impresa," & vbLf & "NO se puede Eliminar"
            Detalle.Enabled = False
            Campos_Limpiar
            Numero.Enabled = True
            Numero.SetFocus
        
        Else
        
'        If rsbuldet!Tipo = "N" Then
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

    RsBulDet.Seek "=", Numero.Text, 1
    If RsBulDet.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Doc_Leer
'        If RsBulDet!Tipo = "N" Then
        If True Then
            Numero.Enabled = False
            
            Detalle.visible = True
            Detalle.Enabled = True
'            Detalle_Especial.Visible = False
            
        Else
        
            Numero.Enabled = False
            
'            Detalle_Especial.Visible = True
'            Detalle_Especial.Enabled = True
            Detalle.visible = False
        
        End If
    End If
    
Case "Autorizando"

    RsBulDet.Seek "=", Numero.Text
    
    If RsBulDet.NoMatch Then
    
        MsgBox Obj & " NO EXISTE"
        
    Else
    
        Doc_Leer
        
        If RsBulDet!Tipo = "N" Then
            Numero.Enabled = False
            
            Detalle.visible = True
            Detalle.Enabled = True
'            Detalle_Especial.Visible = False
            
        Else
        
            Numero.Enabled = False
            
'            Detalle_Especial.Visible = True
'            Detalle_Especial.Enabled = True
            Detalle.visible = False
        
        End If
        
        If RsBulDet!impresa Then
            ' ok se puede autorizar
        Else
            MsgBox Obj & " NO está impresa"
        End If
        
    End If

End Select

End Sub
Private Sub Doc_Leer()
Dim m_resta As Integer
' CABECERA
Fecha.Text = Format(RsBulDet!Fecha, Fecha_Format)
m_Nv = RsBulDet!Nv
Rut.Text = RsBulDet![RUT CLiente]

RsNVc.Seek "=", m_Nv, m_NvArea
If Not RsNVc.NoMatch Then
    ComboNV.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
End If

'CbChofer.Text = NoNulo(rsbuldet![Observación 1])
'CbPatente.Text = NoNulo(rsbuldet![Observación 2])
'Obs(0).Text = NoNulo(rsbuldet![Observación 3])
'Obs(1).Text = NoNulo(rsbuldet![Observación 4])
'Obs(2).Text = NoNulo(rsbuldet![Observación 5])

'DETALLE
RsPd.Index = "NV-Plano-Marca"

'm_Tipo = rsbuldet!Tipo

'Select Case m_Tipo
Select Case "N"

Case "N"

    ' GUIA NORMAL
'    CbTipo.Text = "Normal"
'    GDespecial.Value = 0
'    m_Tipo = "N"
'    RsBulDet.Seek "=", Numero.Text, 1
    If Not RsBulDet.NoMatch Then
        Do While Not RsBulDet.EOF
            If RsBulDet!Numero = Numero.Text Then
            
                i = RsBulDet!linea
                
                Detalle.TextMatrix(i, 1) = RsBulDet!Plano
                Detalle.TextMatrix(i, 2) = RsBulDet!Rev
                Detalle.TextMatrix(i, 3) = RsBulDet!Marca
                
                RsPd.Seek "=", m_Nv, m_NvArea, RsBulDet!Plano, RsBulDet!Marca
                If Not RsPd.NoMatch Then
                    Detalle.TextMatrix(i, 4) = RsPd!Descripcion
                    Detalle.TextMatrix(i, 5) = RsPd![ITO fab]
                    m_resta = IIf(Accion = "Modificando", RsBulDet!Cantidad, 0)
                    Detalle.TextMatrix(i, 6) = RsPd![GD] - m_resta
                End If
                
                Detalle.TextMatrix(i, 7) = RsBulDet!Cantidad
                Detalle.TextMatrix(i, 8) = RsBulDet![PesoUnitario]
'                Detalle.TextMatrix(i, 10) = RsBulDet![Precio Unitario]
                
                Fila_Calcular_Normal i, False
                
            Else
                Exit Do
            End If
            RsBulDet.MoveNext
        Loop
    End If
    
End Select

RsPd.Index = "NV-Plano-Item"

Cliente_Lee Rut.Text

'If m_Tipo = "N" Then
If True Then

    Detalle.Row = 1 ' para q' actualice la primera fila del detalle
    Detalle_Sumar_Normal
    
Else

'    If m_Tipo = "E" Then
    If False Then
'        CbTipo.Text = "Especial"
    Else
    
'        CbTipo.Text = "Pintura"
        
'        Detalle_Especial.TextMatrix(0, 4) = "m2 Unitario"
'        Detalle_Especial.TextMatrix(0, 5) = "m2 TOTAL"
        
    End If

'    Detalle_Especial.Row = 1 ' para q' actualice la primera fila del detalle
'    Detalle_Sumar_Especial
    
End If

End Sub
Private Sub Cliente_Lee(Rut)
RsCl.Seek "=", Rut
If Not RsCl.NoMatch Then
    Razon.Text = RsCl![Razon Social]
'    Direccion.Text = RsCl!Dirección
'    Comuna.Text = NoNulo(RsCl!Comuna)
End If
End Sub
Private Function Doc_Validar() As Boolean
Dim porDespachar As Integer
Doc_Validar = False
If Rut.Text = "" Then
    MsgBox "DEBE ELEGIR CLIENTE"
'    Rut.SetFocus
    btnSearch.SetFocus
    Exit Function
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
        porDespachar = Detalle.TextMatrix(i, 5) - Val(Detalle.TextMatrix(i, 6))
        If porDespachar < Detalle.TextMatrix(i, 7) Then
            MsgBox "Sólo quedan " & porDespachar & " por Despachar", , "ATENCIÓN"
            Detalle.Row = i
            Detalle.col = 7
            Detalle.SetFocus
            Exit Function
        End If
        
        ' peso unitario    8
        ' peso total       9
        ' precio unitario 10
        ' precio total    11
        
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
Dim qry As String
save:

Doc_Detalle_Eliminar

GoTo Detalle
' CABECERA DE GD
With RsBulDet
If Nueva Then
    .AddNew
    !Numero = Numero.Text
'    !Tipo = m_Tipo '"N"
Else

    Doc_Detalle_Eliminar
    
    .Edit
    
End If

!Fecha = Fecha.Text
!Nv = Val(m_Nv)
![RUT CLiente] = Rut.Text
'![Peso Total] = Val(TotalPeso.Caption)
'![Peso Total] = CDbl(TotalPeso.Caption) ' modif 01/01/05
'![Precio Total] = Val(TotalPrecio.Caption)
'![Observación 1] = CbChofer.Text
'![Observación 2] = CbPatente.Text
'![Observación 3] = Obs(0).Text
'![Observación 4] = Obs(1).Text
'![Observación 5] = Obs(2).Text

.Update

End With

Detalle:
' DETALLE DE GD
'If m_Tipo = "N" Then
If True Then

    ' NORMAL
    With RsBulDet
    j = 0
    RsPd.Index = "NV-Plano-Marca"
    For i = 1 To n_filas
        m_Plano = Trim(Detalle.TextMatrix(i, 1))
        If m_Plano <> "" Then
        
            m_Marca = Detalle.TextMatrix(i, 3)
            m_cantidad = Val(Detalle.TextMatrix(i, 7))
            
            .AddNew
            !Numero = Numero.Text
            j = j + 1
            !linea = j
            
            !Nv = m_Nv
            !Fecha = Fecha.Text
            ![RUT CLiente] = Rut.Text
            
            !Plano = m_Plano
            !Rev = Detalle.TextMatrix(i, 2)
            !Marca = m_Marca
            !Cantidad = m_cantidad
    '        RsGDd("Fecha Despacho") = Fecha.Text
            ![PesoUnitario] = m_CDbl(Detalle.TextMatrix(i, 8))
'            ![Precio Unitario] = m_CDbl(Detalle.TextMatrix(i, 10)) '?
    
            .Update
            
'            RsPd.Seek "=", m_Nv, m_Plano, m_Marca
'            If RsPd.NoMatch Then
'                ' no existe marca en el plano
'            Else
'                ' actualiza archivo detalle planos
'                RsPd.Edit
'                RsPd![GD] = RsPd![GD] + m_cantidad
'                RsPd.Update
'            End If
            
        End If
    Next
    RsPd.Index = "NV-Plano-Item"
    End With
    
Else
End If

MousePointer = vbDefault

End Sub
Private Sub Doc_Eliminar()

' elimina cabecera
RsBulDet.Seek "=", Numero.Text
If Not RsBulDet.NoMatch Then

    RsBulDet.Delete
   
End If

' elimina detalle
Doc_Detalle_Eliminar

End Sub
Private Sub Doc_Detalle_Eliminar()
' elimina detalle GD
' al anular detalle GD debe actualizar detalle plano

'If m_Tipo = "N" Then
If True Then

    RsPd.Index = "NV-Plano-Marca"
    RsBulDet.Seek "=", Numero.Text, 1
    If Not RsBulDet.NoMatch Then
        Do While Not RsBulDet.EOF
            If RsBulDet!Numero <> Numero.Text Then Exit Do
'            RsPd.Seek "=", m_Nv, RsBulDet!Plano, RsBulDet!Marca
'            If Not RsPd.NoMatch Then
'                RsPd.Edit
'                RsPd![GD] = RsPd![GD] - RsBulDet!Cantidad
'                RsPd.Update
'            End If
        
            ' borra detalle
            RsBulDet.Delete
        
            RsBulDet.MoveNext
            
        Loop
    End If
    RsPd.Index = "NV-Plano-Item"

Else
    
    
End If

End Sub
Private Sub Campos_Limpiar()
Numero.Text = ""
'm_Tipo = "N"
'CbTipo.Text = "Normal"
'GDespecial.Value = 0
Fecha.Text = Fecha_Vacia
'Fecha.Text = Format(Now, Fecha_Format)
m_Nv = 0
Nv.Text = ""
ComboNV.Text = " "
m_Nv = 0
Rut.Text = ""
Razon.Text = ""
'Direccion.Text = ""
'Comuna.Text = ""

ComboMarca.Clear
ComboPlano.Clear

Detalle_Limpiar Detalle, n_columnas_N
'Detalle_Limpiar Detalle_Especial, n_columnas_E
'CbChofer.Text = ""
'CbPatente.Text = ""
'Obs(0).Text = ""
'Obs(1).Text = ""
'Obs(2).Text = ""
'Obs(2).Visible = False
'TotalPeso.Caption = ""
'TotalPrecio.Caption = ""

'Capturador = False
contador_registros = 0

End Sub
Private Sub Detalle_Limpiar(Detalle As Control, n_columnas As Integer)
Dim fi As Integer, co As Integer
For fi = 1 To n_filas
    For co = 1 To n_columnas
        Detalle.TextMatrix(fi, co) = ""
    Next
Next
Detalle.Row = 1
'Detalle

End Sub
Private Sub Obs_KeyPress(Index As Integer, KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim cambia_titulo As Boolean, n_Copias As Integer, m_Numero As String, NV_Numero As Double, NV_Obra As String, m_ImpresoraNombre As String
cambia_titulo = True
'Accion = ""
Select Case Button.Index
Case 1 ' agregar

    Accion = "Agregando"
    Botones_Enabled 0, 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    
    Numero.Text = Documento_Numero_Nuevo(RsBulDet, "Numero")

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
        
        If False Then
'        n_Copias = 1
'        PrinterNCopias.Numero_Copias = n_Copias
'        PrinterNCopias.Show 1
'        n_Copias = PrinterNCopias.Numero_Copias
    
'        If n_Copias > 0 Then
        If MsgBox("¿ IMPRIMIR ?", vbYesNo) = vbYes Then
'            If EmpOC.Rut = Rut_Eml Then
                'Eml
'                GD_PrintLegal Numero.Text, Mid(ComboNV.Text, 8)
'                GD_PrintLegal_Legal Numero.Text, Mid(ComboNV.Text, 8)
'            Else
                'PyP
'                GD_PrintLegal_Legal Numero.Text, Mid(ComboNV.Text, 8)
'            End If
            
'            Doc_Imprimir
            
            ' grabar "impresa" en GD Cabecera, para que no se pueda modificacar una vez que este impresa
'            RsBulDet.Edit
'            RsBulDet!impresa = True
'            RsBulDet.Update
            
        End If

        'End If
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
        
        End If
        
        ' nueva forma 08/03/2013
        prt_escoger.ImpresoraNombre = ""
        prt_escoger.Show 1
        m_ImpresoraNombre = prt_escoger.ImpresoraNombre

'        If MsgBox("¿ IMPRIME " & Obj & " ?", vbYesNo) = vbYes Then
        MousePointer = vbHourglass
        m_Numero = Numero.Text
        bulto_Prepara m_Numero, Nv.Text, Mid(ComboNV.Text, 8), m_ImpresoraNombre
            
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
            
        MousePointer = vbDefault
                     
        bulto_Print m_Numero
        'MsgBox "aqui"
        
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
            RsBulDet.Edit
            RsBulDet!impresa = False
'            rsbuldet!autorizada = True
            RsBulDet.Update
            
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
'            If CbTipo.Text = "Especial" Then
If False Then
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
        
            If Accion = "Agregando" Then
                Doc_Grabar True
            Else
                Doc_Grabar False
            End If
            
            If MsgBox("¿ IMPRIMIR " & Obj & " ?", vbYesNo) = vbYes Then
            
'                If EmpOC.Rut = Rut_Eml Then
                    'Eml
'                    GD_PrintLegal Numero.Text, Mid(ComboNV.Text, 8)
'                    GD_PrintLegal_Legal Numero.Text, Mid(ComboNV.Text, 8)
'                Else
                    'PyP
'                    GD_PrintLegal_Legal Numero.Text, Mid(ComboNV.Text, 8)
'                End If
                
'                Doc_Imprimir
                
'                RsBulDet.Seek "=", Numero.Text, 1
'                If Not RsBulDet.NoMatch Then
'                    RsBulDet.Edit
'                    RsBulDet!impresa = True
'                    RsBulDet.Update
'                End If
            
            End If
            
            Botones_Enabled 0, 0, 0, 0, 0, 1, 0
            Campos_Limpiar
            Campos_Enabled False
            Numero.Enabled = True
            Numero.SetFocus
            
            If Accion = "Agregando" Then Numero.Text = Documento_Numero_Nuevo(RsBulDet, "Nimero")
            
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

btnSearch.Enabled = Si
'CbTipo.Enabled = Si
'GDespecial.Enabled = Si
Nv.Enabled = Si
ComboNV.Enabled = Si
Detalle.Enabled = Si
'Detalle_Especial.Enabled = Si
'CbChofer.Enabled = Si
'CbPatente.Enabled = Si
'Obs(0).Enabled = Si
'Obs(1).Enabled = Si
'Obs(2).Enabled = Si

Depende_ITOPyG = True

End Sub
Private Sub btnSearch_Click()

Search.Muestra data_file, "Clientes", "RUT", "Razon Social", "Cliente", "Clientes"
Rut.Text = Search.Codigo

If Rut.Text <> "" Then
    RsCl.Seek "=", Rut.Text
    If RsCl.NoMatch Then
        MsgBox "CLIENTE NO EXISTE"
        Rut.SetFocus
    Else
'        Rut.Text = ""
        Razon.Text = Search.Descripcion
'        Direccion.Text = RsCl!Dirección
'        Comuna.Text = NoNulo(RsCl!Comuna)
        If ComboNV.Text <> "" Then Detalle.Enabled = True
    End If
End If
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
        ComboPlano.Width = Int(Detalle.CellWidth * 1.3)
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
'If m_Tipo = "N" Then
    EditKeyCode_N Detalle, txtEditGD, KeyCode, Shift
'Else
'    EditKeyCode_E Detalle_Especial, txtEditGD, KeyCode, Shift
'End If
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
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)
Select Case MSFlexGrid.col
Case 1, 2, 3
'    After_Detalle_Click
Case 4, 5, 6, 8, 9, 11
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
If KeyCode = vbKeyF2 Then
    MSFlexGridEdit Detalle, txtEditGD, 32
End If
End Sub
Private Sub Cursor_Mueve_N(MSFlexGrid As Control)
'MIA
Select Case MSFlexGrid.col
Case 7
'        MSFlexGrid.col = MSFlexGrid.col + 3
Case 10
    If MSFlexGrid.Row + 1 < MSFlexGrid.Rows Then
        MSFlexGrid.col = 1
        MSFlexGrid.Row = MSFlexGrid.Row + 1
    End If
Case Else
    MSFlexGrid.col = MSFlexGrid.col + 1
End Select
End Sub
Private Sub Cursor_Mueve_E(MSFlexGrid As Control)
'MIA
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
Private Sub Fila_Calcular_Normal(Fila As Integer, Actualizar As Boolean)
' actualiza solo linea, y totales generales

n7 = m_CDbl(Detalle.TextMatrix(Fila, 7))
n8 = m_CDbl(Detalle.TextMatrix(Fila, 8))
'n10 = m_CDbl(Detalle.TextMatrix(Fila, 10))

' peso total
Detalle.TextMatrix(Fila, 9) = Format(n7 * n8, num_Formato)
' precio total
'Detalle.TextMatrix(Fila, 11) = Format(n7 * n8 * n10, num_fmtgrl)

If Actualizar Then Detalle_Sumar_Normal

End Sub
Private Sub Detalle_Sumar_Normal()
Dim Tot_Kilos As Double, Tot_Precio As Double
Tot_Kilos = 0
Tot_Precio = 0
For i = 1 To n_filas
    Tot_Kilos = Tot_Kilos + m_CDbl(Detalle.TextMatrix(i, 9))
'    Tot_Precio = Tot_Precio + m_CDbl(Detalle.TextMatrix(i, 11))
Next

'TotalPeso.Caption = Format(Tot_Kilos, num_Formato)
'TotalPrecio.Caption = Format(Tot_Precio, num_fmtgrl)

End Sub
' FIN RUTINAS PARA FLEXGRID
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
Sub MSFlexGridEdit_E(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)

Select Case MSFlexGrid.col
Case 2
    Edt.MaxLength = 3
Case 3
    Edt.MaxLength = 21 '50
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
Private Sub MSComm_OnComm()
Dim EVMsg$
Dim ERMsg$
Dim i As Integer, ch As String, primer_caracter As Integer, suma As Integer
Dim checksum As Integer, resto As Integer, largo As Integer

Dim linea As Integer, registro As String, dv As Integer
Dim poscoma1 As Integer, poscoma2 As Integer, poscoma3 As Integer, codebar As String
'Dim Nv As Integer
Dim m_Plano As String, m_Rev As String, m_Marca As String
Dim posslash1 As Integer, posslash2 As Integer, posslash3 As Integer, j As Integer

Dim nlinea As Integer, m_cantidad As Integer

Dim m_Marca_deNV As Boolean, jj As Integer
Dim Numero_Slash As Long, Numero_Comas As Long

' Bifurca según la propiedad CommEvent.
Select Case MSComm.CommEvent
    ' Mensajes de evento.
    Case comEvReceive
    
'        Dim Buffer As Variant
        Buffer = MSComm.Input
        Reg = Reg & StrConv(Buffer, vbUnicode)
        
        ' verifica si es primer caracter del registro
        ' es un ascii entre 0 y 9
'            For i = 1 To Len(Buffer)
'                Debug.Print Asc(Mid(Buffer, i, 1));
'                If Asc(Mid(Buffer, i, 1)) = 0 Then
'                    Debug.Print "linea0"
'                End If
'            Next

        largo = Len(Reg)
        ch = Mid(Reg, largo)
        If Asc(ch) = 13 Then
        
            Select Case Left(Reg, largo - 1) ' reg menos el enter
            Case "ACK"
            
'                Debug.Print "ack"
                
            Case "OVER" & Chr(13) & "OVER"
            
                ' fin de transmision de datos
                            
'                Debug.Print "over"
                
                ' cierra puerto
                MSComm.PortOpen = False
                
                ' ordena arreglo
                For i = 1 To contador_registros - 1
                    For j = i + 1 To contador_registros
                    
                        If Arr_DatoCapturado(i, 0) > Arr_DatoCapturado(j, 0) Then
                        
                            codebar = Arr_DatoCapturado(i, 0)
                            Arr_DatoCapturado(i, 0) = Arr_DatoCapturado(j, 0)
                            Arr_DatoCapturado(j, 0) = codebar
                            
                            m_cantidad = Arr_DatoCapturado(i, 1)
                            Arr_DatoCapturado(i, 1) = Arr_DatoCapturado(j, 1)
                            Arr_DatoCapturado(j, 1) = m_cantidad
                            
                        End If
                        
                    Next
                Next
                
                
'                For i = 1 To contador_registros
'                    Debug.Print Arr_DatoCapturado(i, 0), Arr_DatoCapturado(i, 1)
'                Next
                
                ' busca si hay lecturas repetidas
                ' es decir, si etiqueta se leyó mas de una vez
'                For i = 1 To contador_registros - 1
                i = 0
                Do While i <= contador_registros - 1
                    i = i + 1
                
'                    For j = i + 1 To contador_registros
                    j = i
                    Do While j <= contador_registros
                        j = j + 1
                    
                        If Arr_DatoCapturado(i, 0) = Arr_DatoCapturado(j, 0) Then
                            ' suma un item mas
                            Arr_DatoCapturado(i, 1) = Arr_DatoCapturado(i, 1) + 1
                            
                            
                            ' "deplaza" arreglo hacia abajo
'                            For k = j To contador_registros - 1
                            k = j - 1
                            Do While k <= contador_registros - 1
                                k = k + 1
                                Arr_DatoCapturado(k, 0) = Arr_DatoCapturado(k + 1, 0)
                            Loop
'                            Next

'                For nlinea = 1 To contador_registros
'                    Debug.Print Arr_DatoCapturado(nlinea, 0), Arr_DatoCapturado(nlinea, 1)
'                Next
                
                            contador_registros = contador_registros - 1
                            j = j - 1
                            
                        End If
                        
'                    Next
                    Loop
'                Next
                Loop
                
                ' muestra arreglo de datos capturados desde scanpal
'                Debug.Print Arr_DatoCapturado(i, 0)

                j = 0
                For i = 1 To contador_registros
                
                    codebar = Arr_DatoCapturado(i, 0)
'                    Debug.Print codebar
'                    poscoma2 = InStr(poscoma1 + 1, registro, ",")
                    
                    posslash1 = InStr(1, codebar, "/")
                    posslash2 = InStr(posslash1 + 1, codebar, "/")
                    posslash3 = InStr(posslash2 + 1, codebar, "/")
                    
                    ' si el codebar tre 3 slash entonces es formato antiguo
                    If posslash3 > posslash2 And posslash2 > posslash1 And posslash1 > 0 Then
                    
                        j = j + 1
                    
                        m_Nv = Int(Left(codebar, posslash1 - 1))
                        m_Plano = Mid(codebar, posslash1 + 1, posslash2 - posslash1 - 1)
                        m_Rev = Mid(codebar, posslash2 + 1, posslash3 - posslash2 - 1)
                        m_Marca = Mid(codebar, posslash3 + 1)
                        
    '                    Debug.Print "nv|" & Nv & "|"
    '                    Debug.Print "plano|" & Plano & "|"
    '                    Debug.Print "rev|" & Rev & "|"
    '                    Debug.Print "marca|" & Marca & "|"
                    
                        If j = 1 Then
                            ' cabecera
                            'lee primer registro y rescata datos para cabecera de GD
    '                        s1 = InStr(1, Reg, "/")
    '                        If s1 > 0 Then
    '                            m_Nv = Left(Reg, s1 - 1)
                                Nv.Text = m_Nv
                                Nv_LostFocus ' para que busque obra
                                
    '                        End If
                            '////////////////
                        End If
                        
                    '    ComboPlano.Visible = True
                        Detalle.Row = j
                        Detalle.col = 1
                        ComboPlano.Text = m_Plano
                        
                        Detalle.col = 3
                        After_Detalle_Click
                        ComboMarca.Text = m_Marca
                        ComboMarca_Click ' para que pueble
                    
                        ' deja cantidad en uno
                        Detalle.col = 7
                        'Detalle = "1"
                        Detalle = Arr_DatoCapturado(i, 1)
                        
                        ' multiplica cantidad x pesounitario
                        Detalle.col = 9
                        Detalle = Arr_DatoCapturado(i, 1) * m_CDbl(Detalle.TextMatrix(i, 8))
                    
                    End If
                    
                    ' nuevo formato de lectura de codigo de barras
                    ' 18/11/08
                    ' si existe 1 solo slash, entonces es formato nuevo
                    If posslash2 = 0 And posslash1 > 0 Then
                    
                        j = j + 1
                        If j = 1 Then
                            If m_Nv = 0 Then
                                MsgBox "Debe Escojer Nota de Venta"
                                ComboNV.SetFocus
                                GoTo Sigue
                            End If
                        End If
'                        Debug.Print "nuevo formato|"; codebar; "|"
                        
'                        m_Nv = Int(Left(codebar, posslash1 - 1))

                        m_Plano = Left(codebar, posslash1 - 2)
                        m_Rev = Mid(codebar, posslash1 - 1, 1)
                        m_Marca = Mid(codebar, posslash1 + 1)
                        
                        Detalle.Row = j
                        Detalle.col = 1
                        
                        ' busca si plano es de nv
                        m_Marca_deNV = False
                        For jj = 1 To ComboPlano.ListCount - 1
                            If m_Plano = ComboPlano.List(jj) Then
                            
                                ComboPlano.Text = m_Plano
                                ' deja cantidad en uno
                                Detalle.col = 7
                                'Detalle_Normal = "1"
                                Detalle = Arr_DatoCapturado(i, 1)
                                                                
                                Detalle.col = 3
                                After_Detalle_Click
                                ComboMarca.Text = m_Marca
                                ComboMarca_Click ' para que pueble
                                
                                m_Marca_deNV = True
                                                                
                            End If
                            
                        Next
                        
                        If Not m_Marca_deNV Then
                            MsgBox "Plano " & m_Plano & " NO es de NV " & m_Nv
                        End If
                        
                    End If
                
                Next
                '/////////////////////////////
                Detalle.Row = 1
'                Detalle.col = 10
                Detalle.SetFocus
Sigue:

                Capturando = False
                
            Case Else
            
                registro = Left(Reg, largo - 2)
'Debug.Print "registro" & registro ' registro completo
                linea = Asc(Left(registro, 1))
                If linea > 9 Then
                    ' es linea 0
                    linea = 0
                Else
                    registro = Mid(registro, 2)
                    largo = largo - 1
                End If
                
                dv = Asc(Mid(registro, largo - 2, 1))
'                Debug.Print linea; "|"; registro; "|"; dv
                
                ' registo: es el registro que me sirve
                ' codigo,fecha,hora
                ' XXX,N,YYYYMMDD
                ' XXX: es el codigo de barras, code128, largo variable
                ' este formato se carga en el ScanPal2 con el programa AG_SP2x.exe, archivo scp.atx
                ' viene en formato "NV/PLANO/REV/MARCA" (antiguo)
                
                
                ' codigo,n,fecha
                ' CODIGOBARRAS,N,YYYYMMDD
                ' CODIGOBARRAS: viene en formato "PLANOR/MARCA" (nuevo)
                ' N: es la cantidad digitada en el scanpal2
                
                Numero_Slash = CharCount(registro, "/")
                Numero_Comas = CharCount(registro, ",")
                
                If Numero_Comas = 2 Then
                
                    ' ok
                    If Numero_Slash = 3 Then
                        ' formato antiguo
                    End If
                    If Numero_Slash = 1 Then
                        ' formato nuevo
                    End If
                    
                    poscoma1 = InStr(1, registro, ",")
                    poscoma2 = InStr(poscoma1 + 1, registro, ",")
                    
                    m_cantidad = Mid(registro, poscoma1 + 1, poscoma2 - poscoma1)
'                    Debug.Print "|" & m_Cantidad & "|"
                
                    codebar = Left(registro, poscoma1 - 1)
                    
                    contador_registros = contador_registros + 1

                    Arr_DatoCapturado(contador_registros, 0) = codebar
                    Arr_DatoCapturado(contador_registros, 1) = m_cantidad
                    
                    ' pide siguiente registro
                    Buffer = "ACK" & vbCr
                    MSComm.Output = Buffer
                    
                Else
                
                    MsgBox "DEBE usar programa DOS en el ScanPal2 para lectura de bultos"
                    Exit Sub
                    
                End If
                                    
            End Select
            
            Reg = ""
            
        End If

'''''''''''            Debug.Print "Recibir - " & StrConv(Buffer, vbUnicode)
        
'        ShowData txtTerm, (StrConv(Buffer, vbUnicode))
        
    Case comEvSend
    Case comEvCTS
        ' terminó transmision desde scanpal2
        EVMsg$ = "Detectado cambio en CTS"
    Case comEvDSR
        EVMsg$ = "Detectado cambio en DSR"
    Case comEvCD
        EVMsg$ = "Detectado cambio en CD"
    Case comEvRing
        EVMsg$ = "El teléfono está sonando"
    Case comEvEOF
        EVMsg$ = "Detectado el final del archivo"

    ' Mensajes de error.
    Case comBreak
        ERMsg$ = "Parada recibida"
    Case comCDTO
        ERMsg$ = "Sobrepasado el tiempo de espera de detección de portadora"
    Case comCTSTO
        ERMsg$ = "Soprepasado el tiempo de espera de CTS"
    Case comDCB
        ERMsg$ = "Error recibiendo DCB"
    Case comDSRTO
        ERMsg$ = "Sobrepasado el tiempo de espera de DSR"
    Case comFrame
        ERMsg$ = "Error de marco"
    Case comOverrun
        ERMsg$ = "Error de sobrecarga"
    Case comRxOver
        ERMsg$ = "Desbordamiento en el búfer de recepción"
    Case comRxParity
        ERMsg$ = "Error de paridad"
    Case comTxFull
        ERMsg$ = "Búfer de transmisión lleno"
    Case Else
        ERMsg$ = "Error o evento desconocido"
End Select

If Len(EVMsg$) Then
    ' Muestra los mensajes de evento en la barra de estado.
'    sbrStatus.Panels("Status").Text = "Estado:" & EVMsg$
            
    ' Activa el cronómetro para que el mensaje de la barra
    ' de estado se borre después de dos segundos.
'    Timer2.Enabled = True
    
ElseIf Len(ERMsg$) Then
    ' Muestra los mensajes de evento en la barra de estado.
'    sbrStatus.Panels("Status").Text = "Estado:" & ERMsg$
    
    ' Muestra los mensajes de error en un cuadro de alerta.
    Beep
    Ret = MsgBox(ERMsg$, 1, "Haga clic en Cancelar para salir, clic en Aceptar para ignorar.")
    
    ' Si el usuaruio hace clic en Cancelar (2)...
    If Ret = 2 Then
        MSComm.PortOpen = False    ' Cierra el puerto y sale.
    End If
    
    ' Activa el cronómetro para que el mensaje de la barra
    ' de estado se borre después de dos segundos.
'    Timer2.Enabled = True

End If
End Sub
Private Sub OLDDoc_Imprimir()
' imprime ITOf
MousePointer = vbHourglass
Dim can_valor As String, can_col As Integer, k As Integer
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer, tab4 As Integer
Dim tab5 As Integer, tab6 As Integer, tab7 As Integer, tab8 As Integer, tab40 As Integer

Dim m_TotalKilos As Double
m_TotalKilos = 0

tab0 = 6 'margen izquierdo
tab1 = tab0 + 0
tab2 = tab1 + 10
tab3 = tab2 + 5
tab4 = tab3 + 10
tab5 = tab4 + 19
tab6 = tab5 + 10
tab7 = tab6 + 10
tab8 = tab7 + 10
tab40 = 50

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt

'For k = 1 To n_Copias

prt.Font.Size = 15
prt.Print Tab(tab0 + 14); "BULTO Nº" & Format(Numero.Text, "000")
'prt.Font.Size = 10
prt.Print ""

' cabecera
prt.Font.Bold = True
prt.Print Tab(tab0); Empresa.Razon
prt.Font.Bold = False
'prt.Print Tab(tab0 + tab40); "FECHA     : " & Fecha.Text

'prt.Print Tab(tab0); "GIRO: " & Empresa.Giro;
prt.Print Tab(tab0); "SEÑOR(ES) : " & Razon

'prt.Print Tab(tab0); Empresa.Direccion;
'prt.Print Tab(tab0 + tab40); "RUT       : " & Rut

'prt.Print Tab(tab0); "Teléfono: " & Empresa.Telefono1 & " - " & Empresa.Comuna;
'prt.Print Tab(tab0 + tab40); "DIRECCIÓN : " & Direccion,

'prt.Print Tab(tab0 + tab40); "COMUNA    : " & Comuna

prt.Print Tab(tab0); "NV        : ";
prt.Font.Bold = True
prt.Print Format(ComboNV.Text, ">")
prt.Font.Bold = False

prt.Font.Size = 12

prt.Print ""
' detalle
prt.Font.Bold = True
prt.Print Tab(tab1); "PLANO";
prt.Print Tab(tab2); "REV";
prt.Print Tab(tab3); "MARCA";
prt.Print Tab(tab4); "DESCRIPCIÓN";
prt.Print Tab(tab5); "CANT";
prt.Print Tab(tab6); "  KG UNIT";
prt.Print Tab(tab7); " KG TOTAL"
prt.Font.Bold = False

prt.Print Tab(tab1); String(73, "-")

j = -1
For i = 1 To n_filas

    can_valor = Detalle.TextMatrix(i, 7)
    
    If Val(can_valor) = 0 Then
    
'        j = j + 1
'        prt.Print Tab(tab1 + j * 5); "    \"
        
    Else
    
        ' PLANO
        prt.Print Tab(tab1); Detalle.TextMatrix(i, 1);
        
        ' REVISIÓN
        prt.Print Tab(tab2); Detalle.TextMatrix(i, 2);
        
        ' MARCA
        prt.Print Tab(tab3); Detalle.TextMatrix(i, 3);
        
        ' DESCRIPCIÓN
        prt.Print Tab(tab4); Left(Detalle.TextMatrix(i, 4), 18);
        
        ' CANTIDAD
        can_valor = Trim(Format(can_valor, "####"))
        can_col = 4 - Len(can_valor)
        prt.Print Tab(tab5 + can_col); can_valor;
        
        ' KG UNITARIO
        can_valor = Trim(Format(m_CDbl(Detalle.TextMatrix(i, 8)), "###,###.0"))
        can_col = 9 - Len(can_valor)
        prt.Print Tab(tab6 + can_col); can_valor;
        
        ' KG TOTAL
        can_valor = Trim(Format(m_CDbl(Detalle.TextMatrix(i, 9)), "###,###.0"))
        can_col = 9 - Len(can_valor)
        prt.Print Tab(tab7 + can_col); can_valor
        
        m_TotalKilos = m_TotalKilos + can_valor
        
    End If
    
Next

prt.Print Tab(tab1); String(73, "-")

prt.Print Tab(tab0 + 46); "TOTAL KILOS : ";
prt.Font.Bold = True
prt.Print Format(m_TotalKilos, "#,###,###.0")
prt.Font.Bold = False
prt.Print ""
'prt.Print Tab(tab0); "OBSERVACIONES :";
'prt.Print Tab(tab0 + 16); Obs(0).Text
'prt.Print Tab(tab0 + 16); Obs(1).Text
'prt.Print Tab(tab0 + 16); Obs(2).Text
'prt.Print Tab(tab0 + 16); Obs(3).Text

For i = 1 To 1
    prt.Print ""
Next

prt.Print Tab(tab0); Tab(14), "S/Guía Nº"
prt.Print Tab(tab0); Tab(14), "          ________________"

prt.EndDoc

'Next

Impresora_Predeterminada "default"

MousePointer = vbDefault

End Sub
Private Sub bulto_Print(bultoNumero)
Cr.WindowTitle = "BULTO Nº " & bultoNumero
Cr.ReportSource = crptReport
Cr.WindowState = crptMaximized
'Cr.WindowBorderStyle = crptFixedSingle
Cr.WindowMaxButton = False
Cr.WindowMinButton = False
Cr.Formulas(0) = "RAZON SOCIAL=""" & Empresa.Razon & """"
'Cr.Formulas(4) = "RUT=""" & "RUT: " & EmpOC.Rut & """"

Cr.DataFiles(0) = repo_file & ".MDB"

Cr.ReportFileName = Drive_Server & Path_Rpt & "bulto.rpt"

Cr.Action = 1

End Sub
Private Sub Privilegios()
If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 1, 0, 0
End If
End Sub
