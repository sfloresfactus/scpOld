VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ITO_PG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ITO PG"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PictureL 
      Height          =   495
      Left            =   13080
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame_Operador 
      Caption         =   "Operador"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   4320
      TabIndex        =   31
      Top             =   1200
      Width           =   5655
      Begin VB.TextBox OperadorRazon 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   33
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton btnOperadorSearch 
         Height          =   300
         Left            =   1920
         Picture         =   "ITO_PG.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   300
      End
      Begin MSMask.MaskEdBox OperadorRut 
         Height          =   300
         Left            =   600
         TabIndex        =   34
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label lblOperadorRut 
         Caption         =   "RUT"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.PictureBox PictureQrCode 
      Height          =   1170
      Left            =   11160
      Picture         =   "ITO_PG.frx":0102
      ScaleHeight     =   74
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   50
      TabIndex        =   30
      Top             =   1440
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton btnBuscar 
      Caption         =   "Busca Marca"
      Height          =   300
      Left            =   1680
      TabIndex        =   29
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtBuscar 
      Height          =   300
      Left            =   360
      MaxLength       =   10
      TabIndex        =   28
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton btnPendientes 
      Caption         =   "Traer Piezas Pendientes"
      Height          =   615
      Left            =   10080
      TabIndex        =   27
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton btnTG 
      Caption         =   "Copiar TG"
      Height          =   375
      Left            =   11520
      TabIndex        =   26
      Top             =   840
      Width           =   1095
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13485
      _ExtentX        =   23786
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
   Begin VB.ComboBox CbTipoGranalla 
      Height          =   315
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   1920
      Width           =   975
   End
   Begin VB.ComboBox CbTrabajadores 
      Height          =   315
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   23
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
      TabIndex        =   21
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
      TabIndex        =   16
      Top             =   5280
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   0
      Left            =   360
      MaxLength       =   50
      TabIndex        =   15
      Top             =   4920
      Width           =   5000
   End
   Begin VB.ComboBox ComboMarca 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   8040
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox ComboPlano 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame_Contratista 
      Caption         =   "Contratista"
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   4320
      TabIndex        =   8
      Top             =   480
      Width           =   5655
      Begin VB.CommandButton btnContratistaSearch 
         Height          =   300
         Left            =   1920
         Picture         =   "ITO_PG.frx":2DFA4
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   300
      End
      Begin VB.TextBox ContratistaRazon 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   240
         Width           =   3015
      End
      Begin MSMask.MaskEdBox ContratistaRut 
         Height          =   300
         Left            =   600
         TabIndex        =   10
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label lblContratistaRut 
         Caption         =   "RUT"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.TextBox txtEditGD 
      Height          =   285
      Left            =   8040
      TabIndex        =   17
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
      Height          =   2685
      Left            =   120
      TabIndex        =   13
      Top             =   1850
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4736
      _Version        =   393216
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
            Picture         =   "ITO_PG.frx":2E0A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_PG.frx":2E1B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_PG.frx":2E2CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_PG.frx":2E3DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_PG.frx":2E4EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_PG.frx":2E600
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_PG.frx":2E712
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label TotalKg 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   5760
      TabIndex        =   25
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Observación"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   14
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
      TabIndex        =   22
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
      TabIndex        =   20
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
Attribute VB_Name = "ITO_PG"
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

' G : galvanizado ahora llamada "reproceso" YO NO VA

' R : granallado, Erwin
' T : produccion pintura , Erwin
' P : pintura

Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button
Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, k As Integer, d As Variant

Private DbD As Database, RsCl As Recordset, RsTra As Recordset

'Private RsSc As Recordset
' contratistas
'Private SqlRsSc As New ADODB.Recordset

Private Dbm As Database, RsNVc As Recordset
Private RsNvPla As Recordset, RsPd As Recordset, RsPdBuscar As Recordset
Private RsITOpgc As Recordset, RsITOpgd As Recordset

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
Private cliRut As String, cliRazon As String, cliNombreFantasia As String
Private AjusteX As Double, AjusteY As Double
Private a_Trabajadores(1, 199) As String, m_Nombre As String
Private conContratista As Boolean
Private conOperador As Boolean
Private a_TipoGranalla(9) As String, m_TotalTiposGranalla As Integer
Private cQrCode As ClsQrCode ' 04/08/12 para codigo qr
Private mProyecto As String, mTag As String
Public Property Let TipoDoc(ByVal New_Tipo As String)
m_TipoDoc = New_Tipo
End Property

Private Sub btnBuscar_Click()

ComboPlano.visible = False
ComboMarca.visible = False

If txtBuscar.Text = "" Then
    MsgBox "Debe digitar pieza a buscar"
    txtBuscar.SetFocus
    Exit Sub
End If

If Nv.Text = "" Then
    MsgBox "Debe digitar NV"
    Nv.SetFocus
    Exit Sub
End If

Dim sql As String, canPendiente As Integer, fi As Integer
Dim can1 As Integer, can2 As Integer

sql = "SELECT * FROM [planos detalle] WHERE nv=" & Nv.Text & " AND marca LIKE '*" & txtBuscar.Text & "*'"
'sql = sql & " AND "
canPendiente = 0
Set RsPdBuscar = Dbm.OpenRecordset(sql)
With RsPdBuscar
Do While Not .EOF
    Select Case m_TipoDoc
    Case "R" ' granallado
        can1 = ![ITO fab]
        can2 = ![ito gr]
    Case "T" ' produccion pintura
        can1 = ![ito gr]
        can2 = ![ITO pp]
    Case "P" ' pintura
        can1 = ![ITO pp]
        can2 = ![ITO pyg]
    End Select
    canPendiente = can1 - can2

    
    If canPendiente > 0 Then
        'Debug.Print !Nv, !Plano, !Marca, ![ito pp], ![ito pyg]
        For fi = 1 To n_filas
        
            If Detalle.TextMatrix(fi, 1) = "" Then
        
                Detalle.TextMatrix(fi, 1) = !Plano
                Detalle.TextMatrix(fi, 2) = !Rev
                Detalle.TextMatrix(fi, 3) = !Marca
                Detalle.TextMatrix(fi, 4) = !Descripcion
                Detalle.TextMatrix(fi, 5) = can1
                Detalle.TextMatrix(fi, 6) = can2
                Detalle.TextMatrix(fi, 8) = ![Superficie]
                Detalle.TextMatrix(fi, 10) = ![Peso]
                
                Exit For
                
            End If
            
        Next
        
    End If
    
    .MoveNext
    
Loop
.Close
End With

txtBuscar.Text = ""

End Sub
Private Sub btnOperadorSearch_Click()

If m_TipoDoc = "T" Then ' produccion pintura
    Search.Muestra data_file, "Trabajadores", "RUT", "ApPaterno]+' '+[ApMaterno]+' '+[Nombres", Obj, Objs, "tipo4"
End If
If m_TipoDoc = "R" Then ' granallado
    Search.Muestra data_file, "Trabajadores", "RUT", "ApPaterno]+' '+[ApMaterno]+' '+[Nombres", Obj, Objs, "tipo5"
End If

OperadorRut.Text = Search.Codigo
OperadorRazon.Text = Search.Descripcion

End Sub

Private Sub btnPendientes_Click()
' trae piezas pendientes dependiendo del tipo de ITO
' debe tener escogida la NV y Contratista
Dim canPendiente As Integer, fi As Integer

If Nv.Text = "" Then
    MsgBox "Debe escoger NV"
    Nv.SetFocus
    Exit Sub
End If

If m_TipoDoc = "R" Then ' granallado

    Detalle_Limpiar Detalle, n_columnas

    ' busca lo que ya esta fabricado y no granallado, de toda la NV
    i = 0
    With RsPd
    .Seek ">=", m_Nv, m_NvArea, "", 1
    If Not .NoMatch Then
        Do While Not .EOF
            If !Nv = m_Nv Then
                canPendiente = ![ITO fab] - ![ito gr]
                If canPendiente > 0 Then
                    fi = fi + 1
                    If fi <= n_filas Then
                        Detalle.TextMatrix(fi, 1) = !Plano
                        Detalle.TextMatrix(fi, 2) = !Rev
                        Detalle.TextMatrix(fi, 3) = !Marca
                        Detalle.TextMatrix(fi, 4) = !Descripcion
                        Detalle.TextMatrix(fi, 5) = ![ITO fab]
                        Detalle.TextMatrix(fi, 6) = ![ito gr]
                        Detalle.TextMatrix(fi, 8) = ![Superficie]
                        Detalle.TextMatrix(fi, 10) = ![Peso]
                    End If
                End If
            Else
                Exit Do
            End If
            .MoveNext
        Loop
    End If
    End With
    
End If

If m_TipoDoc = "T" Then ' produccion pintura

    Detalle_Limpiar Detalle, n_columnas
    
    ' busca lo que ya esta granallado menos produccion pintura, de toda la NV
    i = 0
    With RsPd
    .Seek ">=", m_Nv, m_NvArea, "", 1
    If Not .NoMatch Then
        Do While Not .EOF
            If !Nv = m_Nv Then
                canPendiente = ![ito gr] - ![ITO pp]
                If canPendiente > 0 Then
                    fi = fi + 1
                    If fi <= n_filas Then
                        Detalle.TextMatrix(fi, 1) = !Plano
                        Detalle.TextMatrix(fi, 2) = !Rev
                        Detalle.TextMatrix(fi, 3) = !Marca
                        Detalle.TextMatrix(fi, 4) = !Descripcion
                        Detalle.TextMatrix(fi, 5) = ![ito gr]
                        Detalle.TextMatrix(fi, 6) = ![ITO pp]
                        Detalle.TextMatrix(fi, 8) = ![Superficie]
                        Detalle.TextMatrix(fi, 10) = ![Peso]
                    End If
                End If
            Else
                Exit Do
            End If
            .MoveNext
        Loop
    End If
    End With
    
End If

If m_TipoDoc = "P" Then ' pintura

    Detalle_Limpiar Detalle, n_columnas
    
    ' busca lo que ya esta en produccion pintura menos pintadas, de toda la NV
    i = 0
    With RsPd
    .Seek ">=", m_Nv, m_NvArea, "", 1
    If Not .NoMatch Then
        Do While Not .EOF
            If !Nv = m_Nv Then
                canPendiente = ![ITO pp] - ![ITO pyg]
                If canPendiente > 0 Then
                    fi = fi + 1
                    If fi <= n_filas Then
                        Detalle.TextMatrix(fi, 1) = !Plano
                        Detalle.TextMatrix(fi, 2) = !Rev
                        Detalle.TextMatrix(fi, 3) = !Marca
                        Detalle.TextMatrix(fi, 4) = !Descripcion
                        Detalle.TextMatrix(fi, 5) = ![ITO pp]
                        Detalle.TextMatrix(fi, 6) = ![ITO pyg]
                        Detalle.TextMatrix(fi, 8) = ![Superficie]
                        Detalle.TextMatrix(fi, 10) = ![Peso]
                    End If
                End If
            Else
                Exit Do
            End If
            .MoveNext
        Loop
    End If
    End With
    
End If

End Sub

Private Sub btnTG_Click()
Dim m_TG As String, i As Integer
m_TG = Detalle.TextMatrix(1, 14)
If m_TG = "" Then
    MsgBox "Debe escoger Tipo Granalado en linea 1"
Else
    For i = 1 To n_filas
        If Detalle.TextMatrix(i, 1) <> "" Then
            Detalle.TextMatrix(i, 14) = m_TG
        End If
    Next
End If
End Sub
Private Sub CbTipoGranalla_Click()
Detalle = CbTipoGranalla.Text
CbTipoGranalla.visible = False
End Sub
Private Sub CbTrabajadores_Click()
Detalle = CbTrabajadores.Text
CbTrabajadores.visible = False
End Sub

Private Sub Detalle_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MsgBox "mouse up"
If Button = 2 Then ' boton derecho
    MsgBox Button & "|" & Detalle.Row
End If
End Sub

Private Sub Form_Load()

' abre archivos
Set DbD = OpenDatabase(data_file)
'Set RsSc = DbD.OpenRecordset("Contratistas")
'RsSc.Index = "RUT"

Set RsCl = DbD.OpenRecordset("Clientes")
RsCl.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set RsITOpgc = Dbm.OpenRecordset("ITO PG Cabecera")
RsITOpgc.Index = "Numero"
'RsITOpgc.Index = "Tipo-Numero"

Set RsITOpgd = Dbm.OpenRecordset("ITO PG Detalle")
'RsITOpgd.Index = "Número-Línea"
RsITOpgd.Index = "Tipo-Numero-Linea"

'Dbm.Execute "UPDATE [ITO PG cabecera] SET [rut contratista]='89784800-7' WHERE tipo='R' OR tipo='T'"
'Dbm.Execute "UPDATE [ITO PG detalle] SET [rut contratista]='89784800-7' WHERE tipo='R' OR tipo='T'"
'Dbm.Execute "UPDATE [ITO PG detalle] SET [rut operador]='' WHERE [rut operador] IS NULL"

'DbM.Execute "UPDATE [ITO PG cabecera] SET protocolo=0"

'duplicados
'busca 2856

' puebla nv
' Combo obra
ComboNV.AddItem " "
i = 0
Do While Not RsNVc.EOF

    If Usuario.Nv_Activas = False Then ' todas
        GoTo IncluirNV
    Else
        If Usuario.Nv_Activas And RsNVc!activa Then
            GoTo IncluirNV
        End If
    End If
    If False Then
IncluirNV:

'    If m_TipoDoc = "P" And RsNVc!pintura Then
    If m_TipoDoc = "P" Then
    
        i = i + 1
        a_Nv(i, 0) = RsNVc!Numero
        a_Nv(i, 1) = RsNVc!obra
    
        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
        
    End If
    If m_TipoDoc = "G" And RsNVc!galvanizado Then
    
        i = i + 1
        a_Nv(i, 0) = RsNVc!Numero
        a_Nv(i, 1) = RsNVc!obra
    
        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
        
    End If
    
    If m_TipoDoc = "R" Then  ' granallado
    
        i = i + 1
        a_Nv(i, 0) = RsNVc!Numero
        a_Nv(i, 1) = RsNVc!obra
    
        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
        
    End If
    
    If m_TipoDoc = "T" Then  ' produccion pintura
    
        i = i + 1
        a_Nv(i, 0) = RsNVc!Numero
        a_Nv(i, 1) = RsNVc!obra
    
        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
        
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

Set RsTra = DbD.OpenRecordset("Trabajadores")
RsTra.Index = "RUT"

'n_columnas = 15 + 1
n_columnas = 17

btnTG.visible = False
conContratista = False
conOperador = False

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

    conContratista = True
    conOperador = True

    n_columnas = 17
    
    btnTG.visible = True
    
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
    conContratista = True
    conOperador = True
    Detalle_Config
    n_columnas = 17
End If

Contratista_Visible conContratista
Operador_Visible conOperador

Privilegios

m_NvArea = 0

CbTipoGranalla.visible = False
CbTipoGranalla.Width = 100

CbTrabajadores.visible = False
CbTrabajadores.Width = 1000

End Sub
Private Sub Contratista_Visible(visible As Boolean)
Frame_Contratista.visible = visible
'lblRut.visible = visible
'lblSenores.visible = visible
'lblDireccion.visible = visible
'lblComuna.visible = visible
End Sub
Private Sub Operador_Visible(visible As Boolean)
Frame_Operador.visible = visible
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
    n_filas = 30
End If
If m_TipoDoc = "G" Then
'    Obj = "ITO GALVANIZADO"
'    Objs = "ITOS GALVANIZADO"
    Obj = "ITO REPROCESO"
    Objs = "ITOS REPROCESO"
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

Accion = ""

btnContratistaSearch.visible = False
btnContratistaSearch.ToolTipText = "Busca Contratistas"

btnOperadorSearch.visible = False
btnOperadorSearch.ToolTipText = "Busca Operadores"

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
cliRut = ""
cliRazon = ""
cliNombreFantasia = ""
RsNVc.Seek "=", m_Nv, m_NvArea
If Not RsNVc.NoMatch Then
    RsCl.Seek "=", RsNVc![RUT CLiente]
    cliRazon = "Cliente NO Existe"
    If Not RsCl.NoMatch Then
        cliRut = RsCl!Rut
        cliRazon = RsCl![Razon Social]
        cliNombreFantasia = NoNulo(RsCl![NombreFantasia])
        If cliNombreFantasia = "" Then
            cliNombreFantasia = cliRazon
        End If
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
c_itogr = 0 ' grtanallado
c_itopp = 0 ' prod pintura
c_itopg = 0 ' pintura

' busca marca en plano
RsPd.Seek ">=", m_Nv, m_NvArea, m_Plano, 0
If Not RsPd.NoMatch Then
    Do While Not RsPd.EOF
        If RsPd!Marca = m_Marca Then
        
            c_tot = RsPd![Cantidad Total]
            c_otf = RsPd![OT fab]
            c_itof = RsPd![ITO fab] ' fabricacion
            c_itogr = RsPd![ito gr] ' granallado
            c_itopp = RsPd![ITO pp] ' produccion pintura
            c_itopg = RsPd![ITO pyg] ' pintura
            c_gdgal = RsPd![GD gal] ' ?
            
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
                    & "No está Recibida"
                Detalle.TextMatrix(fil, 3) = ""
                Detalle.SetFocus
                Exit Sub
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
            
            ' verifica que quede algo por procesar
            Select Case m_TipoDoc
            Case "R"
                ' granalla
                If c_itof - c_itogr <= 0 Then
                    Beep
                    MsgBox "La marca """ & m_Marca & """" & vbCr & "Ya se Granalló"
                    Detalle.TextMatrix(fil, 3) = ""
                    Detalle.SetFocus
                    Exit Sub
                End If
            Case "T"
                ' produccion pintura
'                If c_itopp - c_itopg <= 0 Then
                If c_itogr <= c_itopp Then
                    Beep
'                    MsgBox "La marca """ & m_Marca & """" & vbCr & "Ya se Está en Produccion Pintura"
                    MsgBox "En la marca """ & m_Marca & """" & vbCr & "itoGR=" & c_itogr & ", itoPP=" & c_itopp
                    Detalle.TextMatrix(fil, 3) = ""
                    Detalle.SetFocus
                    Exit Sub
                End If
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
            
            If m_TipoDoc = "P" Then
                Detalle.TextMatrix(fil, 5) = c_itopp
                Detalle.TextMatrix(fil, 6) = c_itopg
            End If
            If m_TipoDoc = "G" Then
                Detalle.TextMatrix(fil, 5) = c_gdgal
                Detalle.TextMatrix(fil, 6) = c_itopg
            End If
            If m_TipoDoc = "R" Then
                Detalle.TextMatrix(fil, 5) = c_itof
                Detalle.TextMatrix(fil, 6) = c_itogr - Val(Detalle.TextMatrix(fil, 7))
            End If
            If m_TipoDoc = "T" Then
                Detalle.TextMatrix(fil, 5) = c_itogr
                Detalle.TextMatrix(fil, 6) = c_itopp - Val(Detalle.TextMatrix(fil, 7))
            End If
            
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


If m_TipoDoc = "R" Then ' granallado

    btnPendientes.Left = 10050 ' 4500

    Detalle.TextMatrix(0, 5) = "Cant itoF"            '*
    Detalle.TextMatrix(0, 6) = "Cant itoGr"            '*
    
    Detalle.TextMatrix(0, 14) = "Tipo Grana"
    Detalle.TextMatrix(0, 15) = "Maquina"
    
    Detalle.ColWidth(14) = 700
    Detalle.ColWidth(15) = 700
    Detalle.ColWidth(16) = 550
    Detalle.ColWidth(17) = 2000
    
End If

If m_TipoDoc = "T" Then

    btnPendientes.Left = 10050 ' 4500
    
    Detalle.TextMatrix(0, 5) = "Cant itoGr"            '*
    Detalle.TextMatrix(0, 6) = "Cant itoPp"            '*
    
    Detalle.TextMatrix(0, 14) = "nManos Antic"
    Detalle.TextMatrix(0, 15) = "nManos Termin"
    
    Detalle.ColWidth(14) = 700
    Detalle.ColWidth(15) = 700
    Detalle.ColWidth(16) = 550
    Detalle.ColWidth(17) = 2000
    
End If

If m_TipoDoc = "P" Then ' pintura

    btnPendientes.Left = 10050
    
    Detalle.TextMatrix(0, 5) = "Cant itoPp" ' *
    Detalle.TextMatrix(0, 6) = "Cant itoP"  ' *

    Detalle.ColWidth(14) = 0
    Detalle.ColWidth(15) = 0
    Detalle.ColWidth(16) = 0
    Detalle.ColWidth(17) = 0

End If

If m_TipoDoc = "G" Then

''Detalle.TextMatrix(0, 6) = "Cant ITOG"            '*
    Detalle.TextMatrix(0, 5) = "Cant ITOF"            '*
    Detalle.TextMatrix(0, 6) = "Cant GDgal"            '*
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
Detalle.ColWidth(1) = 2300 ' plano
Detalle.ColWidth(2) = 400
Detalle.ColWidth(3) = 2300 ' marca
Detalle.ColWidth(4) = 1200
Detalle.ColWidth(5) = 500
Detalle.ColWidth(6) = 500
Detalle.ColWidth(7) = 500
Detalle.ColWidth(8) = 600
Detalle.ColWidth(9) = 600
Detalle.ColWidth(10) = 650
Detalle.ColWidth(11) = 800
Detalle.ColWidth(12) = 650
Detalle.ColWidth(13) = 800

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

'itoGenerar

Dim n_Copias As Integer

Select Case Accion

Case "Agregando"

    RsITOpgc.Seek "=", m_TipoDoc, Numero.Text
    
    If RsITOpgc.NoMatch Then
    
        Campos_Enabled True
        Numero.Enabled = False
        Detalle.Enabled = False
        
        Fecha.Text = Format(Now, Fecha_Format)
        
'        GDespecial.SetFocus
        
        btnGrabar.Enabled = True
        btnContratistaSearch.visible = True
        btnOperadorSearch.visible = True
                
    Else
    
        Doc_Leer
        
        If m_TipoDoc = "P" Then
            If RsITOpgc!Tipo = "P" Then
                MsgBox "ITO PINTURA YA EXISTE"
            Else
'                MsgBox "ITO GALVANIZADO YA EXISTE"
                MsgBox "ITO REPROCESO YA EXISTE"
            End If
        Else
            If RsITOpgc!Tipo = "G" Then
'                MsgBox "ITO GALVANIZADO YA EXISTE"
                MsgBox "ITO REPROCESO YA EXISTE"
            Else
                MsgBox "ITO PINTURA YA EXISTE"
            End If
        End If
        Detalle.Enabled = False
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
        
    End If
    
Case "Modificando"

    RsITOpgc.Seek "=", m_TipoDoc, Numero.Text
    If RsITOpgc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        
        Doc_Leer
'        If rsitopgc!Tipo = "N" Then
            Campos_Enabled True
            Numero.Enabled = False
            
'            GDespecial.SetFocus

            btnContratistaSearch.visible = True
            btnOperadorSearch.visible = True
            
            btnGrabar.Enabled = True
'        Else
'            MsgBox "DEBE MODIFICAR COMO GUIA ESPECIAL"
'            Campos_Limpiar
'            Numero.Enabled = True
'            Numero.SetFocus
'        End If
    End If

Case "Eliminando"
    
    RsITOpgc.Seek "=", m_TipoDoc, Numero.Text
    If RsITOpgc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
    
        Doc_Leer
'        If rsitopgc!Tipo = "N" Then
            Numero.Enabled = False
            If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo) = vbYes Then
                Doc_Eliminar
                PlanoDetalle_Actualizar Dbm, Nv.Text, m_NvArea, RsITOpgd, "nv-plano-marca", "ito pyg", m_TipoDoc
            End If
'        Else
'            MsgBox "DEBE ELIMINAR COMO GUIA ESPECIAL"
'        End If
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
    End If

Case "Imprimiendo"

    RsITOpgc.Seek "=", m_TipoDoc, Numero.Text
    If RsITOpgc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
    
        Doc_Leer
        
        Numero.Enabled = False
            
        Detalle.visible = True
        Detalle.Enabled = True
            
    End If
    
End Select

End Sub
Private Sub Trabajadores_Poblar()

Dim sql As String

CbTrabajadores.Clear

If m_TipoDoc = "P" Or m_TipoDoc = "T" Then ' pintura
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
Fecha.Text = Format(RsITOpgc!Fecha, Fecha_Format)
m_Nv = RsITOpgc!Nv
'MsgBox "m_nv" & m_Nv

ContratistaRut.Text = NoNulo(RsITOpgc![Rut contratista])

' busca nv
RsNVc.Index = "Numero"

RsNVc.Seek "=", m_Nv, m_NvArea

mProyecto = NoNulo(RsNVc!Proyecto)
mTag = NoNulo(RsNVc!Tag)

cliRazon = "NV NO Existe"

If Not RsNVc.NoMatch Then

    ' busca nombre de cliente
    cliRut = RsNVc![RUT CLiente]
    RsCl.Seek "=", cliRut
    cliRazon = "Cliente NO Existe"
    If Not RsCl.NoMatch Then
        cliRazon = RsCl![Razon Social]
    End If

    ComboNV.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
    If False Then
        
        If m_TipoDoc = "P" And RsNVc!pintura Then
            ComboNV.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
        End If
        If m_TipoDoc = "G" And RsNVc!galvanizado Then
            ComboNV.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
        End If
        If m_TipoDoc = "R" Then
            ComboNV.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
        End If
        If m_TipoDoc = "T" Then
            ComboNV.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
        End If
        
    End If

End If

RsNVc.Index = Nv_Index ' "Número"

'Obs(0).Text = NoNulo(RsITOpgc![Observación 1])
'Obs(1).Text = NoNulo(RsITOpgc![Observación 2])
Obs(0).Text = NoNulo(RsITOpgc![Observacion 1])
Obs(1).Text = NoNulo(RsITOpgc![Observacion 2])


'DETALLE
RsPd.Index = "NV-Plano-Marca"

RsITOpgd.Seek "=", m_TipoDoc, Numero.Text, 1
If Not RsITOpgd.NoMatch Then

    Do While Not RsITOpgd.EOF
    
'        If m_TipoDoc = "P" Then
        If True Then
        
            If RsITOpgd!Numero = Numero.Text Then
            
                i = RsITOpgd!linea
                
                If i < 2 Then
                    OperadorRut.Text = NoNulo(RsITOpgd![RUT Operador])
                    'OperadorRut.Text = RsITOpgd![RUT Operador]
                    If Len(OperadorRut.Text) > 0 Then
                        RsTra.Seek "=", OperadorRut.Text
                        If Not RsTra.NoMatch Then
                            OperadorRazon.Text = RsTra![appaterno] & " " & RsTra![apmaterno] & " " & RsTra![nombres]
                        End If
                    End If
                End If
                
                Detalle.TextMatrix(i, 1) = RsITOpgd!Plano
                Detalle.TextMatrix(i, 2) = RsITOpgd!Rev
                Detalle.TextMatrix(i, 3) = RsITOpgd!Marca
                Detalle.TextMatrix(i, 7) = RsITOpgd!Cantidad
                
                RsPd.Seek "=", m_Nv, m_NvArea, RsITOpgd!Plano, RsITOpgd!Marca
                
                If Not RsPd.NoMatch Then
                
                    Detalle.TextMatrix(i, 4) = RsPd!Descripcion
                    Detalle.TextMatrix(i, 5) = RsPd![ITO fab]
                    Detalle.TextMatrix(i, 10) = RsPd![Peso]
                    
                    m_resta = IIf(Accion = "Modificando", RsITOpgd!Cantidad, 0)
                    Detalle.TextMatrix(i, 6) = RsPd![ITO pyg] - m_resta
                    
                    If Accion = "Modificando" And m_TipoDoc = "R" Then
                        Detalle.TextMatrix(i, 5) = RsPd![ITO fab]
                        Detalle.TextMatrix(i, 6) = RsPd![ito gr] - Val(Detalle.TextMatrix(i, 7))
                    End If
                    
                    If Accion = "Modificando" And m_TipoDoc = "T" Then
                        Detalle.TextMatrix(i, 5) = RsPd![ito gr]
                        Detalle.TextMatrix(i, 6) = RsPd![ITO pp] - Val(Detalle.TextMatrix(i, 7))
                    End If
                                    
                End If

                Detalle.TextMatrix(i, 8) = RsITOpgd![m2 Unitario]
                Detalle.TextMatrix(i, 12) = RsITOpgd![Precio Unitario]
                                
                If m_TipoDoc = "R" Then
                
                    Detalle.TextMatrix(i, 14) = NoNulo(RsITOpgd![tipo2]) '12
                    Detalle.TextMatrix(i, 15) = NoNulo(RsITOpgd![Maquina]) '13
                    
                    Detalle.TextMatrix(i, 16) = RsITOpgd![Turno] ' 14
                    
                    For j = 0 To 199
                        If a_Trabajadores(0, j) = RsITOpgd![RUT Operador] Then
                            Detalle.TextMatrix(i, 17) = a_Trabajadores(1, j) '15
                            Exit For
                        End If
                    Next
                
                End If
                
                If m_TipoDoc = "T" Then
                
                    Detalle.TextMatrix(i, 14) = RsITOpgd![manos1]
                    Detalle.TextMatrix(i, 15) = RsITOpgd![manos2]
                    
                    Detalle.TextMatrix(i, 16) = RsITOpgd![Turno]
                    
                    For j = 0 To 199
                        If a_Trabajadores(0, j) = RsITOpgd![RUT Operador] Then
                            Detalle.TextMatrix(i, 17) = a_Trabajadores(1, j)
                            Exit For
                        End If
                    Next
                
                End If
                
                Fila_Calcular_Normal i, False
                
            Else
            
                Exit Do
                
            End If
        
        End If
        
        RsITOpgd.MoveNext
        
    Loop
    
End If

RsPd.Index = "NV-Plano-Item"

ContratistaRazon.Text = Contratista_Lee(SqlRsSc, ContratistaRut.Text)

' operador
'OperadorRut.Text = ""

Detalle.Row = 1 ' para q' actualice la primera fila del detalle
Detalle_Sumar_Normal

End Sub
Private Function Doc_Validar() As Boolean
Dim porRecibir As Integer, m_Maquina As String
Doc_Validar = False

If conContratista Then

    If ContratistaRut.Text = "" Then
        MsgBox "DEBE ELEGIR CONTRATISTA"
    '    Rut.SetFocus
        btnContratistaSearch.SetFocus
        Exit Function
    End If

End If
If conOperador Then

    If OperadorRut.Text = "" Then
        MsgBox "DEBE ELEGIR OPERADOR"
    '    Rut.SetFocus
        btnOperadorSearch.SetFocus
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
        
            m_Maquina = Detalle.TextMatrix(i, 14)
            If m_Maquina = "" Then
                MsgBox "Maquina escoger Tipo de Granallado", , "ATENCIÓN"
                Detalle.Row = i
                Detalle.col = 14
                Detalle.SetFocus
                Exit Function
            End If

            m_Maquina = UCase(Detalle.TextMatrix(i, 15))
            If m_Maquina <> "A" And m_Maquina <> "M" And m_Maquina <> "E" Then
                MsgBox "Maquina debe ser A, M ó E", , "ATENCIÓN"
                Detalle.Row = i
                Detalle.col = 15
                Detalle.SetFocus
                Exit Function
            End If
                                                
        End If
        
        If m_TipoDoc = "R" Or m_TipoDoc = "T" Then
        
'            porRecibir = Val(Detalle.TextMatrix(i, 14))
            porRecibir = Val(Detalle.TextMatrix(i, 16))
            If 1 <= porRecibir And porRecibir <= 3 Then
            Else
                MsgBox "Turno debe ser 1, 2 ó 3", , "ATENCIÓN"
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
Dim m_PesoUnitario As Double, m_PesoTotal As Double
Dim qry As String, jj As Integer

m_PesoTotal = 0

If Nueva Then
    Numero.Text = Documento_Numero_Nuevo_PG(m_TipoDoc, RsITOpgc)
Else
    Doc_Detalle_Eliminar
End If

' DETALLE DE ITO

With RsITOpgd
j = 0
RsPd.Index = "NV-Plano-Marca"

For i = 1 To n_filas

    m_Plano = Trim(Detalle.TextMatrix(i, 1))
    m_cantidad = Val(Trim(Detalle.TextMatrix(i, 7)))

    If m_Plano <> "" Then
    If m_cantidad > 0 Then
    
        m_Marca = Detalle.TextMatrix(i, 3)
        m_cantidad = Val(Detalle.TextMatrix(i, 7))
        
        m_PesoUnitario = 0
        ' actualiza cantidad en planos
        RsPd.Seek "=", m_Nv, m_NvArea, m_Plano, m_Marca
        If RsPd.NoMatch Then
            ' no existe marca en el plano
        Else
        
            m_PesoUnitario = RsPd!Peso
            
            ' actualiza archivo detalle planos
            RsPd.Edit
'            If m_TipoDoc = "P" Or m_TipoDoc = "G" Then
'                RsPd![ito pyg] = RsPd![ito pyg] + m_cantidad
'            End If
            If m_TipoDoc = "R" Then
                RsPd![ito gr] = RsPd![ito gr] + m_cantidad
            End If
            If m_TipoDoc = "T" Then
                RsPd![ITO pp] = RsPd![ITO pp] + m_cantidad
            End If
            If m_TipoDoc = "P" Then
                RsPd![ITO pyg] = RsPd![ITO pyg] + m_cantidad
            End If
            
            RsPd.Update
            
        End If
        
        .AddNew
        !Numero = Numero.Text
        j = j + 1
        !linea = j
        
        !Tipo = m_TipoDoc
        
        !Nv = m_Nv
        !Fecha = Fecha.Text
        ![Rut contratista] = ContratistaRut.Text
        
        !Plano = m_Plano
        !Rev = Detalle.TextMatrix(i, 2)
        !Marca = m_Marca
        !Cantidad = m_cantidad
        ![Peso Unitario] = m_PesoUnitario
        ![m2 Unitario] = m_CDbl(Detalle.TextMatrix(i, 8))
        ![Precio Unitario] = m_CDbl(Detalle.TextMatrix(i, 12))


        If m_TipoDoc = "P" Or m_TipoDoc = "R" Or m_TipoDoc = "T" Then
            ' pintura o granallado o Produccion Pintura
            ![RUT Operador] = OperadorRut.Text
        End If


        If m_TipoDoc = "R" Then
        
            ![tipo2] = Detalle.TextMatrix(i, 14) '12
            ![Maquina] = UCase(Detalle.TextMatrix(i, 15)) '13
            
            ![Turno] = m_CDbl(Detalle.TextMatrix(i, 16)) '14
            
            If False Then
                For jj = 0 To 199
                    If a_Trabajadores(1, jj) = Detalle.TextMatrix(i, 17) Then ' 15
                        ![RUT Operador] = a_Trabajadores(0, jj)
                        Exit For
                    End If
                Next
            End If
            
        End If

        If m_TipoDoc = "T" Then ' ito produccion pintura
        
            ![manos1] = m_CDbl(Detalle.TextMatrix(i, 14))
            ![manos2] = m_CDbl(Detalle.TextMatrix(i, 15))
            
            ![Turno] = m_CDbl(Detalle.TextMatrix(i, 16))
            
            If False Then
                For jj = 0 To 199
                    If a_Trabajadores(1, jj) = Detalle.TextMatrix(i, 17) Then
                        ![RUT Operador] = a_Trabajadores(0, jj)
                        Exit For
                    End If
                Next
            End If
            
        End If

        .Update
        
        m_PesoTotal = m_PesoTotal + (m_cantidad * m_PesoUnitario)
        
    End If ' m_Cantidad>0
    End If ' m_plano<>""
    
Next
RsPd.Index = "NV-Plano-Item"
End With
    
save:
' CABECERA DE ITO
With RsITOpgc
If Nueva Then
    .AddNew
'    !Número = Numero.Text
    !Tipo = m_TipoDoc
    !Numero = Numero.Text
Else

'    Doc_Detalle_Eliminar
    
    .Edit
    
End If

!Fecha = Fecha.Text
!Nv = Val(m_Nv)
![Rut contratista] = ContratistaRut.Text
![m2 Total] = m_CDbl(Totalm2.Caption)
![Peso Total] = m_PesoTotal
![Precio Total] = Val(TotalPrecio.Caption)
![Observacion 1] = Obs(0).Text
![Observacion 2] = Obs(1).Text
.Update

End With

Select Case m_TipoDoc
Case "R"
    PlanoDetalle_Actualizar Dbm, Nv.Text, m_NvArea, RsITOpgd, "nv-plano-marca", "ito gr", m_TipoDoc
Case "T"
    PlanoDetalle_Actualizar Dbm, Nv.Text, m_NvArea, RsITOpgd, "nv-plano-marca", "ito pp", m_TipoDoc
'Case Else
Case "P"
    PlanoDetalle_Actualizar Dbm, Nv.Text, m_NvArea, RsITOpgd, "nv-plano-marca", "ito pyg", m_TipoDoc
End Select

Select Case Accion
Case "Agregando"
    Track_Registrar "ITO" & m_TipoDoc, Numero.Text, "AGR"
Case "Modificando"
    Track_Registrar "ITO" & m_TipoDoc, Numero.Text, "MOD"
End Select

MousePointer = vbDefault

End Sub
Private Sub Doc_Eliminar()

' elimina cabecera
RsITOpgc.Seek "=", m_TipoDoc, Numero.Text
If Not RsITOpgc.NoMatch Then
    
    RsITOpgc.Delete

End If

' elimina detalle
Doc_Detalle_Eliminar

End Sub
Private Sub Doc_Detalle_Eliminar()
' elimina detalle ITO
' al anular detalle ITO debe actualizar detalle plano

RsPd.Index = "NV-Plano-Marca"
RsITOpgd.Seek "=", m_TipoDoc, Numero.Text, 1
If Not RsITOpgd.NoMatch Then
    Do While Not RsITOpgd.EOF
        If RsITOpgd!Numero <> Numero.Text Then Exit Do
        RsPd.Seek "=", m_Nv, m_NvArea, RsITOpgd!Plano, RsITOpgd!Marca
        If Not RsPd.NoMatch Then
            RsPd.Edit
            If m_TipoDoc = "P" Or m_TipoDoc = "G" Then
                RsPd![ITO pyg] = RsPd![ITO pyg] - RsITOpgd!Cantidad
            End If
            If m_TipoDoc = "R" Then
                RsPd![ito gr] = RsPd![ito gr] - RsITOpgd!Cantidad
            End If
            RsPd.Update
        End If
    
        ' borra detalle
        RsITOpgd.Delete
    
        RsITOpgd.MoveNext
    Loop
End If
RsPd.Index = "NV-Plano-Item"


End Sub
Private Sub Campos_Limpiar()
Numero.Text = ""
'GDespecial.Value = 0
Fecha.Text = Fecha_Vacia
'Fecha.Text = Format(Now, Fecha_Format)
ComboNV.Text = " "
m_Nv = 0
Nv.Text = "" ' m_Nv
ContratistaRut.Text = ""
OperadorRut.Text = ""
ContratistaRazon.Text = ""
OperadorRazon.Text = ""
'Direccion.Text = ""
'Comuna.Text = ""
mProyecto = ""
mTag = ""
Detalle_Limpiar Detalle, n_columnas
Obs(0).Text = ""
Obs(1).Text = ""
Totalm2.Caption = ""
TotalKg.Caption = ""
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
    
    If False Then
    Accion = "Agregando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    Numero.Text = Documento_Numero_Nuevo_PG(m_TipoDoc, RsITOpgc)
    Numero.Enabled = True
    Numero.SetFocus
    End If
    
    Accion = "Agregando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    
    Campos_Enabled True
    Numero.Enabled = False
    Fecha.SetFocus
    btnGrabar.Enabled = True
    
    btnContratistaSearch.visible = True
    btnOperadorSearch.visible = True
    
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

        'If m_TipoDoc = "P" Or m_TipoDoc = "G" Then
        If m_TipoDoc = "T" Then
            If MsgBox("¿ IMPRIMIR ETIQUETAS ?", vbYesNo) = vbYes Then
                Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Etiquetas")
                'Etiquetas_Imprimir
                Select Case Nv.Text
                Case "3580"
                    ' flsmidth 07/01/2016
                    etiquetaImprimirV160107
                Case "3588"
                    ' flsmidth 09/02/2016
                    etiquetaImprimirV160209
                Case "3607"
                    ' flsmidth 22/02/2016
                    etiquetaImprimirV160222
                Case "3633"
                    ' flsmidth 23/02/2016
                    etiquetaImprimirV160223
                Case "3598"
                    ' flsmidth 26/02/2016
                    etiquetaImprimirV160226
                Case Else
                    etiquetaImprimirV1308
                End Select
    '            Impresora_Predeterminada "default"
            End If
        End If
        
        If cliRut = "96864470-K" Then ' TAKRAF
            If m_TipoDoc = "P" Then
                If MsgBox("¿ IMPRIMIR ETIQUETAS TAKRAF ?", vbYesNo) = vbYes Then
                    Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Etiquetas")
                    etiquetaImprimirTakraf_V1308
                End If
            End If
        End If

        If m_TipoDoc = "P" Then
            If MsgBox("¿ Exportr a EXCEL ?", vbYesNo) = vbYes Then
                excelExportar
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
            
            'If m_TipoDoc = "P" Or m_TipoDoc = "G" Then
            If m_TipoDoc = "T" Then
                If MsgBox("¿ IMPRIMIR ETIQUETAS ?", vbYesNo) = vbYes Then
                    Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Etiquetas")
                    etiquetaImprimirV1308
                End If
            End If
            
            Botones_Enabled 0, 0, 0, 0, 1, 0
            Campos_Limpiar
            Campos_Enabled False
            Numero.Enabled = True
            Numero.SetFocus
            
            'If Accion = "Agregando" Then Numero.Text = Documento_Numero_Nuevo_PG(m_TipoDoc, RsITOpgc)
            
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

btnTG.Enabled = Si

btnContratistaSearch.Enabled = Si
btnOperadorSearch.Enabled = Si

Nv.Enabled = Si
ComboNV.Enabled = Si

btnPendientes.Enabled = Si

txtBuscar.Enabled = Si
btnBuscar.Enabled = Si

Detalle.Enabled = Si
Obs(0).Enabled = Si
Obs(1).Enabled = Si

End Sub
Private Sub btnContratistaSearch_Click()

Dim arreglo(1) As String
arreglo(1) = "razon_social"

ComboPlano.visible = False
ComboMarca.visible = False

sql_Search.Muestra "personas", "RUT", arreglo(), Obj, Objs, "contratista='S'"
ContratistaRut.Text = sql_Search.Codigo
ContratistaRazon.Text = sql_Search.Descripcion

'Search.Muestra data_file, "Contratistas", "RUT", "Razon Social", "Contratista", "Contratistas", "Activo"
'Rut.Text = Search.codigo
'If Rut.Text <> "" Then
'    RsSc.Seek "=", Rut.Text
'    If RsSc.NoMatch Then
'        MsgBox "CONTRATISTA NO EXISTE"
'        Rut.SetFocus
'    Else
'        Razon.Text = Search.descripcion
'        If ComboNV.Text <> "" Then Detalle.Enabled = True
'    End If
'End If

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
If Edt.Text <> "M" And Edt <> "A" And Edt <> "E" Then
'Else
    MsgBox "Maquina debe ser A, M ó E", , "ATENCIÓN"
    Detalle.col = 15 '12
    Detalle.SetFocus
    Maquina_Validar = False
End If
End Function
Private Function Turno_Validar(Colu As Integer, Edt As Control) As Boolean
Turno_Validar = True
If 1 <= Val(Edt) And Val(Edt) <= 3 Then
Else
    MsgBox "Turno debe ser 1, 2 ó 3", , "ATENCIÓN"
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
prt.Print Tab(tab0 + tab40); "SEÑOR(ES) : " & ContratistaRazon,

prt.Print Tab(tab0); Empresa.Direccion;
prt.Print Tab(tab0 + tab40); "RUT       : " & ContratistaRut

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
Private Sub Etiquetas_ImprimirOLD()

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
            
            If False Then
                SetpYX -0.15, Margen_Izquierdo
                Printer.Print "Delgado"
                '//////////////////
                
            Else
            
                SetpYX -0.15, Margen_Izquierdo
                Printer.Print "Delgado";
                
                Printer.Font.Name = "Arial Black" ' oficial
    '            Printer.Font.Name = "Arial"
                Printer.Font.Bold = False
                Printer.Font.Size = 12 '14
                
                SetpYX 0.1, 8
                Printer.Print Fecha.Text
                
            End If
            
            Printer.Font.Name = "Arial Black" ' oficial
'            Printer.Font.Name = "Arial"
            Printer.Font.Bold = False
            Printer.Font.Size = 12 '14
            
            SetpYX li1, Margen_Izquierdo
            Printer.Print "Cliente: "; Left(cliRazon, 23)
            
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
Private Sub etiquetaImprimirV1208_OLD()

' imprime etiquetas version agosto 2012
' cambios en etiqueta debido a peticion FLSMIDTH
' se reemplaza codigo de barras (code128), por QR Code (codigo bidimensional)

Dim li As Integer, nCopias As Integer, Copia As Integer
Dim tab0 As Double, tab1 As Double, mDes As String, QrCodeText As String

tab0 = 0.5
tab1 = 4.5

Set prt = Printer
Set cQrCode = New ClsQrCode

If mProyecto = "" Then
'    mProyecto = Trim(Mid(ComboNV.Text, 8))
    mProyecto = "//"
End If
If mTag = "" Then
    mTag = "///"
End If

For li = 1 To n_filas

    nCopias = Val(Trim(Detalle.TextMatrix(li, 7))) ' Cant a Asignar
'    If n_Copias > 0 Then n_Copias = 1
    
    If nCopias > 0 Then
    
        For Copia = 1 To nCopias

            Prt_Ini
            
            m_obra = Trim(Mid(ComboNV.Text, 7)) 'Detalle.TextMatrix(li, 2)
            m_Plano = Detalle.TextMatrix(li, 1)
            m_Rev = Detalle.TextMatrix(li, 2)
            m_Marca = Detalle.TextMatrix(li, 3)
            mDes = Detalle.TextMatrix(li, 4)
            
            m_Peso = m_CDbl(Detalle.TextMatrix(li, 10))
                        
            '//////////////////////////////////////////////////
            '
            ' CODIGO QR
            '
            ' carga imagen generada desde la web
            ' NombreCliente/Proyecto/Tag/Plano/Descripcion/PesoUnitario/NV
            'QrCodeText = cliNombreFantasia & "/" & mProyecto & "/" & mTag & "/" & m_Plano & "/" & mDes & "/" & m_Peso & "/" & Nv.Text
            
            ' version 27/08/12
            ' proveedor/mina/sector/marca/Descripcion/PesoUnitario/NV/maestranza/nombresector/itemOrdenDeCompra
            ' proveedor/mina/sector => en tabla NV = proyecto
            ' maestranza/nombresector/itemOrdenDeCompra => en tabla NV = tag
            QrCodeText = mProyecto & "/" & m_Marca & "/" & mDes & "/" & m_Peso & "/" & Nv.Text & "/" & mTag
            
            'no esta establecida
            PictureQrCode.Picture = cQrCode.GetPictureQrCode(QrCodeText, PictureQrCode.ScaleWidth, PictureQrCode.ScaleHeight)
            If PictureQrCode.Picture Is Nothing Then MsgBox "SCP: Error al obtener Codigo QR"
            
            '////////////////////////////////////////////////////
            ' para servidor qrcode.com (que no estaba disponible los dias 05 y 06/09/12
            'Printer.PaintPicture PictureQrCode.Picture, tab0, 1, 3.7, 3.7 ' en cms
            '////////////////////////////////////////////////////
            ' para api by Goolge
            ' sAPI = "http://chart.apis.google.com/chart?cht=qr&chs=" & Width & "x" & Height & "&chl=" & GetSafeURL(Unicode2UTF8(sText)) & "&choe=" & Encoding & "&chld=" & ErrCorrectionLevel
            '                                            X1 , Y1, ancho1, alto1
            'Printer.PaintPicture PictureQrCode.Picture, -1, -1.8, 6.4, 9.5 ' ajustados segun lo que se imprime hasta 08/01/2013
            
            On Error GoTo ErrorImpresion
            ' x, y, ancho, alto
            Printer.PaintPicture PictureQrCode.Picture, -0.2, -0.7, 5.4, 8 ' ajustados segun lo que se imprime desde 08/01/2013
            On Error GoTo 0
            '////////////////////////////////////////////////////
            
            
            '////////////////////////
            '
            ' LOGO DELGADO
            '
            prt.Font.Name = "delgado"
            prt.Font.Size = 18
            prt.Font.Bold = True
            SetpYX 0.1, tab0
            prt.Print "Delgado"
            '////////////////////////
            
            '
            ' TEXTOS
            '
            
            prt.Font.Name = "Arial Black"
            prt.Font.Size = 12 '15
            prt.Font.Bold = False
            
            SetpYX 0.1, 8
            Printer.Print Fecha.Text
            
            SetpYX 1, tab1
            prt.Print cliNombreFantasia
            SetpYX 1.6, tab1
            prt.Print m_Marca & " " & mDes
            SetpYX 2.2, tab1
            prt.Print m_Peso; " Kgs"

            SetpYX 2.8, tab1
            If mProyecto <> "//" Then
                prt.Print mProyecto
            End If
            SetpYX 3.4, tab1
            If mTag <> "///" Then
                prt.Print mTag
            End If

            SetpYX 5, tab0
            prt.Print "NV " & Nv.Text & " " & m_obra
            
            ' Get the picture's dimensions in the printer's scale
            ' mode.
            'wid = ScaleX(Picture1.ScaleWidth, Picture1.ScaleMode, Printer.ScaleMode)
            'hgt = ScaleY(Picture1.ScaleHeight, Picture1.ScaleMode, Printer.ScaleMode)
            
            ' Draw the box.
            'Printer.Line (1440, 1440)-Step(wid, hgt), , B
            
            prt.EndDoc

        Next
    End If
Next

Exit Sub

ErrorImpresion:
MsgBox "SCP: Error al tratar de imprimir codigo QR"

End Sub
Private Sub etiquetaImprimirV1308()

' imprime etiquetas estandar (iguales para todos los clientes)

Dim li As Integer, nCopias As Integer, Copia As Integer
Dim tab0 As Double, tab1 As Double, mDes As String, QrCodeText As String
Dim pos As Integer, marca1 As String, marca2 As String

Dim paso As String

tab0 = 0.5
tab1 = 4.5

Set prt = Printer
Set cQrCode = New ClsQrCode

For li = 1 To n_filas

    nCopias = Val(Trim(Detalle.TextMatrix(li, 7))) ' Cant a Asignar
'    If n_Copias > 0 Then n_Copias = 1
    
    If nCopias > 0 Then
    
        For Copia = 1 To nCopias

            Prt_Ini
            
            m_obra = Trim(Mid(ComboNV.Text, 7)) 'Detalle.TextMatrix(li, 2)
            m_Plano = Detalle.TextMatrix(li, 1)
            m_Rev = Detalle.TextMatrix(li, 2)
            m_Marca = Detalle.TextMatrix(li, 3)
            mDes = Detalle.TextMatrix(li, 4)
            
            m_Peso = m_CDbl(Detalle.TextMatrix(li, 10))

            pos = InStrLast(m_Marca, "-")
            If pos > 0 Then
                marca1 = Left(m_Marca, pos - 1)
                marca2 = Mid(m_Marca, pos + 1)
            Else
                marca1 = ""
                marca2 = m_Marca
            End If
                        
            '//////////////////////////////////////////////////
            QrCodeText = mProyecto & "/" & m_Marca & "/" & mDes & "/" & m_Peso & "/" & Nv.Text & "/" & mTag
            'no esta establecida
            PictureQrCode.Picture = cQrCode.GetPictureQrCode(QrCodeText, PictureQrCode.ScaleWidth, PictureQrCode.ScaleHeight)
            If PictureQrCode.Picture Is Nothing Then MsgBox "SCP: Error al obtener Codigo QR"
            '////////////////////////////////////////////////////
            On Error GoTo ErrorImpresion
            Printer.PaintPicture PictureQrCode.Picture, -0.5, -1, 5.4, 8 ' ajustados segun lo que se imprime desde 08/01/2013
            On Error GoTo 0
            '////////////////////////////////////////////////////
            '
            ' LOGO DELGADO
            '
            
            ' OJO OJO
            If Nv.Text = 3036 Then
                paso = marca2
                marca2 = m_Plano
                m_Plano = paso
            End If
            
            '////////////////////////
            If True Then
                prt.Font.Name = "delgado"
                prt.Font.Size = 18
                prt.Font.Bold = True
                SetpYX 0.1, tab0
                prt.Print "Delgado"
            Else
                ' cargo la imagen
                'PictureL.Picture = LoadPicture("F:\scp_1308\imagenes\logoBN.jpg")
                PictureL.Picture = LoadPicture("E:\scp\logocical.jpg")
                ' x, y ,ancho, alto
                'Printer.PaintPicture PictureL.Picture, 0.5, 0.3, 4.1, 1.29
                'Printer.PaintPicture PictureL.Picture, 0.5, 0.3, 3.1, 0.98
                Printer.PaintPicture PictureL.Picture, 0.5, 0.3, 2.48, 0.784
            End If
            '////////////////////////
            '
            ' TEXTOS
            '
            prt.Font.Name = "Arial Black"
            prt.Font.Size = 12 '15
            prt.Font.Bold = False
            
            ' nuevo Rodrigo Nuñez 18/08/15
            SetpYX 0.1, 5
            Printer.Print Numero.Text
            
            SetpYX 0.1, 8
            Printer.Print Fecha.Text
            
            SetpYX 0.8, tab1
            prt.Print cliNombreFantasia
            
            SetpYX 1.4, tab1
            prt.Print m_Plano
            
            SetpYX 2, tab1
            prt.Print marca1

            prt.Font.Size = 24
            prt.Font.Bold = True
            SetpYX 2.5, tab1
            prt.Print marca2
            prt.Font.Size = 12
            prt.Font.Bold = False
            
            SetpYX 3.7, tab1
            prt.Print mDes
            
            SetpYX 4.3, tab1
            prt.Print m_Peso; " Kgs"

            SetpYX 5, tab0
            prt.Print "NV " & Nv.Text & " " & m_obra
                        
            prt.EndDoc

        Next
    End If
Next

Exit Sub

ErrorImpresion:
MsgBox "SCP: Error al tratar de imprimir codigo QR"

End Sub
Private Sub etiquetaImprimirV160107()

' imprime etiquetas estandar solo para NV 3580, cliente FLSmidth
' fecha 07/01/2016

Dim li As Integer, nCopias As Integer, Copia As Integer
Dim tab0 As Double, mDes As String
Dim nParte As String
Dim pos1 As Integer, pos2 As Integer

tab0 = 0.5

Set prt = Printer

For li = 1 To n_filas

    nCopias = Val(Trim(Detalle.TextMatrix(li, 7))) ' Cant a Asignar
'    If n_Copias > 0 Then n_Copias = 1
    
    If nCopias > 0 Then
    
        For Copia = 1 To nCopias

            Prt_Ini
            
            m_obra = Trim(Mid(ComboNV.Text, 7)) 'Detalle.TextMatrix(li, 2)
            m_Plano = Detalle.TextMatrix(li, 1)
            m_Rev = Detalle.TextMatrix(li, 2)
            
            m_Marca = Detalle.TextMatrix(li, 3)
            pos1 = InStr(1, m_Marca, "-")
            pos2 = InStr(pos1 + 1, m_Marca, "-")
            nParte = Mid(m_Marca, pos1 + 1, pos2 - pos1 - 1)
            
            mDes = Detalle.TextMatrix(li, 4)
            
            m_Peso = m_CDbl(Detalle.TextMatrix(li, 10))
           
            prt.Font.Name = "Arial Black"
            prt.Font.Size = 15
            prt.Font.Bold = True
            
            SetpYX 0, tab0
            prt.Print "FLSmidth"
                        
            SetpYX 0, 7
            Printer.Print Fecha.Text

            SetpYX 0.7, tab0
            prt.Print "OC 163842"
            
            SetpYX 1.4, tab0
            prt.Print mDes
            
            SetpYX 2.1, tab0
            prt.Print "Nº de Parte: " & nParte
                        
            SetpYX 2.8, tab0
            prt.Print "Cantidad: 1 de " & Detalle.TextMatrix(li, 7)
            
            SetpYX 3.5, tab0
            prt.Print m_Peso; " Kgs"

            SetpYX 4.2, tab0
            prt.Print "NV " & Nv.Text & " " & m_obra
                        
            prt.EndDoc

        Next
    End If
Next

Exit Sub

ErrorImpresion:
MsgBox "SCP: Error al tratar de imprimir codigo QR"

End Sub
Private Sub etiquetaImprimirV160209()

' imprime etiquetas estandar solo para NV 3588, cliente FLSmidth
' fecha 09/02/2016
' solicitado por Wilfredo Lopez

Dim li As Integer, nCopias As Integer, Copia As Integer
Dim tab0 As Double, mDes As String
Dim nParte As String
Dim pos1 As Integer, pos2 As Integer

tab0 = 0.5

Set prt = Printer

For li = 1 To n_filas

    nCopias = Val(Trim(Detalle.TextMatrix(li, 7))) ' Cant a Asignar
'    If n_Copias > 0 Then n_Copias = 1
    
    If nCopias > 0 Then
    
        For Copia = 1 To nCopias

            Prt_Ini
            
            m_obra = Trim(Mid(ComboNV.Text, 7)) 'Detalle.TextMatrix(li, 2)
            m_Plano = Detalle.TextMatrix(li, 1)
            m_Rev = Detalle.TextMatrix(li, 2)
            
            m_Marca = Detalle.TextMatrix(li, 3)
            pos1 = InStr(1, m_Marca, "-F")
            'pos2 = InStr(pos1 + 1, m_Marca, "-")
            nParte = Mid(m_Marca, pos1 + 1)
            
            mDes = Detalle.TextMatrix(li, 4)
            
            m_Peso = m_CDbl(Detalle.TextMatrix(li, 10))
           
            prt.Font.Name = "Arial Black"
            prt.Font.Size = 15
            prt.Font.Bold = True
            
            SetpYX 0, tab0
            prt.Print "FLSmidth"
                        
            SetpYX 0, 7
            Printer.Print Fecha.Text

            SetpYX 0.7, tab0
            prt.Print "OC 164180"
            
            SetpYX 1.4, tab0
            prt.Print mDes
            
            SetpYX 2.1, tab0
            prt.Print "Nº de Parte: " & nParte
                        
            SetpYX 2.8, tab0
            prt.Print "Cantidad: 1 de " & Detalle.TextMatrix(li, 7)
            
            SetpYX 3.5, tab0
            prt.Print m_Peso; " Kgs"

            SetpYX 4.2, tab0
            prt.Print "NV " & Nv.Text & " " & m_obra
                        
            prt.EndDoc

        Next
    End If
Next

Exit Sub

ErrorImpresion:
MsgBox "SCP: Error al tratar de imprimir codigo QR"

End Sub
Private Sub etiquetaImprimirV160222()

' imprime etiquetas especial solo para NV 3607, cliente FLSmidth
' fecha 22/02/2016
' solicitado por Wilfredo Lopez

Dim li As Integer, nCopias As Integer, Copia As Integer
Dim tab0 As Double, mDes As String
Dim nParte As String
Dim pos1 As Integer, pos2 As Integer

tab0 = 0.5

Set prt = Printer

For li = 1 To n_filas

    nCopias = Val(Trim(Detalle.TextMatrix(li, 7))) ' Cant a Asignar
'    If n_Copias > 0 Then n_Copias = 1
    
    If nCopias > 0 Then
    
        For Copia = 1 To nCopias

            Prt_Ini
            
            m_obra = Trim(Mid(ComboNV.Text, 7)) 'Detalle.TextMatrix(li, 2)
            m_Plano = Detalle.TextMatrix(li, 1)
            m_Rev = Detalle.TextMatrix(li, 2)
            
            m_Marca = Detalle.TextMatrix(li, 3)
            pos1 = InStr(1, m_Marca, "-1000")
            nParte = Mid(m_Marca, pos1 + 5)
            
            mDes = Detalle.TextMatrix(li, 4)
            
            m_Peso = m_CDbl(Detalle.TextMatrix(li, 10))
           
            prt.Font.Name = "Arial Black"
            prt.Font.Size = 15
            prt.Font.Bold = True
            
            SetpYX 0, tab0
            prt.Print "FLSmidth"
                        
            SetpYX 0, 7
            Printer.Print Fecha.Text

            SetpYX 0.7, tab0
            prt.Print "OC 164623"
            
            SetpYX 1.4, tab0
            prt.Print mDes
            
            SetpYX 2.1, tab0
            prt.Print "Nº Parte: " & nParte
                        
            SetpYX 2.8, tab0
            prt.Print "Cantidad: 1 de " & Detalle.TextMatrix(li, 7)
            
            SetpYX 3.5, tab0
            prt.Print m_Peso; " Kgs"

            SetpYX 4.2, tab0
            prt.Print "NV " & Nv.Text & " " & m_obra
                        
            prt.EndDoc

        Next
    End If
Next

Exit Sub

ErrorImpresion:
MsgBox "SCP: Error al tratar de imprimir codigo QR"

End Sub
Private Sub etiquetaImprimirV160223()

' imprime etiquetas especial solo para NV 3633, cliente FLSmidth
' fecha 23/02/2016
' solicitado por Wilfredo Lopez

Dim li As Integer, nCopias As Integer, Copia As Integer
Dim tab0 As Double, mDes As String
Dim nParte As String
Dim pos1 As Integer, pos2 As Integer

tab0 = 0.5

Set prt = Printer

For li = 1 To n_filas

    nCopias = Val(Trim(Detalle.TextMatrix(li, 7))) ' Cant a Asignar
'    If n_Copias > 0 Then n_Copias = 1
    
    If nCopias > 0 Then
    
        For Copia = 1 To nCopias

            Prt_Ini
            
            m_obra = Trim(Mid(ComboNV.Text, 7)) 'Detalle.TextMatrix(li, 2)
            m_Plano = Detalle.TextMatrix(li, 1)
            m_Rev = Detalle.TextMatrix(li, 2)
            
            m_Marca = Detalle.TextMatrix(li, 3)
            nParte = m_Marca ' marca tal cual
            
            mDes = Detalle.TextMatrix(li, 4)
            
            m_Peso = m_CDbl(Detalle.TextMatrix(li, 10))
           
            prt.Font.Name = "Arial Black"
            prt.Font.Size = 15
            prt.Font.Bold = True
            
            SetpYX 0, tab0
            prt.Print "FLSmidth"
                        
            SetpYX 0, 7
            Printer.Print Fecha.Text

            SetpYX 0.7, tab0
            prt.Print "OC 164623"
            
            SetpYX 1.4, tab0
            prt.Print mDes
            
            SetpYX 2.1, tab0
            prt.Print "Nº Parte: " & nParte
                        
            SetpYX 2.8, tab0
            prt.Print "Cantidad: 1 de " & Detalle.TextMatrix(li, 7)
            
            SetpYX 3.5, tab0
            prt.Print m_Peso; " Kgs"

            SetpYX 4.2, tab0
            prt.Print "NV " & Nv.Text & " " & m_obra
                        
            prt.EndDoc

        Next
    End If
Next

Exit Sub

ErrorImpresion:
MsgBox "SCP: Error al tratar de imprimir codigo QR"

End Sub
Private Sub etiquetaImprimirV160226()

' imprime etiquetas especial solo para NV 3598, cliente FLSmidth
' fecha 26/02/2016
' solicitado por Wilfredo Lopez

Dim li As Integer, nCopias As Integer, Copia As Integer
Dim tab0 As Double, mDes As String
Dim nParte As String
Dim pos1 As Integer, pos2 As Integer

tab0 = 0.5

Set prt = Printer

For li = 1 To n_filas

    nCopias = Val(Trim(Detalle.TextMatrix(li, 7))) ' Cant a Asignar
    
    If nCopias > 0 Then
    
        For Copia = 1 To nCopias

            Prt_Ini
            
            m_obra = Trim(Mid(ComboNV.Text, 7)) 'Detalle.TextMatrix(li, 2)
            m_Plano = Detalle.TextMatrix(li, 1)
            m_Rev = Detalle.TextMatrix(li, 2)
            
            m_Marca = Detalle.TextMatrix(li, 3)
            pos1 = InStr(1, m_Marca, "-") ' posicion primer guion
            pos2 = InStr(pos1 + 1, m_Marca, "-") ' posicion segundo guion
            nParte = Mid(m_Marca, pos1 + 1, pos2 - pos1 - 1)
            
            mDes = Detalle.TextMatrix(li, 4)
            
            m_Peso = m_CDbl(Detalle.TextMatrix(li, 10))
           
            prt.Font.Name = "Arial Black"
            prt.Font.Size = 15
            prt.Font.Bold = True
            
            SetpYX 0, tab0
            prt.Print "FLSmidth"
                        
            SetpYX 0, 7
            Printer.Print Fecha.Text

            SetpYX 0.7, tab0
            prt.Print "OC 163447"
            
            SetpYX 1.4, tab0
            prt.Print mDes
            
            SetpYX 2.1, tab0
            prt.Print "Nº Parte: " & nParte
                        
            SetpYX 2.8, tab0
            prt.Print "Cantidad: 1 de " & Detalle.TextMatrix(li, 7)
            
            SetpYX 3.5, tab0
            prt.Print m_Peso; " Kgs"

            SetpYX 4.2, tab0
            prt.Print "NV " & Nv.Text & " " & m_obra
                        
            prt.EndDoc

        Next
    End If
Next

Exit Sub

ErrorImpresion:
MsgBox "SCP: Error al tratar de imprimir"

End Sub
Private Sub etiquetaImprimirTakraf_V1308()

' imprime etiquetas version agosto 2013
' cambios en etiqueta para cliente TRAKRAF

Dim li As Integer, nCopias As Integer, Copia As Integer
Dim tab1 As Double, mDes As String, QrCodeText As String

Dim vt As Double ' vertical tab
Dim incrementoVertical As Double

Dim Texto(9) As String

Dim wid As Integer, hgt As Integer
Dim pos1 As Integer, pos2 As Integer

incrementoVertical = 0.46

tab1 = 4.3

Set prt = Printer
Set cQrCode = New ClsQrCode

For li = 1 To n_filas

    nCopias = Val(Trim(Detalle.TextMatrix(li, 7))) ' Cant a Asignar
'    If n_Copias > 0 Then n_Copias = 1
    
    If nCopias > 0 Then
    
        For Copia = 1 To nCopias

            Prt_Ini
            
            m_obra = Trim(Mid(ComboNV.Text, 7)) 'Detalle.TextMatrix(li, 2)
            m_Plano = Detalle.TextMatrix(li, 1)
            m_Rev = Detalle.TextMatrix(li, 2)
            m_Marca = Detalle.TextMatrix(li, 3)
            mDes = Detalle.TextMatrix(li, 4)
            
            m_Peso = m_CDbl(Detalle.TextMatrix(li, 10))
            
            ' busca oc
            ' xxxx-OC-yyyyyy etc
            ' en la NV 3110, es el segundo grupo de caracteres entre guines
            pos1 = InStr(1, m_Marca, "-")
            If pos1 > 0 Then
                pos2 = InStr(pos1 + 1, m_Marca, "-")
            End If
            
            Texto(8) = "?"
            If pos1 > 0 And pos2 > pos1 Then
                Texto(8) = Mid(m_Marca, pos1 + 1, pos2 - pos1 - 1)
            End If
            
            If True Then
                Texto(1) = cliNombreFantasia
                Texto(2) = "AMPLIACION BOTADERO"
                Texto(3) = "4501296511"
                Texto(4) = mDes
                Texto(5) = m_Peso
                Texto(6) = "BODEGA PROYECTO R.T."
                Texto(7) = m_Marca
                'Texto(8) = "1"
                Texto(9) = m_obra
            Else
                Texto(1) = "CLIENTE: " & cliNombreFantasia
                Texto(2) = "PROYECTO: AMPLIACION BOTADERO"
                Texto(3) = "ORDEN DE COMPRA: 4501296511"
                Texto(4) = "DESCRIPCION: " & mDes
                Texto(5) = "PESO: " & m_Peso
                Texto(6) = "DESTINO: BODEGA PROYECTO R.T."
                Texto(7) = "MARCA: " & m_Marca
                Texto(8) = "ITEM OC: " & Texto(8)
                Texto(9) = "OBRA: " & Nv.Text & " " & m_obra
            End If
            
            '//////////////////////////////////////////////////
            ' CODIGO QR
            QrCodeText = Texto(1) & Texto(2) & Texto(3) & Texto(4) & Texto(5) & Texto(6) & Texto(7) & Texto(8) & Texto(9)
            QrCodeText = Texto(1) & vbLf & Texto(2) & vbLf & Texto(3) & vbLf & Texto(4) & vbLf & Texto(5) & vbLf & Texto(6) & vbLf & Texto(7) & vbLf & Texto(8) & vbLf & Texto(9)
            
            'QrCodeText = Mid(QrCodeText, 1, 120) ' qr ok
            'QrCodeText = Mid(QrCodeText, 1, 130) ' qr ok
            'QrCodeText = Mid(QrCodeText, 1, 132) ' qr ok
            QrCodeText = Mid(QrCodeText, 1, 133) ' qr ok
            'QrCodeText = Mid(QrCodeText, 1, 134) ' qr NO FUNCIONA
            
            'MsgBox Len(QrCodeText)
            
            'no esta establecida
            PictureQrCode.Picture = cQrCode.GetPictureQrCode(QrCodeText, PictureQrCode.ScaleWidth, PictureQrCode.ScaleHeight)
            If PictureQrCode.Picture Is Nothing Then MsgBox "SCP: Error al obtener Codigo QR"
            
            'wid = ScaleX(PictureQrCode.ScaleWidth, PictureQrCode.ScaleMode, Printer.ScaleMode)
            'hgt = ScaleY(PictureQrCode.ScaleHeight, PictureQrCode.ScaleMode, Printer.ScaleMode)
            
            'Debug.Print wid, hgt
            
            On Error GoTo ErrorImpresion
            '                                           x1  , y1, ancho, alto
            'Printer.PaintPicture PictureQrCode.Picture, -0.6, -1.5, 5.4, 8 ' ajustados segun lo que se imprime desde 22/08/2013
            ' ok
            'Printer.PaintPicture PictureQrCode.Picture, -0.4, -1, 4.86, 7.2 ' ajustados segun lo que se imprime desde 22/08/2013
            
            'tab1 = 3.5
            'Printer.PaintPicture PictureQrCode.Picture, -0.3, -0.6, 4.05, 6 ' ajustados segun lo que se imprime desde 22/08/2013
            
            tab1 = 3.9 ' 20/11/13
            Printer.PaintPicture PictureQrCode.Picture, -0.3, -0.6, 4.45, 6.6 ' ajustados segun lo que se imprime desde 20/11/2013
            
            On Error GoTo 0
            '////////////////////////////////////////////////////

            ' TEXTOS
            
            prt.Font.Name = "Arial Black"
            prt.Font.Size = 8
            prt.Font.Bold = False

            vt = 0.2

            If True Then
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print "CLIENTE: " & Texto(1)
    
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print "PROYECTO: " & Texto(2)
    
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print "ORDEN DE COMPRA: " & Texto(3)
    
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print "DESCRIPCION: " & Texto(4)
    
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print "PESO: " & Texto(5)
    
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print "DESTINO: " & Texto(6)
    
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print "MARCA: " & Texto(7)
    
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print "ITEM OC: " & Texto(8)
    
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print "OBRA: " & Texto(9)
                
            Else
                
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print Texto(1)
    
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print Texto(2)
    
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print Texto(3)
    
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print Texto(4)
    
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print Texto(5)
    
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print Texto(6)
    
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print Texto(7)
    
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print Texto(8)
    
                vt = vt + incrementoVertical
                SetpYX vt, tab1
                prt.Print Texto(9)

            End If

            prt.EndDoc

        Next
    End If
Next

Exit Sub

ErrorImpresion:
MsgBox "SCP: Error al tratar de imprimir codigo QR"

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

Cr.WindowTitle = "ITO Nº " & Numero_Ito
Cr.ReportSource = crptReport
Cr.WindowState = crptMaximized
'Cr.WindowBorderStyle = crptFixedSingle
Cr.WindowMaxButton = False
Cr.WindowMinButton = False
Cr.Formulas(0) = "RAZON SOCIAL=""" & EmpOC.Razon & """"
Cr.Formulas(1) = "GIRO=""" & "GIRO: " & EmpOC.Giro & """"
Cr.Formulas(2) = "DIRECCION=""" & EmpOC.Direccion & """"
Cr.Formulas(3) = "TELEFONOS=""" & "Teléfono: " & EmpOC.Telefono1 & " " & EmpOC.Comuna & """"
Cr.Formulas(4) = "RUT=""" & "RUT: " & EmpOC.Rut & """"

'MsgBox Certificado.Value

Cr.DataFiles(0) = repo_file & ".MDB"

'If Tipo = "E" Then
'    CR.Formulas(5) = "COTIZACION=""" & "GUÍA DESP. Nº:" & """"
'    CR.ReportFileName = Drive_Server & Path_Server & EmpOC.Fantasia & "\Oc_Especial.Rpt"
'Else
'    CR.Formulas(5) = "COTIZACION=""" & "COTIZACIÓN Nº:" & """"
    Cr.ReportFileName = Drive_Server & Path_Rpt & "itopg_legal.rpt"
'End If

Cr.Action = 1

End Sub
Private Sub Privilegios()
If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 0, 0
End If
End Sub
Private Sub itoGenerar()
' genera ito en forma automatica
' con todas las piezas que esten pendientes
' se hace por tipo de ito

Dim RsPd As Recordset, Nv As Double

Nv = 2725 '6

Set RsPd = Dbm.OpenRecordset("SELECT * FROM [planos detalle] WHERE nv=" & Nv & " ORDER BY plano,marca")

With RsPd
Do While Not .EOF
'    If ![ito fab] > ![ito gr] Then ' pendiente granallado
    If ![ito gr] > ![ITO pp] Then ' pendiente produccion pintura
        Debug.Print !Nv, !Plano, ![ITO fab], ![ito gr], ![ITO pp]
    End If
    .MoveNext
Loop
End With

End Sub
Private Sub duplicados()
' busca elementos duplicados detro de una misma ito, gr y pp

Dim Numero As Integer, Plano As String, Marca As String, Tipo As String
Tipo = "T"

Set RsPd = Dbm.OpenRecordset("SELECT * FROM [ITO pg detalle] WHERE tipo='" & Tipo & "' ORDER BY numero,plano,marca")
Numero = 0

With RsPd
Do While Not .EOF
    If Numero = !Numero And Plano = !Plano And Marca = !Marca Then
        Debug.Print !Numero, !Fecha, !Plano, !Marca
        Dbm.Execute "DELETE FROM [ITO pg detalle] WHERE tipo='" & Tipo & "' AND numero=" & Numero & " AND plano='" & Plano & "' AND marca='" & Marca & "'"
    Else
        Numero = !Numero
        Plano = !Plano
        Marca = !Marca
    End If
    .MoveNext
Loop
End With

End Sub
Private Sub busca(Nv As Integer)
' busca ito pintura que tengan pocas lineas y pocos componentes
' de una nv
Dim Rs As Recordset
Set Rs = Dbm.OpenRecordset("SELECT * FROM [ito pg detalle] WHERE tipo='P' AND nv=" & Nv & " AND cantidad=1") '" ORDER BY ")
With Rs
Do While Not .EOF
    Debug.Print !Numero
    .MoveNext
Loop
.Close
End With

End Sub
Private Sub excelExportar()
' exporta ito pin actual a archivo CSV

Dim fi As Integer
Dim Archivo As String
Dim sep As String
Dim linea As String
sep = ";"

Archivo = "itoPin-" & Numero.Text & ".csv"

Open Archivo For Output As #1

Print #1, ComboNV.Text
Print #1, ""
Print #1, "ITOP Nº " & sep & Numero.Text & sep & "Fecha" & sep & Fecha.Text
Print #1, ""

For fi = 1 To n_filas
    
    linea = Detalle.TextMatrix(fi, 1) & sep & Detalle.TextMatrix(fi, 2) & sep & Detalle.TextMatrix(fi, 3) & sep & Detalle.TextMatrix(fi, 4) & sep & Detalle.TextMatrix(fi, 7) & sep & Detalle.TextMatrix(fi, 8) & sep & Detalle.TextMatrix(fi, 10)
    Print #1, linea
    
Next
Close #1

MsgBox "Archivo Generado"

End Sub
