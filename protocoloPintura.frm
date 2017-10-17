VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form protocoloPintura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "protocolo Pintura"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameProtocolo 
      Caption         =   "Datos Protocolo Pintura"
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   9255
      Begin VB.TextBox responsableNombre 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1320
         TabIndex        =   34
         Top             =   320
         Width           =   2895
      End
      Begin VB.CommandButton btnSearch 
         Height          =   300
         Left            =   4320
         Picture         =   "protocoloPintura.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   300
         Width           =   300
      End
      Begin VB.Frame FrameProbeta 
         Caption         =   "Proceso c/probeta previa"
         Height          =   615
         Left            =   2520
         TabIndex        =   30
         Top             =   720
         Width           =   2175
         Begin VB.OptionButton OpProbeta 
            Caption         =   "SI"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   32
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton OpProbeta 
            Caption         =   "NO"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   31
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame FrameGranalla 
         Caption         =   "Granalla Mezclada"
         Height          =   615
         Left            =   4800
         TabIndex        =   27
         Top             =   720
         Width           =   1695
         Begin VB.OptionButton OpGranalla 
            Caption         =   "SI"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton OpGranalla 
            Caption         =   "NO"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   28
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame FrameEsquema 
         Caption         =   "Esquema Solicitado"
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   2295
         Begin VB.TextBox esquema 
            Height          =   300
            Left            =   120
            MaxLength       =   20
            TabIndex        =   26
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame FrameCalibre 
         Caption         =   "Calibre/Tipo"
         Height          =   615
         Left            =   6600
         TabIndex        =   23
         Top             =   720
         Width           =   2535
         Begin VB.TextBox calibre 
            Height          =   300
            Left            =   240
            MaxLength       =   20
            TabIndex        =   24
            Top             =   240
            Width           =   2175
         End
      End
      Begin MSMask.MaskEdBox paginaTotal 
         Height          =   300
         Left            =   8640
         TabIndex        =   17
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox paginaNumero 
         Height          =   300
         Left            =   7920
         TabIndex        =   18
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox numeroProtocolo 
         Height          =   300
         Left            =   6120
         TabIndex        =   19
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   "_"
      End
      Begin VB.Label lblResponsable 
         Caption         =   "Responsable"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblNumeroProtocolo 
         Caption         =   "Nº Protocolo"
         Height          =   255
         Left            =   5160
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblPagina 
         Caption         =   "Página            de"
         Height          =   255
         Left            =   7320
         TabIndex        =   20
         Top             =   360
         Width           =   1455
      End
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
   Begin Crystal.CrystalReport Cr 
      Left            =   8640
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.TextBox Nv 
      Height          =   300
      Left            =   840
      TabIndex        =   5
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   1
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   15
      Top             =   6120
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   0
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   14
      Top             =   5760
      Width           =   5000
   End
   Begin VB.Frame Frame_Contratista 
      Caption         =   "Contratista"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   4320
      TabIndex        =   7
      Top             =   480
      Width           =   5055
      Begin VB.TextBox Razon 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   11
         Top             =   360
         Width           =   2895
      End
      Begin MSMask.MaskEdBox Rut 
         Height          =   300
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label lblRut 
         Caption         =   "RUT"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblSenores 
         Caption         =   "SEÑOR(ES)"
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   120
         Width           =   975
      End
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   3240
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
      Left            =   1200
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
      Height          =   2805
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4948
      _Version        =   327680
      Enabled         =   0   'False
      ScrollBars      =   2
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   8520
      Top             =   3120
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
            Picture         =   "protocoloPintura.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "protocoloPintura.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "protocoloPintura.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "protocoloPintura.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "protocoloPintura.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "protocoloPintura.frx":065C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "protocoloPintura.frx":076E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "Observación"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   13
      Top             =   5760
      Width           =   975
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
      Left            =   2640
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "&Nº (ITOp)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Width           =   1335
   End
End
Attribute VB_Name = "protocoloPintura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' protocolo de pintura
' llevan el mismo numero correlativo que las ITOp
' solo una tabla, de cabecera, no lleva detalle (pues lo extrae de ITOp)
Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button
Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, k As Integer, d As Variant

Private DbD As Database, RsCl As Recordset, RsTra As Recordset
Private Dbm As Database, RsNVc As Recordset
Private RsNvPla As Recordset, RsPd As Recordset, RsPdBuscar As Recordset
Private RsITOpgc As Recordset, RsITOpgd As Recordset
Private RsPp As Recordset ' protocolo pintura cabecera

Private TipoDoc As String

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
Private cliRazon As String, cliNombreFantasia As String
'Private a_Trabajadores(1, 199) As String, m_Nombre As String
Private a_TipoGranalla(9) As String, m_TotalTiposGranalla As Integer
Private trabajadorRut As String
' arreglo para un solo trabajador, con sus datos
'Private aTrabajador(9) As String
Private Sub Form_Load()

TipoDoc = "P"
n_filas = 30

' abre archivos
Set DbD = OpenDatabase(data_file)

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

Set RsPp = Dbm.OpenRecordset("protocoloPintura")
RsPp.Index = "numero"

Set RsNvPla = Dbm.OpenRecordset("Planos Cabecera")
RsNvPla.Index = "NV-Plano"

Set RsPd = Dbm.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Item"

Inicializa

n_columnas = 9

Detalle_Config

'Privilegios

m_NvArea = 0

' campos simpre disabled
Nv.Enabled = False
Rut.Enabled = False
Razon.Enabled = False
responsableNombre = False
Obs(0).Enabled = False
Obs(1).Enabled = False

End Sub
Private Sub Inicializa()

Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnDesHacer = Toolbar.Buttons(6)
Set btnGrabar = Toolbar.Buttons(7)

Accion = ""

Campos_Enabled False

Obj = "Protocolo de Pintura"
Objs = "Protocolos de Pintura"

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
Detalle.TextMatrix(0, 5) = "Cant itoP"  ' *
Detalle.TextMatrix(0, 6) = "m2 Uni"         '*
Detalle.TextMatrix(0, 7) = "m2 Tot"            '*
Detalle.TextMatrix(0, 8) = "Peso Uni"
Detalle.TextMatrix(0, 9) = "Peso Total"         '*
'Detalle.TextMatrix(0, 10) = "Precio Uni"
'Detalle.TextMatrix(0, 11) = "Precio Total"         '*

Detalle.ColWidth(0) = 300
Detalle.ColWidth(1) = 1800 '2000 ' plano
Detalle.ColWidth(2) = 400
Detalle.ColWidth(3) = 2000 '2200 ' marca
Detalle.ColWidth(4) = 2200 ' descripcion
Detalle.ColWidth(5) = 500
Detalle.ColWidth(6) = 500
Detalle.ColWidth(7) = 500
Detalle.ColWidth(8) = 600
Detalle.ColWidth(9) = 600
'Detalle.ColWidth(10) = 650
'Detalle.ColWidth(11) = 800

'Detalle.ColWidth(16) = 0 ' peso unitario

ancho = 350 ' con scroll vertical
'ancho = 100 ' sin scroll vertical

'Totalm2.Width = Detalle.ColWidth(11)
For i = 0 To n_columnas
    'If i = 9 Then Totalm2.Left = ancho + Detalle.Left - 350
    'If i = 11 Then TotalKg.Left = ancho + Detalle.Left - 350
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
      
    Detalle.col = 8
    Detalle.CellForeColor = vbRed
    Detalle.col = 9
    Detalle.CellForeColor = vbRed
    
Next

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

    ' primero busco ITOp
    RsITOpgc.Seek "=", TipoDoc, Numero.Text
    
    If RsITOpgc.NoMatch Then
    
        MsgBox "ITO Pintura Nº " & Numero.Text & " NO existe"
        
    Else
        
        ITOpLeer Numero.Text
        
        RsPp.Seek "=", Numero.Text
        
        camposITOpEnabled False
        
        If RsPp.NoMatch Then
    
            camposProtocoloEnabled True
            Numero.Enabled = False
    
            Fecha.Text = Format(Now, Fecha_Format)
            Fecha.SetFocus
    
            btnGrabar.Enabled = True
            
            Doc_Leer
    
        Else
        
            Doc_Leer
            
            MsgBox "PROTOCOLO PINTURA YA EXISTE"
            
            Detalle.Enabled = False
            Campos_Limpiar
            Numero.Enabled = True
            Numero.SetFocus
            
        End If
    
    End If
    
Case "Modificando"

    ' primero busco ITOp
    RsITOpgc.Seek "=", TipoDoc, Numero.Text
    If RsITOpgc.NoMatch Then
        MsgBox "ITO Pintura Nº " & Numero.Text & " NO existe"
    Else
        
        ITOpLeer Numero.Text
        
        RsPp.Seek "=", Numero.Text
        
        camposITOpEnabled False
        
        If RsPp.NoMatch Then
        
            MsgBox "PROTOCOLO PINTURA NO EXISTE"
                    
        Else
        
            Doc_Leer
                       
            Campos_Enabled True
            
            btnGrabar.Enabled = True
          
            Numero.Enabled = False
                                   
        End If
    
    End If

Case "Eliminando"
    
    ' primero busco ITOp
    RsITOpgc.Seek "=", TipoDoc, Numero.Text
    If RsITOpgc.NoMatch Then
        MsgBox "ITO Pintura Nº " & Numero.Text & " NO existe"
    Else
        
        ITOpLeer Numero.Text
        
        RsPp.Seek "=", Numero.Text
        
        camposITOpEnabled False
        
        If RsPp.NoMatch Then
        
            MsgBox "PROTOCOLO PINTURA NO EXISTE"
                    
        Else
        
            Doc_Leer
            
            Numero.Enabled = False
            If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo) = vbYes Then
                Doc_Eliminar
            End If
                                                          
            Campos_Limpiar
            Numero.Enabled = True
            Numero.SetFocus
                                   
        End If
    
    End If

Case "Imprimiendo"
    
    ' primero busco ITOp
    RsITOpgc.Seek "=", TipoDoc, Numero.Text
    If RsITOpgc.NoMatch Then
        MsgBox "ITO Pintura Nº " & Numero.Text & " NO existe"
    Else
        
        ITOpLeer Numero.Text
        
        RsPp.Seek "=", Numero.Text
        
        camposITOpEnabled False
        
        If RsPp.NoMatch Then
        
            MsgBox "PROTOCOLO PINTURA NO EXISTE"
                    
        Else
        
            Doc_Leer
                       
            Numero.Enabled = False
            
            Detalle.visible = True
            Detalle.Enabled = True
                                   
        End If
    
    End If
    
End Select

End Sub
Private Sub Doc_Leer()

' lee solo protocolo pintura

ITOpLeer Numero.Text

protocoloLeer Numero.Text

End Sub
Private Sub ITOpLeer(Numero As Long)
Dim m_resta As Integer

' CABECERA
RsITOpgc.Seek "=", TipoDoc, Numero

If RsITOpgc.NoMatch Then

Else

    m_Nv = RsITOpgc!Nv
    'MsgBox "m_nv" & m_Nv
    
    Rut.Text = NoNulo(RsITOpgc![RUT Contratista])
    
    ' busca nv
    RsNVc.Index = "Numero"
    
    RsNVc.Seek "=", m_Nv, m_NvArea
    
    cliRazon = "NV NO Existe"
    
    If Not RsNVc.NoMatch Then
    
        'obra
        Nv.Text = m_Nv & " " & RsNVc!obra
    
        ' busca nombre de cliente
        RsCl.Seek "=", RsNVc![Rut cliente]
        cliRazon = "Cliente NO Existe"
        If Not RsCl.NoMatch Then
            cliRazon = RsCl![razon social]
        End If
        
    End If
    
    RsNVc.Index = Nv_Index ' "Número"
    
    Obs(0).Text = NoNulo(RsITOpgc![Observacion 1])
    Obs(1).Text = NoNulo(RsITOpgc![Observacion 2])

End If

'DETALLE
RsPd.Index = "NV-Plano-Marca"

RsITOpgd.Seek "=", TipoDoc, Numero, 1
If Not RsITOpgd.NoMatch Then

    Do While Not RsITOpgd.EOF
    
'        If m_TipoDoc = "P" Then
        If True Then
        
            If RsITOpgd!Numero = Numero Then
            
                i = RsITOpgd!linea
                
                Detalle.TextMatrix(i, 1) = RsITOpgd!Plano
                Detalle.TextMatrix(i, 2) = RsITOpgd!Rev
                Detalle.TextMatrix(i, 3) = RsITOpgd!Marca
                Detalle.TextMatrix(i, 5) = RsITOpgd!Cantidad
                
                RsPd.Seek "=", m_Nv, m_NvArea, RsITOpgd!Plano, RsITOpgd!Marca
                
                If Not RsPd.NoMatch Then
                
                    Detalle.TextMatrix(i, 4) = RsPd!Descripcion
                    Detalle.TextMatrix(i, 6) = RsPd![Superficie]
                    Detalle.TextMatrix(i, 8) = RsPd![Peso]
                                    
                End If

                Detalle.TextMatrix(i, 7) = RsITOpgd!Cantidad * RsPd![Superficie]
                Detalle.TextMatrix(i, 9) = RsITOpgd!Cantidad * RsPd![Peso]
                                                
            Else
            
                Exit Do
                
            End If
        
        End If
        
        RsITOpgd.MoveNext
        
    Loop
    
End If

RsPd.Index = "NV-Plano-Item"

Razon.Text = Contratista_Lee(SqlRsSc, Rut.Text)

Detalle.Row = 1 ' para q' actualice la primera fila del detalle
Detalle_Sumar_Normal

End Sub
Private Sub protocoloLeer(Numero As Long)

Dim m_resta As Integer

' CABECERA
With RsPp

.Seek "=", Numero

If Not .NoMatch Then

    Fecha.Text = Format(!Fecha, Fecha_Format)
    
    trabajadorRut = NoNulo(!Responsable)
    responsableNombre.Text = trabajadorBuscarNombreCompleto(trabajadorRut)

    numeroProtocolo.Text = PadR(NoNulo(!numeroProtocolo), 6, "_")
    paginaNumero = PadR(NoNulo(!paginaNumero), 2, "_")
    paginaTotal = PadR(NoNulo(!paginaTotal), 2, "_")
    esquema.Text = NoNulo(!esquema)
    
    OpProbeta(0).Value = !probetaprevia
    OpProbeta(1).Value = Not !probetaprevia

    OpGranalla(0).Value = !granallamezclada
    OpGranalla(1).Value = Not !granallamezclada

    calibre.Text = NoNulo(!calibretipo)

End If

End With

End Sub
Private Function Doc_Validar() As Boolean

Dim porRecibir As Integer, m_Maquina As String
Doc_Validar = False

If Len(responsableNombre) = 0 Then
    MsgBox "Debe elegirn RESPONSABLE"
    btnSearch.SetFocus
    Exit Function
End If

If numeroProtocolo.Text = "" Then
    MsgBox "Debe digitar NUMERO DE PROTOCOLO"
    numeroProtocolo.SetFocus
    Exit Function
End If

If paginaNumero.Text = "" Then
    MsgBox "Debe digitar NUMERO PAGINA"
    paginaNumero.SetFocus
    Exit Function
End If

If paginaTotal.Text = "" Then
    MsgBox "Debe digitar TOTAL DE PAGINAS"
    paginaTotal.SetFocus
    Exit Function
End If

If Len(esquema.Text) = 0 Then
    MsgBox "Debe digitar ESQUEMA SOLICITADO"
    esquema.SetFocus
    Exit Function
End If

If OpProbeta(0).Value = False And OpProbeta(1).Value = False Then
    MsgBox "Debe elegir PROCESO CON PROBETA PREVIA"
    OpProbeta(1).SetFocus
    Exit Function
End If

If OpGranalla(0).Value = False And OpGranalla(1).Value = False Then
    MsgBox "Debe elegir GRANALLA MEZCLADA"
    OpGranalla(1).SetFocus
    Exit Function
End If

If Len(calibre.Text) = 0 Then
    MsgBox "Debe digitar CALIBRE"
    calibre.SetFocus
    Exit Function
End If

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

' CABECERA PP

With RsPp
If Nueva Then
    .AddNew
    !Numero = Numero.Text
Else

'    Doc_Detalle_Eliminar
    
    .Edit
    
End If

!Nv = Val(m_Nv)
!Fecha = Fecha.Text

!numeroProtocolo = numeroProtocolo.Text
!paginaNumero = paginaNumero.Text
!paginaTotal = paginaTotal.Text
!Responsable = trabajadorRut
!esquema = esquema.Text
!probetaprevia = OpProbeta(0).Value
!granallamezclada = OpGranalla(0).Value
!calibretipo = calibre.Text

.Update

End With

Select Case Accion
Case "Agregando"
    Track_Registrar "PP" & "PP", Numero.Text, "AGR"
Case "Modificando"
    Track_Registrar "PP" & "PP", Numero.Text, "MOD"
End Select

MousePointer = vbDefault

End Sub
Private Sub Doc_Eliminar()

' elimina cabecera
With RsPp
.Seek "=", Numero.Text
If Not .NoMatch Then
    .Delete
End If
.Close
End With

End Sub
Private Sub Campos_Limpiar()

Numero.Text = ""
'GDespecial.Value = 0
Fecha.Text = Fecha_Vacia

m_Nv = 0
Nv.Text = "" ' m_Nv
Rut.Text = ""
Razon.Text = ""

' datos protocolo
trabajadorRut = ""
responsableNombre.Text = ""
numeroProtocolo.Text = "______"
paginaNumero.Text = "__"
paginaTotal.Text = "__"
esquema.Text = ""
OpProbeta(0).Value = False
OpProbeta(1).Value = False
OpGranalla(0).Value = False
OpGranalla(1).Value = False
calibre.Text = ""
'//////////////////////////

Detalle_Limpiar Detalle, n_columnas

Obs(0).Text = ""
Obs(1).Text = ""
'Totalm2.Caption = ""
'TotalKg.Caption = ""

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
    'Numero.Text = Documento_Numero_Nuevo_PG(m_TipoDoc, RsITOpgc)
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

        If MsgBox("¿ IMPRIMIR PROTOCOLO ?", vbYesNo) = vbYes Then
            Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Docs")
            Doc_ImprimeRPT
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
                Doc_ImprimeRPT
'                Impresora_Predeterminada "default"
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

camposITOpEnabled Si
camposProtocoloEnabled Si

End Sub
Private Sub camposITOpEnabled(Si As Boolean)

' des/habilita solo compos de la ito pintura

Numero.Enabled = Si

'Nv.Enabled = Si

'Detalle.Enabled = Si

'Obs(0).Enabled = Si
'Obs(1).Enabled = Si

End Sub
Private Sub camposProtocoloEnabled(Si As Boolean)

' des/habilita solo campos del protocolo pintura

Fecha.Enabled = Si

FrameEsquema.Enabled = Si
esquema.Enabled = Si

FrameProbeta.Enabled = Si
OpProbeta(0).Enabled = Si
OpProbeta(1).Enabled = Si

FrameGranalla.Enabled = Si
OpGranalla(0).Enabled = Si
OpGranalla(1).Enabled = Si

FrameCalibre.Enabled = Si
calibre.Enabled = Si

btnSearch.Enabled = Si
numeroProtocolo.Enabled = Si

paginaNumero.Enabled = Si
paginaTotal.Enabled = Si

End Sub
Private Sub btnSearch_Click()

Search.Muestra data_file, "Trabajadores", "RUT", "ApPaterno]+' '+[ApMaterno]+' '+[Nombres", "Trabajador", "Trabajadores" ', "Activo"

trabajadorRut = Search.Codigo

If trabajadorRut <> "" Then
   
    responsableNombre.Text = Search.Descripcion
    
End If

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
Private Sub Detalle_Sumar_Normal()
Dim Tot_m2 As Double, Tot_Kg As Double, Tot_Precio As Double
Tot_m2 = 0
Tot_Kg = 0
Tot_Precio = 0
For i = 1 To n_filas
    Tot_m2 = Tot_m2 + m_CDbl(Detalle.TextMatrix(i, 7))
    Tot_Kg = Tot_Kg + m_CDbl(Detalle.TextMatrix(i, 9))
    'Tot_Precio = Tot_Precio + m_CDbl(Detalle.TextMatrix(i, 13))
Next

'Totalm2.Caption = Format(Tot_m2, "#,###.00")
'TotalKg.Caption = Format(Tot_Kg, "#,###.00")

End Sub
' FIN RUTINAS PARA FLEXGRID
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
Private Sub Doc_ImprimeRPT()

protocoloPreparar Numero.Text, Nv

Cr.WindowTitle = "ITO Nº " & Numero.Text
Cr.ReportSource = crptReport
Cr.WindowState = crptMaximized
'Cr.WindowBorderStyle = crptFixedSingle
Cr.WindowMaxButton = False
Cr.WindowMinButton = False
'Cr.Formulas(0) = "RAZON SOCIAL=""" & EmpOC.Razon & """"

Cr.DataFiles(0) = repo_file & ".MDB"

Cr.ReportFileName = Drive_Server & Path_Rpt & "protocoloPintura.rpt"

Cr.Action = 1

End Sub
Private Sub Privilegios()
If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 0, 0
End If
End Sub
