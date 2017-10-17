VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ITO_PG_Esp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ITO PG Esp"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame_Operador 
      Caption         =   "Operador"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   4320
      TabIndex        =   19
      Top             =   480
      Width           =   3135
      Begin VB.CommandButton btnOperadorSearch 
         Height          =   300
         Left            =   1920
         Picture         =   "ITO_PG_Esp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   300
      End
      Begin VB.TextBox OperadorRazon 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   2295
      End
      Begin MSMask.MaskEdBox OperadorRut 
         Height          =   300
         Left            =   600
         TabIndex        =   22
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
         TabIndex        =   23
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
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
   Begin VB.ComboBox CbTrabajadores 
      Height          =   315
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2760
      Width           =   1815
   End
   Begin VB.ComboBox CbTipoGranalla 
      Height          =   315
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1440
      Width           =   975
   End
   Begin VB.ComboBox ComboNV 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Nv 
      Height          =   300
      Left            =   840
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   3
      Left            =   120
      MaxLength       =   30
      TabIndex        =   4
      Top             =   5700
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   2
      Left            =   120
      MaxLength       =   30
      TabIndex        =   3
      Top             =   5400
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   1
      Left            =   120
      MaxLength       =   30
      TabIndex        =   2
      Top             =   5100
      Width           =   5000
   End
   Begin VB.TextBox txtEdit 
      Height          =   285
      Left            =   8400
      TabIndex        =   5
      Text            =   "txtEdit"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   0
      Left            =   120
      MaxLength       =   30
      TabIndex        =   1
      Top             =   4800
      Width           =   5000
   End
   Begin MSFlexGridLib.MSFlexGrid Detalle 
      Height          =   2805
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4948
      _Version        =   393216
      ScrollBars      =   2
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   3120
      TabIndex        =   12
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
      TabIndex        =   13
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
   Begin MSComctlLib.ImageList ImageList 
      Left            =   9720
      Top             =   1560
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
            Picture         =   "ITO_PG_Esp.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_PG_Esp.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_PG_Esp.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_PG_Esp.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_PG_Esp.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_PG_Esp.frx":065C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_PG_Esp.frx":076E
            Key             =   ""
         EndProperty
      EndProperty
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
      TabIndex        =   16
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lbl 
      Caption         =   "&FECHA"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   15
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "&OBRA"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   615
   End
   Begin VB.Label NVnumero 
      Caption         =   "NVnumero"
      Height          =   255
      Left            =   8160
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Totalm2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   5040
      TabIndex        =   8
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
End
Attribute VB_Name = "ITO_PG_Esp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' tipos de ITO:

' P : pintura
' G : galvanizado
' R : granallado, Erwin
' T : produccion pintura , Erwin

' estos tipos se usan aqui //////////////////////////////////////////
' S : granallado especial
' U : produccion pintura especial

Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button

Private DbD As Database ', RsSc As Recordset
'Private SqlRsSc As New ADODB.Recordset

Private Dbm As Database, RsITOpgc As Recordset, RsITOpgd As Recordset
Private RsNVc As Recordset, RsSc As Recordset

Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, d As Variant

Private n_filas As Integer, n_columnas As Integer
Private prt As Printer
Private n1 As Double, n3 As Double, n4 As Double, n5 As Double
Private linea As String, m_TipoDoc As String, m_NvArea As Integer
Private a_TipoGranalla(9) As String, m_TotalTiposGranalla As Integer
Private a_Trabajadores(1, 199) As String, m_Nombre As String
Private a_Nv(2999, 1) As String, m_Nv As Double ', m_NvArea As Integer
Public Property Let TipoDoc(ByVal New_Tipo As String)
m_TipoDoc = New_Tipo
End Property
Private Sub btnOperadorSearch_Click()
If m_TipoDoc = "S" Then ' granallado
    Search.Muestra data_file, "Trabajadores", "RUT", "ApPaterno]+' '+[ApMaterno]+' '+[Nombres", Obj, Objs, "tipo5"
End If
If m_TipoDoc = "U" Then ' produccion pintura
    Search.Muestra data_file, "Trabajadores", "RUT", "ApPaterno]+' '+[ApMaterno]+' '+[Nombres", Obj, Objs, "tipo4"
End If

OperadorRut.Text = Search.Codigo
OperadorRazon.Text = Search.Descripcion

End Sub
Private Sub CbTipoGranalla_Click()
Detalle = CbTipoGranalla.Text
CbTipoGranalla.visible = False
End Sub
Private Sub CbTrabajadores_Click()
Detalle = CbTrabajadores.Text
CbTrabajadores.visible = False
End Sub
Private Sub ComboNV_Click()
'NVnumero.Caption = Val(Left(ComboNV.Text, 6))
Nv.Text = Val(Left(ComboNV.Text, 6))
End Sub
Private Sub Fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then ComboNV.SetFocus
End Sub
Private Sub Fecha_LostFocus()
d = Fecha_Valida(Fecha, Now)
End Sub
Private Sub Form_Load()

' abre archivos
Set DbD = OpenDatabase(data_file)
Set RsSc = DbD.OpenRecordset("Trabajadores")
RsSc.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)

Set RsITOpgc = Dbm.OpenRecordset("ITO PG Cabecera")
RsITOpgc.Index = "Numero"

Set RsITOpgd = Dbm.OpenRecordset("ITO PG Detalle")
RsITOpgd.Index = "Tipo-Numero-Linea"

Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

' Combo obra
i = 0
ComboNV.AddItem " "
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
        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
    End If
    
    RsNVc.MoveNext
    
Loop

Inicializa
Detalle_Config

Privilegios

m_NvArea = 0

CbTipoGranalla.visible = False
CbTipoGranalla.Width = 100

CbTrabajadores.visible = False
CbTrabajadores.Width = 1000

End Sub
Private Sub Form_Unload(Cancel As Integer)
DataBases_Cerrar
End Sub
Private Sub Inicializa()

Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnDesHacer = Toolbar.Buttons(6)
Set btnGrabar = Toolbar.Buttons(7)

Accion = ""
'old_accion = ""

n_filas = 5
n_columnas = 11

Campos_Enabled False

If m_TipoDoc = "S" Then ' granallado

'    n_columnas = 15 + 1

    Obj = "ITO GRANALLADO ESPECIAL"
    Objs = "ITOS GRANALLADO"
    Me.Caption = Obj
    Trabajadores_Poblar
    
    n_filas = 10

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

If m_TipoDoc = "U" Then
    Obj = "ITO PRODUCCION PINTURA ESPECIAL"
    Objs = "ITOS PRODUCCION PINTURA ESPECIAL"
    Me.Caption = Obj
    Trabajadores_Poblar
    n_filas = 30
    
End If

End Sub
Private Sub Detalle_Config()

Dim i As Integer, ancho As Integer

Detalle.Left = 100
Detalle.WordWrap = True
Detalle.RowHeight(0) = 450
Detalle.Rows = n_filas + 1
Detalle.Cols = n_columnas + 1

Detalle.TextMatrix(0, 0) = ""
Detalle.TextMatrix(0, 1) = "Cant"
Detalle.TextMatrix(0, 2) = "Descripción"
Detalle.TextMatrix(0, 3) = "m2 Uni"
Detalle.TextMatrix(0, 4) = "m2 Tot"
Detalle.TextMatrix(0, 5) = "$ Uni"
Detalle.TextMatrix(0, 6) = "$ Tot"

If m_TipoDoc = "S" Then ' granallado especial

    
    Detalle.TextMatrix(0, 7) = "Tipo Grana"
    Detalle.TextMatrix(0, 8) = "Maquina"
    Detalle.TextMatrix(0, 9) = "" ' columna invisible
    
    Detalle.ColWidth(7) = 700
    Detalle.ColWidth(8) = 700
    Detalle.ColWidth(9) = 0
    
End If

If m_TipoDoc = "U" Then ' produccion pintura especial

    
    Detalle.TextMatrix(0, 7) = "nManos Antic"
    Detalle.TextMatrix(0, 8) = "nManos Termin"
    Detalle.TextMatrix(0, 9) = "m2 Tot"
    
    Detalle.ColWidth(7) = 700
    Detalle.ColWidth(8) = 700
    Detalle.ColWidth(9) = 600
    
End If

Detalle.TextMatrix(0, 10) = "Turno"
Detalle.TextMatrix(0, 11) = "Operador"

Detalle.ColWidth(0) = 300
Detalle.ColWidth(1) = 700
Detalle.ColWidth(2) = 2000
Detalle.ColWidth(3) = 600
Detalle.ColWidth(4) = 600
Detalle.ColWidth(5) = 600
Detalle.ColWidth(6) = 600

Detalle.ColWidth(10) = 550
'Detalle.ColWidth(11) = 2000
Detalle.ColWidth(11) = 0

'Detalle.ColWidth(11) = 0 ' peso unitario

ancho = 350 ' con scroll vertical
'ancho = 100 ' sin scroll vertical

If m_TipoDoc = "S" Then
    Totalm2.Width = Detalle.ColWidth(4)
End If
If m_TipoDoc = "U" Then
    Totalm2.Width = Detalle.ColWidth(9)
End If

For i = 0 To n_columnas
    If m_TipoDoc = "S" Then
        If i = 4 Then Totalm2.Left = ancho + Detalle.Left - 350
    End If
    If m_TipoDoc = "U" Then
        If i = 9 Then Totalm2.Left = ancho + Detalle.Left - 350
    End If
    ancho = ancho + Detalle.ColWidth(i)
Next

Detalle.Width = ancho
Me.Width = ancho + Detalle.Left * 2

For i = 1 To n_filas
    Detalle.TextMatrix(i, 0) = i
Next

If False Then
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
    
Next
End If

txtEdit.Text = ""

Detalle.Enabled = False

End Sub
Private Sub Numero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then After_Enter
End Sub
Private Sub After_Enter()
If Val(Numero.Text) < 1 Then Beep: Exit Sub
Select Case Accion
Case "Agregando"
    RsITOpgc.Seek "=", m_TipoDoc, Numero.Text
    If RsITOpgc.NoMatch Then
        Campos_Enabled True
        Numero.Enabled = False
        Fecha.SetFocus
        btnGrabar.Enabled = True
'        btnSearch.visible = True
    Else
        Doc_Leer
        MsgBox Obj & " YA EXISTE"
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
    End If
Case "Modificando"
    RsITOpgc.Seek "=", m_TipoDoc, Numero.Text
    If RsITOpgc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Botones_Enabled 0, 0, 0, 0, 1, 1
        Doc_Leer
        Campos_Enabled True
        Numero.Enabled = False
    End If

Case "Eliminando"
    RsITOpgc.Seek "=", m_TipoDoc, Numero.Text
    If RsITOpgc.NoMatch Then
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
    
    RsITOpgc.Seek "=", m_TipoDoc, Numero.Text
    If RsITOpgc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Doc_Leer
        Numero.Enabled = False
        Detalle.Enabled = True
    End If
    
End Select

End Sub
Private Sub Doc_Leer()
Dim m_resta As Integer
' CABECERA
Fecha.Text = Format(RsITOpgc!Fecha, Fecha_Format)
NvNumero.Caption = RsITOpgc!Nv
'Rut.Text = rsitopgc![RUT Contratista]

OperadorRut.Text = NoNulo(RsITOpgc![Rut contratista])
RsSc.Seek "=", OperadorRut.Text
If Not RsSc.NoMatch Then
    OperadorRazon = RsSc![appaterno] & " " & RsSc![nombres]
End If

'OtMontaje.Value = IIf(rsitopgc!Montaje, 1, 0)


RsNVc.Seek "=", NvNumero.Caption, m_NvArea
If Not RsNVc.NoMatch Then
    ComboNV.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
    ComboNV_Click
End If

Obs(0).Text = NoNulo(RsITOpgc![Observacion 1])
Obs(1).Text = NoNulo(RsITOpgc![Observacion 2])
Obs(2).Text = NoNulo(RsITOpgc![Observacion 3])
Obs(3).Text = NoNulo(RsITOpgc![Observacion 4])

'DETALLE

RsITOpgd.Seek "=", m_TipoDoc, Numero.Text, 1
If Not RsITOpgd.NoMatch Then

    Do While Not RsITOpgd.EOF
    
        If RsITOpgd!Tipo = m_TipoDoc And RsITOpgd!Numero = Numero.Text Then
        
            i = RsITOpgd!linea
            
            Detalle.TextMatrix(i, 1) = RsITOpgd!Cantidad
            Detalle.TextMatrix(i, 2) = NoNulo(RsITOpgd!Descripcion)
            Detalle.TextMatrix(i, 3) = RsITOpgd![m2 Unitario]
            Detalle.TextMatrix(i, 5) = RsITOpgd![Precio Unitario]
            
            n1 = m_CDbl(Detalle.TextMatrix(i, 1))
            n4 = m_CDbl(Detalle.TextMatrix(i, 4))
            
'            Detalle.TextMatrix(i, 6) = Format(n1 * n4, num_fmtgrl)
            
            If m_TipoDoc = "S" Then
            
                Detalle.TextMatrix(i, 7) = NoNulo(RsITOpgd![tipo2])
                Detalle.TextMatrix(i, 8) = NoNulo(RsITOpgd![Maquina])
                
                Detalle.TextMatrix(i, 10) = RsITOpgd![Turno]
                
                For j = 0 To 199
                    If a_Trabajadores(0, j) = RsITOpgd![Rut operador] Then
                        Detalle.TextMatrix(i, 11) = a_Trabajadores(1, j)
                        Exit For
                    End If
                Next
            
            End If
            
            If m_TipoDoc = "U" Then
            
                Detalle.TextMatrix(i, 7) = RsITOpgd![manos1]
                Detalle.TextMatrix(i, 8) = RsITOpgd![manos2]
                
                Detalle.TextMatrix(i, 10) = RsITOpgd![Turno]
                
                For j = 0 To 199
                    If a_Trabajadores(0, j) = RsITOpgd![Rut operador] Then
                        Detalle.TextMatrix(i, 11) = a_Trabajadores(1, j)
                        Exit For
                    End If
                Next
            
            End If
            
            Fila_Calcular i, False
                        
        Else
            Exit Do
        End If
        RsITOpgd.MoveNext
    Loop
End If

'Contratista_Lee Rut.Text
Detalle.Row = 1 ' para q' actualice la primera fila del detalle

Totales_Actualiza

End Sub
Private Function Doc_Validar() As Boolean

Dim m_Maquina As String

Doc_Validar = False
If Trim(ComboNV.Text) = "" Then
    MsgBox "DEBE ELEGIR OBRA"
    ComboNV.SetFocus
    Exit Function
End If

For i = 1 To n_filas

    ' cant
    If Val(Detalle.TextMatrix(i, 1)) <> 0 Then
    
        If Not CampoReq_Valida(Detalle.TextMatrix(i, 2), i, 2) Then Exit Function
        
        If Not LargoString_Valida(Detalle.TextMatrix(i, 2), 30, i, 2) Then Exit Function
        
        If Not CampoReq_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
        
        If Not Numero_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
        If Not Numero_Valida(Detalle.TextMatrix(i, 5), i, 5) Then Exit Function
        
        If m_TipoDoc = "S" Then ' granalla especial
        
            m_Maquina = UCase(Detalle.TextMatrix(i, 8))
            If m_Maquina <> "A" And m_Maquina <> "M" Then
                MsgBox "Maquina debe ser A ó M", , "ATENCIÓN"
                Detalle.Row = i
                Detalle.col = 8
                Detalle.SetFocus
                Exit Function
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

Dim m_cantidad As Double, jj As Integer

save:
' CABECERA DE ITO
If Nueva Then
    RsITOpgc.AddNew
    RsITOpgc!Tipo = m_TipoDoc
    RsITOpgc!Numero = Numero.Text
Else
    RsITOpgc.Edit
End If

'm_pr = OT_PRecibido(Numero.Text)

RsITOpgc![Rut contratista] = OperadorRut.Text
RsITOpgc!Fecha = Fecha.Text
RsITOpgc!Nv = Nv.Text ' NVnumero.Caption
RsITOpgc![Observacion 1] = Obs(0).Text
RsITOpgc![Observacion 2] = Obs(1).Text
RsITOpgc![Observacion 3] = Obs(2).Text
RsITOpgc![Observacion 4] = Obs(3).Text
RsITOpgc.Update

' DETALLE DE OT

Doc_Detalle_Eliminar

j = 0
With RsITOpgd
For i = 1 To n_filas
    m_cantidad = Val(Detalle.TextMatrix(i, 1))
    If m_cantidad = 0 Then
    Else
        j = j + 1
        
        .AddNew
        !Tipo = m_TipoDoc
        !Numero = Numero.Text
        !linea = j
        !Fecha = Fecha.Text
        !Nv = Nv.Text ' NVnumero.Caption
        ![Rut operador] = OperadorRut.Text
        !Cantidad = m_cantidad
        !Descripcion = Detalle.TextMatrix(i, 2)
        ![m2 Unitario] = m_CDbl(Detalle.TextMatrix(i, 3))
        ![Precio Unitario] = m_CDbl(Detalle.TextMatrix(i, 5))
        
        If m_TipoDoc = "S" Then ' granallado especial
        
            ![tipo2] = Detalle.TextMatrix(i, 7)
            ![Maquina] = UCase(Detalle.TextMatrix(i, 8))
            
            ![Turno] = m_CDbl(Detalle.TextMatrix(i, 10))
            
            'For jj = 0 To 199
            '    If a_Trabajadores(1, jj) = Detalle.TextMatrix(i, 11) Then
            '        ![Rut operador] = a_Trabajadores(0, jj)
            '        Exit For
            '    End If
            'Next
            
        End If

        If m_TipoDoc = "U" Then ' ito produccion pintura especial
        
            ![manos1] = m_CDbl(Detalle.TextMatrix(i, 7))
            ![manos2] = m_CDbl(Detalle.TextMatrix(i, 8))
            
            ![Turno] = m_CDbl(Detalle.TextMatrix(i, 10))
            
            'For jj = 0 To 199
            '    If a_Trabajadores(1, jj) = Detalle.TextMatrix(i, 11) Then
            '        ![Rut operador] = a_Trabajadores(0, jj)
            '        Exit For
            '    End If
            'Next
            
        End If
        
        .Update
        
    End If
Next

End With

End Sub
Private Sub Doc_Eliminar()

' borra CABECERA DE OT
RsITOpgc.Seek "=", m_TipoDoc, Numero.Text
If Not RsITOpgc.NoMatch Then

    RsITOpgc.Delete

End If

Doc_Detalle_Eliminar

End Sub
Private Sub Doc_Detalle_Eliminar()

' elimina detalle
RsITOpgd.Seek "=", m_TipoDoc, Numero.Text, 1
If Not RsITOpgd.NoMatch Then
    Do While Not RsITOpgd.EOF
    
        If RsITOpgd!Tipo <> m_TipoDoc Or RsITOpgd!Numero <> Numero.Text Then Exit Do
    
        ' borra detalle
        RsITOpgd.Delete
    
        RsITOpgd.MoveNext
    Loop
End If

End Sub
Private Sub Campos_Limpiar()
Numero.Text = ""
Fecha.Text = Fecha_Vacia
ComboNV.Text = " "
OperadorRut.Text = ""
OperadorRazon.Text = ""
'Rut.Text = ""
'Razon.Text = ""
'OtMontaje.Value = 0
'CbTipo.Text = " "
Detalle_Limpiar
Obs(0).Text = ""
Obs(1).Text = ""
Obs(2).Text = ""
Obs(3).Text = ""
Totalm2.Caption = "0"
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
Dim m_Nv As Double
m_Nv = m_CDbl(Nv.Text)
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

cambia_titulo = True
'Accion = "" rem accion
Select Case Button.Index
Case 1 ' Agregar
    Accion = "Agregando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    
'    Numero.Text = Documento_Numero_Nuevo(RsITOpgc, "Numero")

    Numero.Text = Documento_Numero_Nuevo_PG(m_TipoDoc, RsITOpgc)
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
            Impresora_Predeterminada "default"
        End If
        
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
                Impresora_Predeterminada "default"
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
'btnSearch.Enabled = Si
Nv.Enabled = Si
ComboNV.Enabled = Si
Frame_Operador.Enabled = Si

Detalle.Enabled = Si
Obs(0).Enabled = Si
Obs(1).Enabled = Si
Obs(2).Enabled = Si
Obs(3).Enabled = Si
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
Case Else
End Select
End Sub
Private Sub Detalle_DblClick()
If Accion = "Imprimiendo" Then Exit Sub
' simula un espacio
MSFlexGridEdit Detalle, txtEdit, 32
End Sub
Private Sub Detalle_GotFocus()
If Accion = "Imprimiendo" Then Exit Sub
Select Case True
Case txtEdit.visible
    Detalle = txtEdit
    txtEdit.visible = False
End Select
End Sub
Private Sub Detalle_LeaveCell()
If Accion = "Imprimiendo" Then Exit Sub
Select Case True
Case txtEdit.visible
    Detalle = txtEdit
    txtEdit.visible = False
End Select
End Sub
Private Sub Detalle_KeyPress(KeyAscii As Integer)
If Accion = "Imprimiendo" Then Exit Sub
MSFlexGridEdit Detalle, txtEdit, KeyAscii
End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode Detalle, txtEdit, KeyCode, Shift
End Sub
Sub EditKeyCode(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
' rutina que se ejecuta con los keydown de Edt
Dim m_fil As Integer, m_col As Integer ', dif As Integer
m_fil = MSFlexGrid.Row
m_col = MSFlexGrid.col
'dif = Val(Detalle.TextMatrix(MSFlexGrid.Row, 5)) - Val(Detalle.TextMatrix(MSFlexGrid.Row, 6))
Select Case KeyCode
Case vbKeyEscape ' Esc
    Edt.visible = False
    MSFlexGrid.SetFocus
Case vbKeyReturn ' Enter
    Select Case m_col
    Case 1, 3 ' Cant y m2
'        If Recibida_Validar(MSFlexGrid.col, dif, Edt) Then
            MSFlexGrid.SetFocus
            DoEvents
            Fila_Calcular m_fil, True
'        End If
    Case 7
        If m_TipoDoc = "U" Then ' manos anticorrosivo
            If Manos_Validar(MSFlexGrid.col, Edt) Then
                MSFlexGrid.SetFocus
                DoEvents
                Fila_Calcular m_fil, True
            End If
        End If
    Case 8
        If m_TipoDoc = "S" Then ' maquina
            If Maquina_Validar(MSFlexGrid.col, Edt) Then
                MSFlexGrid.SetFocus
                DoEvents
            End If
        End If
        If m_TipoDoc = "U" Then ' manos terminacion
            If Manos_Validar(MSFlexGrid.col, Edt) Then
                MSFlexGrid.SetFocus
                DoEvents
                Fila_Calcular m_fil, True
            End If
        End If
    Case 10 ' turno
        If m_TipoDoc = "S" Or m_TipoDoc = "U" Then
'            dif = Val(Detalle.TextMatrix(MSFlexGrid.Row, 14))
            If Turno_Validar(MSFlexGrid.col, Edt) Then
                MSFlexGrid.SetFocus
                DoEvents
            End If
        End If
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
        Fila_Calcular m_fil, True
    End Select
    Cursor_Mueve MSFlexGrid
Case vbKeyUp ' Flecha Arriba
    Select Case m_col
    Case 1, 3 ' Cant y m2 uni
'        If Recibida_Validar(MSFlexGrid.col, dif, Edt) Then
            MSFlexGrid.SetFocus
            DoEvents
            Fila_Calcular m_fil, True
            If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
                MSFlexGrid.Row = MSFlexGrid.Row - 1
            End If
'        End If
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
    Case 1, 3 ' Cant
'        If Recibida_Validar(MSFlexGrid.col, dif, Edt) Then
            MSFlexGrid.SetFocus
            DoEvents
            Fila_Calcular m_fil, True
            If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
                MSFlexGrid.Row = MSFlexGrid.Row + 1
            End If
'        End If
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
Private Sub txtEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)

Select Case MSFlexGrid.col
'Case 1, 2, 3
'    After_Detalle_Click
Case 4, 6
    ' no editables
    Exit Sub
Case 7 ' tipo granalla
    If m_TipoDoc = "S" Then
        CbTipoGranalla.Left = MSFlexGrid.CellLeft + MSFlexGrid.Left
        CbTipoGranalla.Top = MSFlexGrid.CellTop + MSFlexGrid.Top
        CbTipoGranalla.Width = Detalle.ColWidth(7)
        CbTipoGranalla.visible = True
        CbTipoGranalla.SetFocus
    End If
    If m_TipoDoc = "U" Then GoTo Editar
Case 11 ' combo trabajador
    CbTrabajadores.Left = MSFlexGrid.CellLeft + MSFlexGrid.Left
    CbTrabajadores.Top = MSFlexGrid.CellTop + MSFlexGrid.Top
    CbTrabajadores.Width = Detalle.ColWidth(11)
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
    MSFlexGridEdit Detalle, txtEdit, 32
End If
End Sub
Private Sub Linea_Actualizax()
' actualiza solo linea, y totales generales
Dim fi As Integer, co As Integer

fi = Detalle.Row
co = Detalle.col

n1 = m_CDbl(Detalle.TextMatrix(fi, 1))
n3 = m_CDbl(Detalle.TextMatrix(fi, 5))

' m2 tot
Detalle.TextMatrix(fi, 4) = Format(n1 * n3, num_fmtgrl)

Totales_Actualiza

End Sub
Private Sub Totales_Actualiza()
Dim tm2_1 As Double, tm2_2 As Double

tm2_1 = 0
tm2_2 = 0

For i = 1 To n_filas
    tm2_1 = tm2_1 + m_CDbl(Detalle.TextMatrix(i, 4))
    tm2_2 = tm2_2 + m_CDbl(Detalle.TextMatrix(i, 9))
Next

If m_TipoDoc = "S" Then
    Totalm2.Caption = Format(tm2_1, num_Formato)
End If
If m_TipoDoc = "U" Then
    Totalm2.Caption = Format(tm2_2, num_Formato)
End If

End Sub
Private Sub Cursor_Mueve(MSFlexGrid As Control)

If MSFlexGrid.col = 3 Or MSFlexGrid.col = 5 Or MSFlexGrid.col = 9 Then
    MSFlexGrid.col = MSFlexGrid.col + 2
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
Private Sub Doc_Imprimir(nCopias As Integer)
MousePointer = vbHourglass
linea = String(70, "-")
' imprime OT
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer
Dim tab4 As Integer, tab5 As Integer, tab6 As Integer, tab7 As Integer, tab40 As Integer
Dim fc As Integer, fn As Integer, fe As Integer
Dim n_Copia As Integer

' font.size
fc = 10 'comprimida
fn = 12 'normal
fe = 15 'expandida

Dim can_valor As String

'Printer_Set "Documentos" ' OJO seque comentarios 20/02/06
Set prt = Printer
Font_Setear prt

For n_Copia = 1 To nCopias

prt.Font.Size = fe
prt.Print Tab(tab0 - 1); m_TextoIso
If m_TipoDoc = "S" Then
    prt.Print Tab(tab0 + 15); "ITO GRANALLADO ESPECIAL Nº";
End If
If m_TipoDoc = "U" Then
    prt.Print Tab(tab0 + 15); "ITO PRODUCCION PINTURA ESPECIAL Nº";
End If
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
prt.Print Tab(tab0 + tab40); "OPERADOR :" & OperadorRut.Text & " " & OperadorRazon.Text
prt.Print Tab(tab0); "GIRO: " & Empresa.Giro;
prt.Font.Size = fe
prt.Font.Bold = True
'prt.Print Tab(tab0 + 28); Left(Razon, 31)

prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0); Empresa.Direccion;
prt.Font.Size = fc
prt.Print Tab(tab0 + tab40); "OBRA      : "
prt.Print Tab(tab0); "Teléfono: "; Empresa.Telefono1;
prt.Font.Size = fe
prt.Font.Bold = True
prt.Print Tab(tab0 + 28); Left(Format(ComboNV.Text, ">"), 32)

prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0); Empresa.Comuna;
prt.Font.Size = fn

prt.Print ""
prt.Print ""

' detalle
prt.Font.Bold = True
If m_TipoDoc = "S" Then

    tab0 = 3 'margen izquierdo
    tab1 = tab0 + 0   ' cantidad
    tab2 = tab1 + 6   ' descripcion
    tab3 = tab2 + 31  ' m2 uni
    tab4 = tab3 + 9   ' m2 total
    tab40 = 43

    prt.Print Tab(tab1); " CANT";
    prt.Print Tab(tab2); "DESCRIPCIÓN";
    prt.Print Tab(tab3); "   m2 UNI";
    prt.Print Tab(tab4); "     m2 TOT"
    
End If
If m_TipoDoc = "U" Then

    tab0 = 3 'margen izquierdo
    tab1 = tab0 + 0   ' cantidad
    tab2 = tab1 + 6   ' descripcion
    tab3 = tab2 + 25  ' m2 uni
    tab4 = tab3 + 8   ' m2 total
    tab5 = tab4 + 9   ' m ant
    tab6 = tab5 + 6   ' m ter
    tab7 = tab6 + 6   ' m2 total
    tab40 = 43
    
    prt.Print Tab(tab1); " CANT";
    prt.Print Tab(tab2); "DESCRIPCIÓN";
    prt.Print Tab(tab3); " m2 UNI";
    prt.Print Tab(tab4); " m2 TOT";
    prt.Print Tab(tab5); "NºAnt";
    prt.Print Tab(tab6); "NºTer";
    prt.Print Tab(tab7); " m2 TOT"
    
End If
prt.Font.Bold = False

prt.Print Tab(tab1); linea
j = -1
For i = 1 To n_filas

    can_valor = m_CDbl(Detalle.TextMatrix(i, 1))
    
    If can_valor = 0 Then
    
        j = j + 1
        prt.Print Tab(tab1 + j * 3); "  \"
        
    Else
    
        If m_TipoDoc = "S" Then
    
            ' CANTIDAD
            prt.Print Tab(tab1); m_Format(can_valor, "#,##0");
        
            ' DESCRIPCION
            prt.Print Tab(tab2); Left(Detalle.TextMatrix(i, 2), 30);
            
            ' m2 UNITARIO
            prt.Print Tab(tab3); m_Format(Detalle.TextMatrix(i, 3), "#,###.##");
            
            ' m2 TOTAL
            prt.Print Tab(tab4); m_Format(Detalle.TextMatrix(i, 4), "###,###.#")
            
        End If
        
        If m_TipoDoc = "U" Then
    
            ' CANTIDAD
            prt.Print Tab(tab1); m_Format(can_valor, "#,##0");
        
            ' DESCRIPCION
            prt.Print Tab(tab2); Left(Detalle.TextMatrix(i, 2), 24);
            
            ' m2 UNITARIO
            prt.Print Tab(tab3); m_Format(Detalle.TextMatrix(i, 3), "#,###.##");
            
            ' m2 TOTAL
            prt.Print Tab(tab4); m_Format(Detalle.TextMatrix(i, 4), "#,###.#");
            
            ' m ant
            prt.Print Tab(tab5); m_Format(Detalle.TextMatrix(i, 7), "#0");
            
            ' m ter
            prt.Print Tab(tab6); m_Format(Detalle.TextMatrix(i, 8), "#0");
            
            ' m2 TOTAL
            prt.Print Tab(tab7); m_Format(Detalle.TextMatrix(i, 9), "#,###.#")
            
        End If
        
    End If
    
Next

prt.Print Tab(tab1); linea
prt.Font.Bold = True
'If m_TipoDoc = "S" Then
    prt.Print Tab(tab4 - 5); m_Format(Totalm2.Caption, "#,###,###,###.#")
'End If
'If m_TipoDoc = "U" Then
'    prt.Print Tab(tab4 - 5); m_Format(Totalm2.Caption, "#,###,###,###.#")
'End If
prt.Font.Bold = False
prt.Print ""

prt.Print Tab(tab0); "OBSERVACIONES :";
prt.Print Tab(tab0 + 16); Obs(0).Text
prt.Print Tab(tab0 + 16); Obs(1).Text
prt.Print Tab(tab0 + 16); Obs(2).Text
prt.Print Tab(tab0 + 16); Obs(3).Text

For i = 1 To 2
    prt.Print ""
Next

prt.Print Tab(tab0); Tab(14), "__________________", Tab(56), "__________________"
prt.Print Tab(tab0); Tab(14), "       VºBº       ", Tab(56), "       VºBº       "

If n_Copia < nCopias Then
    prt.NewPage
Else
    prt.EndDoc
End If

Next

Impresora_Predeterminada "default"

MousePointer = vbDefault

End Sub
Private Sub Fila_Calcular(Fila As Integer, Actualizar As Boolean)
' actualiza solo linea, y totales generales
Dim n1 As Double, n3 As Double, n4 As Double, n5 As Double, n7 As Integer, n8 As Integer, n9 As Double

n1 = m_CDbl(Detalle.TextMatrix(Fila, 1)) ' cant
n3 = m_CDbl(Detalle.TextMatrix(Fila, 3)) ' m2 uni
n5 = m_CDbl(Detalle.TextMatrix(Fila, 5)) ' $ uni
n7 = m_CDbl(Detalle.TextMatrix(Fila, 7)) ' m ant
n8 = m_CDbl(Detalle.TextMatrix(Fila, 8)) ' m ter

n4 = n1 * n3 ' m2 tot
n9 = n1 * n3 * (n7 + n8) ' m2 tot

Detalle.TextMatrix(Fila, 4) = Format(n4, "#.00")
Detalle.TextMatrix(Fila, 9) = Format(n9, "#.00")

' precio total
Detalle.TextMatrix(Fila, 6) = Format(n1 * n3 * n5, "#")

Totales_Actualiza

'If Actualizar Then Detalle_Sumar_Normal

End Sub
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
If Edt <> "M" Or Edt <> "A" Then
Else
    MsgBox "Maquina debe ser A ó M", , "ATENCIÓN"
    If m_TipoDoc = "S" Then
    Detalle.col = 8
    Else
    Detalle.col = 12
    End If
    Detalle.SetFocus
    Maquina_Validar = False
End If
End Function
Private Function Turno_Validar(Colu As Integer, Edt As Control) As Boolean
Turno_Validar = True
If 1 <= Val(Edt) And Val(Edt) <= 2 Then
Else
    MsgBox "Turno debe ser 1 ó 2", , "ATENCIÓN"
    Detalle.col = 9
    Detalle.SetFocus
    Turno_Validar = False
End If
End Function
Private Sub Trabajadores_Poblar()

Dim sql As String, RsTra As Recordset

CbTrabajadores.Clear

If m_TipoDoc = "S" Then ' granalla especial
    sql = "SELECT * FROM trabajadores WHERE tipo5 ORDER BY appaterno"
End If

If m_TipoDoc = "U" Then ' pintura especial
    sql = "SELECT * FROM trabajadores WHERE tipo4 ORDER BY appaterno"
End If

Set RsTra = DbD.OpenRecordset(sql)

With RsTra
i = 0
Do While Not .EOF
    i = i + 1
    a_Trabajadores(0, i) = !rut
    m_Nombre = !nombres & " " & !appaterno & " " & !apmaterno
    a_Trabajadores(1, i) = m_Nombre
    CbTrabajadores.AddItem m_Nombre
'        Debug.Print !nombres, !appaterno, !apmaterno
    .MoveNext
Loop
.Close
End With
End Sub
Private Sub Privilegios()
If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 0, 0
End If
End Sub
