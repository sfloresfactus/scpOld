VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form OT_Especial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orden De Trabajo"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
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
   Begin VB.ComboBox CbTipo 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin MSMask.MaskEdBox Entrega 
      Height          =   300
      Left            =   6000
      TabIndex        =   16
      Top             =   1635
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
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1650
      Width           =   2295
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   3
      Left            =   120
      MaxLength       =   30
      TabIndex        =   21
      Top             =   5700
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   2
      Left            =   120
      MaxLength       =   30
      TabIndex        =   20
      Top             =   5400
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   1
      Left            =   120
      MaxLength       =   30
      TabIndex        =   19
      Top             =   5100
      Width           =   5000
   End
   Begin VB.Frame Frame 
      Caption         =   "SubContratista"
      Height          =   1095
      Left            =   2400
      TabIndex        =   5
      Top             =   480
      Width           =   4575
      Begin VB.CommandButton btnSearch 
         Height          =   300
         Left            =   2400
         Picture         =   "OT_Especial.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   300
      End
      Begin VB.TextBox Razon 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   4095
      End
      Begin MSMask.MaskEdBox Rut 
         Height          =   300
         Left            =   720
         TabIndex        =   7
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label lbl 
         Caption         =   "&RUT"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbl 
         Caption         =   "SEÑOR(ES)"
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   9
         Top             =   480
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
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   0
      Left            =   120
      MaxLength       =   30
      TabIndex        =   18
      Top             =   4800
      Width           =   5000
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   840
      TabIndex        =   12
      Top             =   1635
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
      TabIndex        =   3
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
      Height          =   2445
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4313
      _Version        =   327680
      ScrollBars      =   2
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   8520
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
            Picture         =   "OT_Especial.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OT_Especial.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OT_Especial.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OT_Especial.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OT_Especial.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OT_Especial.frx":065C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OT_Especial.frx":076E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "F.&ENTREGA"
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   15
      Top             =   1680
      Width           =   975
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
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label NVnumero 
      Caption         =   "NVnumero"
      Height          =   255
      Left            =   8160
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
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
      Left            =   2040
      TabIndex        =   13
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "&FECHA"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "OT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Width           =   375
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
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   375
   End
End
Attribute VB_Name = "OT_Especial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button

'Private DbD As Database , RsSc As Recordset
'Private SqlRsSc As New ADODB.Recordset

Private DbM As Database, RsOTc As Recordset, RsOTd As Recordset, RsTxt As Recordset
'Private RsNVc As Recordset
'Private RsITOd As Recordset

Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, d As Variant

Private n_filas As Integer, n_columnas As Integer
Private prt As Printer
Private n1 As Double, n4 As Double
Private linea As String, TipoDoc As String, m_NvArea As Integer
Private mNv As NotaVenta
Private Sub ComboNV_Click()
NVnumero.Caption = Val(Left(ComboNV.Text, 6))
End Sub
Private Sub Fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then ComboNV.SetFocus
End Sub
Private Sub Fecha_LostFocus()
d = Fecha_Valida(Fecha, Now)
End Sub
Private Sub Entrega_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Detalle.SetFocus
End Sub
Private Sub Entrega_LostFocus()
d = Fecha_Valida(Entrega, Fecha.Text)
End Sub
Private Sub Form_Load()

' abre archivos
'Set DbD = OpenDatabase(data_file)
'Set RsSc = DbD.OpenRecordset("Contratistas")
'RsSc.Index = "RUT"

Set DbM = OpenDatabase(mpro_file)

Set RsOTc = DbM.OpenRecordset("OT Esp Cabecera")
RsOTc.Index = "Numero"

Set RsOTd = DbM.OpenRecordset("OT Esp Detalle")
RsOTd.Index = "Numero-Linea"

Set RsTxt = DbM.OpenRecordset("Detalle Texto")
RsTxt.Index = "Tipo-Numero-Linea"

'Set RsITOd = DbM.OpenRecordset("ITO Esp")
'RsITOd.Index = "OT"

'Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
'RsNVc.Index = Nv_Index ' "Número"

' puebla combo tipo de oc especial
CbTipo.AddItem " "
CbTipo.AddItem "Pintura"
CbTipo.AddItem "Galvanizado"
CbTipo.AddItem "Montaje"
CbTipo.AddItem "Otro"

nvListar Usuario.Nv_Activas

' Combo obra
ComboNV.AddItem " "
For i = 1 To nvTotal
    ComboNV.AddItem Format(aNv(i).Numero, "0000") & " - " & aNv(i).obra
Next

Inicializa
Detalle_Config

If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 0, 0
End If

m_NvArea = 0

End Sub
Private Sub Form_Unload(Cancel As Integer)
DataBases_Cerrar
End Sub
Private Sub Inicializa()

Obj = "ORDEN DE TRABAJO"
Objs = "ÓRDENES DE TRABAJO"
TipoDoc = "OTE"

Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnDesHacer = Toolbar.Buttons(6)
Set btnGrabar = Toolbar.Buttons(7)

Accion = ""
'old_accion = ""

'btnSearch.Visible = False
btnSearch.ToolTipText = "Busca Contatista"
Campos_Enabled False

End Sub
Private Sub Detalle_Config()
Dim i As Integer, ancho As Integer

n_filas = 12
n_columnas = 5

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

Detalle.ColWidth(0) = 250
Detalle.ColWidth(1) = 1000
Detalle.ColWidth(2) = 600
Detalle.ColWidth(3) = 3500
Detalle.ColWidth(4) = 800
Detalle.ColWidth(5) = 800

'Detalle.ColAlignment(2) = 0

ancho = 350 ' con scroll vertical
'ancho = 100 ' sin scroll vertical

TotalPrecio.Width = Detalle.ColWidth(5)
For i = 0 To n_columnas
    If i = 5 Then TotalPrecio.Left = ancho + Detalle.Left - 300
    ancho = ancho + Detalle.ColWidth(i)
Next

Detalle.Width = ancho
Me.Width = ancho + Detalle.Left * 2

' col y row fijas
'Detalle.BackColorFixed = vbCyan

' establece colores a columnas
' columnas    modificables : NEGRAS
' columnas no modificables : ROJAS
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

txtEditOT = ""

End Sub
Private Sub Numero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then After_Enter
End Sub
Private Sub After_Enter()
If Val(Numero.Text) < 1 Then Beep: Exit Sub
Select Case Accion
Case "Agregando"
    RsOTc.Seek "=", Numero
    If RsOTc.NoMatch Then
        Campos_Enabled True
        Numero.Enabled = False
        Fecha.SetFocus
        btnGrabar.Enabled = True
        btnSearch.visible = True
    Else
        OT_Leer
        MsgBox Obj & " YA EXISTE"
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
    End If
Case "Modificando"
    RsOTc.Seek "=", Numero
    If RsOTc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Botones_Enabled 0, 0, 0, 0, 1, 1
        OT_Leer
        Campos_Enabled True
        Numero.Enabled = False
    End If

Case "Eliminando"
    RsOTc.Seek "=", Numero
    If RsOTc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        OT_Leer
'        Numero.Enabled = False
        If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo) = vbYes Then
            OT_Eliminar
        End If
        Campos_Limpiar
        Campos_Enabled False
        Numero.Enabled = True
        Numero.SetFocus
    End If
   
Case "Imprimiendo"
    
    RsOTc.Seek "=", Numero
    If RsOTc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        OT_Leer
        Numero.Enabled = False
        Detalle.Enabled = True
    End If
    
End Select

End Sub
Private Sub OT_Leer()
Dim m_resta As Integer
' CABECERA
Fecha.Text = Format(RsOTc!Fecha, Fecha_Format)
NVnumero.Caption = RsOTc!Nv
Rut.Text = RsOTc![Rut contratista]

mNv = nvLeer(NVnumero.Caption)

If mNv.Numero <> 0 Then

'RsNVc.Seek "=", NVnumero.Caption, m_NvArea
'If Not RsNVc.NoMatch Then
'    ComboNV.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
    On Error Resume Next
    ComboNV.Text = Format(mNv.Numero, "0000") & " - " & mNv.obra
    ComboNV_Click
    On Error GoTo 0
End If


Entrega.Text = Format(RsOTc![Fecha Entrega], Fecha_Format)

'OtMontaje.Value = IIf(RsOTc!Montaje, 1, 0)

Select Case RsOTc!Tipo
Case "P"
    CbTipo.Text = "Pintura"
Case "G"
    CbTipo.Text = "Galvanizado"
Case "M"
    CbTipo.Text = "Montaje"
Case Else
End Select

Obs(0).Text = NoNulo(RsOTc![Observacion 1])
Obs(1).Text = NoNulo(RsOTc![Observacion 2])
Obs(2).Text = NoNulo(RsOTc![Observacion 3])
Obs(3).Text = NoNulo(RsOTc![Observacion 4])

'DETALLE

RsOTd.Seek "=", Numero.Text, 1
If Not RsOTd.NoMatch Then
    Do While Not RsOTd.EOF
        If RsOTd!Numero = Numero.Text Then
        
            i = RsOTd!linea
            
            Detalle.TextMatrix(i, 1) = RsOTd!Cantidad
            Detalle.TextMatrix(i, 2) = RsOTd!unidad
            Detalle.TextMatrix(i, 3) = RsOTd!Descripcion
            Detalle.TextMatrix(i, 4) = RsOTd![Precio Unitario]
            
            n1 = m_CDbl(Detalle.TextMatrix(i, 1))
            n4 = m_CDbl(Detalle.TextMatrix(i, 4))
            
            Detalle.TextMatrix(i, 5) = Format(n1 * n4, num_fmtgrl)
            
        Else
            Exit Do
        End If
        RsOTd.MoveNext
    Loop
End If

'lee texto
'RsTxt.Seek ">=", TipoDoc, Val(Numero.Text), 1
'If Not RsTxt.nomatch Then
'    Do While Not RsTxt.EOF
'        If RsTxt![Tipo Documento] = TipoDoc And RsTxt![Número Documento] = Numero.Text Then
'            i = RsTxt!Línea
'            Detalle.TextMatrix(i, 3) = RsTxt!Texto
'        Else
'            Exit Do
'        End If
'        RsTxt.MoveNext
'    Loop
'End If

'nuevo lee texto
For i = 1 To n_filas
    RsTxt.Seek "=", TipoDoc, Val(Numero.Text), i
    If Not RsTxt.NoMatch Then Detalle.TextMatrix(i, 3) = RsTxt!Texto
Next

Razon.Text = Contratista_Lee(SqlRsSc, Rut.Text)

Detalle.Row = 1 ' para q' actualice la primera fila del detalle

Actualiza

End Sub
Private Function OT_Validar() As Boolean
Dim porAsignar As Integer
OT_Validar = False

If Trim(CbTipo.Text) = "" Then
    MsgBox "DEBE ELEGIR TIPO"
    CbTipo.SetFocus
    Exit Function
End If
If Trim(ComboNV.Text) = "" Then
    MsgBox "DEBE ELEGIR OBRA"
    ComboNV.SetFocus
    Exit Function
End If
If Rut.Text = "" Then
    MsgBox "DEBE ELEGIR CONTRATISTA"
    btnSearch.SetFocus
    Exit Function
End If
If Entrega.Text = Fecha_Vacia Then
    MsgBox "DEBE DIGITAR FECHA DE ENTREGA"
    Entrega.SetFocus
    Exit Function
End If

For i = 1 To n_filas

    ' cant
    If Val(Detalle.TextMatrix(i, 1)) <> 0 Then
    
        If Not CampoReq_Valida(Detalle.TextMatrix(i, 2), i, 2) Then Exit Function
        If Not LargoString_Valida(Detalle.TextMatrix(i, 2), 3, i, 2) Then Exit Function
        
        If Not CampoReq_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
        If Not LargoString_Valida(Detalle.TextMatrix(i, 3), 30, i, 3) Then Exit Function
        
        If Not Numero_Valida(Detalle.TextMatrix(i, 4), i, 4) Then Exit Function
        
    End If
    
Next

OT_Validar = True

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
Private Sub OT_Grabar(Nueva As Boolean)
Dim m_cantidad As Double, m_pr As Double

Rut.Text = SqlRutPadL(Rut.Text)

save:
' CABECERA DE OT
If Nueva Then
    RsOTc.AddNew
    RsOTc!Numero = Numero.Text
Else
    RsOTc.Edit
End If

'm_pr = OT_PRecibido(Numero.Text) ' YA NO agosto 2013

RsOTc!Fecha = Fecha.Text
RsOTc!Nv = NVnumero.Caption
RsOTc![Rut contratista] = Rut.Text
RsOTc![Fecha Entrega] = CDate(Entrega.Text)
'RsOTc!Montaje = OtMontaje.Value
RsOTc!Tipo = Left(CbTipo.Text, 1)
RsOTc![Precio Total] = Val(TotalPrecio.Caption)
RsOTc![Porcentaje Recibido] = m_pr
RsOTc![Observacion 1] = Obs(0).Text
RsOTc![Observacion 2] = Obs(1).Text
RsOTc![Observacion 3] = Obs(2).Text
RsOTc![Observacion 4] = Obs(3).Text
RsOTc.Update

' DETALLE DE OT

OT_Detalle_Eliminar

j = 0
For i = 1 To n_filas
    m_cantidad = Val(Detalle.TextMatrix(i, 1))
    If m_cantidad = 0 Then
        ' puede haber texto
        If Trim(Detalle.TextMatrix(i, 3)) <> "" Then
            j = j + 1
            RsTxt.AddNew
            RsTxt![Tipo Documento] = TipoDoc
            RsTxt![Número Documento] = Numero.Text
            RsTxt!Línea = j
            RsTxt!Texto = Detalle.TextMatrix(i, 3)
            RsTxt.Update
        End If
    Else
        j = j + 1
        
        RsOTd.AddNew
        RsOTd!Numero = Numero.Text
        RsOTd!linea = j
        RsOTd!Fecha = Fecha.Text
'        RsOTd!Montaje = OtMontaje.Value
        RsOTd!Tipo = Left(CbTipo.Text, 1)
        RsOTd!Nv = NVnumero.Caption
        RsOTd![Rut contratista] = Rut.Text
        RsOTd!Cantidad = m_cantidad
        RsOTd!unidad = Detalle.TextMatrix(i, 2)
        RsOTd!Descripcion = Detalle.TextMatrix(i, 3)
        RsOTd![Precio Unitario] = m_CDbl(Detalle.TextMatrix(i, 4))
        RsOTd.Update
        
    End If
Next

End Sub
Private Function OT_PRecibido_YA_NO(N_Ot As Double) As Double
' recalcula porcentaje recibido
Dim pr As Double
pr = 0
' busca itos
'RsITOd.Seek "=", N_Ot
'If Not RsITOd.NoMatch Then
'    Do While Not RsITOd.EOF
'        If RsITOd!OT <> N_Ot Then Exit Do
'        If RsITOd!OT = N_Ot Then
'            pr = pr + RsITOd![Porcentaje Recibido]
'        End If
'        RsITOd.MoveNext
'    Loop
'End If
'OT_PRecibido = pr
End Function
Private Sub OT_Eliminar()

' borra CABECERA DE OT
RsOTc.Seek "=", Numero.Text
If Not RsOTc.NoMatch Then

    RsOTc.Delete

End If

OT_Detalle_Eliminar

End Sub
Private Sub OT_Detalle_Eliminar()

' elimina detalle
RsOTd.Seek "=", Numero.Text, 1
If Not RsOTd.NoMatch Then
    Do While Not RsOTd.EOF
        If RsOTd!Numero <> Numero.Text Then Exit Do
    
        ' borra detalle
        RsOTd.Delete
    
        RsOTd.MoveNext
    Loop
End If

' elimina texto
'RsTxt.Seek ">=", TipoDoc, Numero.Text, 1
'If Not RsTxt.nomatch Then
'    Do While Not RsTxt.EOF
'        If RsTxt![Tipo Documento] <> TipoDoc And RsTxt![Número Documento] = Numero.Text Then Exit Do
'
'        ' borra detalle
'       RsTxt.Delete
'
'        RsTxt.MoveNext
'    Loop
'End If

For i = 1 To n_filas
    RsTxt.Seek "=", TipoDoc, Numero.Text, i
    If Not RsTxt.NoMatch Then RsTxt.Delete
Next

End Sub
Private Function OT_Borrable_YA_NO() As Boolean
' busca en ITOs Especiales para ver si se puede borrar OT
'OT_Borrable = True
RsOTd.Seek "=", Numero.Text, 1
If Not RsOTd.NoMatch Then
    Do While Not RsOTd.EOF
        If RsOTd!Número <> Numero.Text Then Exit Do
        
'        RsITOd.Seek "=", RsOTd!Plano, RsOTd!Marca
'        If Not RsITOd.NoMatch Then
'            OT_Borrable = False
            Exit Function
'        End If
        
        RsOTd.MoveNext
    Loop
End If

End Function
Private Sub Campos_Limpiar()
Numero.Text = ""
Fecha.Text = Fecha_Vacia
ComboNV.Text = " "
Rut.Text = ""
Razon.Text = ""
Entrega.Text = Fecha_Vacia
'OtMontaje.Value = 0
CbTipo.Text = " "
Detalle_Limpiar
Obs(0).Text = ""
Obs(1).Text = ""
Obs(2).Text = ""
Obs(3).Text = ""
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
    
    Numero.Text = Documento_Numero_Nuevo(RsOTc, "Numero")
    
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
            OT_Imprimir n_Copias
            Impresora_Predeterminada "default"
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
    If OT_Validar Then
        If MsgBox("¿ GRABA " & Obj & " ?", vbYesNo) = vbYes Then
        
            If Accion = "Agregando" Then
                OT_Grabar True
            Else
                OT_Grabar False
            End If
            
            n_Copias = 1
            PrinterNCopias.Numero_Copias = n_Copias
            PrinterNCopias.Show 1
            n_Copias = PrinterNCopias.Numero_Copias
            
            If n_Copias > 0 Then
                Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Docs")
                OT_Imprimir n_Copias
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
btnSearch.Enabled = Si
ComboNV.Enabled = Si
Entrega.Enabled = Si
'OtMontaje.Enabled = Si
CbTipo.Enabled = Si
Detalle.Enabled = Si
Obs(0).Enabled = Si
Obs(1).Enabled = Si
Obs(2).Enabled = Si
Obs(3).Enabled = Si
End Sub
Private Sub btnSearch_Click()

Dim arreglo(1) As String
arreglo(1) = "razon_social"

sql_Search.Muestra "personas", "RUT", arreglo(), Obj, Objs, "contratista='S'"
Rut.Text = sql_Search.Codigo
Razon.Text = sql_Search.Descripcion

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
MSFlexGridEdit Detalle, txtEditOT, 32
End Sub
Private Sub Detalle_GotFocus()
If Accion = "Imprimiendo" Then Exit Sub
Select Case True
Case txtEditOT.visible
    Detalle = txtEditOT
    txtEditOT.visible = False
End Select
End Sub
Private Sub Detalle_LeaveCell()
If Accion = "Imprimiendo" Then Exit Sub
Select Case True
Case txtEditOT.visible
    Detalle = txtEditOT
    txtEditOT.visible = False
End Select
End Sub
Private Sub Detalle_KeyPress(KeyAscii As Integer)
If Accion = "Imprimiendo" Then Exit Sub
MSFlexGridEdit Detalle, txtEditOT, KeyAscii
End Sub
Private Sub txtEditOT_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCodeP Detalle, txtEditOT, KeyCode, Shift
End Sub
Sub EditKeyCodeP(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
' rutina que se ejecuta con los keydown de Edt
Dim m_col As Integer
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
        Linea_Actualiza
    End Select
    Cursor_Mueve MSFlexGrid
Case vbKeyUp ' Flecha Arriba
    Select Case m_col
    Case 2, 3
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
    Case 2, 3
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

If MSFlexGrid.col = 5 Then Exit Sub

Select Case MSFlexGrid.col
Case 2
    Edt.MaxLength = 3
Case 3
    Edt.MaxLength = 30
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
Edt.Move MSFlexGrid.CellLeft + MSFlexGrid.Left, MSFlexGrid.CellTop + MSFlexGrid.Top, MSFlexGrid.CellWidth, MSFlexGrid.CellHeight
Edt.visible = True
Edt.SetFocus
'opGrabar True

End Sub
Private Sub Detalle_KeyDown(KeyCode As Integer, Shift As Integer)
If Accion = "Imprimiendo" Then Exit Sub
If KeyCode = vbKeyF2 Then
    MSFlexGridEdit Detalle, txtEditOT, 32
End If
End Sub
Private Sub Linea_Actualiza()
' actualiza solo linea, y totales generales
Dim fi As Integer, co As Integer

fi = Detalle.Row
co = Detalle.col

n1 = m_CDbl(Detalle.TextMatrix(fi, 1))
n4 = m_CDbl(Detalle.TextMatrix(fi, 4))

' precio total
Detalle.TextMatrix(fi, 5) = Format(n1 * n4, num_fmtgrl)

Totales_Actualiza

End Sub
Private Sub Actualiza()
' actualiza todo el detalle y totales generales
Dim fi As Integer
For fi = 1 To n_filas
    If Detalle.TextMatrix(fi, 1) <> "" Then
    
        n1 = m_CDbl(Detalle.TextMatrix(fi, 1))
        n4 = m_CDbl(Detalle.TextMatrix(fi, 4))

        ' precio total
        Detalle.TextMatrix(fi, 5) = Format(n1 * n4, num_fmtgrl)
    End If
Next

Totales_Actualiza

End Sub
Private Sub Totales_Actualiza()
Dim Tot_Precio As Double
Tot_Precio = 0
For i = 1 To n_filas
    Tot_Precio = Tot_Precio + m_CDbl(Detalle.TextMatrix(i, 5))
Next

TotalPrecio.Caption = Format(Tot_Precio, num_Formato)

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
Private Sub OT_Imprimir(nCopias As Integer)
MousePointer = vbHourglass
linea = String(70, "-")
' imprime OT
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer
Dim tab4 As Integer, tab5 As Integer, tab40 As Integer
Dim fc As Integer, fn As Integer, fe As Integer
Dim n_Copia As Integer

tab0 = 3 'margen izquierdo
tab1 = tab0 + 0  ' cantidad
tab2 = tab1 + 14 ' unidad
tab3 = tab2 + 4  ' descripcion
tab4 = tab3 + 31 ' $ uni
tab5 = tab4 + 10 ' $ total
tab40 = 43

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
prt.Print Tab(tab0 + 25); "OT ESPECIAL Nº";
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
prt.Print Tab(tab0 + 28); Left(Razon, 31)

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
prt.Print Tab(tab1); "    CANTIDAD";
prt.Print Tab(tab2); "UNI";
prt.Print Tab(tab3); "DESCRIPCIÓN";
prt.Print Tab(tab4); "    $ UNI";
prt.Print Tab(tab5); "     $ TOT"
prt.Font.Bold = False

prt.Print Tab(tab1); linea
j = -1
For i = 1 To n_filas

    can_valor = m_CDbl(Detalle.TextMatrix(i, 1))
    
    If can_valor = 0 Then
    
        j = j + 1
        prt.Print Tab(tab1 + j * 3); "  \"
        
    Else
    
        ' CANTIDAD
        prt.Print Tab(tab1); m_Format(can_valor, "#,###,##0.00");
    
        ' UNIDAD
        prt.Print Tab(tab2); Detalle.TextMatrix(i, 2);
        
        ' DESCRIPCION
        prt.Print Tab(tab3); Left(Detalle.TextMatrix(i, 3), 30);
        
        ' $ UNITARIO
        prt.Print Tab(tab4); m_Format(Detalle.TextMatrix(i, 4), "#,###,###");
        
        ' $ TOTAL
        prt.Print Tab(tab5); m_Format(Detalle.TextMatrix(i, 5), "##,###,###")
        
    End If
    
Next

prt.Print Tab(tab1); linea
prt.Font.Bold = True
prt.Print Tab(tab4 - 5); m_Format(TotalPrecio, "$#,###,###,###")
prt.Font.Bold = False
prt.Print ""

prt.Print Tab(tab0); "FECHA ENTREGA : "; Entrega.Text
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
