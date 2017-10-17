VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form OT_Mantencion 
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
      TabIndex        =   21
      Top             =   0
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Ingresar Nueva OT"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Modificar OT"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Eliminar OT"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Imprimir OT"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "DesHacer"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Grabar OT"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Mantención de Subcontratistas"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      MouseIcon       =   "OT_Mantencion.frx":0000
   End
   Begin VB.ComboBox CbCC 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   1680
      Width           =   2055
   End
   Begin MSMask.MaskEdBox Entrega 
      Height          =   300
      Left            =   6240
      TabIndex        =   14
      Top             =   1635
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   327680
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   3
      Left            =   120
      MaxLength       =   30
      TabIndex        =   19
      Top             =   5700
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   2
      Left            =   120
      MaxLength       =   30
      TabIndex        =   18
      Top             =   5400
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   1
      Left            =   120
      MaxLength       =   30
      TabIndex        =   17
      Top             =   5100
      Width           =   5000
   End
   Begin VB.Frame Frame 
      Caption         =   "SubContratista"
      Height          =   1095
      Left            =   2520
      TabIndex        =   4
      Top             =   480
      Width           =   4575
      Begin VB.CommandButton btnSearch 
         Height          =   300
         Left            =   2400
         Picture         =   "OT_Mantencion.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   300
      End
      Begin VB.TextBox Razon 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   4095
      End
      Begin MSMask.MaskEdBox Rut 
         Height          =   300
         Left            =   720
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbl 
         Caption         =   "SEÑOR(ES)"
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.TextBox txtEditOT 
      Height          =   285
      Left            =   8040
      TabIndex        =   20
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
      TabIndex        =   16
      Top             =   4800
      Width           =   5000
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   840
      TabIndex        =   11
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
      TabIndex        =   15
      Top             =   2040
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4313
      _Version        =   327680
      ScrollBars      =   2
   End
   Begin VB.Label lbl 
      Caption         =   "F.&ENTREGA"
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   13
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "MANTENCIÓN"
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
      Width           =   1575
   End
   Begin VB.Label TotalPrecio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6720
      TabIndex        =   23
      Top             =   4560
      Width           =   1095
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
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OT_Mantencion.frx":011E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OT_Mantencion.frx":0230
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OT_Mantencion.frx":0342
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OT_Mantencion.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OT_Mantencion.frx":0566
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OT_Mantencion.frx":0678
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OT_Mantencion.frx":078A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   22
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Centro Costo"
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "&FECHA"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   10
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
Attribute VB_Name = "OT_Mantencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button

Private DbD As Database, RsSc As Recordset
Private Dbm As Database, RsCc As Recordset, RsOTc As Recordset, RsOTd As Recordset, RsTxt As Recordset
'Private RsNVc As Recordset, RsITOd As Recordset

Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, d As Variant

Private n_filas As Integer, n_columnas As Integer
Private prt As Printer

Private n1 As Double, n4 As Double
Private linea As String, m_NvArea As Integer
Private TipoDocOt As String ' tipo de documento para tabla ot, "M"
Private TipoDocTxt As String ' tipo docuento para tabla txt, "OTM"
Private m_KCentroCosto As String, a_CC(1, 99) As String
Private Sub Fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then CbCC.SetFocus
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
Set DbD = OpenDatabase(data_file)
Set RsSc = DbD.OpenRecordset("Contratistas")
RsSc.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)

Set RsCc = Dbm.OpenRecordset("TablasMaestras")
RsCc.Index = "tipo-codigo"

Set RsOTc = Dbm.OpenRecordset("OTe Cabecera")
RsOTc.Index = "Tipo-Numero"

Set RsOTd = Dbm.OpenRecordset("OTe Detalle")
RsOTd.Index = "Tipo-Numero-Linea"

Set RsTxt = Dbm.OpenRecordset("Detalle Texto")
RsTxt.Index = "Tipo-Numero-Linea"

'Set RsITOd = Dbm.OpenRecordset("ITOe")
'RsITOd.Index = "OT"

'Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
'RsNVc.Index = Nv_Index ' "Número"


' Combo centros de costo
i = 0
CbCC.AddItem " "
With RsCc
Do While Not .EOF
    CbCC.AddItem !codigo & " " & !descripcion
    i = i + 1
    a_CC(0, i) = !codigo
    a_CC(1, i) = !descripcion
    .MoveNext
Loop
End With

Inicializa
Detalle_Config

m_NvArea = 0

TipoDocOt = "M"
TipoDocTxt = "OTM"

m_KCentroCosto = "CCO"

Privilegios

End Sub
Private Sub Form_Unload(Cancel As Integer)
DataBases_Cerrar
End Sub
Private Sub Inicializa()

Obj = "OT MANTENCIÓN"
Objs = "OTS MANTENCIÓN"
TipoDocTxt = "OTM"

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

    RsOTc.Seek "=", TipoDocOt, Numero.Text
    
    If RsOTc.NoMatch Then
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

    RsOTc.Seek "=", TipoDocOt, Numero.Text
    
    If RsOTc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
    
        Botones_Enabled 0, 0, 0, 0, 1, 1
        
        Doc_Leer
        Campos_Enabled True
        Numero.Enabled = False
    End If

Case "Eliminando"
    RsOTc.Seek "=", TipoDocOt, Numero.Text
    If RsOTc.NoMatch Then
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
    
    RsOTc.Seek "=", TipoDocOt, Numero.Text
    If RsOTc.NoMatch Then
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
Fecha.Text = Format(RsOTc!Fecha, Fecha_Format)

For i = 1 To 99
    If a_CC(0, i) = RsOTc!CentroCosto Then
        CbCC.ListIndex = i
        Exit For
    End If
Next

Rut.Text = RsOTc![RUT Contratista]

'RsCc.Seek "=", m_KCentroCosto, RsOTc!CentroCosto
'If Not RsCc.NoMatch Then
'    CbCC.Text = RsCc!descripcion
'End If

Entrega.Text = Format(RsOTc![Fecha Entrega], Fecha_Format)

Obs(0).Text = NoNulo(RsOTc![Observacion 1])
Obs(1).Text = NoNulo(RsOTc![Observacion 2])
Obs(2).Text = NoNulo(RsOTc![Observacion 3])
Obs(3).Text = NoNulo(RsOTc![Observacion 4])

'DETALLE

RsOTd.Seek "=", TipoDocOt, Numero.Text, 1

If Not RsOTd.NoMatch Then

    Do While Not RsOTd.EOF
    
        If RsOTd!Numero = Numero.Text Then
        
            i = RsOTd!linea
            
            Detalle.TextMatrix(i, 1) = RsOTd!Cantidad
            Detalle.TextMatrix(i, 2) = RsOTd!unidad
            Detalle.TextMatrix(i, 3) = RsOTd!descripcion
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
    RsTxt.Seek "=", TipoDocTxt, Val(Numero.Text), i
    If Not RsTxt.NoMatch Then Detalle.TextMatrix(i, 3) = RsTxt!Texto
Next

Contratista_Lee Rut.Text
Detalle.Row = 1 ' para q' actualice la primera fila del detalle

Actualiza

End Sub
Private Sub Contratista_Lee(Rut)
RsSc.Seek "=", Rut
If Not RsSc.NoMatch Then
    Razon.Text = RsSc![Razon Social]
End If
End Sub
Private Function Doc_Validar() As Boolean
Dim porAsignar As Integer
Doc_Validar = False
If Trim(CbCC.Text) = "" Then
    MsgBox "DEBE ELEGIR CENTRO DE COSTO"
    CbCC.SetFocus
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
Dim m_cantidad As Double, m_pr As Double

save:
' CABECERA DE OT
If Nueva Then
    RsOTc.AddNew
    RsOTc!Tipo = TipoDocOt
    RsOTc!Numero = Numero.Text
Else
    RsOTc.Edit
End If

m_pr = OT_PRecibido(Numero.Text)

RsOTc!Fecha = Fecha.Text
'RsOTc!Nv = NVnumero.Caption
RsOTc!CentroCosto = a_CC(0, CbCC.ListIndex)
RsOTc![RUT Contratista] = Rut.Text
RsOTc![Fecha Entrega] = CDate(Entrega.Text)
'RsOTc!Montaje = OtMontaje.Value
'RsOTc!Tipo = Left(CbTipo.Text, 1)
RsOTc![Precio Total] = Val(TotalPrecio.Caption)
RsOTc![Porcentaje Recibido] = m_pr
RsOTc![Observacion 1] = Obs(0).Text
RsOTc![Observacion 2] = Obs(1).Text
RsOTc![Observacion 3] = Obs(2).Text
RsOTc![Observacion 4] = Obs(3).Text
RsOTc.Update

' DETALLE DE OT

Doc_Detalle_Eliminar

j = 0
For i = 1 To n_filas
    m_cantidad = Val(Detalle.TextMatrix(i, 1))
    If m_cantidad = 0 Then
        ' puede haber texto
        If Trim(Detalle.TextMatrix(i, 3)) <> "" Then
            j = j + 1
            RsTxt.AddNew
            RsTxt![Tipo Documento] = TipoDocTxt
            RsTxt![Numero Documento] = Numero.Text
            RsTxt!linea = j
            RsTxt!Texto = Detalle.TextMatrix(i, 3)
            RsTxt.Update
        End If
    Else
    
        j = j + 1
        
        RsOTd.AddNew
        RsOTd!Numero = Numero.Text
        RsOTd!Tipo = TipoDocOt
        RsOTd!linea = j
        RsOTd!Fecha = Fecha.Text
        
        RsOTd!CentroCosto = a_CC(0, CbCC.ListIndex)
        
        RsOTd![RUT Contratista] = Rut.Text
        RsOTd!Cantidad = m_cantidad
        RsOTd!unidad = Detalle.TextMatrix(i, 2)
        RsOTd!descripcion = Detalle.TextMatrix(i, 3)
        RsOTd![Precio Unitario] = m_CDbl(Detalle.TextMatrix(i, 4))
        RsOTd.Update
        
    End If
    
Next

End Sub
Private Function OT_PRecibido(N_Ot As Double) As Double
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
Private Sub Doc_Eliminar()

' borra CABECERA DE OT
RsOTc.Seek "=", TipoDocOt, Numero.Text

If Not RsOTc.NoMatch Then

    RsOTc.Delete

End If

Doc_Detalle_Eliminar

End Sub
Private Sub Doc_Detalle_Eliminar()

' elimina detalle
RsOTd.Seek "=", TipoDocOt, Numero.Text, 1

If Not RsOTd.NoMatch Then

    Do While Not RsOTd.EOF
    
        If RsOTd!Numero <> Numero.Text Or RsOTd!Tipo <> TipoDocOt Then Exit Do
    
        ' borra detalle
        RsOTd.Delete
    
        RsOTd.MoveNext
        
    Loop
    
End If

' elimina texto
RsTxt.Seek ">=", TipoDocTxt, Numero.Text, 1
If Not RsTxt.NoMatch Then

    Do While Not RsTxt.EOF
    
        If RsTxt![Tipo Documento] <> TipoDocTxt Or RsTxt![Numero Documento] <> Numero.Text Then Exit Do

        ' borra detalle
        RsTxt.Delete

        RsTxt.MoveNext
        
    Loop
    
End If

For i = 1 To n_filas
    RsTxt.Seek "=", TipoDocTxt, Numero.Text, i
    If Not RsTxt.NoMatch Then RsTxt.Delete
Next

End Sub
Private Sub Campos_Limpiar()
Numero.Text = ""
Fecha.Text = Fecha_Vacia
CbCC.Text = " "
Rut.Text = ""
Razon.Text = ""
Entrega.Text = Fecha_Vacia
'CbTipo.Text = " "
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
    
    Numero.Text = Documento_Numero_Nuevo(TipoDocOt)
    
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
'        Botones_Enabled 1, 1, 1, 1, 0, 0
        Campos_Limpiar
        Campos_Enabled False
    Else
        If Accion = "Imprimiendo" Then
            Privilegios
'            Botones_Enabled 1, 1, 1, 1, 0, 0
            Campos_Limpiar
            Campos_Enabled False
        Else
            If MsgBox("¿Seguro que quiere DesHacer?", vbYesNo) = vbYes Then
                Privilegios
'                Botones_Enabled 1, 1, 1, 1, 0, 0
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
btnSearch.Enabled = Si
CbCC.Enabled = Si
Entrega.Enabled = Si
Detalle.Enabled = Si
Obs(0).Enabled = Si
Obs(1).Enabled = Si
Obs(2).Enabled = Si
Obs(3).Enabled = Si
End Sub
Private Sub btnSearch_Click()

Search.Muestra data_file, "Contratistas", "RUT", "Razon Social", "contratista", "contratistas", "Activo"

Rut.Text = Search.codigo
If Rut.Text <> "" Then
    RsSc.Seek "=", Rut
    If RsSc.NoMatch Then
        MsgBox "CONTRATISTA NO EXISTE"
        Rut.SetFocus
    Else
        Razon.Text = Search.descripcion
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
Private Sub Doc_Imprimir(nCopias As Integer)
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
prt.Print Tab(tab0 + 25); "OT MANTENCION Nº";
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
prt.Print Tab(tab0 + tab40); "Centro de Costo : "
prt.Print Tab(tab0); "Teléfono: "; Empresa.Telefono1;
prt.Font.Size = fe
prt.Font.Bold = True
prt.Print Tab(tab0 + 28); CbCC.Text

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
    
        If Trim(Detalle.TextMatrix(i, 3)) <> "" Then
        
            ' DESCRIPCION
            prt.Print Tab(tab3); Left(Detalle.TextMatrix(i, 3), 30);
        
        Else
        
            j = j + 1
            prt.Print Tab(tab1 + j * 3); "  \"
        
        End If
        
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
Private Function Documento_Numero_Nuevo(TipoDoc As String) As Long
' busca nuevo correlativo para ot mantencion
Documento_Numero_Nuevo = 0
'On Error GoTo Sigue

If TipoDoc = "M" Then ' ot mantencion

    Do While Not RsOTc.EOF
        
        If RsOTc!Tipo = "M" Then ' ot mantencion
            Documento_Numero_Nuevo = RsOTc!Numero
        End If
        
        RsOTc.MoveNext
        
    Loop
    
End If
Sigue:
On Error GoTo 0
Documento_Numero_Nuevo = Documento_Numero_Nuevo + 1
End Function
Private Sub Privilegios()
If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 0, 0
End If
End Sub
