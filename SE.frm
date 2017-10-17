VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form SE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servicios Externos"
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
      TabIndex        =   0
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
            Object.ToolTipText     =   "Mantención de Clientes"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      MouseIcon       =   "SE.frx":0000
   End
   Begin MSMask.MaskEdBox EditFecha 
      Height          =   495
      Left            =   8040
      TabIndex        =   35
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      _Version        =   327680
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox EspesorTotal 
      Height          =   300
      Left            =   5880
      TabIndex        =   14
      Top             =   1650
      Width           =   700
   End
   Begin VB.TextBox Nv 
      Height          =   300
      Left            =   960
      TabIndex        =   11
      Top             =   1650
      Width           =   615
   End
   Begin VB.Frame Frame_Terminacion 
      Caption         =   "Terminacion"
      Height          =   1095
      Left            =   5040
      TabIndex        =   21
      Top             =   2040
      Width           =   2295
      Begin VB.TextBox T_NManos 
         Height          =   300
         Left            =   1080
         TabIndex        =   24
         Top             =   700
         Width           =   400
      End
      Begin VB.ComboBox CbTerminacion 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lbl 
         Caption         =   "Nº Manos"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame_Anticorrosivo 
      Caption         =   "Anticorrosivo"
      Height          =   1095
      Left            =   2640
      TabIndex        =   17
      Top             =   2040
      Width           =   2295
      Begin VB.TextBox A_NManos 
         Height          =   300
         Left            =   1080
         TabIndex        =   20
         Top             =   720
         Width           =   400
      End
      Begin VB.ComboBox CbAnticorrosivo 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lbl 
         Caption         =   "Nº Manos"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame_Granallado 
      Caption         =   "Granallado"
      Height          =   855
      Left            =   240
      TabIndex        =   15
      Top             =   2040
      Width           =   2295
      Begin VB.ComboBox CbGranallado 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.ComboBox ComboNV 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1650
      Width           =   1575
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   3
      Left            =   120
      MaxLength       =   30
      TabIndex        =   31
      Top             =   5700
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   2
      Left            =   120
      MaxLength       =   30
      TabIndex        =   30
      Top             =   5400
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   1
      Left            =   120
      MaxLength       =   30
      TabIndex        =   29
      Top             =   5100
      Width           =   5000
   End
   Begin VB.Frame Frame_Cliente 
      Caption         =   "Cliente"
      Height          =   1095
      Left            =   2400
      TabIndex        =   4
      Top             =   480
      Width           =   4575
      Begin VB.CommandButton btnSearch 
         Height          =   300
         Left            =   2040
         Picture         =   "SE.frx":001C
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
         Width           =   1335
         _ExtentX        =   2355
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
   Begin VB.TextBox txtEdit 
      Height          =   285
      Left            =   8040
      TabIndex        =   32
      Text            =   "txtEdit"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   0
      Left            =   120
      MaxLength       =   30
      TabIndex        =   28
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
      TabIndex        =   25
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
      Height          =   1245
      Left            =   120
      TabIndex        =   27
      Top             =   3240
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2196
      _Version        =   327680
      ScrollBars      =   2
   End
   Begin VB.Label lbl 
      Caption         =   "Espesor Total Esquema"
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   13
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label TotalPrecio 
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
            Picture         =   "SE.frx":011E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SE.frx":0230
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SE.frx":0342
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SE.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SE.frx":0566
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SE.frx":0678
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SE.frx":078A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   33
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "&OBRA"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "&FECHA"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   26
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "Servicio Externo"
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
      TabIndex        =   1
      Top             =   480
      Width           =   1935
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
Attribute VB_Name = "SE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button

Private DbD As Database, RsCl As Recordset
Private Dbm As Database, RsSEc As Recordset, RsSEd As Recordset
Private RsNVc As Recordset

Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, d As Variant

Private n_filas As Integer, n_columnas As Integer
Private prt As Printer
Private n3 As Double, n4 As Double
Private linea As String
Private m_Nv As Integer, m_NvArea As Integer
' 0: numero,  1: nombre obra
Private a_Nv(2999, 1) As String
Private Sub Form_Load()

' abre archivos
Set DbD = OpenDatabase(data_file)
Set RsCl = DbD.OpenRecordset("Clientes")
RsCl.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)

Set RsSEc = Dbm.OpenRecordset("SE Cabecera")
RsSEc.Index = "Numero"

Set RsSEd = Dbm.OpenRecordset("SE Detalle")
RsSEd.Index = "Numero-Linea"

Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

CbGranallado.AddItem "SIN GRANALLADO"
CbGranallado.AddItem "BRUSH OFF"
CbGranallado.AddItem "SP10"
CbGranallado.AddItem "SP5"
CbGranallado.AddItem "SP6"

CbAnticorrosivo.AddItem "SIN ANTICORROSIVO"
CbAnticorrosivo.AddItem "EPOXICO"
CbAnticorrosivo.AddItem "ALQUIDICO"
CbAnticorrosivo.AddItem "INORGANICO"

CbTerminacion.AddItem "SIN TERMINACION"
CbTerminacion.AddItem "EPOXICO"
CbTerminacion.AddItem "ALQUIDICO"

' Combo obra
ComboNV.AddItem " "
i = 0
Do While Not RsNVc.EOF
    i = i + 1
    a_Nv(i, 0) = RsNVc!Número
    a_Nv(i, 1) = RsNVc!Obra
    ComboNV.AddItem Format(RsNVc!Número, "0000") & " - " & RsNVc!Obra
    RsNVc.MoveNext
Loop

Inicializa
Detalle_Config

Privilegios

m_NvArea = 0

End Sub
Private Sub Form_Unload(Cancel As Integer)
DataBases_Cerrar
End Sub
Private Sub Inicializa()

Obj = "SERVICIO EXTERNO"
Objs = "SERVICIOS EXTERNOS"

Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnDesHacer = Toolbar.Buttons(6)
Set btnGrabar = Toolbar.Buttons(7)

Accion = ""
'old_accion = ""

EspesorTotal.MaxLength = 5
A_NManos.MaxLength = 1
T_NManos.MaxLength = 1

'btnSearch.Visible = False
btnSearch.ToolTipText = "Busca Cliente"
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
Detalle.TextMatrix(0, 1) = "Fecha Recepcion"
Detalle.TextMatrix(0, 2) = "Descripción"
Detalle.TextMatrix(0, 3) = "Cantidad"
Detalle.TextMatrix(0, 4) = "m2 Uni"
Detalle.TextMatrix(0, 5) = "m2 TOTAL"

Detalle.ColWidth(0) = 250
Detalle.ColWidth(1) = 1000
Detalle.ColWidth(2) = 2500
Detalle.ColWidth(3) = 1000
Detalle.ColWidth(4) = 1200
Detalle.ColWidth(5) = 1300

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
'    Detalle.col = 2
'    Detalle.CellAlignment = flexAlignLeftCenter
'    Detalle.col = 3
'    Detalle.CellAlignment = flexAlignLeftCenter
    Detalle.col = 5
    Detalle.CellForeColor = vbRed
Next

txtEdit.Text = ""
EditFecha.Text = "__/__/__"

End Sub
Private Sub Numero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then After_Enter
End Sub
Private Sub ComboNV_Click()

'Nv.Text = Val(Left(ComboNV.Text, 6))
m_Nv = Val(Left(ComboNV.Text, 6))
If m_Nv = 0 Then
    Nv.Text = ""
Else
    Nv.Text = m_Nv
End If

End Sub
Private Sub Fecha_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Fecha_LostFocus()
d = Fecha_Valida(Fecha, Now)
End Sub
Private Sub After_Enter()
If Val(Numero.Text) < 1 Then Beep: Exit Sub
Select Case Accion
Case "Agregando"
    RsSEc.Seek "=", Numero.Text
    If RsSEc.NoMatch Then
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
    RsSEc.Seek "=", Numero
    If RsSEc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Botones_Enabled 0, 0, 0, 0, 1, 1
        Doc_Leer
        Campos_Enabled True
        Numero.Enabled = False
    End If

Case "Eliminando"
    RsSEc.Seek "=", Numero
    If RsSEc.NoMatch Then
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
    
    RsSEc.Seek "=", Numero
    If RsSEc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Doc_Leer
        Numero.Enabled = False
        Detalle.Enabled = True
    End If
    
End Select

End Sub
Private Sub Doc_Leer()

' CABECERA

With RsSEc

Fecha.Text = Format(!Fecha, Fecha_Format)
Nv.Text = !Nv
Rut.Text = NoNulo(![Rut_Cliente])

RsNVc.Seek "=", Nv.Text, m_NvArea
If Not RsNVc.NoMatch Then
    ComboNV.Text = Format(RsNVc!Número, "0000") & " - " & RsNVc!Obra
    ComboNV_Click
End If

EspesorTotal.Text = !Espesor_Total

Select Case !Granallado
Case "OFF"
    CbGranallado.Text = "BRUSH OFF"
Case "SP10"
    CbGranallado.Text = "SP10"
Case "SP5"
    CbGranallado.Text = "SP5"
Case "SP6"
    CbGranallado.Text = "SP6"
Case Else
    CbGranallado.Text = "SIN GRANALLADO"
End Select

Select Case !Anticorrosivo
Case "E"
    CbAnticorrosivo.Text = "EPOXICO"
Case "A"
    CbAnticorrosivo.Text = "ALQUIDICO"
Case "I"
    CbAnticorrosivo.Text = "INORGANICO"
Case Else
    CbAnticorrosivo.Text = "SIN ANTICORROSIVO"
End Select
A_NManos.Text = !A_NumeroManos

Select Case !Terminacion
Case "E"
    CbTerminacion.Text = "EPOXICO"
Case "A"
    CbTerminacion.Text = "ALQUIDICO"
Case Else
    CbTerminacion.Text = "SIN TERMINACION"
End Select
T_NManos.Text = !T_NumeroManos

Obs(0).Text = NoNulo(![Observacion1])
Obs(1).Text = NoNulo(![Observacion2])
Obs(2).Text = NoNulo(![Observacion3])
Obs(3).Text = NoNulo(![Observacion4])

End With

'DETALLE

With RsSEd
.Seek ">=", Numero.Text, 0
If Not .NoMatch Then
    Do While Not .EOF
    
        If !Numero = Numero.Text Then
        
            i = !linea
            
            Detalle.TextMatrix(i, 1) = !fecha_recepcion
            Detalle.TextMatrix(i, 2) = !Descripcion
            Detalle.TextMatrix(i, 3) = !Cantidad
            Detalle.TextMatrix(i, 4) = !m2_Unitario
            
            n3 = m_CDbl(Detalle.TextMatrix(i, 3))
            n4 = m_CDbl(Detalle.TextMatrix(i, 4))
            
            Detalle.TextMatrix(i, 5) = Format(n3 * n4, num_fmtgrl)
            
        Else
        
            Exit Do
            
        End If
        .MoveNext
    Loop
End If
End With

Cliente_Lee Rut.Text
Detalle.Row = 1 ' para q' actualice la primera fila del detalle

Actualiza

End Sub
Private Sub Cliente_Lee(Rut)
RsCl.Seek "=", Rut
If Not RsCl.NoMatch Then
    Razon.Text = RsCl![Razón Social]
End If
End Sub
Private Function Doc_Validar() As Boolean
Dim porAsignar As Integer
Doc_Validar = False
If Trim(ComboNV.Text) = "" Then
    MsgBox "DEBE ELEGIR OBRA"
    ComboNV.SetFocus
    Exit Function
End If
If Rut.Text = "" Then
    MsgBox "DEBE ELEGIR CLIENTE"
    btnSearch.SetFocus
    Exit Function
End If

For i = 1 To n_filas

    ' cant
    If Val(Detalle.TextMatrix(i, 1)) <> 0 Then
    
        If Not CampoReq_Valida(Detalle.TextMatrix(i, 2), i, 2) Then Exit Function
        If Not LargoString_Valida(Detalle.TextMatrix(i, 2), 20, i, 2) Then Exit Function
        
        If Not CampoReq_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
        If Not LargoString_Valida(Detalle.TextMatrix(i, 3), 10, i, 3) Then Exit Function
        
        If Not Numero_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
        
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

Dim m_cantidad As Double

save:
With RsSEc
' CABECERA
If Nueva Then
    .AddNew
    !Numero = Numero.Text
Else
    .Edit
End If

!Fecha = Fecha.Text
!Nv = Nv.Text
![Rut_Cliente] = Rut.Text

Select Case CbGranallado.Text
Case "BRUSH OFF"
    !Granallado = "OFF"
Case "SP10"
    !Granallado = "SP10"
Case "SP5"
    !Granallado = "SP5"
Case "SP6"
    !Granallado = "SP6"
Case Else
    !Granallado = ""
End Select

If CbAnticorrosivo.Text = "S" Then
    !Anticorrosivo = ""
Else
    !Anticorrosivo = Left(CbAnticorrosivo.Text, 1)
End If
!A_NumeroManos = A_NManos.Text

If CbTerminacion.Text = "S" Then
    !Terminacion = ""
Else
    !Terminacion = Left(CbTerminacion.Text, 1)
End If
!T_NumeroManos = T_NManos.Text

!Espesor_Total = Val(EspesorTotal.Text)

![Observacion1] = Obs(0).Text
![Observacion2] = Obs(1).Text
![Observacion3] = Obs(2).Text
![Observacion4] = Obs(3).Text
.Update

End With

' DETALLE

Doc_Detalle_Eliminar

With RsSEd
j = 0
For i = 1 To n_filas

    m_cantidad = Val(Detalle.TextMatrix(i, 3))
    
    If m_cantidad > 0 Then
    
        j = j + 1
        
        .AddNew
        !Numero = Numero.Text
        !linea = j
        !Fecha = Fecha.Text
        !fecha_recepcion = Detalle.TextMatrix(i, 1)
        !Nv = Nv.Text
        ![Rut_Cliente] = Rut.Text
        !Descripcion = Detalle.TextMatrix(i, 2)
        !Cantidad = m_cantidad
        !m2_Unitario = m_CDbl(Detalle.TextMatrix(i, 4))
        .Update
    
    End If
        
Next
End With

End Sub
Private Sub Doc_Eliminar()

' borra CABECERA
RsSEc.Seek "=", Numero.Text
If Not RsSEc.NoMatch Then

    RsSEc.Delete

End If

Doc_Detalle_Eliminar

End Sub
Private Sub Doc_Detalle_Eliminar()

' elimina detalle
With RsSEd
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
Private Sub Campos_Limpiar()

Numero.Text = ""
Fecha.Text = Fecha_Vacia
Nv.Text = ""
ComboNV.Text = " "
Rut.Text = ""
Razon.Text = ""

EspesorTotal.Text = ""
A_NManos.Text = ""
T_NManos.Text = ""

CbGranallado.ListIndex = 0
CbAnticorrosivo.ListIndex = 0
CbTerminacion.ListIndex = 0

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
    
    Numero.Text = Documento_Numero_Nuevo(RsSEc, "Numero")
    
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
            Doc_Imprimir n_Copias
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
                Doc_Imprimir n_Copias
            End If
            
            Botones_Enabled 0, 0, 0, 0, 1, 0
            Campos_Limpiar
            Campos_Enabled False
            Numero.Enabled = True
            Numero.SetFocus
            
        End If
    End If
Case 8 ' Separador
Case 9 ' clientes
    MousePointer = 11
    Load Clientes
    MousePointer = 0
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
Nv.Enabled = Si
ComboNV.Enabled = Si

EspesorTotal.Enabled = Si

Frame_Granallado.Enabled = Si
CbGranallado.Enabled = Si

Frame_Anticorrosivo.Enabled = Si
CbAnticorrosivo.Enabled = Si
A_NManos.Enabled = Si

Frame_Terminacion.Enabled = Si
CbTerminacion.Enabled = Si
T_NManos.Enabled = Si

Detalle.Enabled = Si

Obs(0).Enabled = Si
Obs(1).Enabled = Si
Obs(2).Enabled = Si
Obs(3).Enabled = Si

End Sub
Private Sub btnSearch_Click()

Search.Muestra data_file, "Clientes", "RUT", "Razon Social", "Cliente", "Clientes"

Rut.Text = Search.Codigo
If Rut.Text <> "" Then
    RsCl.Seek "=", Rut
    If RsCl.NoMatch Then
        MsgBox "CLIENTE NO EXISTE"
        Rut.SetFocus
    Else
        Razon.Text = Search.Descripcion
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
Case 1 ' fecha recep
Case 2 ' descr
Case 3 ' cant
Case 4 ' m2 uni
Case 5 ' m2 tot
Case Else
End Select
End Sub
Private Sub Detalle_DblClick()
If Accion = "Imprimiendo" Then Exit Sub
' simula un espacio
If Detalle.col = 1 Then
    MSFlexGridEdit Detalle, EditFecha, 32  'FECHA
Else
    MSFlexGridEdit Detalle, txtEdit, 32
End If
End Sub
Private Sub Detalle_GotFocus()
If Accion = "Imprimiendo" Then Exit Sub
Select Case True
Case txtEdit.visible
    Detalle = txtEdit
    txtEdit.visible = False
Case EditFecha.visible
    Detalle = EditFecha
    EditFecha.visible = False
End Select
End Sub
Private Sub Detalle_LeaveCell()
If Accion = "Imprimiendo" Then Exit Sub
Select Case True
Case txtEdit.visible
    Detalle = txtEdit
    txtEdit.visible = False
Case EditFecha.visible
    Detalle = EditFecha
    EditFecha.visible = False
End Select
End Sub
Private Sub Detalle_KeyPress(KeyAscii As Integer)
If Accion = "Imprimiendo" Then Exit Sub
If Detalle.col = 1 Then
    MSFlexGridEdit Detalle, EditFecha, KeyAscii
Else
    MSFlexGridEdit Detalle, txtEdit, KeyAscii
End If
End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCodeP Detalle, txtEdit, KeyCode, Shift
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
    Case 1 ' Fecha
'        Detalle.SetFocus
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
Private Sub txtEdit_KeyPress(KeyAscii As Integer)
'If KeyAscii = vbKeyReturn Then KeyAscii = 0
'If Detalle.col = 1 Then
'    MSFlexGridEdit Detalle, EditFecha, KeyAscii
'Else
'    MSFlexGridEdit Detalle, txtEdit, KeyAscii
'End If
End Sub
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)

If MSFlexGrid.col = 5 Then Exit Sub

Select Case MSFlexGrid.col
Case 1
    Edt.MaxLength = 8
Case 2
    Edt.MaxLength = 20
Case Else
    Edt.MaxLength = 10
End Select

If MSFlexGrid.col = 1 Then ' fecha

    Select Case KeyAscii
    Case 0 To 32
        If MSFlexGrid = "" Then
            Edt = "__/__/__"
            Edt.SelStart = 0
        Else
            Edt = Format(MSFlexGrid, "dd/mm/yy")
            Edt.SelStart = 1000
        End If
    Case 48 To 51 ' "0" al "3"
        Edt = Chr(KeyAscii) & "_/__/__"
        Edt.SelStart = 1
    End Select
    Edt.Move MSFlexGrid.CellLeft + MSFlexGrid.Left, MSFlexGrid.CellTop + MSFlexGrid.Top, MSFlexGrid.CellWidth, MSFlexGrid.CellHeight + 50
    Edt.visible = True
    Edt.SetFocus
    
Else
    
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
    
End If

End Sub
Private Sub Detalle_KeyDown(KeyCode As Integer, Shift As Integer)
If Accion = "Imprimiendo" Then Exit Sub
If KeyCode = vbKeyF2 Then
    MSFlexGridEdit Detalle, txtEdit, 32
End If
End Sub
Private Sub Linea_Actualiza()
' actualiza solo linea, y totales generales
Dim fi As Integer, co As Integer

fi = Detalle.Row
co = Detalle.col

n3 = m_CDbl(Detalle.TextMatrix(fi, 3))
n4 = m_CDbl(Detalle.TextMatrix(fi, 4))

' precio total
Detalle.TextMatrix(fi, 5) = Format(n3 * n4, num_fmtgrl)

Totales_Actualiza

End Sub
Private Sub Actualiza()
' actualiza todo el detalle y totales generales
Dim fi As Integer
For fi = 1 To n_filas
    If Detalle.TextMatrix(fi, 1) <> "" Then
    
        n3 = m_CDbl(Detalle.TextMatrix(fi, 3))
        n4 = m_CDbl(Detalle.TextMatrix(fi, 4))

        ' precio total
        Detalle.TextMatrix(fi, 5) = Format(n3 * n4, num_fmtgrl)
        
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

Private Sub EditFecha_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCodeP Detalle, EditFecha, KeyCode, Shift
End Sub

Private Sub Doc_Imprimir(nCopias As Integer)
MousePointer = vbHourglass
linea = String(70, "-")
' imprime
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer
Dim tab4 As Integer, tab5 As Integer, tab40 As Integer
Dim fc As Integer, fn As Integer, fe As Integer
Dim n_Copia As Integer

tab0 = 3 ' margen izquierdo
tab1 = tab0 + 0  ' fecha recepcion
tab2 = tab1 + 9  ' descripcion
tab3 = tab2 + 30 ' cantidad
tab4 = tab3 + 10 ' m2 uni
tab5 = tab4 + 10 ' m2 total
tab40 = 43

' font.size
fc = 10 'comprimida
fn = 12 'normal
fe = 15 'expandida

Dim can_valor As String

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt

For n_Copia = 1 To nCopias

prt.Font.Size = fe
prt.Print Tab(tab0 - 1); m_TextoIso
prt.Print Tab(tab0 + 25); "SERVICIO EXTERNO Nº";
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
prt.Print Tab(tab0 + tab40); "CLIENTE :"
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

prt.Print Tab(tab0); "Espesor Total Esquema "; EspesorTotal.Text
prt.Print Tab(tab0); "Granallado   : " & CbGranallado.Text
prt.Print Tab(tab0); "Anticorrosivo: " & CbAnticorrosivo.Text; Tab(tab3); "NºManos " & A_NManos.Text
prt.Print Tab(tab0); "Terminación  : " & CbTerminacion.Text; Tab(tab3); ; "NºManos " & T_NManos.Text

prt.Print ""
prt.Print ""

' detalle
prt.Font.Bold = True
prt.Print Tab(tab1); "RECEP.";
prt.Print Tab(tab2); "DESCRIPCIÓN";
prt.Print Tab(tab3); "  CANTIDAD";
prt.Print Tab(tab4); "   m2 UNI";
prt.Print Tab(tab5); "   m2 TOT"
prt.Font.Bold = False

prt.Print Tab(tab1); linea
j = -1
For i = 1 To n_filas

    can_valor = m_CDbl(Detalle.TextMatrix(i, 3))
    
    If can_valor = 0 Then
    
        j = j + 1
        prt.Print Tab(tab1 + j * 3); "  \"
        
    Else
    
        ' fecha recp
        prt.Print Tab(tab1); Detalle.TextMatrix(i, 1);
        
        ' DESCRIPCION
        prt.Print Tab(tab2); Left(Detalle.TextMatrix(i, 2), 30);
        
        ' cant
        prt.Print Tab(tab3); m_Format(can_valor, "##,###,##0");
        
        ' m2 UNITARIO
        prt.Print Tab(tab4); m_Format(Detalle.TextMatrix(i, 4), "#,###,###");
        
        ' m2 TOTAL
        prt.Print Tab(tab5); m_Format(Detalle.TextMatrix(i, 5), "##,###,###")
        
    End If
    
Next

prt.Print Tab(tab1); linea
prt.Font.Bold = True
prt.Print Tab(tab5); m_Format(TotalPrecio, "##,###,###")
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

prt.Print Tab(tab0); Tab(14), "__________________", Tab(55), "__________________"
prt.Print Tab(tab0); Tab(14), "       VºBº       ", Tab(55), "       VºBº       "

If n_Copia < nCopias Then
    prt.NewPage
Else
    prt.EndDoc
End If

Next

Impresora_Predeterminada "default"

MousePointer = vbDefault

End Sub
Private Sub Privilegios()
If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 0, 0
End If
End Sub
