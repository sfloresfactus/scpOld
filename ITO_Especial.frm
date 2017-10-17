VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ITO_Especial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vales ITO"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   5610
      _ExtentX        =   9895
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
            Object.ToolTipText     =   "Ingresar Nueva ITO"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Modificar ITO"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Eliminar ITO"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Imprimir ITO"
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
            Object.ToolTipText     =   "Grabar ITO"
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
      MouseIcon       =   "ITO_Especial.frx":0000
   End
   Begin VB.Frame Frame 
      Caption         =   "O&T"
      Height          =   3135
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   4575
      Begin VB.ComboBox ComboOT 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   3855
      End
      Begin MSFlexGridLib.MSFlexGrid Detalle 
         Height          =   1815
         Left            =   600
         TabIndex        =   19
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3201
         _Version        =   327680
      End
      Begin VB.Label Total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   3120
         TabIndex        =   21
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "Total % Recibido"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   20
         Top             =   2760
         Width           =   1215
      End
   End
   Begin MSMask.MaskEdBox aRecibir 
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   4440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   327680
      ForeColor       =   255
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox ComboNV 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   3
      Left            =   240
      MaxLength       =   30
      TabIndex        =   14
      Top             =   5805
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   2
      Left            =   240
      MaxLength       =   30
      TabIndex        =   13
      Top             =   5505
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   1
      Left            =   240
      MaxLength       =   30
      TabIndex        =   12
      Top             =   5205
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   0
      Left            =   240
      MaxLength       =   30
      TabIndex        =   11
      Top             =   4905
      Width           =   5000
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   2160
      TabIndex        =   7
      Top             =   4440
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
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "##########"
      PromptChar      =   "_"
   End
   Begin VB.Label lbl 
      Caption         =   "% a &Recibir"
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   4
      Left            =   3360
      TabIndex        =   8
      Top             =   4440
      Width           =   735
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
      Index           =   7
      Left            =   840
      TabIndex        =   18
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lbl 
      Caption         =   "ITO"
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
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   495
   End
   Begin VB.Label NVnumero 
      Caption         =   "NVnumero"
      Height          =   255
      Left            =   6600
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   6600
      Top             =   360
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
            Picture         =   "ITO_Especial.frx":001C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ITO_Especial.frx":012E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ITO_Especial.frx":0240
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ITO_Especial.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ITO_Especial.frx":0464
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ITO_Especial.frx":0576
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ITO_Especial.frx":0688
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "O&bservaciones"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   10
      Top             =   4635
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "&OBRA"
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "&FECHA"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   6
      Top             =   4440
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
      TabIndex        =   0
      Top             =   840
      Width           =   375
   End
End
Attribute VB_Name = "ITO_Especial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button

Private DbD As Database, RsSc As Recordset
Private Dbm As Database, RsNVc As Recordset, RsOTc As Recordset, RsITO As Recordset

Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, d As Variant
Private n_filas As Integer, n_columnas As Integer
Private prt As Printer
Private a_Rut(1999) As String, m_Rut As String, m_Razon As String, m_ot As Double
Private m_NvArea As Integer
Private Sub Form_Load()

n_filas = 6
n_columnas = 3

' abre archivos
Set DbD = OpenDatabase(data_file)
Set RsSc = DbD.OpenRecordset("Contratistas")
RsSc.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index '"Número"

Set RsOTc = Dbm.OpenRecordset("OT Esp Cabecera")
RsOTc.Index = "Numero"

Set RsITO = Dbm.OpenRecordset("ITO Esp")
RsITO.Index = "Numero"

' Combo obra
ComboNv.AddItem " "

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
        ComboNv.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
    End If
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
Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnDesHacer = Toolbar.Buttons(6)
Set btnGrabar = Toolbar.Buttons(7)

Obj = "VALE ITO"
Objs = "VALES ITO"

Accion = ""
'old_accion = ""

Campos_Enabled False
End Sub
Private Sub Detalle_Config()

Dim i As Integer, ancho As Integer

'Detalle.Left = 100
Detalle.WordWrap = True
Detalle.RowHeight(0) = 450
Detalle.Rows = n_filas + 1
Detalle.Cols = n_columnas + 1

Detalle.TextMatrix(0, 0) = ""
Detalle.TextMatrix(0, 1) = "Nº ITO"
Detalle.TextMatrix(0, 2) = " Fecha "
Detalle.TextMatrix(0, 3) = "% Recib"

Detalle.ColWidth(0) = 250
Detalle.ColWidth(1) = 1000
Detalle.ColWidth(2) = 1000
Detalle.ColWidth(3) = 1000

ancho = 350  ' con scroll vertical

For i = 0 To n_columnas
    ancho = ancho + Detalle.ColWidth(i)
Next

Detalle.Width = ancho

For i = 1 To n_filas
    Detalle.TextMatrix(i, 0) = i
Next

Detalle.Enabled = False

End Sub
Private Sub Numero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    After_Enter
End If
End Sub
Private Sub After_Enter()
If Val(Numero.Text) < 1 Then Beep: Exit Sub
Select Case Accion
Case "Agregando"
    RsITO.Seek "=", Numero
    If RsITO.NoMatch Then
    
        Campos_Enabled True
        Numero.Enabled = False
        Detalle.Enabled = False
        
'        Fecha.Text = Format(Now, Fecha_Format)
'        If Usuario.AccesoTotal Then
'            Fecha.SetFocus
'        Else
'            ComboNV.SetFocus
'        End If
        ComboNv.SetFocus
        
        btnGrabar.Enabled = True
    Else
        ITO_Leer
        MsgBox Obj & " YA EXISTE"
        Detalle.Enabled = False
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
    End If
Case "Modificando"
    RsITO.Seek "=", Numero
    If RsITO.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        Numero.SetFocus
    Else
        Botones_Enabled 0, 0, 0, 0, 1, 1
        ITO_Leer
        Campos_Enabled True
        Numero.Enabled = False
        
'        If Usuario.AccesoTotal Then
'            Fecha.SetFocus
'        Else
'            ComboNV.SetFocus
'        End If
        ComboNv.SetFocus
        
        btnGrabar.Enabled = True
        
    End If
Case "Eliminando"
    RsITO.Seek "=", Numero
    If RsITO.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        Numero.SetFocus
    Else
        Campos_Enabled False
        ITO_Leer
        If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo, "Atención") = vbYes Then
            ITO_Eliminar
        End If
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
    End If
Case "Imprimiendo"
    
    RsITO.Seek "=", Numero
    If RsITO.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        ITO_Leer
        Numero.Enabled = False
        Detalle.Enabled = True
    End If
End Select

End Sub
Private Sub Fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then aRecibir.SetFocus
End Sub
Private Sub Fecha_LostFocus()
d = Fecha_Valida(Fecha, Now)
End Sub
Private Sub ComboNV_Click()
MousePointer = vbHourglass
Dim m_sc As String, j As Integer
NvNumero.Caption = Val(Left(ComboNv.Text, 4))

j = 0
a_Rut(0) = ""
ComboOT.Clear
ComboOT.AddItem " "

RsOTc.MoveFirst
Do While Not RsOTc.EOF
    If RsOTc!Nv = NvNumero.Caption Then
        j = j + 1
        a_Rut(j) = RsOTc![RUT Contratista]
        m_sc = Contratista_Leer(RsOTc![RUT Contratista])
        ComboOT.AddItem Format(RsOTc!Numero, "0000") & ", " & m_sc
    End If
    RsOTc.MoveNext
Loop
Detalle_Limpiar

MousePointer = vbDefault
End Sub
Private Sub ComboOT_Click()
m_Rut = a_Rut(ComboOT.ListIndex)
m_Razon = Mid(ComboOT.Text, 7)
m_ot = Val(Left(ComboOT.Text, 4))
ITOsdeOT
End Sub
Private Sub ITOsdeOT()
Dim fi As Integer, m_Total As Double
' lee itos recibidas para esta ot
RsITO.Index = "OT"
fi = 0
m_Total = 0
Detalle_Limpiar
RsITO.Seek "=", m_ot
If Not RsITO.NoMatch Then
    Do While Not RsITO.EOF
        If RsITO!OT <> m_ot Then Exit Do
        fi = fi + 1
        Detalle.TextMatrix(fi, 1) = RsITO!Número
        Detalle.TextMatrix(fi, 2) = RsITO!Fecha
        Detalle.TextMatrix(fi, 3) = RsITO![Porcentaje Recibido]
        m_Total = m_Total + RsITO![Porcentaje Recibido]
        RsITO.MoveNext
    Loop
End If
Total.Caption = m_Total

RsITO.Index = "Numero"
End Sub
Private Sub ITO_Leer()
Fecha.Text = Format(RsITO!Fecha, Fecha_Format)
NvNumero.Caption = RsITO!Nv

RsNVc.Seek "=", NvNumero.Caption, m_NvArea
If Not RsNVc.NoMatch Then
    ComboNv.Text = Format(RsNVc!Número, "0000") & " - " & RsNVc!obra
End If

m_Rut = RsITO![RUT Contratista]
m_Razon = Contratista_Leer(m_Rut)

aRecibir.Text = Left(Format(RsITO![Porcentaje Recibido]) + "___", 3)
Obs(0).Text = NoNulo(RsITO![Observación 1])
Obs(1).Text = NoNulo(RsITO![Observación 2])
Obs(2).Text = NoNulo(RsITO![Observación 3])
Obs(3).Text = NoNulo(RsITO![Observación 4])

ComboNV_Click
RsOTc.Seek "=", RsITO!OT
If Not RsOTc.NoMatch Then
    ComboOT.Text = Format(RsOTc!Número, "0000") & ", " & m_Razon
End If

Detalle.Row = 1 ' para q' actualice la primera fila del detalle

End Sub
Private Function Contratista_Leer(Rut) As String
Dim txt As String
txt = ""
RsSc.Seek "=", Rut
If Not RsSc.NoMatch Then txt = RsSc![Razon Social]
Contratista_Leer = txt
End Function
Private Function ITO_Validar() As Boolean
ITO_Validar = False

If m_Rut = "" Then
    MsgBox "DEBE ELEGIR OT"
    Exit Function
End If

If Val(Total.Caption) + Val(Replace(aRecibir.Text, "_")) > 100 Then
    MsgBox "Valor a Recibir Erróneo", , "Error"
    Exit Function
End If

ITO_Validar = True

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
    If Val(num) < 0 Then
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
Private Sub ITO_Grabar(Nueva As Boolean)
MousePointer = vbHourglass
Dim m_Total As Double
save:
' CABECERA DE ITO
If Nueva Then
    RsITO.AddNew
    RsITO!Número = Numero.Text
Else
    RsITO.Edit
End If
RsITO!Fecha = Fecha.Text
RsITO!Nv = NvNumero.Caption
RsITO![RUT Contratista] = m_Rut
RsITO!OT = m_ot
RsITO![Porcentaje Recibido] = Replace(aRecibir.Text, "_")
RsITO![Observación 1] = Obs(0).Text
RsITO![Observación 2] = Obs(1).Text
RsITO![Observación 3] = Obs(2).Text
RsITO![Observación 4] = Obs(3).Text
RsITO.Update

' actualiza porcentaje recibido en OT
'//////////////////////
RsITO.Index = "OT"
m_Total = 0
RsITO.Seek "=", m_ot
If Not RsITO.NoMatch Then
    Do While Not RsITO.EOF
        If RsITO!OT <> m_ot Then Exit Do
        m_Total = m_Total + RsITO![Porcentaje Recibido]
        RsITO.MoveNext
    Loop
End If
RsOTc.Seek "=", m_ot
If Not RsOTc.NoMatch Then
    RsOTc.Edit
    RsOTc![Porcentaje Recibido] = m_Total
    RsOTc.Update
End If
RsITO.Index = "Número"
'/////////////////////

MousePointer = vbDefault

End Sub
Private Sub ITO_Eliminar()

' borra CABECERA DE ITO
RsITO.Seek "=", Numero.Text
If Not RsITO.NoMatch Then

    RsITO.Delete
   
End If

End Sub
Private Sub Campos_Limpiar()
Numero.Text = ""
Fecha.Text = Fecha_Vacia
'Fecha.Text = Format(Now, Fecha_Format)
ComboNv.Text = " "
ComboOT.Text = " "
m_Rut = ""
m_Razon = ""
aRecibir = "___"
Detalle_Limpiar
Obs(0).Text = ""
Obs(1).Text = ""
Obs(2).Text = ""
Obs(3).Text = ""
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
Dim cambia_titulo As Boolean

cambia_titulo = True
'Accion = ""
Select Case Button.Index
Case 1 ' agregar
    Accion = "Agregando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    
    Numero.Text = Documento_Numero_Nuevo(RsITO, "Numero")

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
        If MsgBox("¿ Imprime ?", vbYesNo) = vbYes Then
            ITO_Imprimir
        End If
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
    End If
Case 5 ' separador
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

    If ITO_Validar Then
        If MsgBox("¿ GRABAR " & Obj & " ?", vbYesNo) = vbYes Then
        
            If Accion = "Agregando" Then
                ITO_Grabar True
            Else
                ITO_Grabar False
            End If
            
            If MsgBox("¿ IMPRIME " & Obj & " ?", vbYesNo) = vbYes Then
                ITO_Imprimir
            End If
            
            Botones_Enabled 0, 0, 0, 0, 1, 0
            Campos_Limpiar
            Campos_Enabled False
            Numero.Enabled = True
            Numero.SetFocus
            
        End If
    End If
    
    
Case 8 ' separador
Case 9 ' contratista
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
Private Sub Botones_Enabled(btn_Agregar As Boolean, btn_Modificar, _
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

ComboNv.Enabled = Si
ComboOT.Enabled = Si
aRecibir.Enabled = Si
Detalle.Enabled = Si
Obs(0).Enabled = Si
Obs(1).Enabled = Si
Obs(2).Enabled = Si
Obs(3).Enabled = Si
End Sub
Private Sub ITO_Imprimir()
' imprime ITOf
MousePointer = vbHourglass
Dim can_valor As String, can_col As Integer
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer, tab4 As Integer
Dim tab5 As Integer, tab6 As Integer, tab7 As Integer, tab8 As Integer, tab40 As Integer
tab0 = 7 'margen izquierdo
tab1 = tab0 + 0
tab2 = tab1 + 10
tab3 = tab2 + 10
tab4 = tab3 + 10
tab5 = tab4 + 19
tab6 = tab5 + 10
tab7 = tab6 + 10
tab8 = tab7 + 10
tab40 = 50

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt

prt.Font.Size = 15
prt.Print Tab(tab0 + 14); "VALE ITO ESPECIAL Nº" & Format(Numero.Text, "000")
prt.Font.Size = 10
prt.Print ""

' cabecera
prt.Font.Bold = True
prt.Print Tab(tab0); Empresa.Razon;
prt.Font.Bold = False
prt.Print Tab(tab0 + tab40); "FECHA     : " & Fecha.Text

prt.Print Tab(tab0); "GIRO: " & Empresa.Giro;
prt.Print Tab(tab0 + tab40); "SEÑOR(ES) : " & m_Razon,

prt.Print Tab(tab0); Empresa.Direccion;
prt.Print Tab(tab0 + tab40); "RUT       : " & m_Rut

prt.Print Tab(tab0); "Teléfono: " & Empresa.Telefono1 & " - " & Empresa.Comuna;

prt.Print Tab(tab0 + tab40); "OBRA      : ";
prt.Font.Bold = True
prt.Print Format(Mid(ComboNv.Text, 8), ">")
prt.Font.Bold = False

prt.Print ""
' detalle
prt.Font.Bold = True
prt.Print Tab(tab1); "PLANO";
prt.Print Tab(tab2); "REV";
prt.Print Tab(tab3); "MARCA";
prt.Print Tab(tab4); "DESCRIPCIÓN";
prt.Print Tab(tab5); "Nº OT";
prt.Print Tab(tab6); "CANT";
prt.Print Tab(tab7); "  KG UNIT";
prt.Print Tab(tab8); " KG TOTAL"
prt.Font.Bold = False

prt.Print Tab(tab1); String(110, "-")
j = -1
For i = 1 To n_filas

    can_valor = Detalle.TextMatrix(i, 9)
    
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
        
        ' Nº OT
        prt.Print Tab(tab5); Detalle.TextMatrix(i, 7);
        
        ' CANTIDAD
        can_valor = Trim(Format(can_valor, "####"))
        can_col = 4 - Len(can_valor)
        prt.Print Tab(tab6 + can_col); can_valor;
        
        ' KG UNITARIO
        can_valor = Trim(Format(m_CDbl(Detalle.TextMatrix(i, 10)), "###,###.0"))
        can_col = 9 - Len(can_valor)
        prt.Print Tab(tab7 + can_col); can_valor;
        
        ' KG TOTAL
        can_valor = Trim(Format(m_CDbl(Detalle.TextMatrix(i, 11)), "###,###.0"))
        can_col = 9 - Len(can_valor)
        prt.Print Tab(tab8 + can_col); can_valor
        
    End If
    
Next

prt.Print Tab(tab1); String(110, "-")
prt.Print Tab(tab0 + 60); "TOTAL KILOS : ";
prt.Font.Bold = True
prt.Font.Bold = False
prt.Print ""
prt.Print Tab(tab0); "OBSERVACIONES :";
prt.Print Tab(tab0 + 16); Obs(0).Text
prt.Print Tab(tab0 + 16); Obs(1).Text
prt.Print Tab(tab0 + 16); Obs(2).Text
prt.Print Tab(tab0 + 16); Obs(3).Text

For i = 1 To 1
    prt.Print ""
Next

prt.Print Tab(tab0); Tab(14), "__________________", Tab(56), "__________________"
prt.Print Tab(tab0); Tab(14), "       VºBº       ", Tab(56), "       VºBº       "

prt.EndDoc

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
