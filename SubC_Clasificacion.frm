VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form SubC_Clasificacion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5445
   ClientLeft      =   3135
   ClientTop       =   2805
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Agregar"
            Object.Tag             =   "[Agregando]"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "[Modificando]"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "[Eliminando]"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Listar"
            Object.Tag             =   "[Listando]"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Deshacer"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      MouseIcon       =   "SubC_Clasificacion.frx":0000
   End
   Begin VB.TextBox observacion 
      Height          =   300
      Left            =   1320
      TabIndex        =   6
      Top             =   1440
      Width           =   4215
   End
   Begin VB.TextBox descripcion 
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox codigo 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox EditTabla 
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3120
      MousePointer    =   14  'Arrow and Question
      Picture         =   "SubC_Clasificacion.frx":001C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   300
   End
   Begin MSFlexGridLib.MSFlexGrid Tabla 
      Height          =   3000
      Left            =   1320
      TabIndex        =   7
      Top             =   1920
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5292
      _Version        =   327680
   End
   Begin VB.Label lbl 
      Caption         =   "&Observación"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&Descripción"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&Código"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   6960
      Top             =   600
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SubC_Clasificacion.frx":011E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SubC_Clasificacion.frx":0230
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SubC_Clasificacion.frx":0342
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SubC_Clasificacion.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SubC_Clasificacion.frx":0566
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SubC_Clasificacion.frx":0678
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "SubC_Clasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Obj As String, Objs As String, nuevo As Boolean, Accion As String, old_accion As String
Private Db As Database, Rs As Recordset, RsTabla As Recordset
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnListar As Button, btnGrabar As Button, btnDesHacer As Button
Private lc(10) As Integer ' largo de los campos
Private m_Tabla As String, i As Integer
Private n_filas As Integer, n_columnas As Integer
Private Sub Campos_Enabled(Si As Boolean)
codigo.Enabled = Si
descripcion.Enabled = Si
Observacion.Enabled = Si

If Usuario.ReadOnly Then
    Tabla.Enabled = False
Else
    Tabla.Enabled = Si
End If

End Sub
Private Sub btnBuscar_Click()

Search.Muestra data_file, m_Tabla, "Codigo", "Descripcion", Obj, Objs

codigo = Search.codigo
descripcion = Search.descripcion

If codigo <> "" Then
    Rs.Seek "=", codigo
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        codigo.SetFocus
    Else
        After_Enter
    End If
End If
End Sub
Private Sub Codigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If codigo.Text = "" Then
        Beep
    Else
        codigo.Text = UCase(codigo.Text)
        After_Enter
    End If
End If
End Sub
Private Sub After_Enter()


Select Case Accion
Case "Agregar"
    Rs.Seek "=", codigo
    If Rs.NoMatch Then
        Campos_Enabled True
        codigo.Enabled = False
        descripcion.SetFocus
        btnGrabar.Enabled = True
    Else
        Registro_Leer
        MsgBox Obj & " YA EXISTE"
        
        Campos_Limpiar
        Campos_Enabled False
        
        codigo.Enabled = True
        codigo.SetFocus
    End If
Case "Modificar"
    Rs.Seek "=", codigo
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        codigo.SetFocus
    Else
        Campos_Enabled True
        codigo.Enabled = False
        descripcion.SetFocus
        Registro_Leer
        btnGrabar.Enabled = True
        btnBuscar.Visible = False
    End If
Case "Eliminar"
    Rs.Seek "=", codigo
    If Rs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        codigo.SetFocus
    Else
        Campos_Enabled False
        Registro_Leer
        If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo, "Atención") = vbYes Then
            Rs.Delete
        End If
        btnBuscar.Visible = True
        Campos_Limpiar
        codigo.Enabled = True
        codigo.SetFocus
    End If
End Select
End Sub
Private Sub Registro_Grabar(nuevo As Boolean)
save:
If nuevo Then
    Rs.AddNew
    Rs!Código = codigo.Text
Else
    Rs.Edit
End If

Rs!descripcion = descripcion.Text
Rs!Observacion = Observacion.Text
Rs.Update

Tabla_Grabar

Campos_Limpiar
Accion = old_accion

End Sub
Private Sub Tabla_Grabar()
For i = 1 To 10
    RsTabla.Seek "=", codigo.Text, i
    If RsTabla.NoMatch Then
        RsTabla.AddNew
        RsTabla!Clasificación = codigo.Text
        RsTabla!Tramo = i
    Else
        RsTabla.Edit
    End If
    RsTabla!Desde = Val(Tabla.TextMatrix(RsTabla!Tramo, 1))
    RsTabla!Hasta = Val(Tabla.TextMatrix(RsTabla!Tramo, 2))
    RsTabla!valor = Val(Tabla.TextMatrix(RsTabla!Tramo, 3))
    RsTabla.Update
Next
End Sub
Private Sub Registro_Leer()
descripcion.Text = Rs!descripcion
Observacion.Text = Rs!Observacion
Tabla_Leer
End Sub
Private Sub Tabla_Leer()
RsTabla.Seek ">=", codigo.Text, 1
If Not RsTabla.NoMatch Then
    Do While Not RsTabla.EOF
        If RsTabla!Clasificacion <> codigo.Text Then Exit Do
        Tabla.TextMatrix(RsTabla!Tramo, 1) = RsTabla!Desde
        Tabla.TextMatrix(RsTabla!Tramo, 2) = RsTabla!Hasta
        Tabla.TextMatrix(RsTabla!Tramo, 3) = RsTabla!valor
        RsTabla.MoveNext
    Loop
End If
End Sub
Private Sub Form_Load()

Inicializa

Me.Caption = "Mantención de " & Objs
Botones_Enabled True, True, True, True, False, False

' abre archivos
Set Db = OpenDatabase(data_file)
Set Rs = Db.OpenRecordset(m_Tabla)
Rs.Index = "Codigo"

nuevo = False

Set RsTabla = Db.OpenRecordset("Tabla Bono Produccion")
RsTabla.Index = "Clasificacion-Tramo"

Tabla_Config

Campos_Limpiar
Campos_Enabled False

End Sub
Private Sub Inicializa()
Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnListar = Toolbar.Buttons(4)
Set btnGrabar = Toolbar.Buttons(5)
Set btnDesHacer = Toolbar.Buttons(6)

Obj = "Clasificación"
Objs = "Clasificaciones"

m_Tabla = "Clasificacion de Contratistas"

Accion = ""
old_accion = ""

btnBuscar.Visible = False
btnBuscar.ToolTipText = "Busca " & StrConv(Obj, vbProperCase)

' largo del los campos
lc(1) = 10 'codigo
lc(2) = 50 'descripcion
lc(3) = 50 'obs

'With codigo
'    .Mask = String(lc(1), "&")
'    .PromptInclude = False
'End With
codigo.MaxLength = lc(1)

'With descripcion
'    .Mask = ">" & String(lc(2), "&")
'    .PromptInclude = False
'End With
descripcion.MaxLength = lc(2)

'With Observacion
'    .Mask = ">" & String(lc(3), "&")
'    .PromptInclude = False
'End With

Observacion.MaxLength = lc(3)

End Sub
Private Sub Tabla_Config()
Dim ancho As Integer

n_filas = 10
n_columnas = 3

'Detalle.Left = 100
Tabla.WordWrap = True
Tabla.RowHeight(0) = 450
Tabla.Rows = n_filas + 1
Tabla.Cols = n_columnas + 1

Tabla.TextMatrix(0, 0) = ""
Tabla.TextMatrix(0, 1) = "Desde (Ton)"
Tabla.TextMatrix(0, 2) = "Hasta (Ton)"
Tabla.TextMatrix(0, 3) = "Valor ($/Kg)"

Tabla.ColWidth(0) = 250
Tabla.ColWidth(1) = 1200
Tabla.ColWidth(2) = 1200
Tabla.ColWidth(3) = 1300

ancho = Tabla.ColWidth(0) + Tabla.ColWidth(1) + Tabla.ColWidth(2) + Tabla.ColWidth(3) + 100

Tabla.Width = ancho
'Me.Width = ancho + Tabla.Left * 2.5

For i = 1 To n_filas
    Tabla.TextMatrix(i, 0) = i
Next

End Sub
Private Sub Form_Unload(Cancel As Integer)
RsTabla.Close
Rs.Close
Db.Close
End Sub
Private Sub Botones_Enabled(btn_Agregar As Boolean, btn_Modificar As Boolean, btn_Eliminar As Boolean, btn_Listar As Boolean, btn_Grabar As Boolean, btn_DesHacer As Boolean)
Dim i As Integer
btnAgregar.Enabled = btn_Agregar
btnModificar.Enabled = btn_Modificar
btnEliminar.Enabled = btn_Eliminar
btnListar.Enabled = btn_Listar
btnGrabar.Enabled = btn_Grabar
btnDesHacer.Enabled = btn_DesHacer

For i = 1 To 6
    Toolbar.Buttons(i).Value = tbrUnpressed
Next

End Sub
Private Sub Descripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Observacion.SetFocus
End Sub
Private Sub Observacion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then descripcion.SetFocus
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Arch As String
Accion = Button.ToolTipText
Select Case Accion
Case "Agregar"
    Campos_Enabled False
    codigo.Enabled = True
    codigo.SetFocus
    nuevo = True
    old_accion = Accion
Case "Modificar"
    codigo.Enabled = True
    descripcion.Enabled = False
    codigo.SetFocus
    nuevo = False
    old_accion = Accion
    btnBuscar.Visible = True
Case "Eliminar"
    codigo.Enabled = True
    descripcion.Enabled = False
    codigo.SetFocus
'    nuevo = False
    btnBuscar.Visible = True
Case "Listar"
    MousePointer = vbHourglass
'    dbLister.Muestra "", data_file, m_Tabla, Array("Código", 1000, 10, "", 0), _
                                                Array("Descripción", 2000, 20, "", 0), _
                                                Array("Observación", 1000, 10, "", 0)
    MousePointer = vbDefault
Case "Grabar"
    If Valida(nuevo) Then
        Registro_Grabar nuevo
    Else
        Exit Sub
    End If
Case "Deshacer"
    Campos_Limpiar
    Campos_Enabled False
    btnBuscar.Visible = False
End Select

Select Case Button.Index
Case 5  ' btnGrabar
    Campos_Limpiar
    Campos_Enabled False
    codigo.Enabled = True
    codigo.SetFocus
    btnGrabar.Value = tbrUnpressed
    btnGrabar.Enabled = False
    
Case 4 To 6 ' btnDesHacer
    Botones_Enabled True, True, True, True, False, False
    Me.Caption = "Mantención de " & Objs
Case Else
    Botones_Enabled False, False, False, False, False, True
    Me.Caption = "Mantención de " & Objs & " " & Button.Tag
End Select

End Sub
Private Function Valida(nuevo As Boolean) As Boolean
Valida = False
If IsObjBlanco(descripcion, "DESCRIPCIÓN", btnGrabar) Then Exit Function

'' valida que descripcion no esté repetida
If nuevo Then
    Rs.Index = "Descripción"
    Rs.Seek "=", descripcion
    If Not Rs.NoMatch Then
        MsgBox "DESCRIPCIÓN YA EXISTE"
        btnGrabar.Value = tbrUnpressed
        Rs.Index = "Código"
        descripcion.SetFocus
        Exit Function
    End If
    btnGrabar.Value = tbrUnpressed
    Rs.Index = "Código"
End If
''

Valida = True
End Function
Private Sub Campos_Limpiar()
codigo.Text = ""
descripcion.Text = ""
Observacion.Text = ""
For i = 1 To n_filas
    Tabla.TextMatrix(i, 1) = ""
    Tabla.TextMatrix(i, 2) = ""
    Tabla.TextMatrix(i, 3) = ""
Next
End Sub
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'
' RUTINAS PARA EL FLEXGRID
'
'
Private Sub Tabla_Click()
After_Detalle_Click
End Sub
Private Sub After_Detalle_Click()
Dim fil As Integer
fil = Tabla.Row - 1
Select Case Tabla.col
Case 1 ' desde
Case 2 ' hasta
Case Else ' clas n
End Select
End Sub
Private Sub Tabla_DblClick()
' simula un espacio
MSFlexGridEdit Tabla, EditTabla, 32
End Sub
Private Sub Tabla_GotFocus()
Select Case True
Case EditTabla.Visible
    Tabla = EditTabla
    EditTabla.Visible = False
End Select
End Sub
Private Sub Tabla_LeaveCell()
Select Case True
Case EditTabla.Visible
    Tabla = EditTabla
    EditTabla.Visible = False
End Select
End Sub
Private Sub Tabla_KeyPress(KeyAscii As Integer)
MSFlexGridEdit Tabla, EditTabla, KeyAscii
End Sub
Private Sub EditTabla_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCodeP Tabla, EditTabla, KeyCode, Shift
End Sub
Sub EditKeyCodeP(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
' rutina que se ejecuta con los keydown de Edt
Dim m_col As Integer
m_col = MSFlexGrid.col

Select Case KeyCode
Case vbKeyEscape ' Esc
    Edt.Visible = False
    MSFlexGrid.SetFocus
Case vbKeyReturn ' Enter
    MSFlexGrid.SetFocus
    DoEvents
    Cursor_Mueve MSFlexGrid
Case vbKeyUp ' Flecha Arriba
    MSFlexGrid.SetFocus
    DoEvents
    If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
        MSFlexGrid.Row = MSFlexGrid.Row - 1
    End If
Case vbKeyDown ' Flecha Abajo
    MSFlexGrid.SetFocus
    DoEvents
    If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
        MSFlexGrid.Row = MSFlexGrid.Row + 1
    End If
End Select
End Sub
Private Sub EditTabla_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)

Select Case MSFlexGrid.col
Case 1, 2
    Edt.MaxLength = 10
Case Is > 2
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
Edt.Visible = True
Edt.SetFocus
'opGrabar True

End Sub
Private Sub Detalle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    MSFlexGridEdit Tabla, EditTabla, 32
End If
End Sub
Private Sub Cursor_Mueve(MSFlexGrid As Control)
'MIA
If MSFlexGrid.col = n_columnas Then
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
