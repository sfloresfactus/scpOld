VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PlanosRecepcion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepcion de Planos"
   ClientHeight    =   5370
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   9390
   Icon            =   "PlanosRecepcion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Nueva Recepción"
            Object.Tag             =   ""
            ImageIndex      =   1
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Modificar Recepción"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Eliminar Recepción"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Imprimir Recepción"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            Object.Width           =   1e-4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Deshacer"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Grabar Recepción"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      MouseIcon       =   "PlanosRecepcion.frx":030A
   End
   Begin VB.ComboBox CbSc 
      Height          =   315
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Prioridad 
      Height          =   300
      Left            =   1200
      TabIndex        =   10
      Top             =   1440
      Width           =   495
   End
   Begin VB.ComboBox CbCondicion 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   1695
   End
   Begin VB.Frame Frame 
      Caption         =   "Entregado a"
      Height          =   975
      Left            =   3000
      TabIndex        =   11
      Top             =   840
      Width           =   4095
      Begin VB.CheckBox Check4 
         Caption         =   "Inspección"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Producción"
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Adquisiciones Planificación"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Gerencia Operaciones"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.ComboBox ComboNV 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Numero 
      Height          =   300
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtEditP 
      Height          =   285
      Left            =   360
      TabIndex        =   18
      Top             =   4920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Detalle 
      Height          =   2895
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5106
      _Version        =   327680
      BackColorBkg    =   12632256
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   2760
      TabIndex        =   4
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   327680
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Label lbl 
      Caption         =   "Prioridad"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Condición"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Nº"
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
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lbl 
      Caption         =   "FECHA"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "OBRA"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.Label PesoTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   4800
      TabIndex        =   17
      Top             =   4920
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   8640
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PlanosRecepcion.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PlanosRecepcion.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PlanosRecepcion.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PlanosRecepcion.frx":065C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PlanosRecepcion.frx":076E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PlanosRecepcion.frx":0880
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MenuPop 
      Caption         =   "MenuPop"
      Visible         =   0   'False
      Begin VB.Menu FilaInsertar 
         Caption         =   "&Insertar Fila"
      End
      Begin VB.Menu FilaEliminar 
         Caption         =   "&Eliminar Fila"
      End
      Begin VB.Menu FilaBorrarContenido 
         Caption         =   "&Borrar Contenido"
      End
   End
End
Attribute VB_Name = "PlanosRecepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DbD As Database
Private RsSc As Recordset ' contratistas

Private Dbm As Database
Private RsNVc As Recordset 'nota venta cabecera
Private RsPr As Recordset  'planos recepcion

Private Obj As String, Objs As String, Accion As String

Private n_filas As Integer, n_columnas As Integer
Private Titulo As String
Private m_PesoTot As Double
Private i As Integer, j As Integer

Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button, btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button
Private Imprimiendo As Boolean, Imprimir_Abrir As Boolean
Private m_Nv As Double, m_Nuevo As Boolean
'Private m_RutC As String
Private aRutSc(99) As String
Private d As Variant, m_NvArea As Integer

Private Sub CbSc_Click()
i = Detalle.Row
If CbSc.ListIndex > -1 Then
    Detalle.TextMatrix(i, 4) = CbSc.Text
    j = CbSc.ListIndex
    Detalle.TextMatrix(i, 5) = aRutSc(j)
End If

CbSc.visible = False
'Detalle = CbSc.Text

'Cursor_NoMueve
End Sub
Private Sub CbSc_LostFocus()
'MsgBox "CbSC lostfocus"
'Cursor_NoMueve
CbSc_Click
End Sub
Private Sub Form_Load()

Inicializa

Set DbD = OpenDatabase(data_file)

Set RsSc = DbD.OpenRecordset("Contratistas")
RsSc.Index = "Razon Social"
' puebla combo
i = 0
aRutSc(i) = " "
CbSc.AddItem " "
Do While Not RsSc.EOF
    If RsSc!Activo Then
        i = i + 1
        aRutSc(i) = RsSc!Rut
        CbSc.AddItem RsSc![Razon Social]
    End If
    RsSc.MoveNext
Loop
RsSc.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)

Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set RsPr = Dbm.OpenRecordset("Planos Recepcion")
RsPr.Index = "Numero-Linea"

' Combo obra
ComboNV.AddItem " "
Do While Not RsNVc.EOF
    If Usuario.Nv_Activas Then
        If RsNVc!Activa Then
            ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra
        End If
    Else
        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra
    End If
    RsNVc.MoveNext
Loop

CbCondicion.AddItem "Preliminar"
CbCondicion.AddItem "Definitivo"

Imprimiendo = False
Imprimir_Abrir = True

End Sub
Private Sub Inicializa()

Obj = "RECEPCIÓN DE PLANO"
Objs = "RECEPCIONES DE PLANO"

Titulo = "Recepción de Planos"

opGrabar False

Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnDesHacer = Toolbar.Buttons(6)
Set btnGrabar = Toolbar.Buttons(7)

If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 0, 0
End If

Variables_Limpiar

n_filas = 20
n_columnas = 5

Campos_Enabled False

Detalle_Config

Detalle.Enabled = False

End Sub
Private Sub Variables_Limpiar()
m_Nv = 0
'm_RutC = ""
m_Nuevo = False
PesoTotal.Caption = "0"

End Sub
Private Sub Botones_Enabled(btn_Agregar As Boolean, btn_Modificar As Boolean, btn_Eliminar, btn_Imprimir As Boolean, btn_DesHacer As Boolean, btn_Grabar As Boolean)

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
Private Sub Detalle_Config()
Dim i As Integer, ancho As Integer

Detalle.Cols = n_columnas + 1
Detalle.Rows = n_filas + 1
Detalle.Left = 100
Detalle.WordWrap = True
Detalle.RowHeight(0) = 450

' solo seleccion filas
'Detalle.SelectionMode = flexSelectionByRow

Detalle.TextMatrix(0, 0) = ""
Detalle.TextMatrix(0, 1) = "Plano"
Detalle.TextMatrix(0, 2) = "Rev"
Detalle.TextMatrix(0, 3) = "  Kg  Total"
Detalle.TextMatrix(0, 4) = "Contratista"
Detalle.TextMatrix(0, 5) = "" ' rut del contratista

Detalle.ColWidth(0) = 350 '250
Detalle.ColWidth(1) = 1200
Detalle.ColWidth(2) = 600
Detalle.ColWidth(3) = 1200
Detalle.ColWidth(4) = 4000
Detalle.ColWidth(5) = 0

ancho = 350
For i = 0 To n_columnas
    If i = 3 Then
        PesoTotal.Left = ancho - 200
        PesoTotal.Width = Detalle.ColWidth(i)
    End If
    ancho = ancho + Detalle.ColWidth(i)
Next

Detalle.Width = ancho
Me.Width = ancho + Detalle.Left * 2 + 300

'For i = 1 To n_filas
'    Detalle.TextMatrix(i, 0) = i
'    Detalle.Row = i
    ' establece colores a columnas
'    Detalle.col = 4
'    Detalle.CellForeColor = vbRed
'Next

Detalle.Row = 1
Detalle.col = 1

txtEditP.Text = ""

End Sub
Private Sub Campos_Enabled(Si As Boolean)

Numero.Enabled = Si
Fecha.Enabled = Si
ComboNV.Enabled = Si

CbCondicion.Enabled = Si
Prioridad.Enabled = Si

Check1.Enabled = Si
Check2.Enabled = Si
Check3.Enabled = Si
Check4.Enabled = Si

Detalle.Enabled = Si

End Sub
Private Sub Campos_Limpiar()

Numero.Text = ""
Fecha.Text = "__/__/__"
ComboNV.Text = " "

CbCondicion.ListIndex = 0
Prioridad.Text = 0

Check1.Value = False
Check2.Value = False
Check3.Value = False
Check4.Value = False

Detalle_Limpiar

PesoTotal.Caption = ""

End Sub
Private Sub opGrabar(Habilitada As Boolean)
Toolbar.Buttons(7).Enabled = Habilitada
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim cambia_titulo As Boolean, n_Copias As Integer

cambia_titulo = True
'Accion = "" rem accion

'Cursor_Mueve Detalle
' mi propio cursor mueve
If txtEditP.visible = True Then
    Cursor_NoMueve
    Actualiza
End If

If CbSc.visible = True Then
    CbSc_LostFocus
End If

Select Case Button.Index
Case 1 ' Agregar

    Accion = "Agregando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    
    Numero.Text = Documento_Numero_Nuevo(RsPr, "Numero")
    
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
'            Doc_Imprimir n_Copias
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

    If Doc_Validar Then
    
        If MsgBox("¿ GRABA " & Obj & " ?", vbYesNo) = vbYes Then
        
            If Accion = "Agregando" Then
                Doc_Grabar False
            Else
                Doc_Grabar True
            End If
            
            n_Copias = 1
            PrinterNCopias.Numero_Copias = n_Copias
            PrinterNCopias.Show 1
            n_Copias = PrinterNCopias.Numero_Copias
            
            If n_Copias > 0 Then
'                Doc_Imprimir n_Copias
            End If
            
            Botones_Enabled 0, 0, 0, 0, 1, 0
            Campos_Limpiar
            Campos_Enabled False
            Numero.Enabled = True
            Numero.SetFocus
            
        End If
    End If
End Select

If cambia_titulo Then
    If Accion = "" Then
        Me.Caption = "Mantención de " & StrConv(Objs, vbProperCase)
    Else
        Me.Caption = "Mantención de " & StrConv(Objs, vbProperCase) & " [" & Accion & "]"
    End If
End If

End Sub
Private Sub Numero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then After_Enter
End Sub
Private Sub After_Enter()
If Val(Numero.Text) < 1 Then Beep: Exit Sub
Select Case Accion
Case "Agregando"
   
    RsPr.Seek "=", Numero.Text, 1
    If RsPr.NoMatch Then
        Campos_Enabled True
        Numero.Enabled = False
        Fecha.SetFocus
        btnGrabar.Enabled = True
'        btnSearch.Visible = True
    Else
        Doc_Leer
        MsgBox Obj & " YA EXISTE"
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
    End If

Case "Modificando"

    RsPr.Seek "=", Numero.Text, 1
    If RsPr.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Botones_Enabled 0, 0, 0, 0, 1, 1
        Doc_Leer
        Campos_Enabled True
        Numero.Enabled = False
        
        Fecha.SetFocus
        
    End If

Case "Eliminando"

    RsPr.Seek "=", Numero.Text, 1
    If RsPr.NoMatch Then
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
    
    RsPr.Seek "=", Numero.Text, 1
    If RsPr.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Doc_Leer
        Numero.Enabled = False
        Detalle.Enabled = True
    End If
    
End Select

End Sub
Private Sub Fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then ComboNV.SetFocus
End Sub
Private Sub Fecha_LostFocus()
d = Fecha_Valida(Fecha, Now)
End Sub
Private Sub ComboNV_Click()

m_Nv = Val(Left(ComboNV.Text, 6))

End Sub

Private Sub Detalle_Click()
CbSc.visible = False
Select Case Detalle.col
Case 4
    CbSc.Top = Detalle.CellTop + Detalle.Top
    CbSc.Left = Detalle.CellLeft + Detalle.Left
    CbSc.Width = Int(Detalle.CellWidth * 1)
    CbSc.visible = True
End Select
End Sub

Private Sub Doc_Eliminar()
Dim qry As String

' borra plano detalle
qry = "DELETE * FROM [Planos Recepcion] WHERE Numero=CDbl(" & Numero.Text & ")"
Dbm.Execute qry

End Sub
Private Sub Imprimir()
' PIDE NOMBRE del PLANO (Editable) a imprimir

    If MsgBox("¿ IMPRIME PLANO ?", vbYesNo) = vbYes Then
        MousePointer = vbHourglass
        Plano_Imprimir
        MousePointer = vbDefault
    End If
    
    Variables_Limpiar
    Detalle_Limpiar
    Detalle.Enabled = False
    
    Imprimir_Abrir = True
    
End Sub
Private Sub DesHacer()
If btnGrabar.Enabled = True Then
    If MsgBox("¿ ABANDONA PLANO SIN GRABAR CAMBIOS ?", vbYesNo) = vbYes Then
        DesHacer1
    End If
Else
    DesHacer1
End If
Imprimir_Abrir = True
End Sub
Private Sub DesHacer1()
Variables_Limpiar
Detalle_Limpiar
Detalle.Enabled = False
Botones_Enabled 1, 1, 1, 1, 0, 0
End Sub
Private Sub Actualiza()
Dim fi As Integer, co As Integer, num As Double
fi = Detalle.Row
co = Detalle.col

m_PesoTot = 0

For i = 1 To n_filas
    m_PesoTot = m_PesoTot + m_CDbl(Detalle.TextMatrix(i, 3))
Next

PesoTotal.Caption = Format(m_PesoTot, num_Formato)

End Sub
Private Function Doc_Validar()

Doc_Validar = False
j = 0

If ComboNV.ListIndex < 0 Then
    MsgBox "Debe escoger Obra"
    ComboNV.SetFocus
    Exit Function
End If

For i = 1 To n_filas

    If Trim(Detalle.TextMatrix(i, 1)) <> "" Then
    
        ' Rev
        If Trim(Detalle.TextMatrix(i, 2)) = "" Then
            MsgBox "Debe digitar Revisión"
            Detalle.Row = i
            Detalle.col = 2
            Detalle.SetFocus
            Exit Function
        End If
        
        ' Rev
        If Not LargoString_Valida(Detalle.TextMatrix(i, 2), 10, i, 2) Then Exit Function
        
        ' Kilos
        If Trim(Detalle.TextMatrix(i, 3)) = "" Then
            MsgBox "Debe digitar Kgs"
            Detalle.Row = i
            Detalle.col = 3
            Detalle.SetFocus
            Exit Function
        End If
        
        ' KILOS
        If Not Numero_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
        
        j = j + 1
    
    End If
    
Next

If j = 0 Then
    MsgBox "Debe digitar al menos 1 plano"
    Detalle.SetFocus
    Exit Function
End If

Doc_Validar = True

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
    If num <> "" Then
        GoTo Sigue
    End If
Else
    If Val(num) < 0 Then
Sigue:
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
Private Sub Doc_Grabar(Borra As Boolean)

If Borra Then
    ' borra detalle
    Dim qry  As String
    qry = "DELETE * FROM [Planos Recepcion]"
    qry = qry & " WHERE Numero=Cdbl(" & Numero.Text & ")"
    Dbm.Execute qry
End If

' graba detalle
j = 0

For i = 1 To n_filas
    ' verifica marca y cantidad
    If Detalle.TextMatrix(i, 1) <> "" And Val(Detalle.TextMatrix(i, 3)) <> 0 Then
    
        RsPr.AddNew
        
        RsPr!Numero = Numero.Text
        RsPr!Fecha = Fecha.Text
        RsPr!Nv = m_Nv
        j = j + 1
        RsPr!linea = j
        RsPr!Plano = UCase(Detalle.TextMatrix(i, 1))
        RsPr!Rev = UCase(Detalle.TextMatrix(i, 2))
        RsPr![Peso Plano] = m_CDbl(Detalle.TextMatrix(i, 3))
        RsPr![RUT Contratista] = Detalle.TextMatrix(i, 5)
        RsPr![Entrega 1] = Check1.Value
        RsPr![Entrega 2] = Check2.Value
        RsPr![Entrega 3] = Check3.Value
        RsPr![Entrega 4] = Check4.Value
        RsPr!Prioridad = Val(Prioridad.Text)
        RsPr!Condicion = Left(CbCondicion.Text, 1)
        
        RsPr.Update
        
    End If
Next

End Sub
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'
' RUTINAS PARA EL FLEXGRID
'
'
Private Sub Detalle_DblClick()
' simula un espacio
If Imprimiendo Then Exit Sub
MSFlexGridEdit Detalle, txtEditP, 32
End Sub
Private Sub Detalle_GotFocus()
'If Imprimiendo Then Exit Sub
If txtEditP.visible = False Then Exit Sub
Detalle = txtEditP
txtEditP.visible = False
End Sub
Private Sub Detalle_LeaveCell()
If txtEditP.visible = False Then Exit Sub
Detalle = txtEditP
txtEditP.visible = False
End Sub
Private Sub Detalle_KeyPress(KeyAscii As Integer)
If Imprimiendo Then Exit Sub
MSFlexGridEdit Detalle, txtEditP, KeyAscii
End Sub
Private Sub txtEditP_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCodeP Detalle, txtEditP, KeyCode, Shift
End Sub
Private Sub txtEditP_LostFocus()
' pierde foco
'Cursor_Mueve Detalle
If txtEditP.visible Then Cursor_NoMueve
'Actualiza
End Sub
Sub EditKeyCodeP(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)

Select Case KeyCode

Case vbKeyEscape ' Esc

    Edt.visible = False
    MSFlexGrid.SetFocus
    
Case vbKeyReturn ' Enter

    MSFlexGrid.SetFocus
    DoEvents
    
    If Detalle.col = 1 Then
        ' busca si plano ya tiene alguna version anterior
        i = Detalle.Row
        RsPr.Index = "NV-Plano-Rev"
        d = UCase(Detalle.TextMatrix(i, 1))
        RsPr.Seek ">=", m_Nv, d
        
        If Not RsPr.NoMatch Then
        
            Do While Not RsPr.EOF
            
                If m_Nv <> RsPr!Nv Or d <> RsPr!Plano Then Exit Do
                
                Detalle.TextMatrix(i, 2) = RsPr!Rev
                Detalle.TextMatrix(i, 3) = RsPr![Peso Plano]
                
                Detalle.TextMatrix(i, 5) = RsPr![RUT Contratista]
                RsSc.Seek "=", RsPr![RUT Contratista]
                If Not RsSc.NoMatch Then
                    Detalle.TextMatrix(i, 4) = RsSc![Razón Social]
                End If
                
    '            Debug.Print RsPr!NV, RsPr!Plano
    
                RsPr.MoveNext
                
            Loop
            
        End If
        
        RsPr.Index = "Numero-Linea"
        
    End If
    ' //////////////////////////////////////////////////
    
    Actualiza
    
    Cursor_Mueve MSFlexGrid
    
Case 38 ' Flecha Arriba

    MSFlexGrid.SetFocus
    DoEvents
    Actualiza
    If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
        MSFlexGrid.Row = MSFlexGrid.Row - 1
    End If
    
Case 40 ' Flecha Abajo

    MSFlexGrid.SetFocus
    DoEvents
    Actualiza
    If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
        MSFlexGrid.Row = MSFlexGrid.Row + 1
    End If
    
End Select
End Sub
Private Sub txtEditP_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)
If MSFlexGrid.col = 4 Then Exit Sub
Select Case KeyAscii
Case 0 To 32
    Edt = MSFlexGrid
    Edt.SelStart = 1000
Case Else
    Edt = Chr(KeyAscii)
    Edt.SelStart = 1
End Select
'Edt.Move SSTab.Left + MSFlexGrid.CellLeft + MSFlexGrid.Left, SSTab.Top + MSFlexGrid.CellTop + MSFlexGrid.Top, MSFlexGrid.CellWidth, MSFlexGrid.CellHeight
Edt.Move MSFlexGrid.CellLeft + MSFlexGrid.Left, MSFlexGrid.CellTop + MSFlexGrid.Top, MSFlexGrid.CellWidth, MSFlexGrid.CellHeight
Edt.visible = True
Edt.SetFocus
opGrabar True
End Sub
Private Sub Detalle_KeyDown(KeyCode As Integer, Shift As Integer)
If Imprimiendo Then Exit Sub
If KeyCode = vbKeyF2 Then
    MSFlexGridEdit Detalle, txtEditP, 32
End If
End Sub
Private Sub Detalle_RowColChange()
'MIA
'Posicion = "Lín " & Detalle.Row & ", Col " & Detalle.col
End Sub
Private Sub Cursor_Mueve(MSFlexGrid As Control)
'MIA
Select Case MSFlexGrid.col
Case 4
    If MSFlexGrid.Row + 1 < MSFlexGrid.Rows Then
        MSFlexGrid.col = 1
        MSFlexGrid.Row = MSFlexGrid.Row + 1
    End If
Case Else
    MSFlexGrid.col = MSFlexGrid.col + 1
End Select
End Sub
Private Sub Cursor_NoMueve()
i = Detalle.col
Detalle.col = IIf(i = 1, 2, 1)
Detalle.col = i
Actualiza
End Sub
'
' FIN RUTINAS PARA FLEXGRID
'
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
Private Sub Doc_Leer()

Detalle_Limpiar

m_PesoTot = 0

With RsPr
i = 0
.Seek "=", Numero.Text, 1

If Not .NoMatch Then

    ' lee cabecera
    
    m_Nv = !Nv
'    Rut.Text = NoNulo(![RUT Contratista])
    
    RsNVc.Seek "=", m_Nv, m_NvArea
    If Not RsNVc.NoMatch Then
        On Error GoTo NoNv
        ComboNV.Text = Format(RsNVc!Número, "0000") & " - " & RsNVc!Obra
        GoTo Sigue
NoNv:
        ComboNV.AddItem Format(RsNVc!Número, "0000") & " - " & RsNVc!Obra
        ComboNV.Text = Format(RsNVc!Número, "0000") & " - " & RsNVc!Obra
Sigue:
        On Error GoTo 0
    '    ComboNV_Click
    End If
    
    Fecha.Text = !Fecha
    Prioridad.Text = !Prioridad
    
    If !Condicion = "P" Then CbCondicion.Text = "Preliminar"
    If !Condicion = "D" Then CbCondicion.Text = "Definitivo"
    
    Check1.Value = IIf(![Entrega 1], 1, 0)
    Check2.Value = IIf(![Entrega 2], 1, 0)
    Check3.Value = IIf(![Entrega 3], 1, 0)
    Check4.Value = IIf(![Entrega 4], 1, 0)
    
End If

Do While Not .EOF

    If Numero.Text <> !Numero Then Exit Do
    
    i = i + 1
    
    Detalle.TextMatrix(i, 1) = NoNulo(!Plano)
    Detalle.TextMatrix(i, 2) = NoNulo(!Rev)
    Detalle.TextMatrix(i, 3) = Replace(RsPr![Peso Plano], ",", ".")
    
    Detalle.TextMatrix(i, 5) = NoNulo(![RUT Contratista])
    RsSc.Seek "=", NoNulo(![RUT Contratista])
    If Not RsSc.NoMatch Then
        Detalle.TextMatrix(i, 4) = NoNulo(RsSc![Razón Social])
    End If
    
'    Detalle.TextMatrix(i, 9) = NoNulo(RsPr!Observaciones)
    m_PesoTot = m_PesoTot + CDbl(Detalle.TextMatrix(i, 3))
        
    .MoveNext
    
Loop
'Next
End With

PesoTotal.Caption = m_PesoTot

Detalle.Enabled = True
'opGrabar (False)
End Sub
Private Sub Detalle_Limpiar()
Dim j As Integer
For i = 1 To n_filas
    For j = 1 To n_columnas
        Detalle.TextMatrix(i, j) = ""
    Next
Next
'Detalle.Row = 1
'Detalle.col = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim Rs As Recordset
If btnGrabar.Enabled = True Then
    If MsgBox("¿ ABANDONA PLANO SIN GRABAR CAMBIOS ?", vbYesNo) = vbNo Then
        ' NO sale
        Cancel = True
    End If
    ' abandona
Else
    ' sale sin problemas, pues no modificó
    DataBases_Cerrar
End If
End Sub
Private Sub Detalle_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Imprimiendo Then Exit Sub
If Button = 2 Then
'    MsgBox Detalle.ColSel & vbCr & Detalle.col
    If Detalle.ColSel = 5 And Detalle.col = 1 Then
        PopupMenu MenuPop
    End If
End If
End Sub
Private Sub FilaInsertar_Click()
' inserta fila en FlexGrid
Dim fi As Integer, co As Integer, fi_ini As Integer
fi_ini = Detalle.Row
For fi = n_filas To fi_ini + 1 Step -1
    For co = 1 To n_columnas
        Detalle.TextMatrix(fi, co) = Detalle.TextMatrix(fi - 1, co)
    Next
Next
' fila nueva
For co = 1 To n_columnas
    Detalle.TextMatrix(fi_ini, co) = ""
Next
Detalle.col = 1
Detalle.Row = fi_ini
End Sub
Private Sub FilaEliminar_Click()
' elimina fila en FlexGrid
Dim fi As Integer, co As Integer, fi_ini As Integer
fi_ini = Detalle.Row
For fi = fi_ini To n_filas - 1
    For co = 1 To n_columnas
        Detalle.TextMatrix(fi, co) = Detalle.TextMatrix(fi + 1, co)
    Next
Next
' última fila
For co = 1 To n_columnas
    Detalle.TextMatrix(fi, co) = ""
Next

Actualiza

Detalle.col = 1
Detalle.Row = fi_ini
End Sub
Private Sub FilaBorrarContenido_Click()
Dim fi As Integer, co As Integer
fi = Detalle.Row
For co = 1 To n_columnas
    Detalle.TextMatrix(fi, co) = ""
Next
End Sub
Private Sub Detalle_SelChange()
' se produce con cada click
End Sub
Private Sub Plano_Imprimir()
Dim prt As Printer, linea As String
linea = String(99, "-")
' imprime OT
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer, tab4 As Integer
Dim tab5 As Integer, tab6 As Integer, tab7 As Integer, tab8 As Integer, tab9 As Integer
tab0 = 5 'margen izquierdo
tab1 = tab0 + 0
tab2 = tab0 + 6
tab3 = tab0 + 19
tab4 = tab0 + 34
tab5 = tab0 + 44
tab6 = tab0 + 55
tab7 = tab0 + 66
tab8 = tab0 + 79
tab9 = tab0 + 88

Dim can_valor As String, can_col As Integer

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt

'prt.Font.Size = 15
prt.Print Tab(tab0); "PLANO"
prt.Font.Size = 10
prt.Print ""
prt.Print ""

' cabecera

prt.Print ""
' detalle
prt.Print Tab(tab1); "ITEM";
prt.Print Tab(tab2); "MARCA";
prt.Print Tab(tab3); "DESCRIPCIÓN";
prt.Print Tab(tab4); "CANTIDAD";
prt.Print Tab(tab5); "Kg UNITARIO";
prt.Print Tab(tab6); "   Kg TOTAL";
prt.Print Tab(tab7); "  m2 UNITARIO";
prt.Print Tab(tab8); "     m2 TOTAL"

prt.Print Tab(tab1); linea
'j = -1
For i = 1 To n_filas

    can_valor = Detalle.TextMatrix(i, 3)
    
    If Val(can_valor) = 0 Then
    
    '    j = j + 1
    '    prt.Print Tab(tab1 + j * 5); "----\"

    Else
    
        ' ITEM
        prt.Print Tab(tab1); Format(i, "#### ");
        
        ' MARCA
        prt.Print Tab(tab2); Detalle.TextMatrix(i, 1);
        
        ' DESCRIPCION
        prt.Print Tab(tab3); Detalle.TextMatrix(i, 2);
        
        ' CANTIDAD
        can_valor = Trim(Format(can_valor, "####,###"))
        can_col = 8 - Len(can_valor)
        prt.Print Tab(tab4 + can_col); can_valor;
        
        ' KG UNITARIO
        can_valor = Trim(Format(m_CDbl(Detalle.TextMatrix(i, 5)), "#,###,###.0")) ' 18/05/98
        can_col = 11 - Len(can_valor)
        prt.Print Tab(tab5 + can_col); can_valor;
        
        ' KG TOTAL
        can_valor = Trim(Format(m_CDbl(Detalle.TextMatrix(i, 6)), "#,###,###.0")) '18/05/98
        can_col = 11 - Len(can_valor)
        prt.Print Tab(tab6 + can_col); can_valor;
        
        ' m2 UNITARIO
        can_valor = Trim(Format(m_CDbl(Detalle.TextMatrix(i, 7)), "###,###,###.0"))
        can_col = 13 - Len(can_valor)
        prt.Print Tab(tab7 + can_col); can_valor;
        
        ' m2 TOTAL
        can_valor = Trim(Format(m_CDbl(Detalle.TextMatrix(i, 8)), "###,###,###.0"))
        can_col = 13 - Len(can_valor)
        prt.Print Tab(tab8 + can_col); can_valor
        
    End If
    
Next

prt.Print Tab(tab1); linea
prt.Print ""
prt.Print Tab(tab0 + 40); "TOTAL KILOS : " & Format(PesoTotal, "#,###,###.0");
prt.Print ""

For i = 1 To 5
    prt.Print ""
Next

prt.Print Tab(tab0); Tab(14), "__________________", Tab(56), "__________________"
prt.Print Tab(tab0); Tab(14), "       VºBº       ", Tab(56), "       VºBº       "

prt.EndDoc

Impresora_Predeterminada "default"

End Sub
