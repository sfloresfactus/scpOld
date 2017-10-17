VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ValeBodega 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vale de Bodega"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
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
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Ingresar Nuevo Vale"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Modificar Vale"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Eliminar Vale"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Imprimir Vale"
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
            Object.ToolTipText     =   "Grabar Vale"
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
            Object.ToolTipText     =   "Mantención de Contratistas"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Mantención de Productos"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
      MouseIcon       =   "ValeBodega.frx":0000
   End
   Begin VB.TextBox Numero 
      Height          =   300
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox ComboNV 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Frame Frame_Entidad 
      Caption         =   "Entidad"
      Height          =   1215
      Left            =   3480
      TabIndex        =   8
      Top             =   600
      Width           =   5655
      Begin VB.CommandButton btnSearch 
         Height          =   300
         Left            =   1920
         Picture         =   "ValeBodega.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   300
      End
      Begin VB.TextBox Razon 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   5415
      End
      Begin MSMask.MaskEdBox Rut 
         Height          =   300
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
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
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.TextBox txtEdit 
      Height          =   285
      Left            =   8040
      TabIndex        =   11
      Text            =   "txtEdit"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   960
      TabIndex        =   5
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   327680
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid Detalle 
      Height          =   2445
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4313
      _Version        =   327680
      ScrollBars      =   2
   End
   Begin VB.Label TotalPrecio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6720
      TabIndex        =   15
      Top             =   4560
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   2760
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ValeBodega.frx":011E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ValeBodega.frx":0230
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ValeBodega.frx":0342
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ValeBodega.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ValeBodega.frx":0566
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ValeBodega.frx":0678
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ValeBodega.frx":078A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ValeBodega.frx":0AA4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "OBRA"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "FECHA"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "Vale de Bodega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      Width           =   2175
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
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   375
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
Attribute VB_Name = "ValeBodega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button

Private DbD As Database, RsProvee As Recordset, RsTra As Recordset, RsPrd As Recordset
Private Dbm As Database, RsVb As Recordset, RsNVc As Recordset

Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, d As Variant

Private n_filas As Integer, n_columnas As Integer
Private prt As Printer, m_TipoDoc As String
Private n3 As Double, n7 As Double
Private linea As String, m_Nv As Double
Private Detalle_Ancho As Integer, Col9_Ancho As Integer
Private m_NvArea As Integer
Public Property Let TipoDoc(ByVal New_Value As String)
m_TipoDoc = New_Value
End Property
Private Sub Form_Load()
' abre archivos
Set DbD = OpenDatabase(data_file)

Set RsProvee = DbD.OpenRecordset("Proveedores")
RsProvee.Index = "RUT"
Set RsTra = DbD.OpenRecordset("Trabajadores")
RsTra.Index = "RUT"
Set RsPrd = DbD.OpenRecordset("Productos2")
RsPrd.Index = "Codigo"

Set Dbm = OpenDatabase(mpro_file)

Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set RsVb = Dbm.OpenRecordset("MovBodega")
RsVb.Index = "tipo-numero-linea"

' Combo obra
ComboNv.AddItem " "
Do While Not RsNVc.EOF
    If Usuario.Nv_Activas Then
        If RsNVc!Activa Then
            ComboNv.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
        End If
    Else
        ComboNv.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
    End If
    RsNVc.MoveNext
Loop

Inicializa

Detalle_Config

If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 0, 0
End If

m_NvArea = 0

Exit Sub

End Sub
Private Sub Inicializa()

Obj = "VALE DE BODEGA"
Objs = "VALES DE BODEGA"

Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnDesHacer = Toolbar.Buttons(6)
Set btnGrabar = Toolbar.Buttons(7)

Accion = ""
'old_accion = ""

btnSearch.visible = False
btnSearch.ToolTipText = "Busca Contatista"
Campos_Enabled False

TipoDoc = "VC"

End Sub
Private Sub Detalle_Config()
Dim i As Integer

n_filas = 25
' col9  trabajador
' col10 largo especial ?
n_columnas = 10
Col9_Ancho = 2000

Detalle.Left = 100
Detalle.WordWrap = True
'Detalle.RowHeight(0) = 450
Detalle.Cols = n_columnas + 1
Detalle.Rows = n_filas + 1

'Detalle.TextMatrix(0, 0) = ""
'Detalle.TextMatrix(0, 1) = "Código"
'Detalle.TextMatrix(0, 2) = "Descripción"
'Detalle.TextMatrix(0, 3) = "Cantidad"
'Detalle.TextMatrix(0, 4) = "$ Uni"
'Detalle.TextMatrix(0, 5) = "Total"

'Detalle.ColWidth(0) = 250
'Detalle.ColWidth(1) = 1500
'Detalle.ColWidth(2) = 3500
'Detalle.ColWidth(3) = 1000
'Detalle.ColWidth(4) = 1000
'Detalle.ColWidth(5) = 1000

Detalle.TextMatrix(0, 0) = ""
Detalle.TextMatrix(0, 1) = "Código"
Detalle.TextMatrix(0, 2) = "L"
Detalle.TextMatrix(0, 3) = "Cantidad"
Detalle.TextMatrix(0, 4) = "Uni"
Detalle.TextMatrix(0, 5) = "Descripción"
Detalle.TextMatrix(0, 6) = "Largo(mm)"
Detalle.TextMatrix(0, 7) = "$ Uni"
Detalle.TextMatrix(0, 8) = "Total"
Detalle.TextMatrix(0, 9) = "Trabajador"

Detalle.ColWidth(0) = 250
Detalle.ColWidth(1) = 1500
Detalle.ColWidth(2) = 300
Detalle.ColWidth(3) = 800
Detalle.ColWidth(4) = 500
Detalle.ColWidth(5) = 3000
Detalle.ColWidth(6) = 950
Detalle.ColWidth(7) = 950
Detalle.ColWidth(8) = 950
Detalle.ColWidth(9) = Col9_Ancho
Detalle.ColWidth(10) = 0

'Detalle.ColAlignment(2) = 0

Detalle_Ancho = 350 ' con scroll vertical
'ancho = 100 ' sin scroll vertical

TotalPrecio.Width = Detalle.ColWidth(5)
For i = 0 To n_columnas
    If i = 8 Then
        TotalPrecio.Left = Detalle_Ancho + Detalle.Left - 300
        TotalPrecio.Width = Detalle.ColWidth(8)
    End If
    Detalle_Ancho = Detalle_Ancho + Detalle.ColWidth(i)
Next

Detalle.Width = Detalle_Ancho
Me.Width = Detalle_Ancho + 50 + Detalle.Left * 2

' col y row fijas
'Detalle.BackColorFixed = vbCyan

' establece colores a columnas
' columnas    modificables : NEGRAS
' columnas no modificables : ROJAS

For i = 1 To n_filas

    Detalle.TextMatrix(i, 0) = i
    Detalle.Row = i
    Detalle.col = 1
    Detalle.CellAlignment = flexAlignLeftCenter
    Detalle.col = 2
    Detalle.CellForeColor = vbRed
    Detalle.CellAlignment = flexAlignLeftCenter
'    Detalle.col = 3
    Detalle.col = 4
    Detalle.CellForeColor = vbRed
    Detalle.col = 5
    Detalle.CellForeColor = vbRed
    Detalle.CellAlignment = flexAlignLeftCenter
    
Next

txtEdit.Text = ""

Detalle.ColWidth(9) = 0

'Detalle.ScrollTrack = True 'no se nota
'Detalle.TextMatrix(1, 1) = "hola" 'ok

End Sub
Private Sub ComboNV_Click()

MousePointer = vbHourglass

i = 0
m_Nv = Val(Left(ComboNv.Text, 6))

MousePointer = vbDefault

End Sub
Private Sub Fecha_GotFocus()
'
End Sub
Private Sub Fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then ComboNv.SetFocus
End Sub
Private Sub Fecha_LostFocus()
d = Fecha_Valida(Fecha, Now)
End Sub
Private Sub Form_Unload(Cancel As Integer)
DataBases_Cerrar
End Sub
Private Function Archivo_Abrir(Db As Database, archivo As String) As Boolean
' intenta abrir archivo en forma compartida
Archivo_Abrir = True
On Error GoTo Error
Set Db = OpenDatabase(archivo)
Exit Function
Error:
MsgBox "Archivo YA está en uso"
Archivo_Abrir = False
End Function
Private Sub Numero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then After_Enter
End Sub
Private Sub After_Enter()
If Val(Numero.Text) < 1 Then Beep: Exit Sub
Select Case Accion

Case "Agregando"
   
    RsVb.Seek ">=", m_TipoDoc, Numero.Text, 0
'    rsvb.Seek "=", TipoDoc, Numero.Text, 1
'    If rsvb.NoMatch Then
    If RsVb.EOF Then
        GoTo Agregar
    End If
    If m_TipoDoc <> RsVb!Tipo Or Numero.Text <> RsVb!Numero Then
Agregar:
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
'    End If
    
Case "Modificando"

    RsVb.Seek ">=", m_TipoDoc, Numero.Text, 0
'    rsvb.Seek "=", TipoDoc, Numero.Text, 1
'    If rsvb.NoMatch Then
    If m_TipoDoc <> RsVb!Tipo Or Numero.Text <> RsVb!Numero Then
        MsgBox Obj & " NO EXISTE"
    Else
        Botones_Enabled 0, 0, 0, 0, 1, 1
        Doc_Leer
        Campos_Enabled True
        Numero.Enabled = False
        
    End If

Case "Eliminando"

    RsVb.Seek ">=", m_TipoDoc, Numero.Text, 0
'    rsvb.Seek "=", TipoDoc, Numero.Text, 1
'    If rsvb.NoMatch Then
    If m_TipoDoc <> RsVb!Tipo Or Numero.Text <> RsVb!Numero Then
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
    
    RsVb.Seek ">=", m_TipoDoc, Numero.Text, 0
'    rsvb.Seek "=", TipoDoc, Numero.Text, 1
'    If rsvb.NoMatch Then
    If m_TipoDoc <> RsVb!Tipo Or Numero.Text <> RsVb!Numero Then
        MsgBox Obj & " NO EXISTE"
    Else
        Doc_Leer
        Numero.Enabled = False
        Detalle.Enabled = True
    End If
    
End Select

End Sub
Private Sub Doc_Leer()

Dim m_resta As Integer, m_LargoE As Boolean, m_Tipo As String, m_Rut_Guia As String, m_Ruts_Distintos As Boolean, primera As Boolean

m_Tipo = ""
' CABECERA
Fecha.Text = Format(RsVb!Fecha, Fecha_Format)
m_Nv = RsVb!Nv
Rut.Text = NoNulo(RsVb![Rut])

RsNVc.Seek "=", m_Nv, m_NvArea
If Not RsNVc.NoMatch Then
    On Error GoTo NoNv
    ComboNv.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
    GoTo Sigue
NoNv:
    ComboNv.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
    ComboNv.Text = Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
Sigue:
    On Error GoTo 0
'    ComboNV_Click
End If

With RsVb
.Seek ">=", m_TipoDoc, Numero.Text, 0

i = 0
primera = True
m_Ruts_Distintos = False

If Not .NoMatch Then

    If primera Then
'        CbSeccion.ListIndex = 0
        m_Rut_Guia = NoNulo(!Rut)
        primera = False
    End If

    Do While Not .EOF
    
        If m_TipoDoc = !Tipo And Numero.Text = !Numero Then
        
            i = !linea
            
            m_Tipo = NoNulo(!TipoNE)
            
            m_LargoE = False
            Detalle.TextMatrix(i, 1) = ![codigo producto]
            RsPrd.Seek "=", ![codigo producto]
            If Not RsPrd.NoMatch Then
                Detalle.TextMatrix(i, 4) = RsPrd![unidad de medida]
                Detalle.TextMatrix(i, 5) = RsPrd!descripcion
                m_LargoE = RsPrd![Largo Especial]
            End If
            
            Detalle.TextMatrix(i, 2) = NoNulo(![Largo Especial])
            Detalle.TextMatrix(i, 3) = !Cant_Sale
            Detalle.TextMatrix(i, 6) = ![largo]
            Detalle.TextMatrix(i, 7) = ![Precio Unitario]
            Detalle.TextMatrix(i, 10) = m_LargoE
            
            n3 = m_CDbl(Detalle.TextMatrix(i, 3))
            n7 = m_CDbl(Detalle.TextMatrix(i, 7))
            
            Detalle.TextMatrix(i, 8) = Format(n3 * n7, num_Formato)

            If !Rut <> m_Rut_Guia Then
                m_Ruts_Distintos = True
            End If

        Else
        
            Exit Do
            
        End If
        
        .MoveNext
        
    Loop
    
End If

End With

If m_Tipo = "T" Then

    If m_Ruts_Distintos Then
        ' varias lineas distintos rut
'        Opcion(2).Value = True
'        Detalle.Width = Detalle_Ancho + Col9_Ancho
'        Detalle.ColWidth(9) = Col9_Ancho
        
    Else
        Trabajador_Lee Rut.Text
    End If

Else
'    Contratista_Lee Rut.Text
End If

Detalle.Row = 1 ' para q' actualice la primera fila del detalle

Actualiza

End Sub
Private Sub Trabajador_Lee(pRut)
RsTra.Seek "=", pRut
If Not RsTra.NoMatch Then
    Rut.Text = pRut
    Razon.Text = RsTra!appaterno & " " & RsTra!apmaterno & " " & RsTra!nombres
'    btnDetalleTrabajador.visible = True
'    btnDetalleTrabajador.Enabled = True
    btnSearch.visible = True
'    Direccion.Text = RsSc!Dirección
'    Comuna.Text = NoNulo(RsSc!Comuna)
End If
End Sub
Private Function Doc_Validar() As Boolean
Doc_Validar = False

If m_Nv = 0 Then
    MsgBox "DEBE ELEGIR OBRA"
    ComboNv.SetFocus
    Exit Function
End If

Select Case True
Case True 'Opcion(0)
    If Rut.Text = "" Then
        MsgBox "DEBE ELEGIR CONTRATISTA"
        btnSearch.SetFocus
        Exit Function
    End If
Case True 'Opcion(1)
    If Rut.Text = "" Then
        MsgBox "DEBE ELEGIR TRABAJADOR"
        btnSearch.SetFocus
        Exit Function
    End If
Case True 'Opcion(2)
    ' n trabajadores
End Select

For i = 1 To n_filas

    ' codigo prod
    If Trim(Detalle.TextMatrix(i, 1)) <> "" Then
    
        ' cantidad 3
        If Not CampoReq_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
        
        Detalle.Row = i
        
        ' precio unitario
        If Not Numero_Valida(Detalle.TextMatrix(i, 7), i, 7) Then Exit Function
        
'        If Opcion(2).Value = True Then
            ' trabajador x linea
'            If Not CampoReq_Valida(Detalle.TextMatrix(i, 9), i, 9) Then Exit Function
'        End If
    
        
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

save:

Detalle:
' DETALLE DE OT
Doc_Detalle_Eliminar

' graba detalle
With RsVb

For i = 1 To n_filas

    If Trim(Detalle.TextMatrix(i, 1)) <> "" Then
            
        .AddNew
        
        !Tipo = m_TipoDoc
        !Numero = Numero.Text
        !linea = i
        !Fecha = Fecha.Text
        !Nv = m_Nv
        
        If Rut.Text <> "" Then
        
            ![Rut] = Rut.Text
            
        Else
            ' n trabajadores, 1 x linea
            ' busca rut del trabajador
            
        End If
        
        ![codigo producto] = Detalle.TextMatrix(i, 1)
        ![Cant_Sale] = Detalle.TextMatrix(i, 3)
        ![Largo Especial] = Trim(Detalle.TextMatrix(i, 2))
        ![largo] = Val(Detalle.TextMatrix(i, 6))
        ![Precio Unitario] = Detalle.TextMatrix(i, 7)
        
'        If Opcion(0).Value = True Then
'            !TipoNE = " "
'        Else
'            !TipoNE = "T" ' trabajador
'        End If

        .Update
        
    End If
    
Next

End With

End Sub
Private Sub Doc_Eliminar()

' borra CABECERA DE OT
'rsvb.Seek "=", Numero.Text
'If Not rsvb.NoMatch Then

'    rsvb.Delete

'End If

Doc_Detalle_Eliminar

End Sub
Private Sub Doc_Detalle_Eliminar()

'DbAdq.Execute "DELETE * FROM Documentos WHERE tipo='VC' AND numero=" & Numero.Text

'rsvb.Seek "=", Numero.Text, 1
'If Not rsvb.NoMatch Then
'    Do While Not rsvb.EOF
'
'        ' borra detalle
'        rsvb.Delete
'
'        rsvb.MoveNext
'
'    Loop
'End If

End Sub
Private Sub Campos_Limpiar()
Numero.Text = ""
Fecha.Text = Fecha_Vacia
ComboNv.Text = " "

'Opcion(0).Value = True ' la mayoria de las veces va a ser contratista

Rut.Text = ""
Razon.Text = ""
'Opcion(0).Value = True

'CbSeccion.Text = "Todas"

'btnDetalleTrabajador.visible = False
'Direccion.Text = ""
'Comuna.Text = ""
Detalle_Limpiar
'Obs(0).Text = ""
'Obs(1).Text = ""
'Obs(2).Text = ""
'Obs(3).Text = ""
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
Private Sub Opcion_Click(Index As Integer)
Select Case Index
Case 0
    ' contratista
    btnSearch.visible = True
'    btnDetalleTrabajador.visible = False
    btnSearch.ToolTipText = "Busca Contratista"
'    CbSeccion.ListIndex = -1
'    CbSeccion.Enabled = False
'    lblSeccion.visible = False
'    CbSeccion.visible = False
    Detalle.Width = Detalle_Ancho - Col9_Ancho
    Detalle.ColWidth(9) = 0
Case 1
    ' 1 trabajador
    btnSearch.visible = True
'    btnDetalleTrabajador.visible = False
    btnSearch.ToolTipText = "Busca Trabajador"
'    CbSeccion.ListIndex = -1
'    CbSeccion.Enabled = False
    
'    lblSeccion.visible = False
'    CbSeccion.visible = False
    
    Detalle.Width = Detalle_Ancho - Col9_Ancho
    Detalle.ColWidth(9) = 0
Case Else
    ' N trabajadores
    btnSearch.visible = False
'    btnDetalleTrabajador.visible = False
    btnSearch.ToolTipText = ""
'    CbSeccion.Enabled = True
'    lblSeccion.visible = True
'    CbSeccion.visible = True
    Detalle.Width = Detalle_Ancho
    Detalle.ColWidth(9) = Col9_Ancho
End Select
Rut.Text = ""
Razon.Text = ""
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
    
    Numero.Text = Documento_Numero_Nuevo(RsVb, "Numero")
    
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
Case 9 ' Contratistas
    MousePointer = 11
    Load sql_contratistas
    MousePointer = 0
    sql_contratistas.Show 1
    cambia_titulo = False
Case 10 ' Productos
    MousePointer = 11
    Load Productos
    MousePointer = 0
    Productos.Show 1
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

'Opcion(0).Enabled = Si
'Opcion(1).Enabled = Si
'Opcion(2).Enabled = Si

'btnDetalleTrabajador.Enabled = Si

'lblSeccion.Enabled = Si
'CbSeccion.Enabled = Si

'If Si Then btnDetalleTrabajador.Visible = False
ComboNv.Enabled = Si
Detalle.Enabled = Si
'Obs(0).Enabled = Si
'Obs(1).Enabled = Si
'Obs(2).Enabled = Si
'Obs(3).Enabled = Si
End Sub
Private Sub btnSearch_Click()

If True Then
'If Opcion(0).Value = True Then

    Search.Muestra data_file, "Contratistas", "RUT", "Razon Social", "Contratista", "Contratistas", "Activo"
    
    Rut.Text = Search.codigo
    If Rut.Text <> "" Then
        RsProvee.Seek "=", Rut
        If RsProvee.NoMatch Then
            MsgBox "CONTRATISTA NO EXISTE"
            Rut.SetFocus
        Else
            Razon.Text = Search.descripcion
    '        Direccion.Text = RsSc!Dirección
    '        Comuna.Text = NoNulo(RsSc!Comuna)
        End If
    End If

Else

    Search.Muestra data_file, "Trabajadores", "RUT", "ApPaterno]+' '+[ApMaterno]+' '+[Nombres", "Trabajador", "Trabajadores" ', "Activo"
    
    Rut.Text = Search.codigo
    If Rut.Text <> "" Then
        RsTra.Seek "=", Rut
        If RsTra.NoMatch Then
            MsgBox "TRABAJADOR NO EXISTE"
            Rut.SetFocus
        Else
            Razon.Text = Search.descripcion
'            btnDetalleTrabajador.Enabled = True
'            btnDetalleTrabajador.visible = True
    '        Direccion.Text = RsSc!Dirección
    '        Comuna.Text = NoNulo(RsSc!Comuna)
        End If
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
    Case 1 ' plano
'        If Detalle <> "" Then ComboPlano.Text = Detalle
'        ComboPlano.Top = Detalle.CellTop + Detalle.Top
'        ComboPlano.Left = Detalle.CellLeft + Detalle.Left
'        ComboPlano.Width = Int(Detalle.CellWidth * 1.5)
'        ComboPlano.Visible = True
'        ComboMarca.Visible = False
    Case 3 ' marca
'        ComboMarca_Poblar Detalle.TextMatrix(Detalle.Row, 1)
        On Error GoTo Error
'        If Detalle <> "" Then ComboMarca.Text = Detalle
Error:
        On Error GoTo 0
'        ComboMarca.Text = ""
'        ComboMarca.Top = Detalle.CellTop + Detalle.Top
'        ComboMarca.Left = Detalle.CellLeft + Detalle.Left
'        ComboMarca.Width = Int(Detalle.CellWidth * 1.5)
'        ComboPlano.Visible = False
'        ComboMarca.Visible = True
        
    Case 10 ' fecha de entrega
    Case Else
'        ComboPlano.Visible = False
'        ComboMarca.Visible = False
End Select
End Sub
Private Sub Detalle_DblClick()
If Accion = "Imprimiendo" Then Exit Sub
' simula un espacio
'If Detalle.col = 10 Then
'    MSFlexGridEdit Detalle, EditFecha, 32  'FECHA
'Else
    MSFlexGridEdit Detalle, txtEdit, 32
'End If
End Sub
Private Sub Detalle_GotFocus()
If Accion = "Imprimiendo" Then Exit Sub
Select Case True
Case txtEdit.visible
    Detalle = txtEdit.Text
    txtEdit.visible = False
'Case EditFecha.Visible
'    Detalle = EditFecha
'    EditFecha.Visible = False
End Select
End Sub
Private Sub Detalle_LeaveCell()
If Accion = "Imprimiendo" Then Exit Sub
Select Case True
Case txtEdit.visible
    Detalle = txtEdit
    txtEdit.visible = False
'Case EditFecha.Visible
'    Detalle = EditFecha
'    EditFecha.Visible = False
End Select
End Sub
Private Sub Detalle_KeyPress(KeyAscii As Integer)
If Accion = "Imprimiendo" Then Exit Sub
If Detalle.col = 10 Then
'    MSFlexGridEdit Detalle, EditFecha, KeyAscii 'fecha
Else
    MSFlexGridEdit Detalle, txtEdit, KeyAscii
End If
End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCodeP Detalle, txtEdit, KeyCode, Shift
End Sub
Private Sub txtEdit_LostFocus()
'txtEditOT.Visible = False 07/03/98
'EditKeyCodeP Detalle, txtEditOT, vbkeyreturn, 0
' ó
'Detalle.SetFocus
'DoEvents
'Actualiza
End Sub
Sub EditKeyCodeP(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
' rutina que se ejecuta con los keydown de Edt
Dim m_fil As Integer, m_col As Integer
m_fil = MSFlexGrid.Row
m_col = MSFlexGrid.col

Dim dif As Integer
'dif = Val(Detalle.TextMatrix(MSFlexGrid.Row, 5)) - Val(Detalle.TextMatrix(MSFlexGrid.Row, 6))
Select Case KeyCode
Case vbKeyEscape ' Esc
    Edt.visible = False
    MSFlexGrid.SetFocus
Case vbKeyReturn ' Enter
    Select Case m_col
    Case 1 ' Codigo del Producto
    
        ' busca codigo
        
        MSFlexGrid.SetFocus
        DoEvents
        Linea_Actualiza
        
        Codigo2Descripcion Detalle.TextMatrix(m_fil, 1)
        
    Case 10 ' Fecha
'        If Fecha_Valida(Edt) Then
'            MSFlexGrid.SetFocus
'            DoEvents
'            Linea_Actualiza
'        End If
        Detalle.SetFocus
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
        Linea_Actualiza
    End Select
    Cursor_Mueve MSFlexGrid
Case vbKeyUp ' Flecha Arriba
    Select Case m_col
    Case 7 ' Cantidad a Asignar
        If Asignada_Validar(MSFlexGrid.col, dif, Edt) Then
            MSFlexGrid.SetFocus
            DoEvents
            Linea_Actualiza
            If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
                MSFlexGrid.Row = MSFlexGrid.Row - 1
            End If
        End If
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
    Case 7 ' Cantidad a Asignar
        If Asignada_Validar(MSFlexGrid.col, dif, Edt) Then
            MSFlexGrid.SetFocus
            DoEvents
            Linea_Actualiza
            If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
                MSFlexGrid.Row = MSFlexGrid.Row + 1
            End If
        End If
    Case 10 ' Fecha
        If Fecha_Valida(Edt) Then
            MSFlexGrid.SetFocus
            DoEvents
            Linea_Actualiza
            If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
                MSFlexGrid.Row = MSFlexGrid.Row + 1
            End If
        End If
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
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)
Dim m_Codigo As String

Select Case MSFlexGrid.col
Case 1 'codigo
    Edt.MaxLength = 15
Case 2 'largo especial
    Edt.MaxLength = 1
Case 4, 5, 8 'unidad, desc, total
'    no editables
Case Else
    Edt.MaxLength = 10
End Select

Select Case MSFlexGrid.col
Case 2 'largo especial
'    m_Codigo = MSFlexGrid.TextMatrix(MSFlexGrid.Row,10)
    If Not CBool(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 10)) Then Exit Sub
    GoTo Edita
Case 6
    'largo
    'si y no editable
    m_Codigo = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 2)
    If m_Codigo <> "E" Then Exit Sub
    GoTo Edita
Case 4, 5, 8
    ' no editables
    Exit Sub
Case 9 ' combo trabajador
'    CbTrabajadores.Left = MSFlexGrid.CellLeft + MSFlexGrid.Left
'    CbTrabajadores.Top = MSFlexGrid.CellTop + MSFlexGrid.Top
'    CbTrabajadores.visible = True
'    CbTrabajadores.SetFocus
Case Else
Edita:
    Select Case KeyAscii
    Case 0 To 32
        Edt = MSFlexGrid
        Edt.SelStart = 1000
    Case Else
        Edt = Chr(KeyAscii)
        Edt.SelStart = 1
    End Select
    Edt.Move MSFlexGrid.CellLeft + MSFlexGrid.Left, MSFlexGrid.CellTop + MSFlexGrid.Top, MSFlexGrid.CellWidth, MSFlexGrid.CellHeight * 1.2
    Edt.visible = True
    Edt.SetFocus
    'opGrabar True
End Select

Exit Sub

Select Case MSFlexGrid.col
Case 4, 5, 8
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

Select Case KeyCode
Case vbKeyF1
    If Detalle.col = 1 Then CodigoProducto_Buscar
Case vbKeyF2
    MSFlexGridEdit Detalle, txtEdit, 32
End Select
End Sub
Private Sub Linea_Actualiza()
' actualiza solo linea, y totales generales
Dim fi As Integer, co As Integer

fi = Detalle.Row
co = Detalle.col

n3 = m_CDbl(Detalle.TextMatrix(fi, 3))
n7 = m_CDbl(Detalle.TextMatrix(fi, 7))

' precio total
Detalle.TextMatrix(fi, 8) = Format(n3 * n7, num_fmtgrl)

Totales_Actualiza

End Sub
Private Sub Actualiza()
' actualiza todo el detalle y totales generales
Dim fi As Integer
For fi = 1 To n_filas
    If Detalle.TextMatrix(fi, 1) <> "" Then
    
        n3 = m_CDbl(Detalle.TextMatrix(fi, 3))
        n7 = m_CDbl(Detalle.TextMatrix(fi, 7))

        ' precio total
        Detalle.TextMatrix(fi, 8) = Format(n3 * n7, num_fmtgrl)
        
    End If
Next

Totales_Actualiza

End Sub
Private Sub Totales_Actualiza()
Dim Tot_Precio As Double
Tot_Precio = 0
For i = 1 To n_filas
    Tot_Precio = Tot_Precio + m_CDbl(Detalle.TextMatrix(i, 8))
Next

TotalPrecio.Caption = Format(Tot_Precio, num_Format0)

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
'
' FIN RUTINAS PARA FLEXGRID
'
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
Private Sub Detalle_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = 2 Then
    If Detalle.ColSel = 1 And Detalle.col = 1 Then
        'como F1
        CodigoProducto_Buscar
    End If
    If Detalle.ColSel = n_columnas And Detalle.col = 1 Then
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
Detalle.col = 1
Detalle.Row = fi_ini
'Detalle_Sumar
End Sub
Private Sub FilaBorrarContenido_Click()
' borra contenido de la fila en flexgrid
Dim fi As Integer, co As Integer
fi = Detalle.Row
For co = 1 To n_columnas
    Detalle.TextMatrix(fi, co) = ""
Next
'Detalle_Sumar
End Sub
Private Sub Doc_Imprimir(n_Copias As Integer)
MousePointer = vbHourglass
linea = String(78, "-")
' imprime OT
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer, tab4 As Integer
Dim tab5 As Integer, tab6 As Integer, tab7 As Integer, tab8 As Integer, tab9 As Integer
Dim tab10 As Integer, tab40 As Integer
Dim fc As Integer, fn As Integer, fe As Integer, k As Integer
tab0 = 3 'margen izquierdo
tab1 = tab0 'cod
tab2 = tab1 + 17 ' desc
tab3 = tab2 + 36 ' cant
tab4 = tab3 + 5  ' $ uni
tab5 = tab4 + 7  ' $ tot

tab40 = 43

' font.size
fc = 10 'comprimida
fn = 12 'normal
fe = 15 'expandida

Dim can_valor As String

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt

For k = 1 To n_Copias

prt.Font.Size = fe
prt.Print Tab(tab0 - 1); m_TextoIso
prt.Print Tab(tab0 + 25); "VALE DE CONSUMO Nº";
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
prt.Print Tab(tab0 + 28); Left(Razon, 32)

prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0); Empresa.Direccion;
prt.Font.Size = fc
prt.Print Tab(tab0 + tab40); "OBRA      : "
prt.Print Tab(tab0); "Teléfono: "; Empresa.Telefono1;
prt.Font.Size = fe
prt.Font.Bold = True
prt.Print Tab(tab0 + 28); Left(Format(Mid(ComboNv.Text, 8), ">"), 32)

prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0); Empresa.Comuna;
prt.Font.Size = fn

prt.Print ""
prt.Print ""

' detalle
prt.Font.Bold = True
prt.Print Tab(tab1); "COD";
prt.Print Tab(tab2); "Descripción";
prt.Print Tab(tab3); "CANT";
prt.Print Tab(tab4); " $ UNI";
prt.Print Tab(tab5); "   $ TOT"
prt.Font.Bold = False

prt.Print Tab(tab1); linea
j = -1

For i = 1 To n_filas

    can_valor = Detalle.TextMatrix(i, 3)
    
    If Val(can_valor) = 0 Then
    
        j = j + 1
        prt.Print Tab(tab1 + j * 3); "  \"
        
    Else
    
        ' COD PRO
        prt.Print Tab(tab1); Detalle.TextMatrix(i, 1);
        
        ' DESCRIPCION
        prt.Print Tab(tab2); Left(Detalle.TextMatrix(i, 5), 35);
        
        ' CANTIDAD
        prt.Print Tab(tab3); m_Format(can_valor, "#,###");
        
        ' $ UNITARIO
        prt.Print Tab(tab4); m_Format(Detalle.TextMatrix(i, 7), "###,###");
        
        ' $ TOTAL
        prt.Print Tab(tab5); m_Format(Detalle.TextMatrix(i, 8), "#,###,###")
        
    End If
    
Next

prt.Print Tab(tab1); linea
prt.Font.Bold = True
prt.Print Tab(tab5 - 5); m_Format(TotalPrecio, "$###,###,###")
prt.Font.Bold = False
prt.Print ""

'prt.Print Tab(tab0); "OBSERVACIONES :";
'prt.Print Tab(tab0 + 16); Obs(0).Text
'prt.Print Tab(tab0 + 16); Obs(1).Text
'prt.Print Tab(tab0 + 16); Obs(2).Text
'prt.Print Tab(tab0 + 16); Obs(3).Text

For i = 1 To 2
    prt.Print ""
Next

prt.Print Tab(tab0); Tab(14), "__________________", Tab(56), "__________________"
prt.Print Tab(tab0); Tab(14), "       VºBº       ", Tab(56), "       VºBº       "

If k < n_Copias Then
    prt.NewPage
Else
    prt.EndDoc
End If

Next

Impresora_Predeterminada "default"

MousePointer = vbDefault
End Sub
Private Sub CodigoProducto_Buscar()
Dim m_cod As String

MousePointer = vbHourglass
Product_Search.Condicion = ""
Load Product_Search
MousePointer = vbDefault
Product_Search.Show 1
m_cod = Product_Search.CodigoP

Codigo2Descripcion m_cod

End Sub
Private Sub Codigo2Descripcion(CodigoP As String) ', fila As Integer)
' busca descripcion del codigo del producto

Dim m_Upc As Double
If CodigoP = "" Then

    Detalle.TextMatrix(Detalle.Row, 4) = ""
    Detalle.TextMatrix(Detalle.Row, 5) = ""
   
Else

    With RsPrd
    .Seek "=", CodigoP
    If Not .NoMatch Then
    
        Detalle.TextMatrix(Detalle.Row, 1) = CodigoP
        Detalle.TextMatrix(Detalle.Row, 4) = ![unidad de medida]
        Detalle.TextMatrix(Detalle.Row, 5) = !descripcion
        Detalle.TextMatrix(Detalle.Row, 10) = ![Largo Especial]
        
        Detalle.TextMatrix(Detalle.Row, 7) = m_Upc
        
    '    Detalle.TextMatrix(Detalle.Row, 4) = ![Unidad de Medida]
    '    Detalle.TextMatrix(Detalle.Row, 5) = !Descripción
    '    Detalle.TextMatrix(Detalle.Row, 6) = !Largo
    '    Detalle.TextMatrix(Detalle.Row,10) = ![Largo Especial]
        
    '    If Detalle.TextMatrix(Detalle.Row,10) Then
    '        Detalle.col = 2 'foco en largo especial
    '    Else
         Detalle.col = 2 '3
    '    End If

    Else
    
        Detalle.TextMatrix(Detalle.Row, 5) = "-- CODIGO NO EXISTE --"
        Detalle.TextMatrix(Detalle.Row, 10) = False
    
    End If
    
    End With

End If

End Sub
