VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form As_Mantencion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arco Sumergido"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   25
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
      MouseIcon       =   "As_Mantencion.frx":0000
   End
   Begin VB.TextBox Turno 
      Height          =   285
      Left            =   960
      TabIndex        =   34
      Top             =   1680
      Width           =   255
   End
   Begin VB.ComboBox CbContratistas 
      Height          =   315
      Left            =   7440
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Frame Frame_Material 
      Caption         =   "Material de Aporte"
      Height          =   1215
      Left            =   4080
      TabIndex        =   19
      Top             =   5040
      Width           =   3015
      Begin VB.TextBox Alambre 
         Height          =   300
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   23
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Fundente 
         Height          =   300
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lbl 
         Caption         =   "Alambre"
         Height          =   255
         Index           =   11
         Left            =   480
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "Fundente"
         Height          =   255
         Index           =   10
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame_Parametros 
      Caption         =   "Parametros de Trabajo"
      Height          =   1215
      Left            =   240
      TabIndex        =   12
      Top             =   5040
      Width           =   3735
      Begin VB.TextBox Voltaje 
         Height          =   300
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Velocidad 
         Height          =   300
         Left            =   960
         MaxLength       =   10
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Amperes 
         Height          =   300
         Left            =   960
         MaxLength       =   10
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "Velocidad"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "Voltaje"
         Height          =   300
         Index           =   9
         Left            =   1920
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbl 
         Caption         =   "Amperes"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
   End
   Begin MSMask.MaskEdBox NEquipo 
      Height          =   300
      Left            =   2760
      TabIndex        =   31
      Top             =   1680
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   529
      _Version        =   327680
      MaxLength       =   1
      Mask            =   "#"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame 
      Caption         =   "Operador"
      Height          =   1095
      Left            =   2520
      TabIndex        =   3
      Top             =   480
      Width           =   5775
      Begin VB.TextBox Razon2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   27
         Top             =   600
         Width           =   3255
      End
      Begin VB.CommandButton btnSearch2 
         Height          =   300
         Left            =   1920
         Picture         =   "As_Mantencion.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   600
         Width           =   300
      End
      Begin VB.CommandButton btnSearch1 
         Height          =   300
         Left            =   1920
         Picture         =   "As_Mantencion.frx":011E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   300
      End
      Begin VB.TextBox Razon1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   3255
      End
      Begin MSMask.MaskEdBox Rut1 
         Height          =   300
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Rut2 
         Height          =   300
         Left            =   720
         TabIndex        =   28
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label lbl 
         Caption         =   "&RUT"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbl 
         Caption         =   "&RUT"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.TextBox txtEdit 
      Height          =   285
      Left            =   8280
      TabIndex        =   24
      Text            =   "txtEdit"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   960
      TabIndex        =   9
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
      TabIndex        =   2
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
      TabIndex        =   11
      Top             =   2040
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4313
      _Version        =   327680
      ScrollBars      =   2
   End
   Begin VB.Label Label1 
      Caption         =   "V: viga   T: tubular   S: tubest   P: plancha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   4560
      Width           =   3975
   End
   Begin VB.Label lbl 
      Caption         =   "Nº Equipo"
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   30
      Top             =   1680
      Width           =   855
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   8880
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
            Picture         =   "As_Mantencion.frx":0220
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "As_Mantencion.frx":0332
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "As_Mantencion.frx":0444
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "As_Mantencion.frx":0556
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "As_Mantencion.frx":0668
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "As_Mantencion.frx":077A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "As_Mantencion.frx":088C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "&Turno"
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
      TabIndex        =   8
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "Arco Sumergido"
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
      Width           =   1815
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
      TabIndex        =   1
      Top             =   840
      Width           =   375
   End
End
Attribute VB_Name = "As_Mantencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button

Private DbD As Database, RsTra As Recordset, RsSc As Recordset
Private Dbm As Database, RsAs As Recordset, RsNVc As Recordset

Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, d As Variant

Private n_filas As Integer, n_columnas As Integer
Private prt As Printer

Private n1 As Double, n4 As Double
Private linea As String, m_NvArea As Integer
Private TipoDoc As String
Private PesoEspecificoV As Double ' peso especifico del acero para viga
Private PesoEspecificoT As Double ' peso especifico del acero para Tubular
Private PesoEspecificoS As Double ' peso especifico del acero para Tubest
Private PesoEspecificoP As Double ' peso especifico del acero para Plancha
Private a_Contratistas(1, 199) As String
Private Sub CbContratistas_Click()
Detalle = CbContratistas.Text
CbContratistas.visible = False
End Sub
Private Sub Fecha_KeyPress(KeyAscii As Integer)
'If KeyAscii = vbKeyReturn Then ComboNV.SetFocus
End Sub
Private Sub Fecha_LostFocus()
d = Fecha_Valida(Fecha, Now)
End Sub
Private Sub Form_Load()
Dim li As Integer, m_Nombre As String
' abre archivos
Set DbD = OpenDatabase(data_file)

Set RsTra = DbD.OpenRecordset("Trabajadores")
RsTra.Index = "RUT"

Set RsSc = DbD.OpenRecordset("Contratistas")
RsSc.Index = "RUT"

a_Contratistas_Limpiar
CbContratistas.Clear
Set RsSc = DbD.OpenRecordset("SELECT * FROM contratistas WHERE activo ORDER BY [razon social]")
With RsSc
li = 0
Do While Not .EOF
    li = li + 1
    a_Contratistas(0, li) = !Rut
    m_Nombre = ![Razon Social]
    a_Contratistas(1, li) = m_Nombre
    CbContratistas.AddItem m_Nombre
'        Debug.Print !nombres, !appaterno, !apmaterno
    .MoveNext
Loop
.Close
End With



Set Dbm = OpenDatabase(mpro_file)

Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set RsAs = Dbm.OpenRecordset("Arco Sumergido")
RsAs.Index = "Numero-Linea"

Inicializa
Detalle_Config

If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 0, 0
End If

m_NvArea = 0

TipoDoc = "AS" ' se usa ??

PesoEspecificoV = 8 ' viga
PesoEspecificoT = 7.9 ' tubular
PesoEspecificoS = 7.8 ' tubest
PesoEspecificoP = 8 ' plancha

Amperes.MaxLength = 10 ' double
Voltaje.MaxLength = 10 ' double
Velocidad.MaxLength = 10

Fundente.MaxLength = 20
Alambre.MaxLength = 20

CbContratistas.visible = False

End Sub
Private Sub a_Contratistas_Limpiar()
' limoia arreglo de contratistas
For i = 0 To 199
    a_Contratistas(0, i) = ""
    a_Contratistas(1, i) = ""
Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
DataBases_Cerrar
End Sub
Private Sub Inicializa()

Obj = "Arco Sumergido"
Objs = "Arcos Sumergidos"

Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnDesHacer = Toolbar.Buttons(6)
Set btnGrabar = Toolbar.Buttons(7)

Accion = ""
'old_accion = ""

btnSearch1.ToolTipText = "Busca Operador"
btnSearch2.ToolTipText = "Busca Ayudante"

Turno.MaxLength = 1

Campos_Enabled False

End Sub
Private Sub Detalle_Config()
Dim i As Integer, ancho As Integer

n_filas = 12
n_columnas = 19

Detalle.Left = 100
Detalle.WordWrap = True
Detalle.RowHeight(0) = 450
Detalle.Cols = n_columnas + 1
Detalle.Rows = n_filas + 1

Detalle.TextMatrix(0, 0) = ""
Detalle.TextMatrix(0, 1) = "Cant"
Detalle.TextMatrix(0, 2) = "Pz"
Detalle.TextMatrix(0, 3) = "ala1"
Detalle.TextMatrix(0, 4) = "esp1"
Detalle.TextMatrix(0, 5) = "ala2"
Detalle.TextMatrix(0, 6) = "esp2"
Detalle.TextMatrix(0, 7) = "alt"
Detalle.TextMatrix(0, 8) = "esp ala"
Detalle.TextMatrix(0, 9) = "largo"
Detalle.TextMatrix(0, 10) = "Kg Unitario"
Detalle.TextMatrix(0, 11) = "Kg Total"
Detalle.TextMatrix(0, 12) = "NV"
Detalle.TextMatrix(0, 13) = "Kg Fundente"
Detalle.TextMatrix(0, 14) = "Kg Alambre"
Detalle.TextMatrix(0, 15) = "Nº  Cordon"
Detalle.TextMatrix(0, 16) = "Filete"
Detalle.TextMatrix(0, 17) = "Hora Inicio"
Detalle.TextMatrix(0, 18) = "Hora Termino"
Detalle.TextMatrix(0, 19) = "Destino"

Detalle.ColWidth(0) = 300
Detalle.ColWidth(1) = 500
Detalle.ColWidth(2) = 700
Detalle.ColWidth(3) = 500
Detalle.ColWidth(4) = 500
Detalle.ColWidth(5) = 500
Detalle.ColWidth(6) = 500
Detalle.ColWidth(7) = 500
Detalle.ColWidth(8) = 500
Detalle.ColWidth(9) = 600
Detalle.ColWidth(10) = 600
Detalle.ColWidth(11) = 600
Detalle.ColWidth(12) = 700
Detalle.ColWidth(13) = 700
Detalle.ColWidth(14) = 500
Detalle.ColWidth(15) = 500
Detalle.ColWidth(16) = 500
Detalle.ColWidth(17) = 500
Detalle.ColWidth(18) = 500
Detalle.ColWidth(19) = 1500

ancho = 350 ' con scroll vertical

For i = 0 To n_columnas
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
'    Detalle.col = 5
'    Detalle.CellForeColor = vbRed
Next

txtEdit.Text = ""

End Sub
Private Sub Numero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then After_Enter
End Sub
Private Sub After_Enter()
If Val(Numero.Text) < 1 Then Beep: Exit Sub
Select Case Accion
Case "Agregando"

    RsAs.Seek "=", Numero.Text, 1
    
    If RsAs.NoMatch Then
        Campos_Enabled True
        Numero.Enabled = False
        Fecha.SetFocus
        btnGrabar.Enabled = True
        btnSearch1.visible = True
        btnSearch2.visible = True
    Else
        Doc_Leer
        MsgBox Obj & " YA EXISTE"
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
    End If
    
Case "Modificando"

    RsAs.Seek "=", Numero.Text, 1
    
    If RsAs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Botones_Enabled 0, 0, 0, 0, 1, 1
        Doc_Leer
        Campos_Enabled True
        Numero.Enabled = False
    End If

Case "Eliminando"
    RsAs.Seek "=", Numero.Text, 1
    If RsAs.NoMatch Then
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
    
    RsAs.Seek "=", Numero.Text, 1
    If RsAs.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Doc_Leer
        Numero.Enabled = False
        Detalle.Enabled = True
    End If
    
End Select

End Sub
Private Sub Doc_Leer()

'DETALLE

With RsAs

.Seek "=", Numero.Text, 1

If Not .NoMatch Then

    Do While Not RsAs.EOF
    
        If RsAs!Numero = Numero.Text Then
        
            i = RsAs!linea
            
            If i = 1 Then
            
                ' CABECERA
                
                Fecha.Text = Format(RsAs!Fecha, Fecha_Format)
                Turno.Text = !Turno
                NEquipo = !Equipo
                Rut1.Text = ![RUT Operador1]
                Rut2.Text = NoNulo(![RUT Operador2])
                
                Amperes.Text = !Amperes
                Voltaje.Text = !Voltaje
                
                Velocidad.Text = NoNulo(!Velocidad)
                
                Fundente.Text = NoNulo(!Fundente)
                Alambre.Text = NoNulo(!Alambre)
                
            End If
            
            Detalle.TextMatrix(i, 1) = !Cantidad
            Detalle.TextMatrix(i, 2) = !TipoPieza
            Detalle.TextMatrix(i, 3) = !dim1 ' ala1
            Detalle.TextMatrix(i, 4) = !dim2 ' esp1
            Detalle.TextMatrix(i, 5) = !dim3 ' ala2
            Detalle.TextMatrix(i, 6) = !dim4 ' esp2
            Detalle.TextMatrix(i, 7) = !dim5 ' altura
            Detalle.TextMatrix(i, 8) = !dim6 ' esp.al
            Detalle.TextMatrix(i, 9) = !dim7 ' largo
            
            Detalle.TextMatrix(i, 10) = ![PesoUnitario]
            Detalle.TextMatrix(i, 11) = ![PesoTotal]
            
            Detalle.TextMatrix(i, 12) = ![Nv]
            Detalle.TextMatrix(i, 13) = !kgfundente
            Detalle.TextMatrix(i, 14) = !kgalambre
            Detalle.TextMatrix(i, 15) = !numerocordones
            Detalle.TextMatrix(i, 16) = !tamanofilete
            
'            Detalle.TextMatrix(i, 16) = Format(!horainicio, "hh:mm")
'            Detalle.TextMatrix(i, 17) = Format(!horatermino, "hh:mm")
            Detalle.TextMatrix(i, 17) = Decimal2Hora(!horainicio)
            Detalle.TextMatrix(i, 18) = Decimal2Hora(!horatermino)
            
            For j = 0 To 199
                If a_Contratistas(0, j) = ![Rut destino] Then
                    Detalle.TextMatrix(i, 19) = a_Contratistas(1, j)
                    Exit For
                End If
            Next
                        
        Else
        
            Exit Do
            
        End If
        
        .MoveNext
        
    Loop
End If
End With

Trabajador_Lee Razon1, Rut1.Text
Trabajador_Lee Razon2, Rut2.Text

Detalle.Row = 1 ' para q' actualice la primera fila del detalle

'Actualiza

End Sub
Private Sub Trabajador_Lee(ObjText As Control, Rut)
RsTra.Seek "=", Rut
If Not RsTra.NoMatch Then
    ObjText.Text = RsTra![appaterno] & " " & RsTra![nombres]
End If
End Sub
Private Sub ContratistaD_Lee(Rut)
RsSc.Seek "=", Rut
If Not RsSc.NoMatch Then
    Razon1.Text = RsSc![Razon Social]
End If
End Sub
Private Function Doc_Validar() As Boolean
Dim Hora As String

Doc_Validar = False

If Rut1.Text = "" Then
    MsgBox "DEBE ELEGIR OPERADOR"
    btnSearch1.SetFocus
    Exit Function
End If

If Turno.Text = " " Then
    MsgBox "DEBE ELEGIR TURNO"
    Turno.SetFocus
    Exit Function
End If

If NEquipo.Text = "_" Then
    MsgBox "DEBE ELEGIR Nº EQUIPO"
    NEquipo.SetFocus
    Exit Function
End If

For i = 1 To n_filas

    ' cant
    If Val(Detalle.TextMatrix(i, 1)) <> 0 Then
    
        If Not CampoReq_Valida(Detalle.TextMatrix(i, 2), i, 2) Then Exit Function
        
'        If Not LargoString_Valida(Detalle.TextMatrix(i, 2), 3, i, 2) Then Exit Function
        
'        If Not CampoReq_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
'        If Not LargoString_Valida(Detalle.TextMatrix(i, 3), 30, i, 3) Then Exit Function
        
        Select Case Detalle.TextMatrix(i, 2)
        Case "V"
            If Not Numero_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
            If Not Numero_Valida(Detalle.TextMatrix(i, 4), i, 4) Then Exit Function
            If Not Numero_Valida(Detalle.TextMatrix(i, 5), i, 5) Then Exit Function
            If Not Numero_Valida(Detalle.TextMatrix(i, 6), i, 6) Then Exit Function
            If Not Numero_Valida(Detalle.TextMatrix(i, 7), i, 7) Then Exit Function
            If Not Numero_Valida(Detalle.TextMatrix(i, 8), i, 8) Then Exit Function
            If Not Numero_Valida(Detalle.TextMatrix(i, 9), i, 9) Then Exit Function
        Case "T"
            If Not Numero_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
            If Not Numero_Valida(Detalle.TextMatrix(i, 5), i, 5) Then Exit Function
            If Not Numero_Valida(Detalle.TextMatrix(i, 6), i, 6) Then Exit Function
            If Not Numero_Valida(Detalle.TextMatrix(i, 9), i, 9) Then Exit Function
        Case "P"
            If Not Numero_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
            If Not Numero_Valida(Detalle.TextMatrix(i, 4), i, 4) Then Exit Function
            If Not Numero_Valida(Detalle.TextMatrix(i, 9), i, 9) Then Exit Function
        End Select
        
        ' nv
        If Not Numero_Valida(Detalle.TextMatrix(i, 12), i, 12) Then Exit Function
        
        ' kg fundente
        If Not Numero_Valida(Detalle.TextMatrix(i, 13), i, 13) Then Exit Function
        
        ' kg alambre
        If Not Numero_Valida(Detalle.TextMatrix(i, 14), i, 14) Then Exit Function
        
        ' numero cordones, default 1
        If Not Numero_Valida(Detalle.TextMatrix(i, 15), i, 15) Then Exit Function
        
        ' filete
        If Not Numero_Valida(Detalle.TextMatrix(i, 16), i, 15) Then Exit Function
        
        Hora = Trim(Detalle.TextMatrix(i, 17))
        If Hora = "" Then
            MsgBox "Debe digitar Hora Inicio"
            Detalle.Row = i
            Detalle.col = 17
            Exit Function
        End If
        If Not Horas_Validas(Hora) Then
            Detalle.Row = i
            Detalle.col = 17
            Exit Function
        End If
        
        Hora = Trim(Detalle.TextMatrix(i, 18))
        If Hora = "" Then
            MsgBox "Debe digitar Hora Término"
            Detalle.Row = i
            Detalle.col = 18
            Exit Function
        End If
        If Not Horas_Validas(Hora) Then
            Detalle.Row = i
            Detalle.col = 18
            Exit Function
        End If
        
        If Trim(Detalle.TextMatrix(i, 19)) = "" Then
            MsgBox "Debe elegir Destino"
            Detalle.Row = i
            Detalle.col = 18
            Exit Function
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

Dim m_cantidad As Double, k As Integer, li As Integer

Doc_Detalle_Eliminar

With RsAs
j = 0

For li = 1 To n_filas

    m_cantidad = Val(Detalle.TextMatrix(li, 1))
    
    If m_cantidad = 0 Then
    Else
    
        .AddNew
        !Numero = Numero.Text
        !linea = li
        !Fecha = Fecha.Text
        !Turno = Turno.Text
        !Equipo = NEquipo.Text
        ![RUT Operador1] = Rut1.Text
        ![RUT Operador2] = Rut2.Text
        
        !Cantidad = m_cantidad
        ![TipoPieza] = Detalle.TextMatrix(li, 2)
        
        !dim1 = m_CDbl(Detalle.TextMatrix(li, 3)) ' ala1
        !dim2 = m_CDbl(Detalle.TextMatrix(li, 4)) ' esp1
        !dim3 = m_CDbl(Detalle.TextMatrix(li, 5)) ' ala2
        !dim4 = m_CDbl(Detalle.TextMatrix(li, 6)) ' esp2
        !dim5 = m_CDbl(Detalle.TextMatrix(li, 7)) ' altura
        !dim6 = m_CDbl(Detalle.TextMatrix(li, 8)) ' esp alma
        !dim7 = m_CDbl(Detalle.TextMatrix(li, 9)) ' largo
        
        !PesoUnitario = m_CDbl(Detalle.TextMatrix(li, 10)) ' kg uni
        !PesoTotal = m_CDbl(Detalle.TextMatrix(li, 11)) ' kg tot

        RsAs!Nv = Detalle.TextMatrix(li, 12)
        
        ![kgfundente] = m_CDbl(Detalle.TextMatrix(li, 13))
        ![kgalambre] = m_CDbl(Detalle.TextMatrix(li, 14))
        ![numerocordones] = m_CDbl(Detalle.TextMatrix(li, 15))
        ![tamanofilete] = m_CDbl(Detalle.TextMatrix(li, 16))
        
        ![horainicio] = Hora2Decimal(Detalle.TextMatrix(li, 17))
        ![horatermino] = Hora2Decimal(Detalle.TextMatrix(li, 18))
        
        ' busca rut del contratista
        For k = 0 To 199
            If a_Contratistas(1, k) = Detalle.TextMatrix(li, 19) Then
                ![Rut destino] = a_Contratistas(0, k)
                Exit For
            End If
        Next
        
        !Amperes = m_CDbl(Amperes.Text)
        !Voltaje = m_CDbl(Voltaje.Text)
        
        !Velocidad = Velocidad.Text
        
        !Fundente = Fundente.Text
        !Alambre = Alambre.Text
        
        .Update
        
    End If
    
Next
End With

End Sub
Private Sub Doc_Eliminar()

' borra CABECERA DE OT

Doc_Detalle_Eliminar

End Sub
Private Sub Doc_Detalle_Eliminar()

' elimina detalle
RsAs.Seek ">=", Numero.Text, 0

If Not RsAs.NoMatch Then

    Do While Not RsAs.EOF
    
        If RsAs!Numero <> Numero.Text Then Exit Do
    
        ' borra detalle
        RsAs.Delete
    
        RsAs.MoveNext
        
    Loop
    
End If

End Sub
Private Sub Campos_Limpiar()

Numero.Text = ""
Fecha.Text = Fecha_Vacia
Turno = " "
NEquipo = "_"

Rut1.Text = ""
Razon1.Text = ""
Rut2.Text = ""
Razon2.Text = ""

Detalle_Limpiar

Amperes.Text = ""
Voltaje.Text = ""
Velocidad.Text = ""
Fundente.Text = ""
Alambre.Text = ""

End Sub
Private Sub Detalle_Limpiar()
Dim fi As Integer, co As Integer
For fi = 1 To n_filas
    For co = 1 To n_columnas
        Detalle.TextMatrix(fi, co) = ""
    Next
Next
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
    
    Numero.Text = Documento_Numero_Nuevo(RsAs, "numero")
    
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
btnSearch1.Enabled = Si
btnSearch2.Enabled = Si
Turno.Enabled = Si
NEquipo.Enabled = Si

Detalle.Enabled = Si

Amperes.Enabled = Si
Voltaje.Enabled = Si
Velocidad.Enabled = Si
Fundente.Enabled = Si
Alambre.Enabled = Si

End Sub
Private Sub btnSearch1_Click()

Search.Muestra data_file, "trabajadores", "RUT", "ApPaterno]+' '+[ApMaterno]+' '+[Nombres", "Operador", "Operadores", "Activo"

Rut1.Text = Search.codigo
If Rut1.Text <> "" Then
    RsTra.Seek "=", Rut1.Text
    If RsTra.NoMatch Then
        MsgBox "CONTRATISTA NO EXISTE"
        Rut1.SetFocus
    Else
        Razon1.Text = Search.descripcion
    End If
End If

End Sub
Private Sub btnSearch2_Click()

Search.Muestra data_file, "trabajadores", "RUT", "ApPaterno]+' '+[ApMaterno]+' '+[Nombres", "Ayudante", "Ayudantes", "Activo"

Rut2.Text = Search.codigo
If Rut2.Text <> "" Then
    RsTra.Seek "=", Rut2.Text
    If RsTra.NoMatch Then
        MsgBox "CONTRATISTA NO EXISTE"
        Rut2.SetFocus
    Else
        Razon2.Text = Search.descripcion
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
EditKeyCodeP Detalle, txtEdit, KeyCode, Shift
End Sub
Sub EditKeyCodeP(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
' rutina que se ejecuta con los keydown de Edt
Dim m_col As Integer, m_fil As Integer, m_Pieza As String

m_fil = MSFlexGrid.Row
m_col = MSFlexGrid.col

Select Case KeyCode
Case vbKeyEscape ' Esc
    Edt.visible = False
    MSFlexGrid.SetFocus
Case vbKeyReturn ' Enter
    MSFlexGrid.SetFocus
    DoEvents
    Select Case m_col
    Case 2 ' pz
    
        m_Pieza = UCase(Detalle.TextMatrix(m_fil, 2))
        If m_Pieza <> "V" And m_Pieza <> "T" And m_Pieza <> "S" And m_Pieza <> "P" Then
            MsgBox "Debe ser V, T, S ó P"
            Detalle.TextMatrix(m_fil, 2) = ""
        Else
            Detalle.TextMatrix(m_fil, 2) = m_Pieza
            Linea_Actualiza
        End If
        
    Case 12 ' nv
    Case Else
        Linea_Actualiza
    End Select
    Cursor_Mueve MSFlexGrid
Case vbKeyUp ' Flecha Arriba
    Select Case m_col
    Case 12
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
    Case 12
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
Private Sub txtEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)

If MSFlexGrid.col = 11 Then Exit Sub

Select Case MSFlexGrid.col
Case 1 ' cant
    Edt.MaxLength = 3
Case 2 ' pz
    Edt.MaxLength = 1
Case 3 ' ala1
    Edt.MaxLength = 5
Case 4 ' esp1
    Edt.MaxLength = 2
Case 5 ' ala2
    Edt.MaxLength = 5
Case 6 ' esp2
    Edt.MaxLength = 2
Case 7 ' altura
    Edt.MaxLength = 5
Case 8 ' esp.al
    Edt.MaxLength = 2
Case 9 ' largo
    Edt.MaxLength = 5
Case 19 ' combo contratista
    CbContratistas.Left = MSFlexGrid.CellLeft + MSFlexGrid.Left
    CbContratistas.Top = MSFlexGrid.CellTop + MSFlexGrid.Top
    CbContratistas.visible = True
    CbContratistas.SetFocus
    Exit Sub
Case Else
    Edt.MaxLength = 5
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
    MSFlexGridEdit Detalle, txtEdit, 32
End If
End Sub
Private Sub Linea_Actualiza()

' actualiza solo linea, y totales generales

Dim fi As Integer, Cant As Integer, TipoPieza As String
Dim ala1 As Long, esp1 As Long, ala2 As Long, esp2 As Long, altura As Long, espal As Long, largo As Long, pu As Double

fi = Detalle.Row

Cant = m_CDbl(Detalle.TextMatrix(fi, 1))
TipoPieza = Detalle.TextMatrix(fi, 2)

ala1 = m_CDbl(Detalle.TextMatrix(fi, 3)) ' ala1
esp1 = m_CDbl(Detalle.TextMatrix(fi, 4)) ' esp1
ala2 = m_CDbl(Detalle.TextMatrix(fi, 5)) ' ala2
esp2 = m_CDbl(Detalle.TextMatrix(fi, 6)) ' esp2
altura = m_CDbl(Detalle.TextMatrix(fi, 7)) ' altura alma
espal = m_CDbl(Detalle.TextMatrix(fi, 8)) ' esp.alma
largo = m_CDbl(Detalle.TextMatrix(fi, 9)) ' largo

' peso unitario
pu = 0
Select Case TipoPieza
Case "V" ' viga
    pu = (ala1 * esp1 + ala2 * esp2 + (altura - esp1 - esp2) * espal) * PesoEspecificoV * largo / 1000000
Case "T" 'tubular
    pu = (2 * ala1 + 2 * ala2 - 8 * esp2) * esp2 * PesoEspecificoT * largo / 1000000
Case "S" 'tubest
    pu = largo * 32.5 / 1000
Case "P" ' plancha
    pu = ala1 * esp1 * PesoEspecificoP * largo / 1000000
End Select

Detalle.TextMatrix(fi, 10) = Format(pu, num_fmtgrl)
Detalle.TextMatrix(fi, 11) = Format(Cant * pu, num_fmtgrl)

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

End Sub
Private Sub Cursor_Mueve(MSFlexGrid As Control)

Dim m_fil As Integer, m_col As Integer, m_Tipo As String

m_fil = MSFlexGrid.Row
m_col = MSFlexGrid.col
m_Tipo = MSFlexGrid.TextMatrix(m_fil, 2)

If MSFlexGrid.col = 19 Then
    If MSFlexGrid.Row + 1 < MSFlexGrid.Rows Then
        MSFlexGrid.col = 1
        MSFlexGrid.Row = MSFlexGrid.Row + 1
    End If
Else

    If m_col = 9 Then ' largo
        MSFlexGrid.col = MSFlexGrid.col + 3
    Else
        Select Case m_Tipo
        Case "V" ' viga
            MSFlexGrid.col = MSFlexGrid.col + 1
        Case "T"
            Select Case m_col
            Case 3
                MSFlexGrid.col = MSFlexGrid.col + 2
            Case 6
                MSFlexGrid.col = MSFlexGrid.col + 3
            Case Else
                MSFlexGrid.col = MSFlexGrid.col + 1
            End Select
        Case "P" ' plancha
            Select Case m_col
            Case 4
                MSFlexGrid.col = 9
            Case Else
                MSFlexGrid.col = MSFlexGrid.col + 1
            End Select
        Case Else
            MSFlexGrid.col = MSFlexGrid.col + 1
        End Select
        
    End If
    
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
prt.Print Tab(tab0 + 28); Left(Razon1.Text, 31)

prt.Font.Bold = False
prt.Font.Size = fc
prt.Print Tab(tab0); Empresa.Direccion;
prt.Font.Size = fc
prt.Print Tab(tab0 + tab40); "OBRA      : "
prt.Print Tab(tab0); "Teléfono: "; Empresa.Telefono1;
prt.Font.Size = fe
prt.Font.Bold = True
'prt.Print Tab(tab0 + 28); Left(Format(Mid(ComboNV.Text, 8), ">"), 32)

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
'prt.Print Tab(tab4 - 5); m_Format(TotalPrecio, "$#,###,###,###")
prt.Font.Bold = False
prt.Print ""

'prt.Print Tab(tab0); "FECHA ENTREGA : "; Entrega.Text
prt.Print Tab(tab0); "OBSERVACIONES :";

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
'Private Function Horas_Validas(Hora As Object, Optional Asume) As Boolean
Private Function Horas_Validas(Hora As String) As Boolean
' valida hora en formato hh:mm hasta las 99:59
' Fecha : Objeto TextBox
' Asume : Hora que se asume si se presiona ENTER, si Asume es vacío  entonces, hora="__:__"
Dim s As String, f As Date, pos As Integer, HH As String, mm As String

Horas_Validas = True
s = Replace(Hora, "_", "")
'If s = ":" Then
'    If Not IsMissing(Asume) Then Hora = Format(Asume, "hh:mm")
'    Exit Function
'End If

pos = InStr(1, s, ":")
If pos = 0 Then
    'no trae los dos puntos
    GoTo HoraMala
End If

HH = Trim(Left(s, pos - 1))
If Len(HH) = 0 Then
    GoTo HoraMala
End If
HH = Val(HH)

mm = Trim(Mid(s, pos + 1))
If Len(mm) = 0 Then
    GoTo HoraMala
End If
mm = Val(Mid(s, pos + 1))

If HH > 99 Then GoTo HoraMala
If mm > 59 Then GoTo HoraMala

s = Format(HH, "00") & ":" & Format(mm, "00")
'Hora = Format(f, "hh:mm")
'Hora = Format(s, "00:00")
Hora = s

Exit Function

HoraMala:
Horas_Validas = False
MsgBox "HORA NO VÁLIDA"
'Hora.visible = True
'Hora.SetFocus
End Function
