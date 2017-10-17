VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form ITO_Fabricacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vales ITO"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   3600
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.CommandButton btnArchivoElegir 
      Caption         =   "Capturador"
      Height          =   300
      Left            =   6480
      TabIndex        =   27
      Top             =   1200
      Width           =   1095
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11220
      _ExtentX        =   19791
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
            Object.ToolTipText     =   "Mantención de Subcontratistas"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox Nv 
      Height          =   300
      Left            =   2040
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox ComboOT 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox ComboNV 
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   3
      Left            =   120
      MaxLength       =   30
      TabIndex        =   23
      Top             =   5565
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   2
      Left            =   120
      MaxLength       =   30
      TabIndex        =   22
      Top             =   5265
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   1
      Left            =   120
      MaxLength       =   30
      TabIndex        =   21
      Top             =   4965
      Width           =   5000
   End
   Begin VB.ComboBox ComboMarca 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   8040
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox ComboPlano 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame 
      Caption         =   "Contratista"
      Height          =   735
      Left            =   3240
      TabIndex        =   10
      Top             =   400
      Width           =   5055
      Begin VB.CommandButton btnSearch 
         Height          =   300
         Left            =   1920
         Picture         =   "ITO_Fabricacion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   300
      End
      Begin VB.TextBox Razon 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   15
         Top             =   360
         Width           =   2655
      End
      Begin MSMask.MaskEdBox Rut 
         Height          =   300
         Left            =   600
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbl 
         Caption         =   "SEÑOR(ES)"
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   14
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox txtEditITO 
      Height          =   285
      Left            =   8040
      TabIndex        =   17
      Text            =   "txtEditITO"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   0
      Left            =   120
      MaxLength       =   30
      TabIndex        =   20
      Top             =   4665
      Width           =   5000
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   840
      TabIndex        =   6
      Top             =   1200
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
      TabIndex        =   4
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
   Begin MSFlexGridLib.MSFlexGrid Detalle 
      Height          =   2805
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4948
      _Version        =   327680
      Enabled         =   0   'False
      ScrollBars      =   2
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   8040
      Top             =   1320
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
            Picture         =   "ITO_Fabricacion.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_Fabricacion.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_Fabricacion.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_Fabricacion.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_Fabricacion.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_Fabricacion.frx":065C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITO_Fabricacion.frx":076E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1215
      Left            =   5280
      TabIndex        =   28
      Top             =   4680
      Width           =   4095
   End
   Begin VB.Label lbl 
      Caption         =   "FABRICACIÓN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lbl 
      Caption         =   "ITO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.Label TotalKilos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6120
      TabIndex        =   25
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   24
      Top             =   4400
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "OBRA"
      Height          =   255
      Index           =   7
      Left            =   3240
      TabIndex        =   8
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "FECHA"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1260
      Width           =   615
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
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
End
Attribute VB_Name = "ITO_Fabricacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Ws As Workspace ' para trasacciones
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button
'Private DbD As Database ', RsSc As Recordset

Private SqlRsSc As New ADODB.Recordset ' para buscar id contratista

Private Dbm As Database, RsITOc As Recordset, RsITOd As Recordset
Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, d As Variant
Private RsNvPla As Recordset, RsPd As Recordset
'Private RsNVc As Recordset
Private mNv As NotaVenta
Private RsOTc As Recordset, RsOTd As Recordset, RsITOpd As Recordset
Private n_filas As Integer, n_columnas As Integer
Private Rev(2999) As String
Private prt As Printer, n_Copias As Integer
Private DbH As Database, RsITOcH As Recordset
' 0: numero,  1: nombre obra, 2: esquema de pintura o galvanizado
Private a_Nv(2999, 2) As String, m_Nv As Double, m_NvArea As Integer
Private Const NumeroCampos As Integer = 11
Private aCampos(NumeroCampos) As String ' para funcion split
Private Sub btnArchivoElegir_Click()
ArchivoAbrir
End Sub
Private Sub ArchivoAbrir()

Dim mPath As String, mPathArchivo As String, mArchivo As String, p As Integer

cd.DialogTitle = "Buscar Carpeta"
cd.Filter = "Texto (*.txt)|*.txt|Todos los Archivos (*.*)|*.*"

' busca ultima ruta
mPath = GetSetting("scp", "gd", "ruta")
If mPath = "" Then mPath = "C:"

mPath = Directorio(mPath)

cd.InitDir = mPath
cd.ShowOpen

mPathArchivo = cd.filename

If mPathArchivo = "" Then
    Exit Sub
End If

mPathArchivo = cd.filename

' separa path y archivo
p = InStrLast(mPathArchivo, "\")
If p > 0 Then

    ' guarda ultima ruta usada
    SaveSetting "scp", "gd", "ruta", mPathArchivo

    mPath = Left(mPathArchivo, p)
    mArchivo = Mid(mPathArchivo, p + 1)
    
'    lblCarpeta.Caption = m_Path
    
    ArchivoLeer mPath, mArchivo
    
End If

End Sub
Private Sub ArchivoLeer(Path As String, Archivo As String)
' lee archivo del honey
' abre archivo
Dim li As Integer, mPlano As String, mMarca As String
Dim RsPaso As Recordset, arreglo(99, 9) As String
Dim i As Integer, j As Integer, k As Integer, Reg As String

li = 0

' lee archivo y lo deja en arreglo
Open Path & Archivo For Input As #1
Do While Not EOF(1)
    Line Input #1, Reg
'    Debug.Print Reg

    If Len(Reg) > 1 Then

        split Reg, "/", aCampos, NumeroCampos
        
        li = li + 1
        
        arreglo(li, 0) = aCampos(6) ' nv
        arreglo(li, 1) = aCampos(2) ' plano
        arreglo(li, 2) = aCampos(3) ' marca
        arreglo(li, 3) = 1 ' cantidad
        arreglo(li, 4) = aCampos(1) ' Id contratista
        
    End If
    
Loop
Close #1
'//////////////////////////////////
' buscas marcas repetidas
For i = 1 To li - 1

    For j = i + 1 To li
    
        If arreglo(i, 1) = arreglo(j, 1) Then
            
            ' suma cantidad
            arreglo(i, 3) = arreglo(i, 3) + 1
            
            ' elimina fila j y desplaza filas hacia arriba
            For k = j To li
                arreglo(k, 0) = arreglo(k + 1, 0) ' nv
                arreglo(k, 1) = arreglo(k + 1, 1) ' plano
                arreglo(k, 2) = arreglo(k + 1, 2) ' marca
                arreglo(k, 3) = arreglo(k + 1, 3) ' cant
                arreglo(k, 4) = arreglo(k + 1, 4) ' id contratista
            Next
            li = li - 1
            
            j = j - 1
            
        End If
    Next

Next
'For i = 1 To li
'    Debug.Print i & "|" & arreglo(i, 1) & "|" & arreglo(i, 2)
'Next
'//////////////////////////////////

If li > n_filas Then
    MsgBox "Hay mas de " & n_filas & " lecturas, solo se muestran las primeras " & n_filas
    li = n_filas
End If

' lleva arreglo a pantalla (grilla)
For i = 1 To li
  
    If i = 1 Then
        
        ' pone nv
        Nv.Text = arreglo(i, 0)
        Nv_LostFocus
        
        ' busca contratista
        Rut.Text = ContratistaBuscarXId(arreglo(i, 4))
        
        
    End If
    
    mPlano = arreglo(i, 1)
    mMarca = arreglo(i, 2)
    Detalle.TextMatrix(i, 1) = mPlano
    
    MarcaBuscar mMarca, False, i
    
Next

End Sub

Private Sub ComboNV_Click()

MousePointer = vbHourglass

ComboPlano.visible = False
ComboMarca.visible = False

'If rut.Text = "" Then
'    MsgBox "no hay contratista"
'Else
'    MsgBox "hay contratista"
'End If
'Exit Sub ' OOOJJJOOO

i = 0
m_Nv = Val(Left(ComboNV.Text, 6))
Nv.Text = m_Nv
ComboPlano.Clear

ComboPlano.AddItem " " ' para "borrar" linea
Rev(i) = " "

RsNvPla.Seek ">=", m_Nv, ""
If Not RsNvPla.NoMatch Then
    Do While Not RsNvPla.EOF
        If RsNvPla!Nv = m_Nv Then
            ComboPlano.AddItem RsNvPla!Plano
            i = i + 1
            Rev(i) = RsNvPla!Rev
        Else
            Exit Do
        End If
        RsNvPla.MoveNext
    Loop
End If

Detalle_Limpiar
ComboMarca.Clear

If Rut.Text <> "" Then Detalle.Enabled = True

MousePointer = vbDefault
End Sub
Private Sub PlanosPendientesBuscar()

' puebla combobox con planos pendientes de recibir del contratista escogido

' busca planos pendientes de toda la nv



' busca planos pendientes de nv + contratista

End Sub
Private Sub ComboPlano_Click()
' supuesto: el numero del plano es único para toda nv
Dim old_plano As String, fil As Integer, np As String, indice_plano As Integer, Marca_Unica_enPlano As String
Marca_Unica_enPlano = ""

old_plano = Detalle

fil = Detalle.Row
np = ComboPlano.Text
indice_plano = 0
ComboMarca.Clear
'RsPd.Seek "=", m_Nv, m_NvArea, np, 1
RsPd.Seek ">=", m_Nv, m_NvArea, np, 0
If Not RsPd.NoMatch Then
    Do While Not RsPd.EOF
        If RsPd!Plano = np Then
            indice_plano = indice_plano + 1
            ComboMarca.AddItem RsPd!Marca
            Debug.Print RsPd!Marca
            If indice_plano = 1 Then Marca_Unica_enPlano = RsPd!Marca
        Else
            Exit Do
        End If
        RsPd.MoveNext
    Loop
End If

ComboPlano.visible = False
Detalle = ComboPlano.Text

If Detalle <> old_plano Then
    For i = 2 To n_columnas
        Detalle.TextMatrix(fil, i) = ""
    Next
End If

Detalle.TextMatrix(fil, 2) = Rev(ComboPlano.ListIndex)

'If indice_plano = 1 Then
'    ' plano tiene una sola marca
'    ComboMarca.ListIndex = 0 ' para que pueble grid
'    Detalle.TextMatrix(fil, 3) = Marca_Unica_enPlano ' ComboMarca.Text
'End If

End Sub
Private Sub ComboPlano_LostFocus()
ComboPlano.visible = False
End Sub
Private Sub ComboMarca_Poblar(Plano As String)
' llena combo marcas
ComboMarca.Clear
'RsPd.Seek "=", m_Nv, m_NvArea, Plano, 1
RsPd.Seek ">=", m_Nv, m_NvArea, Plano, 0
If Not RsPd.NoMatch Then
    Do While Not RsPd.EOF
        If RsPd!Nv = m_Nv And RsPd!Plano = Plano Then
            ComboMarca.AddItem RsPd!Marca
        Else
            Exit Do
        End If
        RsPd.MoveNext
    Loop
End If
End Sub
Private Sub ComboMarca_Click()
MarcaBuscar ComboMarca.Text, True, Detalle.Row
End Sub
Private Sub MarcaBuscar(Marca As String, combo As Boolean, Fila As Integer)
' busca marca en archivo mdb
Dim m_Plano As String, m_Marca As String ', fil As Integer
Dim c_ot As Integer, c_itof As Integer

Dim Marca_Recibida As Boolean
Marca_Recibida = False

'fil = Detalle.Row
ComboMarca.visible = False
m_Plano = Detalle.TextMatrix(Fila, 1)
m_Marca = Marca

'///
' verifica si Plano-Marca ya están en esta ITO
If combo Then
    For i = 1 To n_filas
        If m_Plano = Detalle.TextMatrix(i, 1) And m_Marca = Detalle.TextMatrix(i, 3) Then
            Beep
            MsgBox "MARCA YA EXISTE EN ITO"
            Detalle.Row = i
            Detalle.col = 3
            Detalle.SetFocus
            Exit Sub
        End If
    Next
End If
'///

'Detalle.TextMatrix(fila, 3) = m_Marca
'Detalle = m_Plano
' busca marca en plano
RsPd.Seek ">=", m_Nv, m_NvArea, m_Plano, 0
If Not RsPd.NoMatch Then

    Do While Not RsPd.EOF
    
        If RsPd!Marca = m_Marca Then
            
            c_ot = RsPd![OT fab]
            c_itof = RsPd![ITO fab]
            
            '/////
            ' verifica que esté asignado algo
            If c_ot = 0 Then
                Beep
                MsgBox "Marca """ & m_Marca & """" & vbCr & _
                       "No está asignada"
                Detalle.TextMatrix(Fila, 3) = ""
                Detalle.SetFocus
                Exit Sub
            End If
            '/////
            
            '/////
            ' verifica que quede algo por recibir (en general)
            ' para itos nuevas
            If Accion = "Agregando" Then
                If c_ot - c_itof <= 0 Then
                    Beep
                    MsgBox "Marca """ & m_Marca & """" & vbCr & _
                           "Ya fue recibida"
                    Detalle.TextMatrix(Fila, 3) = ""
                    Detalle.SetFocus
                    Exit Sub
                End If
            End If
            
            '/////
        
            Detalle.TextMatrix(Fila, 2) = RsPd!Rev
            Detalle.TextMatrix(Fila, 3) = RsPd!Marca
            Detalle.TextMatrix(Fila, 4) = RsPd!Descripcion
            Detalle.TextMatrix(Fila, 5) = c_ot
            Detalle.TextMatrix(Fila, 6) = c_itof
            If Not combo Then
                Detalle.TextMatrix(Fila, 9) = 1 ' ojo recibe una sola
                Detalle.Row = Fila
            End If
            Detalle.TextMatrix(Fila, 10) = Replace(RsPd![Peso], ",", ".")
            Linea_Actualiza
            Exit Do
        End If
        RsPd.MoveNext
    Loop
End If

' busca OTs donde se asignó esta marca (y al Contratista especificado)
i = 0
RsOTd.Seek ">=", m_Nv, m_NvArea, m_Plano, m_Marca
If Not RsOTd.NoMatch Then

    Do While Not RsOTd.EOF
    
        If m_Nv <> RsOTd!Nv Or m_Plano <> RsOTd!Plano Or m_Marca <> RsOTd!Marca Then Exit Do
        
        RsOTc.Seek "=", RsOTd!Numero
        If Not RsOTc.NoMatch Then
        
            If Rut.Text = RsOTc![Rut contratista] Then
            
                i = i + 1
                
                If i = 1 Then
                
                    If Accion = "Agregando" Then
                    
                        If RsOTd![Cantidad] > RsOTd![Cantidad Recibida] Then
                        
                            Detalle.TextMatrix(Fila, 7) = RsOTd!Numero   'NºOT
                            Detalle.TextMatrix(Fila, 8) = RsOTd!Cantidad 'Cant
                            Detalle.TextMatrix(Fila, 12) = RsOTd![Cantidad Recibida]
                            Detalle.TextMatrix(Fila, 13) = RsOTd!Fecha   'fecha OT
                            
                            ' si queda una por recibir => la recibe
    '                        If Val(Detalle.TextMatrix(fil, 8)) - Val(Detalle.TextMatrix(fil, 12)) = 1 Then
    '                            Detalle.TextMatrix(fil, 9) = 1
    '                        End If
                            
                            Marca_Recibida = False
                            
                        Else
                        
                            ' marca ya está recibida
                            i = 0
                            Marca_Recibida = True
                            
                        End If
                        
                    End If
                    
                    If Accion = "Modificando" Then
                        Detalle.TextMatrix(Fila, 7) = RsOTd!Numero   'NºOT
                        Detalle.TextMatrix(Fila, 8) = RsOTd!Cantidad 'Cant
                        Detalle.TextMatrix(Fila, 13) = RsOTd!Fecha   'fecha OT
'                        Detalle.TextMatrix(fila, 12) = RsOTd![Cantidad Recibida]
                    End If
                
                Else
                    ' hay más de una OT
                    If RsOTd![Cantidad] > RsOTd![Cantidad Recibida] Then
                        Linea_Nueva Fila
                    End If
'                    Detalle.Textmatrix(fil, 7)) = Detalle.Textmatrix(fil, 7)) & "/" & RsOTd("Número")
'                    Detalle.Textmatrix(fil, 8)) = Detalle.Textmatrix(fil, 8)) + RsOTd("Cantidad")
                End If
            End If
        End If
        RsOTd.MoveNext
    Loop
    
    If Marca_Recibida Then
    
        Beep
        MsgBox "Marca """ & m_Marca & """ ya está Recibida" & vbCr & "para este Contratista"
        For j = 3 To n_columnas
            Detalle.TextMatrix(Fila, j) = ""
        Next
        Detalle.col = 3
        Exit Sub
    End If
    
    If i = 0 Then
        Beep
        MsgBox "Marca """ & m_Marca & """ no fue Asignada" & vbCr & "a este Contratista"
        For j = 3 To n_columnas
            Detalle.TextMatrix(Fila, j) = ""
        Next
        Detalle.col = 3
    End If
    
End If

End Sub

Private Sub ComboMarca_LostFocus()
ComboMarca.visible = False
End Sub
Private Sub Linea_Nueva(old_fila As Integer)
Dim i As Integer, new_fila As Integer
' copia fila old en nueva
new_fila = old_fila + 1
For i = 1 To n_columnas
    Select Case i
    Case 7
        Detalle.TextMatrix(new_fila, 7) = RsOTd!Numero
    Case 8
        Detalle.TextMatrix(new_fila, 8) = RsOTd!Cantidad
    Case 12
        Detalle.TextMatrix(new_fila, 12) = RsOTd![Cantidad Recibida]
    Case Else
        Detalle.TextMatrix(new_fila, i) = Detalle.TextMatrix(old_fila, i)
    End Select
Next
End Sub
Private Sub Fecha_GotFocus()
ComboPlano.visible = False
ComboMarca.visible = False
End Sub
Private Sub Fecha_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Fecha_LostFocus()
d = Fecha_Valida(Fecha, Now)
End Sub

Private Sub Form_Load()

'Set Ws = Workspaces(0)

n_filas = 12
n_columnas = 13 '12 '11

' abre archivos
'Set DbD = OpenDatabase(data_file)
'Set RsSc = DbD.OpenRecordset("Contratistas")
'RsSc.Index = "RUT"
Set Dbm = OpenDatabase(mpro_file)

If Not Usuario.ObrasTerminadas Then

    ' si usuario esta en obras en proceso, abre movs historico
    Dim hist_file As String
    hist_file = Movs_Path(Empresa.Rut, True)
    Set DbH = OpenDatabase(hist_file)
    Set RsITOcH = DbH.OpenRecordset("ITO Fab Cabecera")
    RsITOcH.Index = "Numero"
    
End If

Set RsOTc = Dbm.OpenRecordset("OT Fab Cabecera")
RsOTc.Index = "Numero"
Set RsOTd = Dbm.OpenRecordset("OT Fab Detalle")
RsOTd.Index = "NV-Plano-Marca"

Set RsITOc = Dbm.OpenRecordset("ITO Fab Cabecera")
RsITOc.Index = "Numero"

Set RsITOd = Dbm.OpenRecordset("ITO Fab Detalle")
RsITOd.Index = "Numero-Linea"

'Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
'RsNVc.Index = Nv_Index ' "Número"

nvListar Usuario.Nv_Activas

' Combo obra
ComboNV.AddItem " "
For i = 1 To nvTotal
    a_Nv(i, 0) = aNv(i).Numero
    a_Nv(i, 1) = aNv(i).Obra
    If aNv(i).galvanizado Then a_Nv(i, 2) = "Galvanizado"
    If aNv(i).pintura Then a_Nv(i, 2) = "Pintura"
    ComboNV.AddItem Format(aNv(i).Numero, "0000") & " - " & aNv(i).Obra
Next

Set RsNvPla = Dbm.OpenRecordset("Planos Cabecera")
'RsNvPla.Index = "Nota-Línea"
RsNvPla.Index = "NV-Plano"

Set RsPd = Dbm.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Item"

Inicializa
Detalle_Config

Privilegios

m_NvArea = 0

' ojo
'Doc_Actualizar 2155, Dbm
'Doc_Actualizar 2125, Dbm
'Doc_Actualizar 2126, Dbm
'Doc_Actualizar 2064, Dbm

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

btnSearch.visible = False
btnSearch.ToolTipText = "Busca Contatista"
Campos_Enabled False
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
Detalle.TextMatrix(0, 2) = "Rev"                 '*
Detalle.TextMatrix(0, 3) = "Marca"
Detalle.TextMatrix(0, 4) = "Descripción"         '*
Detalle.TextMatrix(0, 5) = "Total Asig."         '*
Detalle.TextMatrix(0, 6) = "Total Reci."         '*
Detalle.TextMatrix(0, 7) = "Nº OT"               '*
Detalle.TextMatrix(0, 8) = "Asignada"            '*
Detalle.TextMatrix(0, 9) = "  a Reci."
Detalle.TextMatrix(0, 10) = "Peso Unitario"      '*
Detalle.TextMatrix(0, 11) = "Peso TOTAL"         '*
'etalle.Textmatrix(0, 12) = "Recibido de OT      '*
'etalle.Textmatrix(0, 13) = "fecha OT

Detalle.ColWidth(0) = 250
Detalle.ColWidth(1) = 2300 ' plano
Detalle.ColWidth(2) = 400
Detalle.ColWidth(3) = 2300 ' marca
Detalle.ColWidth(4) = 1200
Detalle.ColWidth(5) = 500
Detalle.ColWidth(6) = 500
Detalle.ColWidth(7) = 800
Detalle.ColWidth(8) = 500
Detalle.ColWidth(9) = 500
Detalle.ColWidth(10) = 700
Detalle.ColWidth(11) = 800

' recibido de una OT especifica
' columna invisible
Detalle.ColWidth(12) = 0
Detalle.ColWidth(13) = 0

ancho = Detalle.Left + 250 ' con scroll vertical

TotalKilos.Width = Detalle.ColWidth(11)
For i = 0 To n_columnas
    If i = 11 Then TotalKilos.Left = ancho + Detalle.Left - 350
    ancho = ancho + Detalle.ColWidth(i)
Next

Detalle.Width = ancho
Me.Width = ancho + Detalle.Left * 2

For i = 1 To n_filas
    Detalle.TextMatrix(i, 0) = i
Next

' col y row fijas
'Detalle.BackColorFixed = vbCyan

' establece colores a columnas
' columnas    modificables : NEGRAS
' columnas no modificables : ROJAS
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
    Detalle.col = 7
    Detalle.CellForeColor = vbRed
    Detalle.col = 8
    Detalle.CellForeColor = vbRed
    
    Detalle.col = 10
    Detalle.CellForeColor = vbRed
    Detalle.col = 11
    Detalle.CellForeColor = vbRed
Next

txtEditITO = ""

'Top = Form_CentraY(Me)
'Left = Form_CentraX(Me)

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

    If Not Usuario.ObrasTerminadas Then
        ' busca en historico
        RsITOcH.Seek "=", Numero.Text
        If Not RsITOcH.NoMatch Then
            MsgBox Obj & " YA EXISTE" & Chr(10) & "EN OBRAS TERMINADAS"
            Detalle.Enabled = False
            Campos_Limpiar
            Numero.Enabled = True
            Numero.SetFocus
            
            Exit Sub
            
        End If
    End If

    RsITOc.Seek "=", Numero.Text
    
    If RsITOc.NoMatch Then
    
        Campos_Enabled True
        Numero.Enabled = False
        Detalle.Enabled = False
        
        Fecha.Text = Format(Now, Fecha_Format)
        If Usuario.AccesoTotal Then
            Fecha.SetFocus
        Else
            ComboNV.SetFocus
        End If
        
        btnGrabar.Enabled = True
        btnSearch.visible = True
    Else
        Doc_Leer
        MsgBox Obj & " YA EXISTE"
        Detalle.Enabled = False
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
        
    End If
    
Case "Modificando"

    RsITOc.Seek "=", Numero.Text
    
    If RsITOc.NoMatch Then
        MsgBox Obj & " NO EXISTE"
        Numero.SetFocus
    Else
        Botones_Enabled 0, 0, 0, 0, 1, 1
        Doc_Leer
        btnSearch.visible = True
        Campos_Enabled True
        Numero.Enabled = False
        If Usuario.AccesoTotal Then
            Fecha.SetFocus
        Else
            ComboNV.SetFocus
        End If
        btnGrabar.Enabled = True
        btnSearch.visible = True
        
    End If

Case "Eliminando"

    Dim eliminarIgual As Boolean
    eliminarIgual = False

    RsITOc.Seek "=", Numero.Text
    
    If RsITOc.NoMatch Then
    
        If MsgBox(Obj & " NO EXISTE, Elimina de todas maneras ?", vbYesNo) = vbYes Then
            eliminarIgual = False
            GoTo Eliminar
        Else
            Numero.SetFocus
        End If
    
    Else
    
Eliminar:

        Campos_Enabled False
        
        If Not eliminarIgual Then
            Doc_Leer
        End If
        
        If Not ITOp_Buscar Then
        
            If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo, "Atención") = vbYes Then
            
                Doc_Eliminar
    '            Doc_Actualizar Nv.Text, Dbm
                OTfdet_CantidadRecibida_Actualizar Nv.Text
                PlanoDetalle_Actualizar Dbm, Nv.Text, m_NvArea, RsITOd, "nv-plano-marca", "ito fab"
                
            End If
        
        End If
        
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
    End If
    
Case "Imprimiendo"
    
    RsITOc.Seek "=", Numero.Text
    If RsITOc.NoMatch Then
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
On Error Resume Next
Fecha.Text = Format(RsITOc!Fecha, Fecha_Format)
m_Nv = RsITOc!Nv
Nv.Text = m_Nv
Rut.Text = RsITOc![Rut contratista]

mNv = nvLeer(m_Nv)

If m_Nv <> 0 Then
    ComboNV.Text = Format(mNv.Numero, "0000") & " - " & mNv.Obra
End If

'rsitoc("Tipo OT")
Obs(0).Text = NoNulo(RsITOc![Observacion 1])
Obs(1).Text = NoNulo(RsITOc![Observacion 2])
Obs(2).Text = NoNulo(RsITOc![Observacion 3])
Obs(3).Text = NoNulo(RsITOc![Observacion 4])

'DETALLE
RsPd.Index = "NV-Plano-Marca"

RsITOd.Seek "=", Numero.Text, 1

If Not RsITOd.NoMatch Then

    Do While Not RsITOd.EOF
    
        If RsITOd!Numero = Numero.Text Then
        
            i = RsITOd!linea
            
            Detalle.TextMatrix(i, 1) = RsITOd!Plano
            Detalle.TextMatrix(i, 2) = RsITOd!Rev
            Detalle.TextMatrix(i, 3) = RsITOd!Marca
            
            RsPd.Seek "=", m_Nv, m_NvArea, RsITOd!Plano, RsITOd!Marca
            If Not RsPd.NoMatch Then
                Detalle.TextMatrix(i, 4) = RsPd!Descripcion
                Detalle.TextMatrix(i, 5) = RsPd![OT fab]
                Detalle.TextMatrix(i, 6) = RsPd![ITO fab] - RsITOd!Cantidad
            End If
            
            Detalle.TextMatrix(i, 7) = RsITOd![Numero OT]
            RsOTd.Seek "=", m_Nv, m_NvArea, RsITOd!Plano, RsITOd!Marca
            If Not RsOTd.NoMatch Then
               Do While Not RsOTd.EOF
                  If RsOTd!Numero = RsITOd![Numero OT] And RsOTd!Plano = RsITOd!Plano And RsOTd!Marca = RsITOd!Marca Then
                     Detalle.TextMatrix(i, 8) = RsOTd!Cantidad
                    Detalle.TextMatrix(i, 13) = RsOTd!Fecha ' fecha OT
                  End If
                RsOTd.MoveNext
               Loop
            End If
            
            Detalle.TextMatrix(i, 9) = RsITOd!Cantidad
            Detalle.TextMatrix(i, 10) = RsITOd![Peso Unitario]
            Detalle.TextMatrix(i, 11) = Format(Val(Detalle.TextMatrix(i, 9)) * Val(Detalle.TextMatrix(i, 10)), num_Format0)
            
            
        Else
            Exit Do
        End If
        RsITOd.MoveNext
    Loop
End If

RsPd.Index = "NV-Plano-Item"

Razon.Text = Contratista_Lee(SqlRsSc, Rut.Text)

Detalle.Row = 1 ' para q' actualice la primera fila del detalle
Actualiza

End Sub
Private Function Doc_Borrable() As Boolean
' incompleta
Doc_Borrable = True
'RsPd.Seek ">=", Plano_Numero, 0 ojo
If Not RsPd.NoMatch Then
    Do While Not RsPd.EOF
'        If RsPd!Plano <> Plano_Numero Then Exit Do
        
        'OT
        If RsPd![OT fab] <> 0 Then
            Doc_Borrable = False
            Exit Function
        End If
        
        RsPd.MoveNext
        
    Loop
End If
End Function
Private Function Doc_Validar() As Boolean
Dim porRecibir As Integer
Doc_Validar = False
If Rut.Text = "" Then
    MsgBox "DEBE ELEGIR CONTRATISTA"
    btnSearch.SetFocus
    Exit Function
End If

For i = 1 To n_filas

    ' plano
    If Trim(Detalle.TextMatrix(i, 1)) <> "" Then
    
        ' revision           2
        
        ' marca              3
        If Not CampoReq_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
        
        ' descripcion        4
        
        ' tot cant asignada  5
        ' tot cant recibida  6
        
        ' n ot               7
        ' cantidad asig ot   8
        
        ' cantidad a recibir 9
        If Not Numero_Valida(Detalle.TextMatrix(i, 9), i, 9) Then Exit Function
        
        ' [can asig ot]-[tot can recibida]>=[can a recibir]
'        porRecibir = Detalle.Textmatrix(i, 5)) - Detalle.Textmatrix(i, 6))
        porRecibir = Val(Detalle.TextMatrix(i, 8)) - Val(Detalle.TextMatrix(i, 12))
        If porRecibir < Detalle.TextMatrix(i, 9) Then
            MsgBox "Sólo quedan " & porRecibir & " por Recibir", , "ATENCIÓN"
            Detalle.Row = i
            Detalle.col = 9
            Detalle.SetFocus
            Exit Function
        End If
        
        ' peso unitario 10
        ' peso total    11
        
        ' fecha de OT 13
        If Detalle.TextMatrix(i, 13) <> "" Then
        If CDate(Fecha.Text) < CDate(Detalle.TextMatrix(i, 13)) Then
            MsgBox "Fecha de ITO actual es anterior a" & vbLf & "OT " & Detalle.TextMatrix(i, 7) & " que es del " & Detalle.TextMatrix(i, 13)
            Fecha.SetFocus
'            SendKeys "{Home}+{End}"
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
Private Sub Doc_Grabar(Nueva As Boolean)

MousePointer = vbHourglass

Dim m_Plano As String, m_Rev As String, m_Marca As String, m_cantidad As Integer, m_ot As Double
Dim m_PrecioITOc As Double, m_PrecioUnitario As Double, m_PesoUnitario As Double

'Ws.BeginTrans

' 19/11/12 comenté el "On Error GoTo Error" porque no esta grabando en Track
' On Error GoTo Error

' DETALLE DE ITO

Rut.Text = SqlRutPadL(Rut.Text)

m_PrecioITOc = 0

'If Not Nueva Then Doc_Detalle_Eliminar

If Nueva Then
    Numero.Text = Documento_Numero_Nuevo(RsITOc, "Numero")
Else
    Doc_Detalle_Eliminar
End If

j = 0
RsPd.Index = "NV-Plano-Marca"
For i = 1 To n_filas

    m_Plano = Trim(Detalle.TextMatrix(i, 1))
    
    If m_Plano <> "" Then
    
        m_Marca = Detalle.TextMatrix(i, 3)
        m_cantidad = Val(Detalle.TextMatrix(i, 9))
        m_ot = Val(Detalle.TextMatrix(i, 7))
        
        If m_cantidad <> 0 Then
        
            RsITOd.AddNew
            RsITOd!Numero = Numero.Text
            j = j + 1
            RsITOd!linea = j
            
            RsITOd!Fecha = Fecha.Text
            RsITOd!Nv = m_Nv
            RsITOd![Rut contratista] = Rut.Text
            
            RsITOd!Plano = m_Plano
            m_Rev = Detalle.TextMatrix(i, 2)
            RsNvPla.Seek "=", m_Nv, m_NvArea, m_Plano
            If Not RsNvPla.NoMatch Then
                m_Rev = RsNvPla![Rev]
            End If
            RsITOd!Rev = m_Rev
            
            RsITOd!Marca = m_Marca
            RsITOd!Cantidad = m_cantidad
            
'            If IsDate(Fecha.Text) Then
                RsITOd![Fecha Recepcion] = Fecha.Text
'            End If
            
            m_PesoUnitario = m_CDbl(Detalle.TextMatrix(i, 10))
            RsITOd![Peso Unitario] = m_PesoUnitario
            RsITOd![Numero OT] = m_ot
    
            RsITOd.Update
            
        If False Then
            RsPd.Seek "=", m_Nv, m_NvArea, m_Plano, m_Marca
            If RsPd.NoMatch Then
                ' no existe marca en el plano
            Else
                ' actualiza archivo detalle planos
                RsPd.Edit
                RsPd![ITO fab] = RsPd![ITO fab] + m_cantidad
                RsPd.Update
            End If
        End If
        
'            If False Then
            If True Then
            
               ' actualiza OT
               ' (cantidad recibida)
               m_PrecioUnitario = 0
               RsOTd.Seek "=", m_Nv, m_NvArea, m_Plano, m_Marca 'nº ot
               If Not RsOTd.NoMatch Then
               
                   Do While Not RsOTd.EOF
                   
                       If m_Nv <> RsOTd!Nv Or m_Plano <> RsOTd!Plano Or m_Marca <> RsOTd!Marca Then Exit Do
                       
                       If RsOTd!Numero = m_ot Then
                       
                           m_PrecioUnitario = RsOTd![Precio Unitario]
                           
'                           RsOTd.Edit
'                           RsOTd![Cantidad Recibida] = RsOTd![Cantidad Recibida] + m_cantidad
'                           RsOTd![Precio Unitario] = m_PrecioUnitario
'                           RsOTd.Update
                           
                       End If
                       
                       RsOTd.MoveNext
                       
                   Loop
                   
               End If
               
            End If
            
            m_PrecioITOc = m_PrecioITOc + m_PrecioUnitario * m_cantidad * m_PesoUnitario
            
        End If
        
    End If
    
Next

' CABECERA DE ITO
' cabacera va despues de detalle para saber el precio total del la ito
With RsITOc
If Nueva Then
    .AddNew
    !Numero = Numero.Text
Else
    .Edit
End If
!Fecha = Fecha.Text
!Nv = m_Nv
![Rut contratista] = Rut.Text
![Peso Total] = TotalKilos.Caption
![Precio Total] = m_PrecioITOc ' agregado el 07/03/06
![Observacion 1] = Obs(0).Text
![Observacion 2] = Obs(1).Text
![Observacion 3] = Obs(2).Text
![Observacion 4] = Obs(3).Text
.Update
End With

RsPd.Index = "NV-Plano-Item"

PlanoDetalle_Actualizar Dbm, Nv.Text, m_NvArea, RsITOd, "nv-plano-marca", "ito fab"

OTfdet_CantidadRecibida_Actualizar Nv.Text

'Ws.CommitTrans

Select Case Accion
Case "Agregando"
    Track_Registrar "ITOf", Numero.Text, "AGR"
Case "Modificando"
    Track_Registrar "ITOf", Numero.Text, "MOD"
End Select

GoTo Fin
Error:
'Ws.Rollback

Fin:

MousePointer = vbDefault

End Sub
Private Function ITOp_Buscar() As Boolean

' R : granallado, Erwin
' T : produccion pintura , Erwin
' P : pintura

Dim fi As Integer, msg As String
Dim cantITOgr As Integer ' ito granallado
Dim cantITOpp As Integer ' ito granallado
Dim cantITOp As Integer ' ito granallado

cantITOgr = 0
cantITOpp = 0
cantITOp = 0

' busca si hay cantidades recibidas en ITOpg

ITOp_Buscar = False

Set RsITOpd = Dbm.OpenRecordset("ITO pg detalle")
With RsITOpd
.Index = "nv-plano-marca"

RsITOd.Seek "<=", Numero.Text, 0
Do While Not RsITOd.EOF
    If Numero.Text <> RsITOd!Numero Then Exit Do
    .Seek "=", m_Nv, 0, RsITOd!Plano, RsITOd!Marca
    If Not .NoMatch Then
        Do While Not .EOF
            If m_Nv <> !Nv Or RsITOd!Plano <> !Plano Or RsITOd!Marca <> !Marca Then
                Exit Do
            End If
            Select Case RsITOpd!Tipo
            Case "R"
                cantITOgr = cantITOgr + !Cantidad
            Case "T"
                cantITOpp = cantITOpp + !Cantidad
            Case "P"
                cantITOp = cantITOp + !Cantidad
            End Select
            .MoveNext
        Loop
    End If
'    If Cantidad < RsITOd!Cantidad Then
'    If Cantidad <= RsITOd!Cantidad Then
    msg = "No se puede eliminar ITO Fabricación" & vbLf
    If cantITOgr > 0 Then ' ya se ha recibido en ITOpg
        'MsgBox "No se puede eliminar ITO Fabricación" & vbLf & "porque ya se hizo ITO " & a_Nv(ComboNV.ListIndex, 2) & vbLf & "Plano: " & RsITOd!Plano & ", Marca: " & RsITOd!Marca, , "Advertencia"
        msg = msg & "porque ya se hizo ITO Granallado" & vbLf
    End If
    If cantITOpp > 0 Then ' ya se ha recibido en ITOpg
        msg = msg & "porque ya se hizo ITO Produccion Pintura" & vbLf
    End If
    If cantITOp > 0 Then ' ya se ha recibido en ITOpg
        msg = msg & "porque ya se hizo ITO Pintura" & vbLf
    End If
    If cantITOgr + cantITOpp + cantITOp > 0 Then
        msg = msg & "Plano: " & RsITOd!Plano & ", Marca: " & RsITOd!Marca
        MsgBox msg, , "Advertencia"
        ITOp_Buscar = True
    End If
    RsITOd.MoveNext
Loop
.Close
End With

End Function
Private Sub Doc_Eliminar()

'Ws.BeginTrans

'On Error GoTo Error

' borra CABECERA DE ITO
RsITOc.Seek "=", Numero.Text
If Not RsITOc.NoMatch Then

    RsITOc.Delete
   
End If

Doc_Detalle_Eliminar

OTfdet_CantidadRecibida_Actualizar Nv.Text

GoTo Fin
Error:
On Error GoTo 0
'Ws.Rollback

Fin:
'Ws.CommitTrans

End Sub
Private Sub Doc_Detalle_Eliminar()

' DETALLE DE ITO
' al anular detalle ITO debe actualizar detalle plano, y OTs

'RsPd.Index = "NV-Plano-Marca"

Dbm.Execute "DELETE FROM [ito fab detalle] WHERE numero=" & Numero.Text

If False Then
RsITOd.Seek "<=", Numero.Text, 0

If Not RsITOd.NoMatch Then

    Do While Not RsITOd.EOF
    
        If RsITOd!Numero <> Numero.Text Then Exit Do
        
    If False Then
        ' actualiza plano detalle
        RsPd.Seek "=", m_Nv, m_NvArea, RsITOd!Plano, RsITOd!Marca
        If Not RsPd.NoMatch Then
            RsPd.Edit
            RsPd![ITO fab] = RsPd![ITO fab] - RsITOd!Cantidad
            RsPd.Update
        End If
    End If


If False Then
        ' actualiza OT detalle
        RsOTd.Seek "=", m_Nv, m_NvArea, RsITOd!Plano, RsITOd!Marca
        If Not RsOTd.NoMatch Then
            RsOTd.Edit
            RsOTd![Cantidad Recibida] = RsOTd![Cantidad Recibida] - RsITOd!Cantidad
            RsOTd.Update
        End If
End If
        ' borra detalle
        RsITOd.Delete
    
        RsITOd.MoveNext
        
    Loop
End If
End If
'RsPd.Index = "NV-Plano-Item"

End Sub
Private Sub Campos_Limpiar()
Numero.Text = ""
'Fecha.Text = fecha_vacia
Fecha.Text = Format(Now, Fecha_Format)
Nv.Text = ""
ComboNV.Text = " "
Rut.Text = ""
Razon.Text = ""
'Direccion.Text = ""
'Comuna.Text = ""
Detalle_Limpiar
Obs(0).Text = ""
Obs(1).Text = ""
Obs(2).Text = ""
Obs(3).Text = ""
TotalKilos.Caption = ""
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
Private Sub Obs_GotFocus(Index As Integer)
ComboPlano.visible = False
ComboMarca.visible = False
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

ComboPlano.visible = False
ComboMarca.visible = False

cambia_titulo = True
'Accion = ""
Select Case Button.Index
Case 1 ' agregar

    Accion = "Agregando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
'    Campos_Enabled False
    
'    Numero.Text = Documento_Numero_Nuevo(RsITOc, "Numero")

'    Numero.Enabled = True
'    Numero.SetFocus
    
    Campos_Enabled True
    Numero.Enabled = False
    Fecha.SetFocus
    btnGrabar.Enabled = True
    btnSearch.visible = True
    
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
    
            n_Copias = 1
            PrinterNCopias.Numero_Copias = n_Copias
            PrinterNCopias.Show 1
            n_Copias = PrinterNCopias.Numero_Copias
            
'            If MsgBox("¿ IMPRIME " & Obj & " ?", vbYesNo) = vbYes Then
            If n_Copias > 0 Then
            
                Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Docs")
                Doc_Imprimir '
'                OT_Imprimir n_Copias
                Impresora_Predeterminada "default"
                
            End If
    
'        If MsgBox("¿ Imprime ?", vbYesNo) = vbYes Then
'            ITO_Imprimir
'        End If
        
        
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
        
    End If
Case 5 ' separador
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
    
        If MsgBox("¿ GRABAR " & Obj & " ?", vbYesNo) = vbYes Then
        
            If Accion = "Agregando" Then
                Doc_Grabar True
            Else
                Doc_Grabar False
            End If
'            Doc_Actualizar Nv.Text, Dbm
            OTfdet_CantidadRecibida_Actualizar Nv.Text
            
            n_Copias = 1
            PrinterNCopias.Numero_Copias = n_Copias
            PrinterNCopias.Show 1
            n_Copias = PrinterNCopias.Numero_Copias
            
'            If MsgBox("¿ IMPRIME " & Obj & " ?", vbYesNo) = vbYes Then
            If n_Copias > 0 Then
            
                Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Docs")
                Doc_Imprimir '
'                OT_Imprimir n_Copias
                Impresora_Predeterminada "default"
                
            End If
            
            Botones_Enabled 0, 0, 0, 0, 1, 0
            Campos_Limpiar
            
            If Accion = "Agregando" Then
                Campos_Enabled True
                Numero.Enabled = False
                Fecha.SetFocus
                btnGrabar.Enabled = True
                btnSearch.visible = True
            Else
                Campos_Enabled False
                Numero.Enabled = True
                Numero.SetFocus
            End If
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
Private Sub Privilegios()
If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 0, 0
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

'Rut.Enabled = Si
Nv.Enabled = Si
ComboNV.Enabled = Si
btnArchivoElegir.Enabled = Si

Detalle.Enabled = Si
Obs(0).Enabled = Si
Obs(1).Enabled = Si
Obs(2).Enabled = Si
Obs(3).Enabled = Si

End Sub
Private Sub btnSearch_Click()

Dim arreglo(1) As String
arreglo(1) = "razon_social"

ComboPlano.visible = False
ComboMarca.visible = False

sql_Search.Muestra "personas", "RUT", arreglo(), Obj, Objs, "contratista='S'"
Rut.Text = sql_Search.Codigo

Rut.Text = SqlRutPadL(Rut.Text)

Razon.Text = sql_Search.Descripcion

If ComboNV.Text <> "" Then Detalle.Enabled = True

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
ComboPlano.visible = False
ComboMarca.visible = False
Select Case Detalle.col
    Case 1 ' plano
        If Detalle <> "" Then ComboPlano.Text = Detalle
        ComboPlano.Top = Detalle.CellTop + Detalle.Top
        ComboPlano.Left = Detalle.CellLeft + Detalle.Left
        ComboPlano.Width = Int(Detalle.CellWidth * 1.5)
        ComboPlano.visible = True
        ComboMarca.visible = False
    Case 3 ' marca
        On Error Resume Next
        ComboMarca_Poblar Detalle.TextMatrix(Detalle.Row, 1)
        If Detalle <> "" Then ComboMarca.Text = Detalle
        ComboMarca.Top = Detalle.CellTop + Detalle.Top
        ComboMarca.Left = Detalle.CellLeft + Detalle.Left
        ComboMarca.Width = Int(Detalle.CellWidth * 1.5)
        ComboPlano.visible = False
        ComboMarca.visible = True
        On Error GoTo 0
    Case Else
End Select
End Sub
Private Sub Detalle_DblClick()
If Accion = "Imprimiendo" Then Exit Sub
MSFlexGridEdit Detalle, txtEditITO, 32
End Sub
Private Sub Detalle_GotFocus()
If txtEditITO.visible Then
    Detalle = txtEditITO
    txtEditITO.visible = False
End If
End Sub
Private Sub Detalle_LeaveCell()
If txtEditITO.visible Then
    Detalle = txtEditITO
    txtEditITO.visible = False
End If
End Sub
Private Sub Detalle_KeyPress(KeyAscii As Integer)
If Accion = "Imprimiendo" Then Exit Sub
MSFlexGridEdit Detalle, txtEditITO, KeyAscii
End Sub
Private Sub txtEditITO_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCodeP Detalle, txtEditITO, KeyCode, Shift
End Sub
Sub EditKeyCodeP(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
' rutina que se ejecuta con los keydown de Edt
Dim m_col As Integer, dif As Integer
m_col = MSFlexGrid.col
dif = Val(Detalle.TextMatrix(MSFlexGrid.Row, 8)) - Val(Detalle.TextMatrix(MSFlexGrid.Row, 12))
Select Case KeyCode
Case vbKeyEscape ' Esc
    Edt.visible = False
    MSFlexGrid.SetFocus
Case vbKeyReturn ' Enter
    Select Case m_col
    Case 9 ' Cantidad a Recibir
        If Recibida_Validar(MSFlexGrid.col, dif, Edt) Then
            MSFlexGrid.SetFocus
            DoEvents
            Linea_Actualiza
        End If
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
        Linea_Actualiza
    End Select
Case vbKeyUp ' Flecha Arriba
    Select Case m_col
    Case 9 ' Cantidad a Recibir
        If Recibida_Validar(MSFlexGrid.col, dif, Edt) Then
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
'        Linea_Actualiza
        If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
            MSFlexGrid.Row = MSFlexGrid.Row - 1
        End If
    End Select
Case vbKeyDown ' Flecha Abajo
    Select Case m_col
    Case 9 ' Cantidad a Recibir
        If Recibida_Validar(MSFlexGrid.col, dif, Edt) Then
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
'        Linea_Actualiza
        If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
            MSFlexGrid.Row = MSFlexGrid.Row + 1
        End If
    End Select
End Select
End Sub
Private Function Recibida_Validar(Colu As Integer, porRecibir As Integer, Edt As Control) As Boolean
' verifica que Ctotal-CAsignada >= CAAsignar
Recibida_Validar = True
If Colu <> 9 Then Exit Function
If porRecibir < Val(Edt) Then
    MsgBox "Sólo quedan " & porRecibir & " por Recibir", , "ATENCIÓN"
    Recibida_Validar = False
End If
End Function
Private Sub txtEditITO_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)
Select Case MSFlexGrid.col
Case 1, 3
'    After_Detalle_Click
Case 2, 4, 5, 6, 7, 8, 10, 11
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
If KeyCode = vbKeyF2 Then
    MSFlexGridEdit Detalle, txtEditITO, 32
End If
End Sub
'Private Sub Detalle_RowColChange()
'MIA
'Posicion = "Lín " & Detalle.Row & ", Col " & Detalle.col
'End Sub

'Private Sub Cursor_Mueve(MSFlexGrid As Control)
''MIA
'Select Case MSFlexGrid.col
'Case 6
'    If MSFlexGrid.Row + 1 < MSFlexGrid.Rows Then
'        MSFlexGrid.col = 1
'        MSFlexGrid.Row = MSFlexGrid.Row + 1
'    End If
'Case Else
'    MSFlexGrid.col = MSFlexGrid.col + 1
'End Select
'End Sub
Private Sub Linea_Actualiza()
' actualiza solo linea, y totales generales
Dim fi As Integer, co As Integer
Dim n9 As Double, n10 As Double

fi = Detalle.Row
co = Detalle.col

n9 = m_CDbl(Detalle.TextMatrix(fi, 9))
n10 = m_CDbl(Detalle.TextMatrix(fi, 10))

' peso total
Detalle.TextMatrix(fi, 11) = Format(n9 * n10, num_Formato)

Totales_Actualiza

End Sub
Private Sub Actualiza()
' actualiza todo el detalle y totales generales
Dim fi As Integer
For fi = 1 To n_filas
    If Detalle.TextMatrix(fi, 1) <> "" Then
        ' peso total
        Detalle.TextMatrix(fi, 11) = Format(m_CDbl(Detalle.TextMatrix(fi, 9)) * m_CDbl(Detalle.TextMatrix(fi, 10)), num_Formato)
    End If
Next

Totales_Actualiza

End Sub
Private Sub Totales_Actualiza()
Dim Tot_Kilos As Double, Tot_Precio As Double
Tot_Kilos = 0
For i = 1 To n_filas
    Tot_Kilos = Tot_Kilos + m_CDbl(Detalle.TextMatrix(i, 11))
Next

TotalKilos.Caption = Format(Tot_Kilos, num_Formato)

End Sub
'
' FIN RUTINAS PARA FLEXGRID
'
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
Private Sub Doc_Imprimir()
' imprime ITOf
MousePointer = vbHourglass
Dim can_valor As String, can_col As Integer, k As Integer
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

For k = 1 To n_Copias

prt.Font.Size = 15
prt.Print Tab(tab0 + 14); "VALE ITO FABRICACIÓN Nº" & Format(Numero.Text, "000")
prt.Font.Size = 10
prt.Print ""

' cabecera
prt.Font.Bold = True
prt.Print Tab(tab0); Empresa.Razon;
prt.Font.Bold = False
prt.Print Tab(tab0 + tab40); "FECHA     : " & Fecha.Text

prt.Print Tab(tab0); "GIRO: " & Empresa.Giro;
prt.Print Tab(tab0 + tab40); "SEÑOR(ES) : " & Razon,

prt.Print Tab(tab0); Empresa.Direccion;
prt.Print Tab(tab0 + tab40); "RUT       : " & Rut

prt.Print Tab(tab0); "Teléfono: " & Empresa.Telefono1 & " - " & Empresa.Comuna;
'prt.Print Tab(tab0 + tab40); "DIRECCIÓN : " & Direccion,

'prt.Print Tab(tab0 + tab40); "COMUNA    : " & Comuna

prt.Print Tab(tab0 + tab40); "OBRA      : ";
prt.Font.Bold = True
prt.Print Format(Mid(ComboNV.Text, 8), ">")
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
prt.Print Format(TotalKilos, "#,###,###.0")
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

If k < n_Copias Then
    prt.NewPage
Else
    prt.EndDoc
End If

Next

Impresora_Predeterminada "default"

MousePointer = vbDefault

End Sub
Private Sub OTfdet_CantidadRecibida_Actualizar(Nv As Double)

' actualiza cantidad recibida en OTf Det

'Dim m_Cr As Integer
Dim m_Contratista As String

'Dim RsOTdet As Recordset, RsITOdet As Recordset
'Set Dbm = OpenDatabase(movs_file)
'Set RsOTdet = Dbm.OpenRecordset("ot fab detalle")
'RsOTdet.Index = "nv-plano-marca"
'Set RsITOdet = Dbm.OpenRecordset("ito fab detalle")
RsITOd.Index = "nv-plano-marca"

Dbm.Execute "UPDATE [ot fab detalle] SET [cantidad recibida]=0 WHERE nv=" & Nv

With RsOTd

.Seek ">=", Nv
Do While Not .EOF

    If Nv <> !Nv Then Exit Do
            
    'If !marca = "F24D" Then
    ''If !Numero = 13519 Then
    'MsgBox ""
    'End If
       
     m_Contratista = ![Rut contratista]
    '      m_Cr = 0
       ' busca itof
        RsITOd.Seek "=", Nv, m_NvArea, !Plano, !Marca
       
            If Not RsITOd.NoMatch Then
       
                Do While Not RsITOd.EOF
          
                    If Nv <> RsITOd!Nv Or !Plano <> RsITOd!Plano Or !Marca <> RsITOd!Marca Then Exit Do
             
                    If m_Contratista = RsITOd![Rut contratista] Then
             
                        If RsITOd![Numero OT] = !Numero Then
             
                        .Edit
                        ![Cantidad Recibida] = ![Cantidad Recibida] + RsITOd!Cantidad
                        .Update
                 
                        End If
                 
                    End If
             
                    RsITOd.MoveNext

             Loop

        End If

    .MoveNext

Loop

End With

RsITOd.Index = "numero-linea"

End Sub
Private Function ContratistaBuscarXId(ID As String) As String
' entrega el rut del contratista
Dim sql As String
sql = "SELECT * FROM personas WHERE dato1='" & ID & "'"

With SqlRsSc
.Open sql, CnxSqlServer_scp0
If .EOF Then
    ContratistaBuscarXId = ""
    Razon.Text = ""
Else
    ContratistaBuscarXId = !Rut
    Razon.Text = !razon_social
End If
.Close
End With

End Function
