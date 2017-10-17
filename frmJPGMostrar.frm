VERSION 5.00
Begin VB.Form frmJPGMostrar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Muestra Foto"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnSiguiente 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton btnAnterior 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label indicador 
      Alignment       =   2  'Center
      Caption         =   "0 / 0"
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
      Left            =   3000
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Image Imagen 
      Height          =   5235
      Left            =   240
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6690
   End
End
Attribute VB_Name = "frmJPGMostrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Path As String, m_Arch As String
Private m_Numero As String, m_Tipo As String, n_TotalFotos As Integer, i As Integer, m_IndiceFotoActual As Integer

' fotos subidas por el emisor
Private a_EFotos(99) As String ' arreglo con nombres de archivos de fotos subidos, solo formato jpg
Public Property Let Numero(ByVal nuevo_valor As String)
m_Numero = nuevo_valor
End Property
Public Property Let Tipo(ByVal nuevo_valor As String)
m_Tipo = nuevo_valor
End Property
Private Sub Form_Load()

m_Path = Drive_Server & Path_Mdb & "nc_files\"

If m_Tipo = "E" Then
    Emisor_Cargar
End If
If m_Tipo = "R" Then
    Receptor_Cargar
End If

End Sub
Private Sub Emisor_Cargar()

' carga fotos del emisor
Me.Caption = "Imagen NC Nº " & m_Numero

n_TotalFotos = 0
Archivos_Subidos_BuscarE

If n_TotalFotos > 0 Then

    m_IndiceFotoActual = 1
    
    Muestra
 
Else

    indicador.Caption = "NO FOTO"
    btnAnterior.Enabled = False
    BtnSiguiente.Enabled = False

End If

If False Then
    ' reubica controles segun tamaño de pantalla
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    
    'Image.Height = Me.Height - 1000
    'Image.Width = Me.Width - 1000
    
    'Picture.Move 0, 0, 100, 100
    
    'Picture.Width = Me.Width - 1000
    
    indicador.Top = Me.Height - 800
    indicador.Left = (Me.Width / 2) - (indicador.Width / 2)
    
    btnAnterior.Top = indicador.Top
    btnAnterior.Left = indicador.Left - 700
    
    BtnSiguiente.Top = indicador.Top
    BtnSiguiente.Left = indicador.Left + 1500
End If

End Sub
Private Sub Receptor_Cargar()

' carga fotos del receptor
Me.Caption = "Imagen NC Nº " & m_Numero

n_TotalFotos = 0
Archivos_Subidos_BuscarR

If n_TotalFotos > 0 Then

    m_IndiceFotoActual = 1
    
    Muestra
 
Else

    indicador.Caption = "NO FOTO"
    btnAnterior.Enabled = False
    BtnSiguiente.Enabled = False

End If

End Sub
Private Function Botones_Enabled() As Boolean
'
End Function
Private Function Muestra()

MousePointer = vbHourglass

indicador.Caption = m_IndiceFotoActual & " / " & n_TotalFotos

' des/habilita botones
Select Case n_TotalFotos
Case 1
    btnAnterior.Enabled = False
    BtnSiguiente.Enabled = False
Case Else

    Select Case m_IndiceFotoActual
    Case 1
        btnAnterior.Enabled = False
        BtnSiguiente.Enabled = True
    Case n_TotalFotos
        btnAnterior.Enabled = True
        BtnSiguiente.Enabled = False
    Case Else
        btnAnterior.Enabled = True
        BtnSiguiente.Enabled = True
    End Select
    
End Select

m_Arch = a_EFotos(m_IndiceFotoActual)
If Archivo_Existe(m_Path, m_Arch) Then
   Imagen = LoadPicture(m_Path & m_Arch)
Else
    Imagen.Picture = Nothing
End If

MousePointer = vbDefault

End Function
Private Sub BtnAnterior_Click()
' muestra foto anterior
m_IndiceFotoActual = m_IndiceFotoActual - 1
If m_IndiceFotoActual = 0 Then
    m_IndiceFotoActual = 1
End If
Muestra
End Sub
Private Sub BtnSiguiente_Click()
' muestra siguiente foto
m_IndiceFotoActual = m_IndiceFotoActual + 1
If m_IndiceFotoActual > n_TotalFotos Then
    m_IndiceFotoActual = n_TotalFotos
End If
Muestra
End Sub
Private Sub Archivos_Subidos_BuscarE()
' busca nombres de archivos subidos por el emisor
Dim ArchivoNombre As String
n_TotalFotos = 0
'lblAdjuntos.Caption = ""
ArchivoNombre = Dir(Drive_Server & Path_Mdb & "nc_files\" & "nc_" & m_Numero & "_*.JPG", vbArchive)
'D:\acr3006-dualpro\scp\mdb\nc_files
If ArchivoNombre <> "" Then
    Do
        n_TotalFotos = n_TotalFotos + 1
        a_EFotos(n_TotalFotos) = ArchivoNombre
'        lblAdjuntos.Caption = lblAdjuntos.Caption & " " & ArchivoNombre
        ArchivoNombre = Dir()
    Loop Until ArchivoNombre = ""
End If

End Sub
Private Sub Archivos_Subidos_BuscarR()
' busca nombres de archivos subidos por el receptor o el encargado
Dim ArchivoNombre As String
n_TotalFotos = 0
'lblAdjuntos.Caption = ""
ArchivoNombre = Dir(Drive_Server & Path_Mdb & "nc_files\" & "ncr_" & m_Numero & "_*.JPG", vbArchive)
'D:\acr3006-dualpro\scp\mdb\nc_files
If ArchivoNombre <> "" Then
    Do
        n_TotalFotos = n_TotalFotos + 1
        a_EFotos(n_TotalFotos) = ArchivoNombre
'        lblAdjuntos.Caption = lblAdjuntos.Caption & " " & ArchivoNombre
        ArchivoNombre = Dir()
    Loop Until ArchivoNombre = ""
End If

End Sub
