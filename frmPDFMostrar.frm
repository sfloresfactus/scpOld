VERSION 5.00
Begin VB.Form frmPDFMostrar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Muestra PDF"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnVer 
      Caption         =   "Ver"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmPDFMostrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Path As String, m_Arch As String
Private m_Numero As String, m_Tipo As String, n_TotalPDFs As Integer, i As Integer, m_IndicePDFActual As Integer

' fotos subidas por el emisor
Private a_EPdfs(99) As String ' arreglo con nombres de archivos de fotos subidos, solo formato jpg
Public Property Let Numero(ByVal nuevo_valor As String)
m_Numero = nuevo_valor
End Property
Public Property Let Tipo(ByVal nuevo_valor As String)
m_Tipo = nuevo_valor
End Property
Private Sub Form_Load()

btnVer(0).visible = False

m_Path = Drive_Server & Path_Mdb & "nc_files\"

Archivos_Subidos_Buscar m_Tipo

End Sub
Private Sub Archivos_Subidos_Buscar(Tipo As String)
' busca nombres de archivos subidos por el receptor o el encargado
Dim ArchivoNombre As String
n_TotalPDFs = 0
'lblAdjuntos.Caption = ""

If Tipo = "E" Then
    ArchivoNombre = Dir(Drive_Server & Path_Mdb & "nc_files\" & "nc_" & m_Numero & "_*.PDF", vbArchive)
End If
If Tipo = "R" Then
    ArchivoNombre = Dir(Drive_Server & Path_Mdb & "nc_files\" & "ncr_" & m_Numero & "_*.PDF", vbArchive)
End If

'D:\acr3006-dualpro\scp\mdb\nc_files

btnVer(0).visible = False
If ArchivoNombre <> "" Then
    Do
        n_TotalPDFs = n_TotalPDFs + 1
        
        If n_TotalPDFs = 1 Then
            btnVer(0).visible = True
        Else
            Load btnVer(n_TotalPDFs - 1)
            btnVer(n_TotalPDFs - 1).Top = btnVer(0).Top - 450 + n_TotalPDFs * 450
        End If
        
        btnVer(n_TotalPDFs - 1).visible = True
        btnVer(n_TotalPDFs - 1).Caption = ArchivoNombre
        
        Me.Height = btnVer(0).Top + 900 + n_TotalPDFs * 450
        
        a_EPdfs(n_TotalPDFs) = ArchivoNombre
'        lblAdjuntos.Caption = lblAdjuntos.Caption & " " & ArchivoNombre
        ArchivoNombre = Dir()
    
    Loop Until ArchivoNombre = ""

End If

End Sub
Private Sub btnVer_Click(Index As Integer)
Dim mvarURL
Dim Archivo As String, intranet

' en el servidor acr3006-dualpro
' entrar al ISS, existe una carpeta virtual (y oculta) llamada "docs", que tiene la siguiente ruta:
' h:\scp\mdd\nv_files
' aqui estan los documentos dque los usuarios suben como evidencias de las NC

intranet = ReadIniValue(Path_Local & "scp.ini", "Path", "intranet_server")

Archivo = btnVer(Index).Caption
'mvarURL = "http://acr3006-dualpro/nc_files/" & archivo
'mvarURL = "http://acr3006-dualpro/docs/" & Archivo
mvarURL = intranet & "docs/" & Archivo
GoURL (mvarURL)
End Sub
