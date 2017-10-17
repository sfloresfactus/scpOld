VERSION 5.00
Begin VB.Form qrCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QR Code"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton CmdFile 
      Caption         =   "Desde Archivo"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox TxtFile 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   4335
   End
   Begin VB.CommandButton CmdCodificTxt 
      Caption         =   "Codificar Texto"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton CmdDecodificar 
      Caption         =   "Decodificar Imagen"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton CmdURL 
      Caption         =   "Desde URL"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox TxtUrl 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "http://qrcode.es/wp-content/uploads/2007/09/28w.jpg"
      Top             =   3960
      Width           =   4335
   End
   Begin VB.CommandButton CmdWebCam 
      Caption         =   "WebCam"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   3960
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3810
      Left            =   6000
      Picture         =   "qrCode.frx":0000
      ScaleHeight     =   250
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   250
      TabIndex        =   1
      Top             =   120
      Width           =   3810
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   120
      MaxLength       =   700
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "qrCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------
'Autor:  Leandro Ascierto
'Web:    leandroascierto.com
'Date:   09/09/2011
'----------------------------------
Dim cQrCode As ClsQrCode
Dim prt As Printer
Dim AjusteX As Double, AjusteY As Double
Private Sub btnImprimir_Click()
Dim wid As Integer, hgt As Integer
Dim tab0 As Double, tab1 As Double

tab0 = 0.5
tab1 = 4.7

Set prt = Printer

Prt_Ini

'////////////////////////
'
' LOGO DELGADO
'
prt.Font.Name = "delgado"
prt.Font.Size = 16
SetpYX 0.1, tab0
prt.Print "Delgado"
'////////////////////////


'//////////////////////////////////////////////////
'
' CODIGO QR
'
' Print the picture.
'Printer.PaintPicture Picture1.Picture, 1440, 1440
'Printer.PaintPicture Picture1.Picture, 1, 1 ' ok
'Printer.PaintPicture Picture1.Picture, 1, 1, 100, 200 ' muy grande
Printer.PaintPicture Picture1.Picture, tab0, 1, 3.4, 3.4 ' en cms
'////////////////////////////////////////////////////
'
' TEXTOS
'
prt.Font.Name = "Arial"
prt.Font.Size = 15

SetpYX 0.6, tab1
prt.Print "FLSMIDTH"
SetpYX 1.2, tab1
prt.Print "CASERONES"
SetpYX 1.8, tab1
prt.Print "NV 2999"
SetpYX 2.4, tab1
prt.Print "2201-CV-001"
SetpYX 3, tab1
prt.Print "BA1 BARANDA"
SetpYX 3.6, tab1
prt.Print "140,7 Kgs"


' Get the picture's dimensions in the printer's scale
' mode.
'wid = ScaleX(Picture1.ScaleWidth, Picture1.ScaleMode, Printer.ScaleMode)
'hgt = ScaleY(Picture1.ScaleHeight, Picture1.ScaleMode, Printer.ScaleMode)

' Draw the box.
'Printer.Line (1440, 1440)-Step(wid, hgt), , B

prt.EndDoc

End Sub

Private Sub CmdCodificTxt_Click()
    Picture1.Picture = cQrCode.GetPictureQrCode(Text1.Text, Picture1.ScaleWidth, Picture1.ScaleHeight)
    If Picture1.Picture Is Nothing Then MsgBox "Error!"
    'Picture1.Picture = cQrCode.GetPictureQrCode(Text1.Text, 200, 200, "UTF-8", "L", vbRed, vbBlue, 3)
End Sub

Private Sub CmdDecodificar_Click()
    Dim strDecode As String
    If cQrCode.DecodeFromPicture(Picture1.Picture, strDecode) Then
        MsgBox strDecode
    Else
        MsgBox "Error!"
    End If
End Sub

Private Sub CmdFile_Click()
    Dim strDecode As String
    If cQrCode.DecodeFromFile(TxtFile.Text, strDecode) Then
        MsgBox strDecode
    Else
        MsgBox "Error!"
    End If
End Sub

Private Sub CmdURL_Click()
    Dim strDecode As String
    If cQrCode.DecodeFromUrl(TxtUrl.Text, strDecode) Then
        MsgBox strDecode
    Else
        MsgBox "Error!"
    End If
End Sub

Private Sub CmdWebCam_Click()
'    FrmWebCam.Show , Me
End Sub

Private Sub Form_Load()
    Set cQrCode = New ClsQrCode
    TxtFile.Text = App.Path & "\casquillo_gorra.jpg"
    Text1.Text = "FLSMIDTH/CASERONES/2201-CV-001/140,7/2999"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cQrCode = Nothing
End Sub
Private Sub Prt_Ini()
Set prt = Printer
Printer.ScaleMode = 1 ' twips : 576 twips x cm
Printer.ScaleMode = 7 ' centimetros
End Sub
Private Sub SetpYX(y As Double, x As Double)
Printer.CurrentY = AjusteY + y
Printer.CurrentX = AjusteX + x
End Sub
