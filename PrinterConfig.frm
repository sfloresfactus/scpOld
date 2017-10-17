VERSION 5.00
Begin VB.Form PrinterConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración Impresora"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame_Doc 
      ForeColor       =   &H00FF0000&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.ComboBox CBprinters 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1920
         Width           =   5295
      End
      Begin VB.ComboBox CBprinters 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   5295
      End
      Begin VB.ComboBox CBprinters 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label lbl 
         Caption         =   "Impresora &ETIQUETAS"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label lbl 
         Caption         =   "Impresora &DOCUMENTOS"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label lbl 
         Caption         =   "Impresora &REPORTES"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "PrinterConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private i As Integer, nPrinter As Integer
'Private Dbc As Database, RsCfgP As Recordset
'Private Path_Local As String
Private a_font(1, 99) As String
Private Sub Form_Load()
Dim impre As String
'Set Dbc = OpenDatabase(myPC_file)
'Set RsCfgP = Dbc.OpenRecordset("Configuración Impresoras")

' lista impresoras
For i = 0 To Printers.Count - 1
    impre = Printers(i).DeviceName
    CBprinters(0).AddItem impre ' rpt
    CBprinters(1).AddItem impre ' docs
    CBprinters(2).AddItem impre ' etiq
Next

On Error Resume Next

'impre = Trim(RsCfgP![Documentos Printer])
impre = ReadIniValue(Path_Local & "scp.ini", "Printer", "Rpt")
If Len(impre) > 0 Then CBprinters(0).Text = impre

'impre = Trim(RsCfgP![Documentos Font])
'If Len(impre) > 0 Then CBfonts(0).Text = impre

'impre = Trim(RsCfgP![Informes Printer])
impre = ReadIniValue(Path_Local & "scp.ini", "Printer", "Docs")
If Len(impre) > 0 Then CBprinters(1).Text = impre

'impre = Trim(RsCfgP![Informes Font])
'If Len(impre) > 0 Then CBfonts(1).Text = impre

impre = ReadIniValue(Path_Local & "scp.ini", "Printer", "Etiquetas")
If Len(impre) > 0 Then CBprinters(2).Text = impre

End Sub
Private Sub CBprinters_Click(Index As Integer)
'Font_Llena Index
End Sub
Private Sub btnAceptar_Click()

'RsCfgP.Edit
'RsCfgP![Documentos Printer] = CBprinters(0).Text
'RsCfgP![Documentos Font] = CBfonts(0).Text
'RsCfgP![Informes Printer] = CBprinters(1).Text
'RsCfgP![Informes Font] = CBfonts(1).Text
'RsCfgP.Update
'Dbc.Close

WriteIniValue Path_Local & "scp.ini", "Printer", "Rpt", CBprinters(0).Text
WriteIniValue Path_Local & "scp.ini", "Printer", "Docs", CBprinters(1).Text
WriteIniValue Path_Local & "scp.ini", "Printer", "Etiquetas", CBprinters(2).Text

Unload Me

End Sub
Private Sub btnCancelar_Click()
Unload Me
End Sub
