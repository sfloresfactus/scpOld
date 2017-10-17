VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrinterNCopias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame_NCopias 
      Caption         =   "Nº Copias"
      Height          =   975
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   1095
      Begin MSMask.MaskEdBox nCopias 
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         _Version        =   327680
         PromptChar      =   "_"
      End
   End
   Begin VB.Label Impresora 
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   3135
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Caption         =   "Impresora :"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "PrinterNCopias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_NCopias As Integer
Private Cancelar As Boolean
Public Property Get Numero_Copias() As Integer
Numero_Copias = m_NCopias
End Property
Public Property Let Numero_Copias(ByVal vNewValue As Integer)
m_NCopias = vNewValue
End Property
Private Sub Form_Load()

Impresora.Caption = ReadIniValue(Path_Local & "scp.ini", "Printer", "Docs")

nCopias.Mask = "##"
nCopias.Text = "1_"

'UpDown.Value = m_NCopias

Cancelar = True
End Sub
Private Sub btnAceptar_Click()
'm_NCopias = UpDown.Value
m_NCopias = Replace(nCopias.Text, "_", "")
Cancelar = False
Unload Me
End Sub
Private Sub btnCancelar_Click()
m_NCopias = 0
Cancelar = True
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Cancelar Then m_NCopias = 0
End Sub
