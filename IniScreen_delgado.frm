VERSION 5.00
Begin VB.Form IniScreen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3855
   ClientLeft      =   2565
   ClientTop       =   2400
   ClientWidth     =   6735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "IniScreen_delgado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3675
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   6465
      Begin VB.Image imgLogo 
         Height          =   1455
         Left            =   480
         Picture         =   "IniScreen_delgado.frx":000C
         Top             =   240
         Width           =   5685
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1997-2013"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   2
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   1
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label lblOSVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Versión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5280
         TabIndex        =   3
         Top             =   2760
         Width           =   810
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "para Windows NT, 2000, Xp, 7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2745
         TabIndex        =   4
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sistema de Control de Producción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   6195
      End
   End
End
Attribute VB_Name = "IniScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
End Sub
Private Sub Form_Load()
'lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
'lblOSVersion = SysInfo.OSPlatform & " " & SysInfo.OSVersion
lblOSVersion = "Versión 2.0"
'lblProductName.Caption = App.Title
End Sub
Private Sub Frame1_Click()
Unload Me
End Sub

