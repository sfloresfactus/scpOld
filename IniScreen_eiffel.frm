VERSION 5.00
Begin VB.Form IniScreen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4335
   ClientLeft      =   2565
   ClientTop       =   2400
   ClientWidth     =   6735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "IniScreen_eiffel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   6465
      Begin VB.Image imgLogo 
         Height          =   3525
         Left            =   120
         Picture         =   "IniScreen_eiffel.frx":000C
         Top             =   360
         Width           =   4110
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright 1997-2012"
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
         TabIndex        =   2
         Top             =   3360
         Width           =   2295
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
         Left            =   5400
         TabIndex        =   3
         Top             =   3000
         Width           =   810
      End
      Begin VB.Label lblPlatform 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NT, 2000, Xp, 7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4185
         TabIndex        =   4
         Top             =   2040
         Width           =   2175
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
         Height          =   1395
         Left            =   4200
         TabIndex        =   5
         Top             =   360
         Width           =   2115
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

