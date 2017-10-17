VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ip_utiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utilitario IP"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   327680
      Appearance      =   1
      MouseIcon       =   "ip_utiles.frx":0000
   End
   Begin VB.CommandButton btnLimpiar 
      Caption         =   "&Limpiar"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox resultado 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   960
      Width           =   4335
   End
   Begin VB.CommandButton btnAccion 
      Caption         =   "&Comenzar"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "ip_utiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nIP As Integer, sIP As String
Private mascara As String, n As Integer
Private Db As Database, Rs As Recordset, prt As Printer
Private Sub btnImprimir_Click()
Set prt = Printer
prt.Print "Listado de IP "; Format(Date); " "; Time
prt.Print " "
prt.Print " Nº  IP        EQUIPO"
prt.Print resultado.Text
prt.EndDoc
End Sub

Private Sub Form_Load()
'Set Db = OpenDatabase("c:\scp\acr3006-dualpro\red")
'Set Rs = Db.OpenRecordset("puntos")
'Rs.Index = "equipos"
mascara = "192.168.0."
'resultado.Enabled = False
'resultado.MultiLine = True
pb.Min = 0
pb.max = 100
pb.visible = False
End Sub
Private Sub btnAccion_Click()
MousePointer = vbHourglass

Dim Nombre_Equipo As String, Nombre_Usuario As String

pb.visible = True
Debug.Print "Tu IP " & GetIPAddress ' ok
'Debug.Print "Tu IP " & GetIPAddress("bodega") ' ok
Debug.Print "host " & GetIPHostName

'Debug.Print Ping("192.168.0.12", "1", False) ' ok

n = 0
For nIP = 1 To 100
    sIP = mascara & Trim(Str(nIP))
    If Ping(sIP, "0", False) = 0 Then
        n = n + 1
'        Debug.Print n, sIP, GetHostFromIP(sIP)

        ' busca nombre de usuario
        Nombre_Equipo = GetHostFromIP(sIP)
'        Rs.Seek "=", Nombre_Equipo
        
'        If Rs.NoMatch Then
            Nombre_Usuario = "no encontrado"
            Nombre_Usuario = ""
'        Else
'            Nombre_Usuario = NoNulo(Rs!Usuario)
'        End If
        
        resultado.Text = resultado.Text & PadL(Str(n), 2) & " " & PadL(sIP, 13) & " " & Nombre_Equipo & " " & Nombre_Usuario & vbCrLf
        
    End If
    pb.Value = nIP
    Me.Refresh
Next
pb.visible = False
MousePointer = vbDefault
End Sub
Private Sub btnLimpiar_Click()
resultado.Text = ""
End Sub
