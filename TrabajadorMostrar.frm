VERSION 5.00
Begin VB.Form TrabajadorMostrar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha del Trabajador"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame_Elementos 
      Caption         =   "Elementos de Seguridad"
      Height          =   4095
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   5175
      Begin VB.Label dato13 
         Caption         =   "dato13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   31
         Top             =   3600
         Width           =   3495
      End
      Begin VB.Label dato12 
         Caption         =   "dato12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   30
         Top             =   3240
         Width           =   3495
      End
      Begin VB.Label dato11 
         Caption         =   "dato11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Label dato10 
         Caption         =   "dato10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   28
         Top             =   2520
         Width           =   3495
      End
      Begin VB.Label dato9 
         Caption         =   "dato9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   2160
         Width           =   3495
      End
      Begin VB.Label dato8 
         Caption         =   "dato8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   26
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label dato7 
         Caption         =   "dato7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3720
         TabIndex        =   25
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label dato6 
         Caption         =   "dato6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   24
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label dato5 
         Caption         =   "dato5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label dato4 
         Caption         =   "dato4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3720
         TabIndex        =   22
         Top             =   720
         Width           =   735
      End
      Begin VB.Label dato3 
         Caption         =   "dato3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label dato2 
         Caption         =   "dato2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3720
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label dato1 
         Caption         =   "dato1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   1440
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Talla Chaqueta"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Guante"
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "Nº Calzado"
         Height          =   255
         Index           =   8
         Left            =   2760
         TabIndex        =   16
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "Tipo Calzado"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "Talla Pantalon"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Tipo Lente"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Tapon Oido"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "Color Casco"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "Polera"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "T.Agua"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lbl 
         Caption         =   "Bota"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   8
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lbl 
         Caption         =   "Parka"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   7
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lbl 
         Caption         =   "Termico"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   6
         Top             =   3600
         Width           =   1095
      End
   End
   Begin VB.Label lbl 
      Caption         =   "Sección"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label seccion 
      Caption         =   "seccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label nombres 
      Caption         =   "nombres"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label ruttrabajador 
      Caption         =   "rut"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "RUT"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "TrabajadorMostrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DbD As Database, RsT As Recordset
Private m_Rut As String, a_Seccion(1, 20), i As Integer
Public Property Let Rut(ByVal New_Value As String)
m_Rut = New_Value
End Property
Private Sub Form_Load()

a_Seccion(0, 1) = "ADQ"
a_Seccion(1, 1) = "Adquisicion"
a_Seccion(0, 2) = "ARS"
a_Seccion(1, 2) = "Arco Sumergido"
a_Seccion(0, 3) = "BOD"
a_Seccion(1, 3) = "Bodega"
a_Seccion(0, 4) = "CON"
a_Seccion(1, 4) = "Contabilidad"
a_Seccion(0, 5) = "CHO"
a_Seccion(1, 5) = "Chofer"
a_Seccion(0, 6) = "DES"
a_Seccion(1, 6) = "Despacho"
a_Seccion(0, 7) = "GER"
a_Seccion(1, 7) = "Gerencia"
a_Seccion(0, 8) = "GRU"
a_Seccion(1, 8) = "Gruero"
a_Seccion(0, 9) = "GUI"
a_Seccion(1, 9) = "Guillotina"
a_Seccion(0, 10) = "MEL"
a_Seccion(1, 10) = "Mantencion Electrica"
a_Seccion(0, 11) = "OXI"
a_Seccion(1, 11) = "Oxicorte"
a_Seccion(0, 12) = "PPL"
a_Seccion(1, 12) = "Patio Plancha"
a_Seccion(0, 13) = "PIN"
a_Seccion(1, 13) = "Pintura"
a_Seccion(0, 14) = "PLA"
a_Seccion(1, 14) = "Plasma"
a_Seccion(0, 15) = "PMA"
a_Seccion(1, 15) = "Prep. Material"
a_Seccion(0, 16) = "PRO"
a_Seccion(1, 16) = "Producción"
a_Seccion(0, 17) = "OPE"
a_Seccion(1, 17) = "Operaciones"

Set DbD = OpenDatabase(data_file)
Set RsT = DbD.OpenRecordset("Trabajadores")

With RsT
.Index = "rut"
.Seek "=", m_Rut

If .NoMatch Then
    ' no existe
Else
    
    ruttrabajador.Caption = m_Rut
    nombres.Caption = !appaterno & " " & !apmaterno & " " & !nombres
    seccion.Caption = ""

    dato1.Caption = NoNulo(!dato1)
    dato2.Caption = NoNulo(!dato2)
    dato3.Caption = NoNulo(!dato3)
    dato4.Caption = NoNulo(!dato4)
    dato5.Caption = NoNulo(!dato5)
    dato6.Caption = NoNulo(!dato6)
    dato7.Caption = NoNulo(!dato7)
    dato8.Caption = NoNulo(!dato8)
    dato9.Caption = NoNulo(!dato9)
    dato10.Caption = NoNulo(!dato10)
    dato11.Caption = NoNulo(!dato11)
    dato12.Caption = NoNulo(!dato12)
    dato13.Caption = NoNulo(!dato13)

    For i = 1 To 17
        If a_Seccion(0, i) = !clase1 Then
            seccion.Caption = a_Seccion(1, i)
        End If
    Next

End If

.Close
End With

End Sub
