VERSION 5.00
Begin VB.Form As_TablaBono 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Impuestos"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame_Tubest 
      Caption         =   "TABLA TUBEST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2535
      Left            =   5880
      TabIndex        =   34
      Top             =   120
      Width           =   2775
      Begin VB.TextBox Valor 
         Height          =   300
         Index           =   13
         Left            =   1920
         TabIndex        =   47
         Top             =   1680
         Width           =   600
      End
      Begin VB.TextBox Hasta 
         Height          =   300
         Index           =   14
         Left            =   1080
         TabIndex        =   49
         Top             =   2040
         Width           =   600
      End
      Begin VB.TextBox Desde 
         Height          =   300
         Index           =   14
         Left            =   240
         TabIndex        =   48
         Top             =   2040
         Width           =   600
      End
      Begin VB.TextBox Valor 
         Height          =   300
         Index           =   12
         Left            =   1920
         TabIndex        =   44
         Top             =   1320
         Width           =   600
      End
      Begin VB.TextBox Hasta 
         Height          =   300
         Index           =   13
         Left            =   1080
         TabIndex        =   46
         Top             =   1680
         Width           =   600
      End
      Begin VB.TextBox Desde 
         Height          =   300
         Index           =   13
         Left            =   240
         TabIndex        =   45
         Top             =   1680
         Width           =   600
      End
      Begin VB.TextBox Valor 
         Height          =   300
         Index           =   11
         Left            =   1920
         TabIndex        =   41
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox Hasta 
         Height          =   300
         Index           =   12
         Left            =   1080
         TabIndex        =   43
         Top             =   1320
         Width           =   600
      End
      Begin VB.TextBox Desde 
         Height          =   300
         Index           =   12
         Left            =   240
         TabIndex        =   42
         Top             =   1320
         Width           =   600
      End
      Begin VB.TextBox Valor 
         Height          =   300
         Index           =   10
         Left            =   1920
         TabIndex        =   38
         Top             =   600
         Width           =   600
      End
      Begin VB.TextBox Hasta 
         Height          =   300
         Index           =   11
         Left            =   1080
         TabIndex        =   40
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox Desde 
         Height          =   300
         Index           =   11
         Left            =   240
         TabIndex        =   39
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox Valor 
         Height          =   300
         Index           =   14
         Left            =   1920
         TabIndex        =   50
         Top             =   2040
         Width           =   600
      End
      Begin VB.TextBox Hasta 
         Height          =   300
         Index           =   10
         Left            =   1080
         TabIndex        =   37
         Top             =   600
         Width           =   600
      End
      Begin VB.TextBox Desde 
         Height          =   300
         Index           =   10
         Left            =   240
         TabIndex        =   36
         Top             =   600
         Width           =   600
      End
      Begin VB.Label lbl 
         Caption         =   "DESDE       HASTA       VALOR"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame_Tubular 
      Caption         =   "TABLA TUBULARES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2535
      Left            =   3000
      TabIndex        =   17
      Top             =   120
      Width           =   2775
      Begin VB.TextBox Desde 
         Height          =   300
         Index           =   9
         Left            =   240
         TabIndex        =   31
         Top             =   2040
         Width           =   600
      End
      Begin VB.TextBox Hasta 
         Height          =   300
         Index           =   9
         Left            =   1080
         TabIndex        =   32
         Top             =   2040
         Width           =   600
      End
      Begin VB.TextBox Valor 
         Height          =   300
         Index           =   9
         Left            =   1920
         TabIndex        =   33
         Top             =   2040
         Width           =   600
      End
      Begin VB.TextBox Desde 
         Height          =   300
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   600
      End
      Begin VB.TextBox Hasta 
         Height          =   300
         Index           =   5
         Left            =   1080
         TabIndex        =   20
         Top             =   600
         Width           =   600
      End
      Begin VB.TextBox Valor 
         Height          =   300
         Index           =   5
         Left            =   1920
         TabIndex        =   21
         Top             =   600
         Width           =   600
      End
      Begin VB.TextBox Desde 
         Height          =   300
         Index           =   6
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox Hasta 
         Height          =   300
         Index           =   6
         Left            =   1080
         TabIndex        =   23
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox Valor 
         Height          =   300
         Index           =   6
         Left            =   1920
         TabIndex        =   24
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox Desde 
         Height          =   300
         Index           =   7
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   600
      End
      Begin VB.TextBox Hasta 
         Height          =   300
         Index           =   7
         Left            =   1080
         TabIndex        =   26
         Top             =   1320
         Width           =   600
      End
      Begin VB.TextBox Valor 
         Height          =   300
         Index           =   7
         Left            =   1920
         TabIndex        =   27
         Top             =   1320
         Width           =   600
      End
      Begin VB.TextBox Desde 
         Height          =   300
         Index           =   8
         Left            =   240
         TabIndex        =   28
         Top             =   1680
         Width           =   600
      End
      Begin VB.TextBox Hasta 
         Height          =   300
         Index           =   8
         Left            =   1080
         TabIndex        =   29
         Top             =   1680
         Width           =   600
      End
      Begin VB.TextBox Valor 
         Height          =   300
         Index           =   8
         Left            =   1920
         TabIndex        =   30
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lbl 
         Caption         =   "DESDE       HASTA       VALOR"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "TABLA VIGAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.TextBox Hasta 
         Height          =   300
         Index           =   0
         Left            =   1080
         TabIndex        =   3
         Top             =   600
         Width           =   600
      End
      Begin VB.TextBox Hasta 
         Height          =   300
         Index           =   1
         Left            =   1080
         TabIndex        =   6
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox Hasta 
         Height          =   300
         Index           =   2
         Left            =   1080
         TabIndex        =   9
         Top             =   1320
         Width           =   600
      End
      Begin VB.TextBox Hasta 
         Height          =   300
         Index           =   3
         Left            =   1080
         TabIndex        =   12
         Top             =   1680
         Width           =   600
      End
      Begin VB.TextBox Hasta 
         Height          =   300
         Index           =   4
         Left            =   1080
         TabIndex        =   15
         Top             =   2040
         Width           =   600
      End
      Begin VB.TextBox Desde 
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   600
      End
      Begin VB.TextBox Desde 
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox Desde 
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   600
      End
      Begin VB.TextBox Desde 
         Height          =   300
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   600
      End
      Begin VB.TextBox Desde 
         Height          =   300
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   600
      End
      Begin VB.TextBox Valor 
         Height          =   300
         Index           =   0
         Left            =   1920
         TabIndex        =   4
         Top             =   600
         Width           =   600
      End
      Begin VB.TextBox Valor 
         Height          =   300
         Index           =   1
         Left            =   1920
         TabIndex        =   7
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox Valor 
         Height          =   300
         Index           =   2
         Left            =   1920
         TabIndex        =   10
         Top             =   1320
         Width           =   600
      End
      Begin VB.TextBox Valor 
         Height          =   300
         Index           =   3
         Left            =   1920
         TabIndex        =   13
         Top             =   1680
         Width           =   600
      End
      Begin VB.TextBox Valor 
         Height          =   300
         Index           =   4
         Left            =   1920
         TabIndex        =   16
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label lbl 
         Caption         =   "DESDE       HASTA       VALOR"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7440
      TabIndex        =   52
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5880
      TabIndex        =   51
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "** VALOR expresado en pesos"
      Height          =   375
      Left            =   3000
      TabIndex        =   54
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "* DESDE y HASTA expresados en Kg/m"
      Height          =   495
      Left            =   120
      TabIndex        =   53
      Top             =   2760
      Width           =   2415
   End
End
Attribute VB_Name = "As_TablaBono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private i As Integer, m_Fecha As String
Private DbD As Database, RsAsTB As Recordset
Private m_Hasta As Integer
Private Sub Form_Load()

m_Hasta = 4

Set DbD = OpenDatabase(data_file)
Set RsAsTB = DbD.OpenRecordset("Tabla Bono Arco Sumergido")
RsAsTB.Index = "nv-estructura-item"

Inicializa

End Sub
Private Sub Inicializa()
' lee
With RsAsTB

.Seek "=", 0, "V", 1
If Not .NoMatch Then
    Do While Not .EOF
        If !estructura = "V" Then
            i = !Item - 1
            Desde(i).Text = !Desde
            Hasta(i).Text = !Hasta
            Valor(i).Text = !Valor
        End If
        .MoveNext
    Loop
End If

.Seek "=", 0, "T", 1
If Not .NoMatch Then
    Do While Not .EOF
        If !estructura = "T" Then
            i = !Item + 4
            Desde(i).Text = !Desde
            Hasta(i).Text = !Hasta
            Valor(i).Text = !Valor
        End If
        .MoveNext
    Loop
End If

.Seek "=", 0, "S", 1
If Not .NoMatch Then
    Do While Not .EOF
        If !estructura = "S" Then
            i = !Item + 9
            Desde(i).Text = !Desde
            Hasta(i).Text = !Hasta
            Valor(i).Text = !Valor
        End If
        .MoveNext
    Loop
End If

End With

End Sub
Private Sub Desde_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Hasta(Index).SetFocus
End Sub
Private Sub Hasta_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Valor(Index).SetFocus
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub Valor_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If Index = m_Hasta - 1 Then Index = 0 Else Index = Index + 1
    Desde(Index).SetFocus
End If
End Sub
Private Sub btnAceptar_Click()
'grabar

With RsAsTB

For i = 1 To 5
    .Seek "=", 0, "V", i
    If .NoMatch Then
        .AddNew
        !Nv = 0
        !estructura = "V"
        !Item = i
    Else
        .Edit
    End If
    !Desde = Val(Desde(i - 1).Text)
    !Hasta = Val(Hasta(i - 1).Text)
    !Valor = m_CDbl(Valor(i - 1).Text)
    .Update
Next

For i = 6 To 10
    .Seek "=", 0, "T", i - 5
    If .NoMatch Then
        .AddNew
        !Nv = 0
        !estructura = "T"
        !Item = i - 5
    Else
        .Edit
    End If
    !Desde = Val(Desde(i - 1).Text)
    !Hasta = Val(Hasta(i - 1).Text)
    !Valor = m_CDbl(Valor(i - 1).Text)
    .Update
Next

For i = 11 To 15
    .Seek "=", 0, "S", i - 10
    If .NoMatch Then
        .AddNew
        !Nv = 0
        !estructura = "S"
        !Item = i - 10
    Else
        .Edit
    End If
    !Desde = Val(Desde(i - 1).Text)
    !Hasta = Val(Hasta(i - 1).Text)
    !Valor = m_CDbl(Valor(i - 1).Text)
    .Update
Next

End With

Unload Me

End Sub
Private Sub btnCancelar_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
RsAsTB.Close
DbD.Close
End Sub
