VERSION 5.00
Begin VB.Form PlanoMover 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspasa Plano a otra Obra"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox NV_Destino 
      Height          =   1620
      Left            =   5880
      TabIndex        =   10
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancela"
      Height          =   375
      Left            =   7800
      TabIndex        =   9
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.ListBox ListaPlanos 
      Height          =   1620
      Left            =   3360
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.ListBox NV_Origen 
      Height          =   1620
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Observacion 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   3360
      TabIndex        =   15
      Top             =   3120
      Width           =   5295
   End
   Begin VB.Label Rev 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4080
      TabIndex        =   14
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Plano 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4080
      TabIndex        =   13
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lbl 
      Caption         =   "NV DE DESTINO"
      Height          =   255
      Index           =   2
      Left            =   5880
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label NVDestino 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5880
      TabIndex        =   11
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label lbl 
      Caption         =   "OBSERVACIÓN"
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "P&LANOS"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "NV DE ORIGEN"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "REV"
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   6
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lbl 
      Caption         =   "PLANO"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   5
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label NvOrigen 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   3135
   End
End
Attribute VB_Name = "PlanoMover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Db As Database, RsNv As Recordset, RsPc As Recordset, RsPd As Recordset
Private RsOTfd As Recordset
Private NVnum_Ori As Double, NVnum_Des As Double, m_NvArea As Integer
Private Sub Form_Load()
Dim i As Integer
Set Db = OpenDatabase(mpro_file)
Set RsNv = Db.OpenRecordset("NV Cabecera")
RsNv.Index = "Numero"
Set RsPc = Db.OpenRecordset("Planos Cabecera")
RsPc.Index = "NV-Plano"
Set RsPd = Db.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Item"

Set RsOTfd = Db.OpenRecordset("OT Fab Detalle")

Variables_Limpiar

'i = -1
'On Error GoTo Sigue
With RsNv
.MoveFirst
Do While Not RsNv.EOF
'    i = i + 1
    NV_Origen.AddItem Format(!Numero, "0000") & " - " & !Obra
    NV_Destino.AddItem Format(!Numero, "0000") & " - " & !Obra
    RsNv.MoveNext
Loop
End With
Sigue:
On Error GoTo 0

m_NvArea = 0

End Sub
Private Sub Variables_Limpiar()
Plano.Caption = ""
Rev.Caption = ""
Observacion.Caption = ""
End Sub
Private Sub ListaPlanos_Click()
Dim p1 As Integer, p2 As Integer, mi_Plano As String
mi_Plano = ListaPlanos.Text
p1 = InStr(1, mi_Plano, ",")
p2 = InStr(p1 + 1, mi_Plano, ",")

Plano.Caption = Left(mi_Plano, p1 - 1)
Observacion.Caption = Mid(mi_Plano, p2 + 2)

Rev.Caption = Mid(mi_Plano, p1 + 2, p2 - p1 - 2)

End Sub
Private Sub NV_Origen_Click()
ListaPlanos.Clear
NVnum_Ori = 0
NVnum_Des = 0
NvOrigen.Caption = ""
NVDestino.Caption = ""
Plano.Caption = ""
Rev.Caption = ""
Observacion.Caption = ""
End Sub
Private Sub NV_Origen_DblClick()
NVnum_Ori = Left(NV_Origen.Text, 4)
Nota_Seleccionar NV_Origen.ListIndex
End Sub
Private Sub Nota_Seleccionar(s As Integer)
ListaPlanos.Clear

RsPc.Seek ">", NVnum_Ori, m_NvArea, 0
If Not RsPc.NoMatch Then
    Do While Not RsPc.EOF
        If RsPc!Nv = NVnum_Ori Then
            If RsPc!Editable Then
                ListaPlanos.AddItem RsPc!Plano & ", " & RsPc!Rev & ", " & RsPc!Observacion
            End If
        End If
        RsPc.MoveNext
    Loop
End If

NvOrigen.Caption = NV_Origen.Text

End Sub
Private Sub NV_Destino_Click()
NVDestino.Caption = NV_Destino.Text
NVnum_Des = Left(NVDestino.Caption, 4)
End Sub
Private Sub NV_Destino_DblClick()
NVDestino.Caption = NV_Destino.Text
NVnum_Des = Left(NVDestino.Caption, 4)
End Sub
Private Sub btnCancelar_Click()
Variables_Limpiar
Unload Me
End Sub
Private Sub btnOk_Click()

If Validar Then
'    NVnum_Ori
'    NVnum_Des
'    Plano.Caption
    
    If MsgBox("Seguro?", vbYesNoCancel) = vbYes Then
        
        ' modifica plano cabecera
        With RsPc
        .Seek "=", NVnum_Ori, m_NvArea, Plano.Caption
        If Not .NoMatch Then
            .Edit
            !Nv = NVnum_Des
            .Update
        End If
        End With
        
        ' modifica plano detalle
        With RsPd
        .Seek ">=", NVnum_Ori, m_NvArea, Plano.Caption, 1
        If Not .NoMatch Then
            Do While Not .EOF
                If NVnum_Ori <> !Nv Or Plano.Caption <> !Plano Then Exit Do
                .Edit
                !Nv = NVnum_Des
                .Update
                .MoveNext
            Loop
        End If
        End With
        
        NV_Origen_DblClick
        
    End If
End If

End Sub
Private Function Validar() As Boolean
MousePointer = vbHourglass
Validar = False

If NvOrigen.Caption = "" Then
    Beep
    MsgBox "Debe Elegir NV de Origen"
    NV_Origen.SetFocus
    MousePointer = vbDefault
    Exit Function
End If

If Plano.Caption = "" Then
    Beep
    MsgBox "Debe Elegir Plano"
    ListaPlanos.SetFocus
    MousePointer = vbDefault
    Exit Function
End If

If NVDestino.Caption = "" Then
    Beep
    MsgBox "Debe Elegir NV de Destino"
    NV_Destino.SetFocus
    MousePointer = vbDefault
    Exit Function
End If

If NvOrigen.Caption = NVDestino.Caption Then
    Beep
    MsgBox "NV de Destino debe ser" & vbCr & "distinta que NV de Origen"
    NV_Destino.SetFocus
    MousePointer = vbDefault
    Exit Function
End If

' verifica que NO exista plano con el mismo nombre en NV destino
RsPc.Seek "=", NVnum_Des, m_NvArea, , Plano.Caption
If Not RsPc.NoMatch Then
    Beep
    MsgBox "YA EXISTE un Plano" & vbCr & Plano.Caption & vbCr & "En NV de Destino"
    MousePointer = vbDefault
    Exit Function
End If

' verifica que plano NO tenga Otf
With RsOTfd
.MoveFirst
Do While Not .EOF
    If !Nv = NVnum_Ori And !Plano = Plano.Caption Then
        Beep
        MsgBox "Plano ya tiene OT Fab Nº" & !Numero
        MousePointer = vbDefault
        Exit Function
    End If
    .MoveNext
Loop
End With

Validar = True
MousePointer = vbDefault
End Function
