VERSION 5.00
Begin VB.Form ObrasDesActivar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspasa Obras"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAccion 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton btnAccion 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.ListBox ListaObras2 
      Height          =   2790
      Left            =   3720
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   2415
   End
   Begin VB.ListBox ListaObras1 
      Height          =   2790
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lbl 
      Caption         =   "OBRAS EN PROCESO"
      DragIcon        =   "ObrasDesActivar.frx":0000
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lbl 
      Caption         =   "OBRAS TERMINADAS"
      DragIcon        =   "ObrasDesActivar.frx":0442
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "ObrasDesActivar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Traspasa NV desde archivo Movs a Hist y viceversa
' se Usa Transacciones
Option Explicit
Private m_Nv As Double
Private Db_H As Database, H_NVcab As Recordset
Private Db_M As Database, M_NVcab As Recordset
Private Sub Form_Load()

Set Db_H = OpenDatabase(Drive_Server & Path_Mdb & "ScpHist")
Set H_NVcab = Db_H.OpenRecordset("NV Cabecera")
H_NVcab.Index = "Numero"

Set Db_M = OpenDatabase(Drive_Server & Path_Mdb & "ScpMovs")
Set M_NVcab = Db_M.OpenRecordset("NV Cabecera")
M_NVcab.Index = "Numero"

Do While Not H_NVcab.EOF
    ListaObras1.AddItem Format(H_NVcab!Numero, "0000") & " - " & H_NVcab!obra
    H_NVcab.MoveNext
Loop
Do While Not M_NVcab.EOF
    ListaObras2.AddItem Format(M_NVcab!Numero, "0000") & " - " & M_NVcab!obra
    M_NVcab.MoveNext
Loop

End Sub
Private Sub btnAccion_Click(Index As Integer)
Dim NvaTraspasar As String
If Index = 0 Then ' de "en proceso" a historico

    NvaTraspasar = Trim(ListaObras2.Text)
    If NvaTraspasar = "" Then
        MsgBox "Debe Escoger NV de OBRAS EN PROCESO"
        Exit Sub
    End If

    Traspaso "P2T"

Else ' de historico a "en proceso"

    NvaTraspasar = Trim(ListaObras1.Text)
    If NvaTraspasar = "" Then
        MsgBox "Debe Escoger NV de OBRAS TERMINADAS"
        Exit Sub
    End If

    Traspaso "T2P"

End If
End Sub
'Private Sub ListaObras1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If Button = 1 Then
'    ListaObras1.DragIcon = lbl(0).DragIcon
'    ListaObras1.Drag
'End If
'End Sub
'Private Sub ListaObras1_DragDrop(Source As Control, x As Single, Y As Single)
Private Sub Traspaso(Tipo As String)

If Tipo = "P2T" Then

    If MsgBox("OBRA : " & ListaObras2.Text & vbCr & "SERÁ TRASPASADA A" & vbCr & "OBRAS TERMINADAS" & vbCr & "¿ ESTÁ SEGURO ?", vbYesNoCancel, "ATENCIÓN") = vbYes Then
        
        MousePointer = vbHourglass
        
        m_Nv = Left(ListaObras2.Text, 4)
        
        Copiando.Nv = m_Nv
        Copiando.Tipo = "P2T"
        Copiando.Show 1
        
        MousePointer = vbDefault
        
        If Copiando.Nv = 0 Then
            Beep
            MsgBox "No es posible Traspasar Obra" & vbCr & "Se Produjo un Error"
        Else
            ListaObras1.AddItem ListaObras2.Text
            ListaObras2.RemoveItem (ListaObras2.ListIndex)
        End If
    '    MsgBox "Traspaso Finalizó con Éxito"
    End If
    
Else

    If MsgBox("OBRA : " & ListaObras1.Text & vbCr & "SERÁ TRASPASADA A" & vbCr & "OBRAS EN PROCESO" & vbCr & "¿ ESTÁ SEGURO ?", vbYesNoCancel, "ATENCIÓN") = vbYes Then
        
        MousePointer = vbHourglass
        
        m_Nv = Left(ListaObras1.Text, 4)
        
        Copiando.Nv = m_Nv
        Copiando.Tipo = "T2P"
        Copiando.Show 1
        
        MousePointer = vbDefault
        
        If Copiando.Nv = 0 Then
            Beep
            MsgBox "No es posible Traspasar Obra" & vbCr & "Se Produjo un Error"
        Else
            ListaObras2.AddItem ListaObras1.Text
            ListaObras1.RemoveItem (ListaObras1.ListIndex)
        End If
    '    MsgBox "Traspaso Finalizó con Éxito"
    End If
    
End If


End Sub
'Private Sub ListaObras1_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
'Select Case State
'Case 0
'    ListaObras2.DragIcon = lbl(1).DragIcon
'Case 1
'    ListaObras2.DragIcon = lbl(0).DragIcon
'End Select
'End Sub
'Private Sub ListaObras2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If Button = 1 Then
'    ListaObras2.DragIcon = lbl(0).DragIcon
'    ListaObras2.Drag
'End If
'End Sub
'Private Sub ListaObras2_DragDrop(Source As Control, x As Single, Y As Single)
'If Source.Name = "ListaObras2" Then Exit Sub
''MsgBox "de uno"
'End Sub
Private Sub Form_Unload(Cancel As Integer)
DataBases_Cerrar
'Db_H.Close
'Db_M.Close
End Sub
