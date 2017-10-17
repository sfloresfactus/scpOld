VERSION 5.00
Begin VB.Form ObrasActivar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar Obras"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameOrden 
      Caption         =   "NV Ordenadas por"
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   3495
      Begin VB.OptionButton OpOrdenNV 
         Caption         =   "Nombre"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OpOrdenNV 
         Caption         =   "Número"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "NV"
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   1920
      TabIndex        =   3
      Top             =   0
      Width           =   1695
      Begin VB.OptionButton Nv_Op 
         Caption         =   "&Todas"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Nv_Op 
         Caption         =   "&Solo Activas"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "OBRAS"
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.OptionButton Obras_Op 
         Caption         =   "&Terminadas"
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Obras_Op 
         Caption         =   "En &Proceso"
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton Boton 
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Boton 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "ObrasActivar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Db As Database
'Private Rs As Recordset
Private sql As String
Private Rs As ADODB.Recordset
Private Sub Form_Load()
'Set Db = OpenDatabase(Syst_file, False, False, ";pwd=eml")
'Set Rs = Db.OpenRecordset("Usuarios")
'Rs.Index = "Nombre"
Set Rs = New ADODB.Recordset
'Set Rs = cnxSqlServer.OpenRecordset("SELECT * FROM usuarios WHERE nombre='" & Usuario.nombre & "'")

sql = "SELECT * FROM usuarios"

'Set Rs = cnxSqlServer.OpenRecordset("SELECT * FROM usuarios") ' con adodb no coinciden los tipos

Rs.Open sql, CnxSqlServer_scp0, adOpenDynamic

If False Then
With Rs
Do While Not .EOF
    Debug.Print !nombre
    .MoveNext
Loop
End With
End If

'Rs.Seek "=", Usuario.nombre
'If Not Rs.NoMatch Then
If Not Rs.EOF Then

    Usuario.ObrasTerminadas = IIf(NoNulo(Rs![nv_terminadas]) = "S", True, False)
    Usuario.Nv_Activas = IIf(Rs![Nv_Activas] = "S", True, False)
    Usuario.Nv_Orden = NoNulo(Rs!Nv_Orden)
    
    If Usuario.Nv_Orden = "A" Then
        Nv_Index = "Nombre"
'        OpOrdenNV(0).Value = False
        OpOrdenNV(1).Value = True
    Else
        Nv_Index = "Numero"
        OpOrdenNV(0).Value = True
    End If
    
End If

Obras_Op(0).Value = Usuario.ObrasTerminadas
Obras_Op(1).Value = Not Usuario.ObrasTerminadas

Nv_Op(0).Value = Usuario.Nv_Activas
Nv_Op(1).Value = Not Usuario.Nv_Activas

OpOrdenNV(1).Value = Not OpOrdenNV(0).Value

Rs.Close

End Sub
Private Sub Boton_Click(Index As Integer)

If Index = 0 Then

    ' aceptar
    Usuario.ObrasTerminadas = Obras_Op(0).Value
    Usuario.Nv_Activas = Nv_Op(0).Value
    
    sql = "UPDATE usuarios SET "
    sql = sql & "nv_terminadas = '" & IIf(Usuario.ObrasTerminadas, "S", "N") & "',"
    sql = sql & "nv_activas = '" & IIf(Usuario.Nv_Activas, "S", "N") & "',"
    
    If OpOrdenNV(0).Value = True Then
        Usuario.Nv_Orden = "N"
        sql = sql & "nv_orden = 'N'"
        Nv_Index = "Numero"
    Else
        Usuario.Nv_Orden = "A"
        sql = sql & "nv_orden = 'A'"
        Nv_Index = "Nombre"
    End If
    
'    sql = sql & " WHERE nombre='" & UsrNombre.Text & "'"
    sql = sql & " WHERE nombre='" & Usuario.nombre & "'"
    
    CnxSqlServer_scp0.Execute sql

    Unload Me
    
Else

    ' cancelar
    Unload Me
    
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Rs.Close
'Db.Close
End Sub
