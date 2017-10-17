VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Copiando 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiando..."
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   327680
      Appearance      =   1
      MouseIcon       =   "Copiando.frx":0000
   End
End
Attribute VB_Name = "Copiando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Ws As Workspace, Error As Boolean
Private Db_H As Database, RsH As Recordset
Private Db_M As Database, RsM As Recordset
Private gauge As Integer
Private m_Nv As Double, m_Tipo As String
Public Property Get Nv() As Double
Nv = m_Nv
End Property
Public Property Let Nv(nuevo As Double)
m_Nv = nuevo
End Property
Public Property Let Tipo(nuevo As String)
m_Tipo = nuevo
End Property
Private Sub Form_Load()
If Archivo_Existe(Drive_Server & Path_Mdb, "filecopy.avi") Then
'    Animacion.Open (Drive_Server & Path_Server & "filecopy.avi")
'    Animacion.Play
End If
Me.Refresh
Error = False
gauge = 0
ProgressBar.Min = 1
'ProgressBar.max = 16
Set Ws = Workspaces(0)

If m_Tipo = "P2T" Then
    Set Db_H = Ws.OpenDatabase(Drive_Server & Path_Mdb & "ScpHist")
    Set Db_M = Ws.OpenDatabase(Drive_Server & Path_Mdb & "ScpMovs")
Else
    Set Db_H = Ws.OpenDatabase(Drive_Server & Path_Mdb & "ScpMovs")
    Set Db_M = Ws.OpenDatabase(Drive_Server & Path_Mdb & "ScpHist")
End If

Ws.BeginTrans

End Sub
Private Sub Form_Activate()

Dim nTabla As Integer, NombreTabla As String, NumerodeTablas As Integer
'Buscar

NumerodeTablas = Db_M.TableDefs.Count

ProgressBar.max = NumerodeTablas

ProgressBar.max = NumerodeTablas
For nTabla = 0 To NumerodeTablas - 1
    NombreTabla = Db_M.TableDefs(nTabla).Name
    If Left(NombreTabla, 4) <> "MSys" And NombreTabla <> "Detalle Texto" And NombreTabla <> "TablasMaestras" Then
        Debug.Print NombreTabla
        Traspasa NombreTabla
        ProgressBar.Value = nTabla + 1
    End If
Next

If False Then ' rutina ya no corre porque ya traspasa arriba (ver arriba)

Traspasa "Arco Sumergido"

Traspasa "Bultos"
'Traspasa "Detalle Texto" ' no tiene nv

' GD
Traspasa "GD Cabecera"
Traspasa "GD Detalle"
Traspasa "GD Especial Detalle"

Traspasa "ITO Esp"
' ITO FABRICACION
Traspasa "ITO Fab Cabecera"
Traspasa "ITO Fab Detalle"

Traspasa "ITO PG Cabecera"
Traspasa "ITO PG Detalle"

Traspasa "ITOe"

Traspasa "MovBodega"

' NV CABECERA
Traspasa "NV Cabecera", "Numero"
Traspasa "NV Detalle"

Traspasa "OT Esp Cabecera"
Traspasa "OT Esp Detalle"

' OT FABRICACION
Traspasa "OT Fab Cabecera"
Traspasa "OT Fab Detalle"

Traspasa "OTe Cabecera"
Traspasa "OTe Detalle"

' Planos Cabecera
Traspasa "Planos Cabecera", "NV"
' Planos Detalle
Traspasa "Planos Detalle"

Traspasa "Planos Recepcion"

Traspasa "SE Cabecera"

'Traspasa "Tablas Maestras" , no tiene nv

End If

If Error Then
    ' si se produjo algún error se deshacen los cambios
    Ws.Rollback
    m_Nv = 0
Else
   ' si no hubo error se graban los cambios
    Ws.CommitTrans
End If

Db_H.Close
Db_M.Close

Unload Me

End Sub
Private Sub Traspasa(Tabla As String, Optional CampoNV As String)

If Error Then Exit Sub

Set RsH = Db_H.OpenRecordset(Tabla)
Set RsM = Db_M.OpenRecordset(Tabla)

Tabla_TraspasaNV RsM, RsH, CampoNV

RsH.Close
RsM.Close
End Sub
Private Sub Tabla_TraspasaNV(RsO As Recordset, RsD As Recordset, Optional CampoNV As String)
' copia todos los campos de recordset origen a destino
Dim i As Long, nc As Integer, NombreCampo As String
'On Error GoTo Error
nc = RsO.Fields.Count

If RsO.Name = "NV Cabecera" Then
    CampoNV = "Numero"
End If

If CampoNV = "" Then CampoNV = "NV"

Do While Not RsO.EOF

    If RsO(CampoNV) = m_Nv Then
    
        RsD.AddNew
        For i = 0 To nc - 1
            NombreCampo = RsO.Fields(i).Name
            RsD(NombreCampo) = RsO(NombreCampo)
        Next
        RsD.Update
        
        RsO.Delete

    End If
    RsO.MoveNext
    
Loop

gauge = gauge + 1
'ProgressBar.Value = gauge
Me.Refresh
Exit Sub

Error:
'MsgBox "Tabla: " & RsO.Name & Chr(10) & "Nº " & RsO.Fields(0).Value & Chr(10) & "NV: " & m_NV & Chr(10) & "NV: " & RsD(CampoNV)
Error = True
End Sub
Private Sub Buscar()
' busca documentos con el mismo correlativo en origen y destino
' es decir en ScpMovs y ScpHist

'Set RsM = Db_M.OpenRecordset("ITO Fab Cabecera")
Set RsM = Db_M.OpenRecordset("OT Fab Cabecera")
Set RsH = Db_H.OpenRecordset("OT Fab Cabecera")
RsH.Index = "Número"

Do While Not RsM.EOF
    
    RsH.Seek "=", RsM!Número
    If Not RsH.NoMatch Then
        Debug.Print RsM!Número, RsM!Nv, RsH!Nv
    End If
    RsM.MoveNext
    
Loop

gauge = gauge + 1
ProgressBar.Value = gauge
Me.Refresh
Exit Sub
Error:
Error = True
End Sub
