Attribute VB_Name = "scpNuevo"
Option Explicit
Public scpNew_aNv(999, 3) As String
Public scpNew_aNv_size As Integer ' tamaño del arreglo
Public scpNew_aCeCo(999, 1) As String
Public scpNew_aCeCo_size As Integer ' tamaño del arreglo
Dim i As Integer
'
' creado el 24/09/2015
Public Sub scpNew_NVLeer()
Dim sql As String
Dim Rs As New ADODB.Recordset
sql = "SELECT * FROM vw_nv ORDER BY nv"
'MsgBox "1"
Rs.Open sql, CnxSqlServer_delgado1405
'MsgBox "2"

'Debug.Print "dentro del scpNewCentroCostoLeer"
i = -1
With Rs
Do While Not .EOF
    i = i + 1
    'Debug.Print ![Nv]; " "; ![obra]; " "; ![ccCodigo]; " "; ![ccDescripcion]
    scpNew_aNv(i, 0) = ![Nv]
    scpNew_aNv(i, 1) = ![obra]
    scpNew_aNv(i, 2) = ![ccCodigo]
    scpNew_aNv(i, 3) = ![ccDescripcion]
    .MoveNext
Loop
scpNew_aNv_size = i + 1
.Close
End With

End Sub
Public Sub scpNew_CentroCostoLeer()
Dim sql As String
Dim Rs As New ADODB.Recordset
sql = "SELECT * FROM tb_centrocosto"
Rs.Open sql, CnxSqlServer_delgado1405

'Debug.Print "dentro del scpNewCentroCostoLeer"
i = -1
With Rs
Do While Not .EOF
    i = i + 1
    'Debug.Print ![Nv]; " "; ![obra]; " "; ![ccCodigo]; " "; ![ccDescripcion]
    scpNew_aCeCo(i, 0) = ![Codigo]
    scpNew_aCeCo(i, 1) = ![Descripcion]
    .MoveNext
Loop
scpNew_aCeCo_size = i + 1
.Close
End With

End Sub
Public Function scpNew_Nv2ccCodigo(Nv As Integer)
' busca indice de arreglo scpNew_aCeCo(3,999)
Dim indice As Integer, i As Integer
indice = -1 ' no encontrado
If Nv = 3541 Then
    Debug.Print
End If
For i = 0 To scpNew_aNv_size
    If Val(scpNew_aNv(i, 0)) = Nv Then
        indice = i
        i = 999
    End If
Next
scpNew_Nv2ccCodigo = indice
End Function
