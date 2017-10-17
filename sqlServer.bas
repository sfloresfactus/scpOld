Attribute VB_Name = "sqlServer"
Option Explicit
' rutinas para manejo de BD sqlServer
Private RsPaso As New ADODB.Recordset
Private sql As String
Public Function nv_Leer(aNotaVenta) As Integer
' lee notas de venta y las trae a arreglo de memoria "aNv"

Dim nr As String, i As String
nv_Leer = 0

With RsPaso
sql = "SELECT * FROM tb_nv ORDER BY nv DESC"
.Open sql, CnxSqlServer_delgado1405
i = 0
Do While Not .EOF
    i = i + 1
    aNotaVenta(i, 0) = !Nv
    aNotaVenta(i, 1) = !obra
    .MoveNext
Loop
.Close
End With

nv_Leer = i

End Function
Public Function nv_Buscar(Numero As Integer) As NotaVenta
Dim Nv As NotaVenta
Nv.Numero = 0
Nv.centroCostoCodigo = "NV NO ENCONTRADA => no cc"
sql = "SELECT * FROM dbo.vw_nv WHERE nv = " & Numero
'sql = "SELECT dbo.tb_nv.nv, dbo.tb_nv.negocioNumero, dbo.tb_negocios.descripcion AS negocioDescripcion, dbo.fn_fechaSql2Esp(dbo.tb_nv.fecha) AS fecha,"
'sql = sql & "dbo.tb_negocios.clienteRut, dbo.vw_clientes.razonSocial, dbo.tb_nv.tipo, dbo.tb_nv.obra, dbo.tb_nv.pesoEstimado, dbo.tb_nv.recubrimiento,"
'sql = sql & "dbo.tb_recubrimientos.descripcion AS recubrimientoDescripcion, dbo.fn_fechaSql2Esp(dbo.tb_nv.fechaInicio) AS fechaInicio,"
'sql = sql & "dbo.fn_fechaSql2Esp(dbo.tb_nv.fechaTermino) AS fechaTermino, dbo.tb_nv.activa, dbo.tb_nv.observacion1, dbo.tb_nv.listaPernosIncluida,"
'sql = sql & "dbo.tb_nv.listaPernosRecibida, dbo.tb_negocios.centroCosto AS ccCodigo, dbo.tb_centroCosto.descripcion AS ccDescripcion "
'sql = sql & "FROM dbo.tb_nv "
'sql = sql & "INNER JOIN dbo.tb_recubrimientos ON dbo.tb_nv.recubrimiento = dbo.tb_recubrimientos.codigo "
'sql = sql & "INNER JOIN dbo.tb_negocios ON dbo.tb_nv.negocioNumero = dbo.tb_negocios.numero "
'sql = sql & "INNER JOIN dbo.vw_clientes ON dbo.tb_negocios.clienteRut = dbo.vw_clientes.rut "
'sql = sql & "INNER JOIN dbo.tb_centroCosto ON dbo.tb_negocios.centroCosto = dbo.tb_centroCosto.codigo "
'sql = sql & "WHERE nv = " & Numero

'MsgBox "3"
CnxSqlServer_scp0.Execute (sql)
'MsgBox "4"

With RsPaso
.Open sql, CnxSqlServer_scp0
If Not .EOF Then
    Nv.Numero = !Numero
    Nv.obra = !obra
    Nv.centroCostoCodigo = !ccCodigo
    Nv.centroCostoDescripcion = !ccDescripcion
End If
End With
nv_Buscar = Nv
End Function
