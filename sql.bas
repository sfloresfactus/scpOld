Attribute VB_Name = "sql"
Option Explicit
Private Const NvTabla As String = "nv cabecera"
'Private Const NvTabla As String = "nv cabecera2"

Public Type NotaVenta
    Numero As Double
    Fecha As Date
    rutCliente As String
    obra As String
    Tipo As String
    galvanizado As Boolean
    pintura As Boolean
    fechaInicio As Date
    fechaTermino As Date
    ListaPernosIncluida As Boolean
    ListaPernosRecibida As Boolean
    Activa As Boolean
    observacion1 As String
    observacion2 As String
    observacion3 As String
    observacion4 As String
    
    ' incorporado 27/06/2014
    negocioNumero As Integer
    negocioDescripcion As String
    centroCostoCodigo As String
    centroCostoDescripcion As String
    
End Type
' acceso a rutinas de DAO via access o sqlServer
Public Const trabajadorCnx As String = "access"
Public Const contratistaCnx As String = "access"
Public Const clienteCnx As String = "access"
Public Sub Rs_Abrir(Rs As ADODB.Recordset, ByVal sql As String)
' abre recordset, pero antes verifica si esta abierto
'Debug.Print "1|" & Rs.State & "|"
'If Rs.Source <> "" Then Rs.Close
If Rs.State = 1 Then
    ' abierto ?
    Rs.Close
End If
Rs.Open sql, CnxSqlServer_scp0
'Debug.Print "|" & Rs.ActiveConnection & "|"
'Debug.Print "2|" & Rs.State & "|"
End Sub
Public Function Registro_Existe(TablaNombre As String, Condicion As String) As Boolean
' verifica si existe registro, con un solo campo clave
Dim RsPaso As New ADODB.Recordset
Dim sql As String
sql = "SELECT * FROM " & TablaNombre
sql = sql & " WHERE " & Condicion
'RsPaso.Open sql, CnxSqlServer
With RsPaso
.Open sql, CnxSqlServer_scp0
If .EOF Then
    Registro_Existe = False
Else
    Registro_Existe = True
End If
.Close
End With
End Function
Public Function Registro_Existe_OLD(Cnx, TablaNombre As String, CampoNombre As String, clave As String, EsTexto As Boolean) As Boolean
' verifica si existe registro, con un solo campo clave
Dim RsPaso As New ADODB.Recordset
Dim sql As String, comilla As String
If EsTexto Then
    comilla = "'"
Else
    comilla = ""
End If
sql = "SELECT * FROM " & TablaNombre
sql = sql & " WHERE " & CampoNombre & " = " & comilla & clave & comilla
'RsPaso.Open sql, CnxSqlServer
With RsPaso
.Open sql, Cnx
If .EOF Then
    Registro_Existe_OLD = False
Else
    Registro_Existe_OLD = True
End If
.Close
End With
End Function
Public Sub Registro_Agregar(Cnx, TablaNombre As String, Campos, Valores, NumeroCampos As Integer)
' agrega registro a BD
' Campos, Valores y Tipos son arreglos ( matrices )
Dim sql As String, nc As Integer, largo As Integer

sql = "INSERT INTO " & TablaNombre & " ("

For nc = 1 To NumeroCampos
    sql = sql & Campos(nc, 0) & "," ' coma(nc, NumeroCampos)
Next
' saca ultima coma
largo = Len(sql)
sql = Left(sql, largo - 1)

sql = sql & ") VALUES ("

For nc = 1 To NumeroCampos
'    sql = sql & Tipos(nc) & Valores(nc) & Tipos(nc) & "," 'coma(nc, NumeroCampos)
    sql = sql & Campos(nc, 1) & Valores(nc) & Campos(nc, 1) & ","
Next
' saca ultima coma
largo = Len(sql)
sql = Left(sql, largo - 1)

sql = sql & ")"

Cnx.Execute sql

Debug.Print sql

End Sub
Public Sub Registro_Modificar(Cnx, TablaNombre As String, Campos, Valores, VaoNoVa, NumeroCampos As Integer) ', Condicion As String)
' modifica registro a BD
' Campos es arreglo ( nombrecampo,comilla )
' Valores es arreglo vienen los valores a modifica-grabar
' VaoNoVa arreglo de valores booleanos, true significa que el campo se graba, false significa que el campo es la condicion para modificar
Dim sql As String, nc As Integer, largo As Integer, Cuenta_Condiciones As Integer, Condicion As String
Condicion = ""
Cuenta_Condiciones = 0

'If Condicion = "" Then
'End If

sql = "UPDATE " & TablaNombre & " SET "

For nc = 1 To NumeroCampos
    If VaoNoVa(nc) Then
        sql = sql & Campos(nc, 0) & "=" & Campos(nc, 1) & Valores(nc) & Campos(nc, 1) & ","
    Else
        ' aqui se arma la condicion
        Cuenta_Condiciones = Cuenta_Condiciones + 1
        If Cuenta_Condiciones > 1 Then
            Condicion = Condicion & " AND "
        End If
        Condicion = Condicion & Campos(nc, 0) & "=" & Campos(nc, 1) & Valores(nc) & Campos(nc, 1)
    End If
Next
' saca ultima coma
largo = Len(sql)
sql = Left(sql, largo - 1)

sql = sql & " WHERE " & Condicion

Cnx.Execute sql

'Debug.Print sql

End Sub
Public Sub Arreglo_Limpiar(arreglo, NumerodeCampos As Integer, Valor)
' limpia arreglo
' valor indica el valor que queda en arreglo
Dim nc As Integer
For nc = 1 To NumerodeCampos
    arreglo(nc) = Valor
Next
End Sub
Public Function Registro_Existe_K2_OLD(Cnx, TablaNombre As String, CampoClave1 As String, clave1 As String, CampoClave2 As String, clave2 As String) As Boolean
' verifica si existe registro con 2 campos clave
Dim RsPaso As New ADODB.Recordset
Dim sql As String, comilla As String
comilla = "'"
'If EsTexto Then
'    comilla = "'"
'Else
'    comilla = ""
'End If
sql = "SELECT * FROM " & TablaNombre
sql = sql & " WHERE " & CampoClave1 & " = " & comilla & clave1 & comilla
sql = sql & " AND " & CampoClave2 & " = " & comilla & clave2 & comilla

'RsPaso.Open sql, CnxSqlServer
With RsPaso
.Open sql, Cnx
If .EOF Then
    Registro_Existe_K2_OLD = False
Else
    Registro_Existe_K2_OLD = True
End If
.Close
End With

End Function
Public Function Registro_Existe_Kn_OLD(ByVal Cnx As String, TablaNombre As String, Claves, Valores) As Boolean
' verifica si existe registro con n campos clave
Dim RsPaso As New ADODB.Recordset
Dim sql As String, comilla As String, i As Integer
comilla = "'"

'Debug.Print UBound(Claves)

For i = 0 To UBound(Claves)
    Debug.Print Claves(i)
Next

'If EsTexto Then
'    comilla = "'"
'Else
'    comilla = ""
'End If
sql = "SELECT * FROM " & TablaNombre
'sql = sql & " WHERE " & CampoClave1 & " = " & comilla & clave1 & comilla
'sql = sql & " AND " & CampoClave2 & " = " & comilla & clave2 & comilla

'RsPaso.Open sql, CnxSqlServer
With RsPaso
.Open sql, Cnx
If .EOF Then
    Registro_Existe_Kn_OLD = False
Else
    Registro_Existe_Kn_OLD = True
End If
.Close
End With

End Function
'Public Sub Registro_Eliminar(Cnx, TablaNombre As String, Campos, Valores, VaoNoVa, NumeroCampos As Integer)
Public Sub Registro_Eliminar(Cnx, TablaNombre As String, Numero As Integer)
' elimina registro a BD
' Campos es arreglo ( nombrecampo,comilla )
' Valores es arreglo vienen los valores a eliminar
' la condicion viene dada por los campos VaoNoVA=FALSE
Dim sql As String, nc As Integer, largo As Integer, Cuenta_Condiciones As Integer, Condicion As String
Cuenta_Condiciones = 0

sql = "DELETE " & TablaNombre & " WHERE e_numero=" & Numero

Cnx.Execute sql

End Sub
Public Function sql_Documento_Numero_Nuevo(Cnx, TablaNombre As String, TipoDoc As String, CampoNombre As String) As Long
' busca nuevo correlativo para documento "TipoDoc" en la tabla "TablaNombre"

Dim RsPaso As New ADODB.Recordset
Dim sql As String, comilla As String

sql_Documento_Numero_Nuevo = 0

On Error GoTo Sigue

sql = "SELECT MAX(" & CampoNombre & ") AS ultimo FROM " & TablaNombre
RsPaso.Open sql, Cnx

sql_Documento_Numero_Nuevo = RsPaso!ultimo

Sigue:

On Error GoTo 0
sql_Documento_Numero_Nuevo = sql_Documento_Numero_Nuevo + 1

RsPaso.Close

End Function
Public Function sql_rut_trim() As String
' funcion trasitoria
' hace trim a rut
End Function
Public Function SqlRutPadL(ByVal rut As String) As String
' 14/06/10
' funcion transitoria
' hace pad left a rut
SqlRutPadL = PadL(rut, 10)
End Function
Public Sub nvListar(nvActivas As Boolean)

' lee "todas" las nv y las trae a arreglo en memoria

' nvActivas indica si quiero solo las nv activas
' true: solo activas
' false: todas
Dim sql
nvTotal = 0

' con access
If NvCnx = "Access" Then

    Dim Dbm As Database
    Dim RsNVc As Recordset

    If nvActivas Then
        sql = "SELECT * FROM [" & NvTabla & "] WHERE activa ORDER BY numero"
    Else
        sql = "SELECT * FROM [" & NvTabla & "] ORDER BY numero"
    End If

    Set Dbm = OpenDatabase(mpro_file)
    Set RsNVc = Dbm.OpenRecordset(sql)
    
    With RsNVc
    Do While Not .EOF
        nvTotal = nvTotal + 1
        aNv(nvTotal).Numero = !Numero
        aNv(nvTotal).obra = !obra
        .MoveNext
    Loop
    End With
    
    RsNVc.Close
    Dbm.Close
    
End If

If NvCnx = "Sql" Then
    Dim RsPaso As New ADODB.Recordset
    With RsPaso
    If nvActivas Then
        sql = "SELECT * FROM nv WHERE activa ORDER BY nv"
    Else
        sql = "SELECT * FROM nv ORDER BY nv"
    End If
    .Open sql, CnxSqlServer_scp0
    Do While Not .EOF
        nvTotal = nvTotal + 1
        aNv(nvTotal).Numero = !Nv
        aNv(nvTotal).obra = !obra
        .MoveNext
    Loop
    .Close
    End With
End If

End Sub
Public Function nvLeer(Nv As Double) As NotaVenta
' lee una sola NV

nvLeer.Numero = 0
Dim sql As String

If NvCnx = "Access" Then
    
    Dim Dbm As Database
    Dim RsNVc As Recordset
    
    Set Dbm = OpenDatabase(mpro_file)
    
    sql = "SELECT * FROM [" & NvTabla & "] WHERE numero=" & Nv
    Set RsNVc = Dbm.OpenRecordset(sql)

    With RsNVc
    If RsNVc.RecordCount > 0 Then
        nvLeer.Numero = Nv
        nvLeer.Fecha = !Fecha
        nvLeer.rutCliente = ![RUT CLiente]
        nvLeer.obra = !obra
        nvLeer.Tipo = NoNulo(!Tipo)
        nvLeer.galvanizado = !galvanizado
        nvLeer.pintura = !pintura
        nvLeer.fechaInicio = ![fecha inicio]
        nvLeer.fechaTermino = ![fecha termino]
        nvLeer.ListaPernosIncluida = ![Lista Pernos Incluida]
        nvLeer.ListaPernosRecibida = ![Lista Pernos Recibida]
        nvLeer.Activa = !Activa
        nvLeer.observacion1 = NoNulo(![Observacion 1])
        nvLeer.observacion2 = NoNulo(![Observacion 2])
        nvLeer.observacion3 = NoNulo(![Observacion 3])
        nvLeer.observacion4 = NoNulo(![Observacion 4])
    End If
    End With

End If

If NvCnx = "Sql" Then

    Dim RsPaso As New ADODB.Recordset
'    set RsPaso=
    sql = "SELECT * FROM nv WHERE nv=" & Nv
    With RsPaso
    .Open sql, CnxSqlServer_scp0
    If Not .EOF Then
        nvLeer.Numero = Nv
        nvLeer.Fecha = !Fecha
        nvLeer.rutCliente = ![Rut_Cliente]
        nvLeer.obra = !obra
        nvLeer.Tipo = !Tipo
        nvLeer.galvanizado = IIf(!galvanizado = "S", True, False)
        nvLeer.pintura = IIf(!pintura = "S", True, False)
        nvLeer.fechaInicio = ![Fecha_Inicio]
        nvLeer.fechaTermino = ![Fecha_Termino]
        nvLeer.ListaPernosIncluida = IIf(![Lista_Pernos_Incluida] = "S", True, False)
        nvLeer.ListaPernosRecibida = IIf(![Lista_Pernos_Recibida] = "S", True, False)
        nvLeer.Activa = IIf(!Activa = "S", True, False)
        nvLeer.observacion1 = ![observacion1]
        nvLeer.observacion2 = NoNulo(![observacion2])
        nvLeer.observacion3 = NoNulo(![observacion3])
        nvLeer.observacion4 = NoNulo(![observacion4])
    
    End If
    .Close
    End With

End If

End Function
Public Sub nvGrabar(Nv As NotaVenta, Nueva As Boolean)
' graba nv
Dim sql

If NvCnx = "Access" Then
    
    Dim Dbm As Database
    
    Set Dbm = OpenDatabase(mpro_file)

    If Nueva Then
        
        sql = "INSERT INTO [" & NvTabla & "] ("
        sql = sql & "numero,"          ' 1
        sql = sql & "fecha,"           ' 2
        sql = sql & "[rut cliente],"   ' 3
        sql = sql & "obra,"            ' 4
        sql = sql & "tipo,"            ' 5
        sql = sql & "galvanizado,"     ' 6
        sql = sql & "pintura,"         ' 7
        sql = sql & "[fecha inicio],"  ' 8
        sql = sql & "[fecha termino],"   ' 9
        sql = sql & "[lista pernos incluida]," ' 10
        sql = sql & "[lista pernos recibida]," ' 11
        sql = sql & "activa,"          ' 12
        sql = sql & "[observacion 1]," ' 13
        sql = sql & "[observacion 2]," ' 14
        sql = sql & "[observacion 3]," ' 15
        sql = sql & "[observacion 4]"  ' 16
        sql = sql & ") VALUES ("
        sql = sql & Nv.Numero & ","              ' 1
        sql = sql & "'" & Nv.Fecha & "',"        ' 2
        sql = sql & "'" & Nv.rutCliente & "',"   ' 3
        sql = sql & "'" & Nv.obra & "',"         ' 4
        sql = sql & "'" & Nv.Tipo & "',"         ' 5
        sql = sql & IIf(Nv.galvanizado, 1, 0) & "," ' 6
        sql = sql & IIf(Nv.pintura, 1, 0) & ","  ' 7
        sql = sql & "'" & Nv.fechaInicio & "',"  ' 8
        sql = sql & "'" & Nv.fechaTermino & "'," ' 9
        sql = sql & IIf(Nv.ListaPernosIncluida, 1, 0) & "," ' 10
        sql = sql & IIf(Nv.ListaPernosRecibida, 1, 0) & "," ' 11
        sql = sql & IIf(Nv.Activa, 1, 0) & ","   ' 12
        sql = sql & "'" & Nv.observacion1 & "'," ' 13
        sql = sql & "'" & Nv.observacion2 & "'," ' 14
        sql = sql & "'" & Nv.observacion3 & "'," ' 15
        sql = sql & "'" & Nv.observacion4 & "'"  ' 16
        sql = sql & ")"

    Else
        
        sql = "UPDATE [" & NvTabla & "] SET "
        sql = sql & " fecha='" & Nv.Fecha & "',"
        sql = sql & " [rut cliente]='" & Nv.rutCliente & "',"
        sql = sql & " obra='" & Nv.obra & "',"
        sql = sql & " tipo='" & Nv.Tipo & "',"
        sql = sql & " galvanizado=" & IIf(Nv.galvanizado, 1, 0) & ","
        sql = sql & " pintura=" & IIf(Nv.pintura, 1, 0) & ","
        sql = sql & " [fecha inicio]='" & Nv.fechaInicio & "',"
        sql = sql & " [fecha termino]='" & Nv.fechaTermino & "',"
        sql = sql & " [lista pernos incluida]=" & IIf(Nv.ListaPernosIncluida, 1, 0) & ","
        sql = sql & " [lista pernos recibida]=" & IIf(Nv.ListaPernosRecibida, 1, 0) & ","
        sql = sql & " activa=" & IIf(Nv.Activa, 1, 0) & ","
        sql = sql & " [observacion 1]='" & Nv.observacion1 & "',"
        sql = sql & " [observacion 2]='" & Nv.observacion2 & "',"
        sql = sql & " [observacion 3]='" & Nv.observacion3 & "',"
        sql = sql & " [observacion 4]='" & Nv.observacion4 & "'"
        sql = sql & " WHERE numero=" & Nv.Numero
    End If
            
    Dbm.Execute sql

End If

If NvCnx = "Sql" Then
    
    If Nueva Then
        
        sql = "INSERT INTO nv ("
        sql = sql & "nv,"          ' 1
        sql = sql & "fecha,"           ' 2
        sql = sql & "[rut_cliente],"   ' 3
        sql = sql & "obra,"            ' 4
        sql = sql & "tipo,"            ' 5
        sql = sql & "galvanizado,"     ' 6
        sql = sql & "pintura,"         ' 7
        sql = sql & "[fecha_inicio],"  ' 8
        sql = sql & "[fecha_termino],"   ' 9
        sql = sql & "[lista_pernos_incluida]," ' 10
        sql = sql & "[lista_pernos_recibida]," ' 11
        sql = sql & "activa,"          ' 12
        sql = sql & "[observacion1]," ' 13
        sql = sql & "[observacion2]," ' 14
        sql = sql & "[observacion3]," ' 15
        sql = sql & "[observacion4]"  ' 16
        sql = sql & ") VALUES ("
        sql = sql & Nv.Numero & ","              ' 1
        sql = sql & "'" & Nv.Fecha & "',"        ' 2
        sql = sql & "'" & Nv.rutCliente & "',"   ' 3
        sql = sql & "'" & Nv.obra & "',"         ' 4
        sql = sql & "'" & Nv.Tipo & "',"         ' 5
        sql = sql & IIf(Nv.galvanizado, 1, 0) & "," ' 6
        sql = sql & IIf(Nv.pintura, 1, 0) & ","  ' 7
        sql = sql & "'" & Nv.fechaInicio & "',"  ' 8
        sql = sql & "'" & Nv.fechaTermino & "'," ' 9
        sql = sql & IIf(Nv.ListaPernosIncluida, 1, 0) & "," ' 10
        sql = sql & IIf(Nv.ListaPernosRecibida, 1, 0) & "," ' 11
        sql = sql & IIf(Nv.Activa, 1, 0) & ","   ' 12
        sql = sql & "'" & Nv.observacion1 & "'," ' 13
        sql = sql & "'" & Nv.observacion2 & "'," ' 14
        sql = sql & "'" & Nv.observacion3 & "'," ' 15
        sql = sql & "'" & Nv.observacion4 & "'"  ' 16
        sql = sql & ")"

    Else
        sql = "UPDATE nv SET "
        sql = sql & " fecha='" & Nv.Fecha & "',"
        sql = sql & " [rut_cliente]='" & Nv.rutCliente & "',"
        sql = sql & " obra='" & Nv.obra & "',"
        sql = sql & " tipo='" & Nv.Tipo & "',"
        sql = sql & " galvanizado=" & IIf(Nv.galvanizado, 1, 0) & ","
        sql = sql & " pintura=" & IIf(Nv.pintura, 1, 0) & ","
        sql = sql & " [fecha_inicio]='" & Nv.fechaInicio & "',"
        sql = sql & " [fecha_termino]='" & Nv.fechaTermino & "',"
        sql = sql & " [lista_pernos_incluida]=" & IIf(Nv.ListaPernosIncluida, 1, 0) & ","
        sql = sql & " [lista_pernos_recibida]=" & IIf(Nv.ListaPernosRecibida, 1, 0) & ","
        sql = sql & " activa=" & IIf(Nv.Activa, 1, 0) & ","
        sql = sql & " [observacion1]='" & Nv.observacion1 & "',"
        sql = sql & " [observacion2]='" & Nv.observacion2 & "',"
        sql = sql & " [observacion3]='" & Nv.observacion3 & "',"
        sql = sql & " [observacion4]='" & Nv.observacion4 & "'"
        sql = sql & " WHERE nv=" & Nv.Numero
    End If
           
    CnxSqlServer_scp0.Execute sql

End If

End Sub
Public Function nv2Obra(Nv As Double) As String
' busca NV, devuelve campo "obra"
Dim obra As String
obra = ""

If NvCnx = "Access" Then
    
    Dim Dbm As Database
    Dim RsNVc As Recordset
    Dim sql
    
    Set Dbm = OpenDatabase(mpro_file)
    
    sql = "SELECT * FROM [" & NvTabla & "] WHERE numero=" & Nv
    Set RsNVc = Dbm.OpenRecordset(sql)

    If RsNVc.RecordCount > 0 Then
        obra = RsNVc!obra
    End If

End If

nv2Obra = obra

End Function
Public Function nvNueva() As Double
' busca siguiente nv libre
Dim Nv As Double
Nv = 0
Dim sql

If NvCnx = "Access" Then
    
    Dim Dbm As Database
    Dim RsNVc As Recordset
    
    Set Dbm = OpenDatabase(mpro_file)
    
    sql = "SELECT MAX(numero) as ultimo FROM [" & NvTabla & "]"
    Set RsNVc = Dbm.OpenRecordset(sql)

    If RsNVc.RecordCount > 0 Then
        Nv = RsNVc!ultimo
    End If

End If

If NvCnx = "Sql" Then
    
    Dim RsPaso As New ADODB.Recordset
    sql = "SELECT MAX(nv) as ultimo FROM nv"
    With RsPaso
    .Open sql, CnxSqlServer_scp0
    If Not .EOF Then
        Nv = !ultimo
    End If
    .Close
    End With

End If

nvNueva = Nv + 1

End Function
Public Sub nvEliminar(Nv As Double)

' elimina NV
Dim sql As String
If NvCnx = "Access" Then
    
    Dim Dbm As Database
    Set Dbm = OpenDatabase(mpro_file)
    
    sql = "DELETE FROM [" & NvTabla & "] WHERE numero=" & Nv
    Dbm.Execute sql

End If

If NvCnx = "Sql" Then
    sql = "DELETE FROM nv WHERE nv=" & Nv
    CnxSqlServer_scp0.Execute sql
End If

End Sub
Public Function Contratista_Lee(SqlRsSc As ADODB.Recordset, ByVal rut) As String
Dim Razon As String
rut = Trim(rut)
SqlRsSc.Open "SELECT * FROM personas WHERE contratista='S' AND rut='" & rut & "'", CnxSqlServer_scp0
If SqlRsSc.EOF Then
'    Razon = "NO Encontrado"
    Razon = ""
Else
    Razon = SqlRsSc![razon_social]
End If
SqlRsSc.Close
Contratista_Lee = Razon
End Function
Public Function Acceso_Lee(SqlRsAcceso As ADODB.Recordset, ByVal Documento) As String
Dim Codigo As String
SqlRsAcceso.Open "SELECT * FROM acceso WHERE documento='" & Documento & "'", CnxSqlServer_scp0
If SqlRsAcceso.EOF Then
    Codigo = "NO Encontrado"
Else
    Codigo = SqlRsAcceso![clave]
End If
Acceso_Lee = Codigo
SqlRsAcceso.Close
End Function
Public Sub Planos_Revisiones_Grabar(ByVal Nv As Double, ByVal Plano As String, _
 ByVal Revision As String, ByVal Marca As String, ByVal Cantidad As Integer, ByVal Descripcion As String, _
 ByVal Peso As Double, ByVal Superficie As Double)
' para KARINA HENRIQUEZ
' graba revisiones en archivo historico de revisiones
'tabla scp0.planos_detalle_revisiones
Dim m_Tabla As String, sql As String
Dim RsPaso As New ADODB.Recordset

Dim s_PUNI As String, s_SUNI As String

Dim email As String

m_Tabla = "planos_detalle_revisiones"

' primero veo si existe revision
sql = "SELECT * FROM " & m_Tabla
sql = sql & " WHERE nv=" & Nv
sql = sql & " AND plano='" & Plano & "'"
sql = sql & " AND marca='" & Marca & "'"
sql = sql & " AND rev='" & Revision & "'"

RsPaso.Open sql, CnxSqlServer_scp0
If RsPaso.EOF Then

    s_PUNI = Replace(str(Peso), ",", ".")
    s_SUNI = Replace(str(Superficie), ",", ".")
    
'    Open Path_Local & "listaemails.txt" For Input As #2
'    Do While Not EOF(2)
'        Line Input #2, email
'        Email_Generar email, CbNv.Text, Plano, Revision, Marca, Descripcion, Trim(Str(Cantidad)), s_PUNI
'    Loop
'    Close #2

    sql = "INSERT INTO " & m_Tabla & " ("
    sql = sql & " nv,"          '  1
    sql = sql & " nvarea,"      '  2
    sql = sql & " plano,"       '  3
    sql = sql & " rev,"         '  4
    sql = sql & " cantidad_total," ' 5
    sql = sql & " marca,"       '  6
    sql = sql & " descripcion," '  7
    sql = sql & " peso,"        '  8
    sql = sql & " superficie,"  '  9
    sql = sql & " fecha"        ' 10
    sql = sql & ") VALUES ("
    sql = sql & "" & Nv & ","     '  1
    sql = sql & "" & "0" & ","         '  2
    sql = sql & "'" & Plano & "',"     '  3
    sql = sql & "'" & Revision & "',"  '  4
    sql = sql & "" & Cantidad & ","        '  5
    sql = sql & "'" & Marca & "',"     '  6
    sql = sql & "'" & Descripcion & "'," '7
    sql = sql & "" & s_PUNI & ","        '  8
    sql = sql & "" & s_SUNI & ","  '  9
    sql = sql & "'" & Format(Date, "yyyy-mm-dd") & "')" ' 10

    CnxSqlServer_scp0.Execute sql
    
End If

End Sub
Public Function Plano_Borrable(RsPd As Recordset, Nv As Integer, NvArea As Integer, Plano As String) As Boolean
Plano_Borrable = True
RsPd.Seek ">=", Nv, NvArea, Plano, 0
If Not RsPd.NoMatch Then

    Do While Not RsPd.EOF
    
        If RsPd!Nv <> Nv Or RsPd!Plano <> Plano Then Exit Do
        
        If RsPd![OT fab] <> 0 Then
            Plano_Borrable = False
            Exit Function
        End If
        
        RsPd.MoveNext
        
    Loop
End If
End Function
Public Sub Plano_Eliminar(Dbm As Database, Nv, Plano, Marca)
' elimina plano
Dim qry As String, condi As String

condi = " WHERE NV=" & Nv & " AND Plano='" & Plano & "'" & " AND Marca='" & Marca & "'"
qry = "DELETE * FROM [Planos Detalle]"
qry = qry & condi
Dbm.Execute qry

condi = " WHERE NV=" & Nv & " AND Plano='" & Plano & "'"
qry = "DELETE * FROM [Planos Cabecera]"
qry = qry & condi
Dbm.Execute qry

End Sub
Public Sub Marcas_Agregar(a_Marcas, m_MarcaIndice As Integer, Plano As String, Marca As String)
Dim i As Integer
' busca si existe marca
For i = 1 To m_MarcaIndice
    If a_Marcas(i, 0) = Plano And a_Marcas(i, 1) = Marca Then
        Exit Sub
    End If
Next

m_MarcaIndice = m_MarcaIndice + 1

a_Marcas(m_MarcaIndice, 0) = Plano
a_Marcas(m_MarcaIndice, 1) = Marca

End Sub
Public Function fecha2aaaammdd(Fecha As String) As String
' transforma fecha de formato dd/mm/aa a aaaammdd
fecha2aaaammdd = "20" & Mid(Fecha, 7, 2) & Mid(Fecha, 4, 2) & Left(Fecha, 2)
End Function
Public Function trabajadorBuscarNombreCompleto(rut As String) As String
'  devuelve un string con el nombre completo del trabajador
trabajadorBuscarNombreCompleto = ""
If Len(rut) = 0 Then Exit Function
If trabajadorCnx = "access" Then
    Dim DbD As Database, Rs As Recordset
    Set DbD = OpenDatabase(data_file)
    Set Rs = DbD.OpenRecordset("Trabajadores")
    With Rs
    .Index = "RUT"
    .Seek "=", PadL(rut, 10)
    If .NoMatch Then
        trabajadorBuscarNombreCompleto = "NO ENCONTRADO"
    Else
        trabajadorBuscarNombreCompleto = !appaterno & " " & !apmaterno & " " & !nombres
    End If
    .Close
    End With
    DbD.Close
End If
End Function
Public Sub trabajadorBuscarDatos(rut As String, arreglo As Variant)
'  devuelve arreglo con TODOS los datos de trabajador
If trabajadorCnx = "access" Then
    Dim DbD As Database, RsTra As Recordset
    Set DbD = OpenDatabase(data_file)
    Set RsTra = DbD.OpenRecordset("Trabajadores")
    With RsTra
    .Index = "RUT"
    .Seek "=", PadL(rut, 10)
    If .NoMatch Then
        arreglo(0) = "NO ENCONTRADO"
    Else
        arreglo(0) = !rut
        arreglo(1) = !Sexo
        arreglo(2) = !appaterno
        arreglo(3) = !apmaterno
        arreglo(4) = !nombres
    End If
    .Close
    End With
    DbD.Close
End If
End Sub
Public Function contratistaBuscarRazon(rut As String) As String
' devuelve un string con la razon social del cliente
contratistaBuscarRazon = ""
If Len(rut) = 0 Then Exit Function
If trabajadorCnx = "access" Then
    Dim DbD As Database, Rs As Recordset
    Set DbD = OpenDatabase(data_file)
    Set Rs = DbD.OpenRecordset("contratistas")
    With Rs
    .Index = "RUT"
    .Seek "=", PadL(rut, 10)
    If .NoMatch Then
        contratistaBuscarRazon = "NO ENCONTRADO"
    Else
        contratistaBuscarRazon = ![Razon Social]
    End If
    .Close
    End With
    DbD.Close
End If
End Function
Public Function clienteBuscarRazon(rut As String) As String
' devuelve un string con la razon social del cliente
clienteBuscarRazon = ""
If Len(rut) = 0 Then Exit Function
If trabajadorCnx = "access" Then
    Dim DbD As Database, Rs As Recordset
    Set DbD = OpenDatabase(data_file)
    Set Rs = DbD.OpenRecordset("clientes")
    With Rs
    .Index = "RUT"
    .Seek "=", PadL(rut, 10)
    If .NoMatch Then
        clienteBuscarRazon = "NO ENCONTRADO"
    Else
        clienteBuscarRazon = ![Razon Social]
    End If
    .Close
    End With
    DbD.Close
End If
End Function
Public Function cuentasContablesCargar(aCuCo) As Integer

' lee "todas" las cuents contables y las trae a arreglo en memoria

Dim sql, nr As String, i As String
cuentasContablesCargar = 0

Dim RsPaso As New ADODB.Recordset
With RsPaso
sql = "SELECT * FROM maestros WHERE tipo='CUCO' ORDER BY codigo"
.Open sql, CnxSqlServer_scp0
i = 0
Do While Not .EOF
    i = i + 1
    aCuCo(i, 0) = !Codigo
    aCuCo(i, 1) = !Descripcion
    aCuCo(i, 2) = !dato1 ' imputable S o N
    aCuCo(i, 3) = !orden
    .MoveNext
Loop
.Close
End With

cuentasContablesCargar = i

End Function
Public Function centroCostoBuscarIndice(Codigo As String)
Dim i As Integer
Codigo = NoNulo(Codigo)
centroCostoBuscarIndice = 0
If Codigo = "" Then Exit Function
For i = 1 To 999
    If Codigo = aCeCo(i, 0) Then
        centroCostoBuscarIndice = i
        Exit For
    End If
Next
End Function
Public Function centroCostoBuscarDescripcion(Codigo As String)
Dim i As Integer
Codigo = NoNulo(Codigo)
centroCostoBuscarDescripcion = ""
If Codigo = "" Then Exit Function
For i = 1 To 999
    If Codigo = aCeCo(i, 0) Then
        centroCostoBuscarDescripcion = aCeCo(i, 1)
        Exit For
    End If
Next
End Function
Public Function centrosCostoCargar(aCeCo) As Integer

' lee todos los centros de costo y los trae a arreglo en memoria

Dim sql, nr As String, i As String
centrosCostoCargar = 0

Dim RsPaso As New ADODB.Recordset
With RsPaso
sql = "SELECT * FROM maestros WHERE tipo='CECO' ORDER BY orden"
.Open sql, CnxSqlServer_scp0
i = 0
Do While Not .EOF
    i = i + 1
    aCeCo(i, 0) = !Codigo
    aCeCo(i, 1) = !Descripcion
    aCeCo(i, 2) = !dato1 ' imputable S o N
    aCeCo(i, 3) = !orden
    .MoveNext
Loop
.Close
End With

centrosCostoCargar = i

End Function
Public Function cuentaContableBuscarDescripcion(Codigo As String)
Dim i As Integer
Codigo = NoNulo(Codigo)
cuentaContableBuscarDescripcion = ""
If Codigo = "" Then Exit Function
For i = 1 To 999
    If Codigo = aCuCo(i, 0) Then
        cuentaContableBuscarDescripcion = aCuCo(i, 1)
        Exit For
    End If
Next
End Function
Public Function cuentaContableBuscarIndice(Codigo As String)
Dim i As Integer
Codigo = NoNulo(Codigo)
cuentaContableBuscarIndice = 0
If Codigo = "" Then Exit Function
For i = 1 To 999
    If Codigo = aCuCo(i, 0) Then
        cuentaContableBuscarIndice = i
        Exit For
    End If
Next
End Function
