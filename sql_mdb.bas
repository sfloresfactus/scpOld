Attribute VB_Name = "sql_mdb"
Option Explicit
Public Sub Rs_Abrir_MDB(Db As Database, Rs As Recordset, ByVal sql As String)
' abre recordset, pero antes verifica si esta abierto
'If Rs.Name <> "" Then
'    Rs.Close
'End If
Set Rs = Db.OpenRecordset(sql)
End Sub
Public Function Registro_Existe_MDB(Cnx, TablaNombre As String, CampoNombre As String, clave As String, EsTexto As Boolean) As Boolean
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
    Registro_Existe_MDB = False
Else
    Registro_Existe_MDB = True
End If
.Close
End With
End Function
Public Sub Registro_Agregar_MDB(Cnx, TablaNombre As String, Campos, Valores, NumeroCampos As Integer)
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
Public Sub Registro_Modificar_MDB(Cnx, TablaNombre As String, Campos, Valores, VaoNoVa, NumeroCampos As Integer) ', Condicion As String)
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
Public Function Registro_Existe_K2_MDB(Cnx, TablaNombre As String, CampoClave1 As String, clave1 As String, CampoClave2 As String, clave2 As String) As Boolean
' verifica si existe registro con dos campos clave
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
    Registro_Existe_K2_MDB = False
Else
    Registro_Existe_K2_MDB = True
End If
.Close
End With
End Function
Public Sub Registro_Eliminar_MDB(Db As Database, TablaNombre As String, Condition As String)
' elimina registro a BD
' Campos es arreglo ( nombrecampo,comilla )
' Valores es arreglo vienen los valores a eliminar
' la condicion viene dada por los campos VaoNoVA=FALSE
Dim sql As String, nc As Integer, largo As Integer, Cuenta_Condiciones As Integer, Condicion As String
Cuenta_Condiciones = 0

sql = "DELETE " & TablaNombre

For nc = 1 To 1 ' NumeroCampos

'    If VaoNoVa(nc) Then  ' ojo, decia vaonova sin NOT
If True Then
    Else
        ' aqui se arma la condicion
        Cuenta_Condiciones = Cuenta_Condiciones + 1
        If Cuenta_Condiciones > 1 Then
            Condicion = Condicion & " AND "
        End If
'        Condicion = Condicion & Campos(nc, 0) & "=" & Campos(nc, 1) & Valores(nc) & Campos(nc, 1)
    End If
    
Next
' saca ultima coma
largo = Len(sql)
'sql = Left(sql, largo - 1)

sql = sql & " WHERE " & Condicion

'Cnx.Execute sql

End Sub
Public Function sql_Documento_Numero_Nuevo_MDB(Cnx, TablaNombre As String, TipoDoc As String, CampoNombre As String) As Long
' busca nuevo correlativo para documento "TipoDoc" en la tabla "TablaNombre"

Dim RsPaso As New ADODB.Recordset
Dim sql As String, comilla As String

sql_Documento_Numero_Nuevo_MDB = 0

On Error GoTo Sigue

sql = "SELECT MAX(" & CampoNombre & ") AS ultimo FROM " & TablaNombre
RsPaso.Open sql, Cnx

sql_Documento_Numero_Nuevo_MDB = RsPaso!ultimo

Sigue:

On Error GoTo 0
sql_Documento_Numero_Nuevo_MDB = sql_Documento_Numero_Nuevo_MDB + 1

RsPaso.Close

End Function
Public Sub planoMarcaModificar()
Dim Db As Database
Set Db = DBEngine.OpenDatabase(mpro_file)

' 5 tablas afectadas:
' planos detalle
' ot fab detalle
' ito fab detalle
' ito pg detalle
' gd detalle

'db.TableDefs("GD detalle").Fields("plano").Size = 50
'Db.Execute "ALTER TABLE [gd detalle] ALTER COLUMN plano TEXT(50)"
Db.Execute "ALTER TABLE [gd detalle] MODIFY plano TEXT(50)"
Db.Close
'COLUMN tipo de campo[(tamaño)

Set Db = Nothing

End Sub
