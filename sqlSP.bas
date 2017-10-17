Attribute VB_Name = "sqlSP"
Option Explicit
Public Sub procedimientoAlmacenado1()

' sin parametros ok

Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim cmd As ADODB.Command
Dim str_empid, strconnect As String

strconnect = "Provider=SQLNCLI;Data Source=CONTROLSERVER\BDMS;Initial Catalog=delgado1303"
con.Open strconnect, "scp_is", "aqmdla"

str_empid = ""

Set cmd = New ADODB.Command
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "pa_test"

'cmd.Parameters.Append cmd.CreateParameter("pa_testa", adVarChar, adParamInput, 6, str_empid)

Set Rs = cmd.Execute

If Not Rs.EOF Then
    'txt_firstname = rs.Fields(0)
    'txt_title = rs.Fields(1)
    'txt_address = rs.Fields(2)
End If

Set cmd.ActiveConnection = Nothing

End Sub
Public Sub procedimientoAlmacenado2()

' con parametro 2 de entrada y uno se salida

Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim cmd As ADODB.Command
Dim strconnect As String

strconnect = "Provider=SQLNCLI;Data Source=CONTROLSERVER\BDMS;Initial Catalog=delgado1303"
con.Open strconnect, "scp_is", "aqmdla"

Set cmd = New ADODB.Command
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "pa_suma"

cmd.Parameters.Append cmd.CreateParameter("n1", adInteger, adParamInput, , 11)
cmd.Parameters.Append cmd.CreateParameter("n2", adInteger, adParamInput, , 22)
cmd.Parameters.Append cmd.CreateParameter("resultado", adInteger, adParamOutput, 10)

cmd.Execute

Debug.Print cmd.Parameters(2).Value

Set cmd.ActiveConnection = Nothing

End Sub
Public Sub procedimientoAlmacenado3()

' con parametros 1 de entrada y 3 de salida

Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim cmd As ADODB.Command
Dim strconnect As String

strconnect = "Provider=SQLNCLI;Data Source=CONTROLSERVER\BDMS;Initial Catalog=delgado1303"
con.Open strconnect, "scp_is", "aqmdla"

Set cmd = New ADODB.Command
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "pa_nv"

' parametro de entrada
cmd.Parameters.Append cmd.CreateParameter("nv", adInteger, adParamInput, , 1)

' parametros de salida
cmd.Parameters.Append cmd.CreateParameter("obra", adVarChar, adParamOutput, 50)
cmd.Parameters.Append cmd.CreateParameter("fecha", adVarChar, adParamOutput, 10)
cmd.Parameters.Append cmd.CreateParameter("razonSocial", adVarChar, adParamOutput, 50)

cmd.Execute

Debug.Print "|" & cmd.Parameters(1).Value & "|"
Debug.Print "|" & cmd.Parameters(2).Value & "|"
Debug.Print "|" & cmd.Parameters(3).Value & "|"

Set cmd.ActiveConnection = Nothing

End Sub
Public Sub procedimientoAlmacenado4()

' con 4 parametros de entrada

Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim cmd As ADODB.Command
Dim strconnect As String

strconnect = "Provider=SQLNCLI;Data Source=CONTROLSERVER\BDMS;Initial Catalog=scpayd1307"
'strconnect = "Provider=SQLNCLI;Data Source=192.168.0.199\BDMS;Initial Catalog=aydsa1307"
con.Open strconnect, "scp_is", "aqmdla"

Set cmd = New ADODB.Command
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "pa_personaJuridicaAgregar"

cmd.Parameters.Append cmd.CreateParameter("rut", adVarChar, adParamInput, 10, "89784800-7")
cmd.Parameters.Append cmd.CreateParameter("razonSocial", adVarChar, adParamInput, 50, "DELGADO S.A.")
cmd.Parameters.Append cmd.CreateParameter("giro", adVarChar, adParamInput, 50, "Fabricación y Montaje")
cmd.Parameters.Append cmd.CreateParameter("direccion", adVarChar, adParamInput, 50, "Las Acacias 02500")

cmd.Execute

Set cmd.ActiveConnection = Nothing

End Sub
Public Sub procedimientoAlmacenado5()

' con 4 parametros de entrada
' puebla clientes en dos tablas: tb_personaJuridica y tb_clientes

Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim cmd As ADODB.Command
Dim cmd2 As ADODB.Command
Dim strconnect As String

' ////// access /////////////////////////////////////////////////////
Dim DbD As Database, RsCl As Recordset
Dim DbM As Database, RsNv As Recordset
Set DbD = OpenDatabase(data_file)
Set RsCl = DbD.OpenRecordset("clientes")
RsCl.Index = "rut"
Set DbM = OpenDatabase(mpro_file)
Set RsNv = DbM.OpenRecordset("nv cabecera")
'RsNv.Index = "numero"
Dim Rut As String
Dim razonSocial As String
Dim Giro As String
Dim Direccion As String
' ///////////////////////////////////////////////////////////////////

strconnect = "Provider=SQLNCLI;Data Source=CONTROLSERVER\BDMS;Initial Catalog=scpayd1307"
con.Open strconnect, "scp_is", "aqmdla"

Set cmd = New ADODB.Command
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "pa_personaJuridicaAgregar"

cmd.Parameters.Append cmd.CreateParameter("rut", adVarChar, adParamInput, 10, "")
cmd.Parameters.Append cmd.CreateParameter("razonSocial", adVarChar, adParamInput, 50, "")
cmd.Parameters.Append cmd.CreateParameter("giro", adVarChar, adParamInput, 50, "")
cmd.Parameters.Append cmd.CreateParameter("direccion", adVarChar, adParamInput, 50, "")

' para segundo procedimiento
Set cmd2 = New ADODB.Command
cmd2.ActiveConnection = con
cmd2.CommandType = adCmdStoredProc
cmd2.CommandText = "pa_clienteAgregar"
cmd2.Parameters.Append cmd2.CreateParameter("rut", adVarChar, adParamInput, 10, "")
' //////////////////////////

If True Then

    Do While Not RsNv.EOF
    
        ' busca cliente
        Rut = RsNv![RUT CLiente]
        razonSocial = ""
        Giro = ""
        Direccion = ""
        RsCl.Seek "=", Rut
        If Not RsCl.NoMatch Then
            razonSocial = mayuscula2(RsCl![Razon Social])
            Giro = mayuscula2(RsCl!Giro)
            Direccion = mayuscula2(RsCl!Direccion)
        End If
        Rut = Trim(Rut)
        
        If True Then
            Debug.Print razonSocial & "|" & Giro & "|" & Direccion
        End If
        
        cmd.Parameters("rut").Value = Rut
        cmd.Parameters("razonSocial").Value = razonSocial
        cmd.Parameters("giro").Value = Giro
        cmd.Parameters("direccion").Value = Direccion
    
        cmd.Execute
        
        ' y el segundo procedimiento
        cmd2.Parameters("rut").Value = Rut
        cmd2.Execute
                
        RsNv.MoveNext
    
    Loop

End If

Set cmd.ActiveConnection = Nothing
Set cmd2.ActiveConnection = Nothing

End Sub
Public Sub proveedoresAgregar()
' agrega proveedores de acuerdo a OC
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim cmd As ADODB.Command
Dim cmd2 As ADODB.Command
Dim strconnect As String

' ////// access /////////////////////////////////////////////////////
Dim DbD As Database, RsCl As Recordset
Dim DbM As Database, RsNv As Recordset
Set DbD = OpenDatabase(data_file)
Set RsCl = DbD.OpenRecordset("proveedores")
RsCl.Index = "rut"
Set DbM = OpenDatabase(Madq_file)
Set RsNv = DbM.OpenRecordset("oc cabecera")
'RsNv.Index = "numero"
Dim Rut As String
Dim razonSocial As String
Dim Giro As String
Dim Direccion As String
' ///////////////////////////////////////////////////////////////////

strconnect = "Provider=SQLNCLI;Data Source=CONTROLSERVER\BDMS;Initial Catalog=scpayd1307"
con.Open strconnect, "scp_is", "aqmdla"

Set cmd = New ADODB.Command
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "pa_personaJuridicaAgregar"

cmd.Parameters.Append cmd.CreateParameter("rut", adVarChar, adParamInput, 10, "")
cmd.Parameters.Append cmd.CreateParameter("razonSocial", adVarChar, adParamInput, 50, "")
cmd.Parameters.Append cmd.CreateParameter("giro", adVarChar, adParamInput, 50, "")
cmd.Parameters.Append cmd.CreateParameter("direccion", adVarChar, adParamInput, 50, "")

' para segundo procedimiento
Set cmd2 = New ADODB.Command
cmd2.ActiveConnection = con
cmd2.CommandType = adCmdStoredProc
cmd2.CommandText = "pa_proveedorAgregar"
cmd2.Parameters.Append cmd2.CreateParameter("rut", adVarChar, adParamInput, 10, "")
' //////////////////////////

If True Then

    Do While Not RsNv.EOF
    
        ' busca cliente
        Rut = RsNv![RUT Proveedor]
        razonSocial = ""
        Giro = ""
        Direccion = ""
        RsCl.Seek "=", Rut
        If Not RsCl.NoMatch Then
            razonSocial = mayuscula2(RsCl![Razon Social])
            Giro = mayuscula2(RsCl!Giro)
            Direccion = mayuscula2(RsCl!Direccion)
        End If
        Rut = Trim(Rut)
        
        If True Then
            Debug.Print razonSocial & "|" & Giro & "|" & Direccion
        End If
        
        cmd.Parameters("rut").Value = Rut
        cmd.Parameters("razonSocial").Value = razonSocial
        cmd.Parameters("giro").Value = Giro
        cmd.Parameters("direccion").Value = Direccion
    
        cmd.Execute
        
        ' y el segundo procedimiento
        cmd2.Parameters("rut").Value = Rut
        cmd2.Execute
                
        RsNv.MoveNext
    
    Loop

End If

Set cmd.ActiveConnection = Nothing
Set cmd2.ActiveConnection = Nothing

End Sub
Public Sub contratistasAgregar()
' agrega contratistas de acuerdo a OTf

Dim modificar As Boolean
modificar = True

Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim cmd As ADODB.Command
Dim cmd2 As ADODB.Command
Dim strconnect As String

' ////// access /////////////////////////////////////////////////////
Dim DbD As Database, RsRut As Recordset
Dim DbM As Database, RsMov As Recordset
Set DbD = OpenDatabase(data_file)
Set RsRut = DbD.OpenRecordset("contratistas")
RsRut.Index = "rut"
Set DbM = OpenDatabase(mpro_file)
Set RsMov = DbM.OpenRecordset("ot fab cabecera")
Dim Rut As String
Dim razonSocial As String
Dim Giro As String
Dim Direccion As String
Dim Comuna As Integer
' ///////////////////////////////////////////////////////////////////

strconnect = "Provider=SQLNCLI;Data Source=CONTROLSERVER\BDMS;Initial Catalog=scpayd1307"
con.Open strconnect, "scp_is", "aqmdla"

Set cmd = New ADODB.Command
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "pa_personaJuridicaAgregar"
If modificar Then
    cmd.CommandText = "pa_personaJuridicaModificar"
End If

cmd.Parameters.Append cmd.CreateParameter("rut", adVarChar, adParamInput, 10, "")
cmd.Parameters.Append cmd.CreateParameter("razonSocial", adVarChar, adParamInput, 50, "")
cmd.Parameters.Append cmd.CreateParameter("giro", adVarChar, adParamInput, 50, "")
cmd.Parameters.Append cmd.CreateParameter("direccion", adVarChar, adParamInput, 50, "")
If modificar Then
    cmd.Parameters.Append cmd.CreateParameter("comuna", adInteger, adParamInput, , 0)
End If

' para segundo procedimiento
Set cmd2 = New ADODB.Command
cmd2.ActiveConnection = con
cmd2.CommandType = adCmdStoredProc
cmd2.CommandText = "pa_contratistaAgregar"
cmd2.Parameters.Append cmd2.CreateParameter("rut", adVarChar, adParamInput, 10, "")
' //////////////////////////
With RsRut
Do While Not RsMov.EOF

    ' busca cliente
    Rut = RsMov![Rut contratista]
    razonSocial = ""
    Giro = ""
    Direccion = ""
    Comuna = 0
    .Seek "=", Rut
    If Not .NoMatch Then
        razonSocial = mayuscula2(![Razon Social])
        Giro = mayuscula2(!Giro)
        Direccion = mayuscula2(!Direccion)
        Comuna = !orden
    End If
    Rut = Trim(Rut)
    
    If True Then
        Debug.Print razonSocial & "|" & Giro & "|" & Direccion
    End If
    
    cmd.Parameters("rut").Value = Rut
    cmd.Parameters("razonSocial").Value = razonSocial
    cmd.Parameters("giro").Value = Giro
    cmd.Parameters("direccion").Value = Direccion
    If modificar Then
        cmd.Parameters("comuna").Value = Comuna
    End If

    cmd.Execute
    
    ' y el segundo procedimiento
    'cmd2.Parameters("rut").Value = Rut
    'cmd2.Execute
            
    RsMov.MoveNext

Loop
End With

Set cmd.ActiveConnection = Nothing
Set cmd2.ActiveConnection = Nothing

End Sub
Private Function mayuscula1(Texto As String) As String
'Dim cadena As String
mayuscula1 = UCase(Mid(Texto, 1, 1)) + LCase(Mid(Texto, 2))
End Function
Private Function mayuscula2(Texto As String) As String
' transforma un texto todo a minusculas, excepto las iniciales, que las transforma a mayusculas
Dim arreglo(9) As String, i As Integer
split Texto, " ", arreglo, 9
mayuscula2 = ""
For i = 1 To 9
    mayuscula2 = mayuscula2 & mayuscula1(arreglo(i)) & " "
Next
mayuscula2 = Trim(mayuscula2)
End Function
Public Sub SP_nvUpdate(nv As Integer, fecha As String, rutCliente As String, tipo As String, obra As String, recubrimiento As String, fechaInicio As String, fechaTermino As String, listaPernosIncluida As String, listaPernosRecibida As String, activa As String, observacion1 As String, observacion2 As String, observacion3 As String, observacion4 As String)

' para ejecutar procedimiento almacenado en SqlServer
' afecta tablas:
' tb_nv y tb_clientes
' si nv existe=>la actualiza
' si no existe la inserta
' 02/09/13

Dim Provider As String, CatalogScpNew As String, DataSource As String

Provider = ReadIniValue(Path_Local & "scp.ini", "Sql Server", "Provider")
CatalogScpNew = ReadIniValue(Path_Local & "scp.ini", "Sql Server", "Catalog ScpNew")
DataSource = ReadIniValue(Path_Local & "scp.ini", "Sql Server", "Data Source")

Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim cmd As ADODB.Command
Dim strconnect As String

strconnect = "Provider=" & Provider & ";Data Source=" & DataSource & ";Initial Catalog=" & CatalogScpNew
con.Open strconnect, "scp_is", "aqmdla"

Set cmd = New ADODB.Command
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "pa_nvUpdate"

cmd.Parameters.Append cmd.CreateParameter("nv", adInteger, adParamInput, , nv)
cmd.Parameters.Append cmd.CreateParameter("fecha", adVarChar, adParamInput, 10, fecha)
cmd.Parameters.Append cmd.CreateParameter("rutCliente", adVarChar, adParamInput, 10, rutCliente)
cmd.Parameters.Append cmd.CreateParameter("tipo", adVarChar, adParamInput, 2, tipo)
cmd.Parameters.Append cmd.CreateParameter("obra", adVarChar, adParamInput, 30, obra)
cmd.Parameters.Append cmd.CreateParameter("recubrimiento", adVarChar, adParamInput, 1, recubrimiento)
cmd.Parameters.Append cmd.CreateParameter("fechaInicio", adVarChar, adParamInput, 10, fechaInicio)
cmd.Parameters.Append cmd.CreateParameter("fechaTermino", adVarChar, adParamInput, 10, fechaTermino)
cmd.Parameters.Append cmd.CreateParameter("listaPernosIncluida", adVarChar, adParamInput, 1, listaPernosIncluida)
cmd.Parameters.Append cmd.CreateParameter("listaPernosRecibida", adVarChar, adParamInput, 1, listaPernosRecibida)
cmd.Parameters.Append cmd.CreateParameter("activa", adVarChar, adParamInput, 1, activa)
cmd.Parameters.Append cmd.CreateParameter("observacion1", adVarChar, adParamInput, 255, observacion1)
cmd.Parameters.Append cmd.CreateParameter("observacion2", adVarChar, adParamInput, 255, observacion2)
cmd.Parameters.Append cmd.CreateParameter("observacion3", adVarChar, adParamInput, 255, observacion3)
cmd.Parameters.Append cmd.CreateParameter("observacion4", adVarChar, adParamInput, 255, observacion4)

cmd.Execute

'Debug.Print cmd.Parameters(2).Value

Set cmd.ActiveConnection = Nothing

End Sub

