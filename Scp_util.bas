Attribute VB_Name = "Scp_util"
Option Explicit

Public Drive_Local As String
Public Path_Local As String
Public Drive_Server As String
Public Path_Mdb As String, Path_Rpt As String
Public Path_Server As String

Public Syst_file As String
Public data_file As String
Public mpro_file As String ' mov producccion
Public mpro2_file As String ' mov 2
'Public EMLmpro_file As String ' mov producccion de EML
Public Madq_file As String ' mov adquisiciones
Public Hadq_file As String ' hist adquisiciones
Public mvta_file As String ' mov ventas
Public repo_file As String
Public stat_file As String 'status, usuarios conectados
Public scc_file As String

'Public fecha_mask As String, fecha_format As String, fecha_vacia As String
Public num_mask As String, num_Format0 As String
Public num_Formato As String
'Public num_fmtgrl As String

' valores de la empresa
Public Empresa As EnterPrice ' empresa general
Public EmpOC As EnterPrice   ' empresa seleccionada para ordenes de compra

' marcas que estan en plano a modificar o a eliminar
Public Marcas(99, 2) As String
'Public Usr_Name As String, Usr_ReadOnly As Boolean, Usr_Indice As Long, Usr_ObrasTerminadas As Boolean
Public Usuario As Usr

Public Type param
    Iva As Double
End Type

Public Parametro As param
'Public Const Rut_Eml As String = "89784800-7"

Public m_TextoIso As String

Public PrimeraEtiqueta As Boolean

Public Nv_Index As String
Public Const MaximoNumeroFilas As Integer = 40

Private salir As Boolean

Public aCuCo(999, 3) As String ' arreglo de cuentas contables: 0:codigo, 1:descripcion, 2:imputable
Public cuentasContablesTotal As Integer
Public aCeCo(999, 3) As String ' arreglo de centros de costo: 0:codigo, 1:descripcion, 2:imputable 3:orden (porque centros de costos deben orednarse por "orden")
Public centrosCostoTotal As Integer

'//////// parametros para sql server ////////////////////////////////

Public CnxSqlServer_scp0 As New ADODB.Connection
Public CnxSqlServer_delgado1405 As New ADODB.Connection
Public SqlRsSc As New ADODB.Recordset

Public CnxSqlAccess As Connection
'Public Const Usando_ADO As Boolean = True ' indica si estoy usando ADO
Public Const Usando_SQL As Boolean = True
Public sql_Fecha_Formato As String
Public a_Den_s(0, 7) As String
'////////////////////////////////////////////////////////////////////
' variables para sql
Public Tabla_PlanosCabecera As String
Public Tabla_PlanosDetalle As String
Public Win7 As Boolean
' agosto 2013, por nv 3110, takraf, marcas y planos con largo de 27 caracteres
Public Const planoLargo As Integer = 30
Public Const marcaLargo As Integer = 30
'////////////////////////////////////////////////////////////////////
Public Sub densidad_poblar()
' puebla arreglo con densidades por cliente sedgman
a_Den_s(0, 7) = "Super Heavy"
a_Den_s(0, 6) = "Heavy"
a_Den_s(0, 5) = "Medium"
a_Den_s(0, 4) = "Light"
a_Den_s(0, 3) = "Grating ARS 6"
a_Den_s(0, 2) = "Handrails"
a_Den_s(0, 1) = "Stair Treads ARS 6"
End Sub
'////////////////////////////////////////////////////////////////////
Sub Main()
'Dim Dbm As Database, Rs As Recordset
'Dim qry As String
'Set Dbm = OpenDatabase("d:\eml\ejemplo")
'qry = "SELECT Clientes.Nombre AS Razon,Comunas.Nombre AS Comuna FROM Clientes INNER JOIN Comunas ON Clientes.Comuna=Comunas.Codigo"
'Set Rs = Dbm.OpenRecordset(qry)
'Rs.MoveFirst
'Do While Not Rs.EOF
'    Debug.Print Rs!Razon, Rs!Comuna
'    Rs.MoveNext
'Loop

' no quiere funcionar el registro
'App.StartLogging "d:\eml", 2
'Debug.Print "|" & App.LogPath & "|"
'Debug.Print App.LogMode
'App.LogEvent ("ini")

'MsgBox App.Path

If Left(UCase(App.Path), 17) = "\\ACR3006-DUALPRO" Then
    MsgBox "PROGRAMA NO PUEDE SER EJECUTADO DESDE SERVIDOR"
    Exit Sub
End If

If App.PrevInstance Then
    MsgBox "PROGRAMA YA ESTÁ EN EJECUCIÓN"
    Exit Sub
End If

' verifica si existe impresora en el sistema
If Printers.Count = 0 Then
    MsgBox "NO Existen IMPRESORAS instaladas" & vbLf & "Comuniquese con el Administrador del Sistema"
    Exit Sub
End If

Dim Formato As String, formatoFechaCorrecto As Boolean
Formato = Get_Locale(LOCALE_SSHORTDATE)
Formato = LCase(Formato)
'Debug.Print formato
formatoFechaCorrecto = False
Select Case Formato
Case "dd/mm/yy"
'    Debug.Print "ok"
    formatoFechaCorrecto = True
Case "dd/mm/yyyy"
'    Debug.Print "ok"
    formatoFechaCorrecto = True
'Case "dd-mm-yyyy"
'    Debug.Print "dd-mm-yyyy ok"
'    formatoFechaCorrecto = True
Case Else
'    Debug.Print "mal"
    formatoFechaCorrecto = False
End Select
If Not formatoFechaCorrecto Then
    MsgBox "Formato de Fecha es " & Formato & vbLf & "el formato correcto debe ser dd/mm/aa" & vbLf & vbLf & "Configure en Panel de Control -> Configuracion Regional -> Fecha"
    Exit Sub
End If

GoTo Sigue
'App.LogEvent ("Inicio Scp") ' escribe en el registro de NT
Debug.Print "Comments        : " & App.Comments
Debug.Print "CompanyName     : " & App.CompanyName
Debug.Print "EXEName         : " & App.EXEName
Debug.Print "FileDescription : " & App.FileDescription
Debug.Print "HelpFile        : " & App.HelpFile
Debug.Print "hInstance       : " & App.hInstance
Debug.Print "LegalCopyright  : " & App.LegalCopyright
Debug.Print "LegalTrademarks : " & App.LegalTrademarks
'Debug.Print "LogEvent        : " & App.LogEvent
Debug.Print "LogMode         : " & App.LogMode
Debug.Print "LogPath         : " & App.LogPath
Debug.Print "Major           : " & App.Major
Debug.Print "Minor           : " & App.Minor
Debug.Print "NonModalAllowed : " & App.NonModalAllowed
Debug.Print "Path            : " & App.Path
Debug.Print "PrevInstance    : " & App.PrevInstance
Debug.Print "ProductName     : " & App.ProductName
Debug.Print "Revision        : " & App.Revision
'Debug.Print "StartLogging    : " & App.StartLogging
Debug.Print "StartMode       : " & App.StartMode
Debug.Print "TaskVisible     : " & App.TaskVisible
Debug.Print "ThreadID        : " & App.ThreadID
Debug.Print "Title           : " & App.Title
Debug.Print "UnattendedApp   : " & App.UnattendedApp
Sigue:

salir = False
'If Archivo_Existe("D:\EML\", "SCP.EXE") Then
'f = Archivo_Fecha("D:\EML\SCP.EXE")

Dim i As Long, j As Long
IniScreen.Show 0
IniScreen.Refresh

Inicializa

If salir Then
    Exit Sub
End If

Load Login
Login.Show 1
Unload Login

End Sub
Private Sub Inicializa()

Dim Db As Database, Rs As Recordset
Dim m_Provider As String, m_InitialCatalog As String, m_DataSource As String
Dim mCatalogScpNew As String

Menu_Columnas = 11 '9
Menu_Filas = 25 '22 '21

Drive_Local = ""
Path_Local = App.Path & "\"

'myPC_file = Drive_Local & Path_Local & "ScpCnfgPC"

'Set Db = OpenDatabase(myPC_file)
'Set Rs = Db.OpenRecordset("Directorios")

'Drive_Server = Rs![Drive Server]
'Path_Server = Rs![Path Server]

'Db.Close

If ReadIniValue(Path_Local & "scp.ini", "Default", "win7") = "true" Then
    Win7 = True
Else
    Win7 = False
End If


' nuevo, con archivo ini
GoTo Nuevo_ScpIni
'Path_Server = ReadIniValue(Path_Local & "scp.ini", "Default", "Path_Server")
'If Path_Server = "" Then
    ' \scp\
'    WriteIniValue Path_Local & "scp.ini", "Default", "Path_Server", "\scp\"
'End If
Drive_Server = ReadIniValue(Path_Local & "scp.ini", "Default", "Drive_Server")
If Drive_Server = "" Then
    ' \\claudia-dv
    ' ó    C:
'    WriteIniValue Path_Local & "scp.ini", "Default", "Drive_Server", "C:\"
'    WriteIniValue Path_Local & "scp.ini", "Default", "Drive_Server", "\\claudia-dv"
End If
If ReadIniValue(Path_Local & "scp.ini", "Default", "Version") = "" Then
    ' primara vez que usa archivo ini
'    WriteIniValue Path_Local & "scp.ini", "Default", "Version", "1.0"
End If

'WriteIniValue Path_Local & "scp.ini", "Impresion", "InpX", 105
'WriteIniValue Path_Local & "scp.ini", "Impresion", "InpY", 5

Nuevo_ScpIni:

' sql server
'''Set wrkODBC = CreateWorkspace("NuevoWorkspaceODBC", "admin", "", dbUseODBC)
' instalar:
'Set cnxSqlServer = wrkODBC.OpenConnection("odbc_sql", dbDriverNoPrompt, , "ODBC;DATABASE=scp0;UID=scp_is;PWD=aqmdla;")

m_Provider = ReadIniValue(Path_Local & "scp.ini", "Sql Server", "Provider")
m_InitialCatalog = ReadIniValue(Path_Local & "scp.ini", "Sql Server", "Initial Catalog")
mCatalogScpNew = ReadIniValue(Path_Local & "scp.ini", "Sql Server", "Catalog ScpNew")
'Debug.Print "mCatalogScpNew|" & mCatalogScpNew; "|"
m_DataSource = ReadIniValue(Path_Local & "scp.ini", "Sql Server", "Data Source")

If False Then
Debug.Print "|" & m_Provider & "|"
Debug.Print "|" & m_InitialCatalog & "|"
Debug.Print "|" & m_DataSource & "|"
End If

On Error GoTo NoSql
'/////////////////////////////////////////////////
'm_DataSource = "192.168.0.70" ' solo por 13/05/16
'/////////////////////////////////////////////////
'CnxSqlServer_scp0.Open "Provider=" & m_Provider & ";" & "Initial Catalog=" & m_InitialCatalog & ";" & "Data Source=" & m_DataSource & ";user id = scp_is; password = aqmdla;"
CnxSqlServer_scp0.Open "Provider=" & m_Provider & ";" & "Initial Catalog=" & m_InitialCatalog & ";" & "Data Source=" & m_DataSource & ";user id = scp_is; password = dsa1605*;"
'CnxSqlServer_delgado1405.Open "Provider=" & m_Provider & ";" & "Initial Catalog=" & mCatalogScpNew & ";" & "Data Source=" & m_DataSource & ";user id = scp_is; password = aqmdla;"
CnxSqlServer_delgado1405.Open "Provider=" & m_Provider & ";" & "Initial Catalog=" & mCatalogScpNew & ";" & "Data Source=" & m_DataSource & ";user id = scp_is; password = dsa1605*;"
'ojo cambiar cnxsqlserverscc
On Error GoTo 0
'///////////

Drive_Server = ReadIniValue(Path_Local & "scp.ini", "Path", "Drive_Server")
Path_Mdb = ReadIniValue(Path_Local & "scp.ini", "Path", "Path_Mdb")
Path_Rpt = ReadIniValue(Path_Local & "scp.ini", "Path", "Path_Rpt")

' verifica si tiene conexion con servidor
'Debug.Print "|" & Drive_Server & Path_Server & "|"
'Debug.Print "|" & Dir(Drive_Server & Path_Server, vbDirectory) & "|"
'On Error GoTo NoConectado:
'Debug.Print "|" & Dir(Drive_Server & "\scpx\*.*") & "|"
'On Error GoTo 0
'If Dir(Path_Server, vbDirectory) Then
'    MsgBox "No Hay Conexion con Servidor"
'    Scp_Exit
'    salir = True
'    Exit Sub
'End If

If Version_Validar = False Then
    Beep
    MsgBox "DEBE ACTUALIZAR SU VERSIÓN" & vbCr & "DEL PROGRAMA", , "Inicio"
    ' sale del programa
'    Scp_Exit
'    salir = True
'    Exit Sub
End If

Syst_file = Drive_Server & Path_Mdb & "ScpCnfgSys" ' config para todo el sistema **
'myPC_file = Drive_Local & Path_Local & "ScpCnfgPC"   ' config para el PC
data_file = Drive_Server & Path_Mdb & "ScpData" ' archivos de maestros
repo_file = Drive_Local & Path_Local & "ScpRepo"   ' archivos de reportes
stat_file = Drive_Server & Path_Mdb & "ScpStatus" ' archivos de status
mpro_file = Drive_Server & Path_Mdb & "ScpMovs" ' obra en proceso
mpro2_file = Drive_Server & Path_Mdb & "ScpMov2" ' obra en proceso
'EMLmpro_file = mpro_file
'madq_file = Drive_Server & Path_Server & "ScpMadq" ' movs adquisiciones
'madq_file = Drive_Server & Path_Server & EmpOC.Fantasia & "\" & "ScpMadq" '03/02/1999
mvta_file = Drive_Server & Path_Mdb & "VtasMovs" ' movs ventas
scc_file = Drive_Server & Path_Mdb & "scc" ' scc DANILO sistema de control de corte 22/02/06

If Archivo_Existe(Drive_Server & Path_Mdb, "enmantencion.txt") Then
    MsgBox "SISTEMA EN MANTENCIÓN 1," & vbCr & "INTÉNTELO MÁS TARDE", , "Inicio"
    Scp_Exit
    salir = True
    Exit Sub
End If

If Not Archivo_Existe(Drive_Server & Path_Mdb, "ScpMovs.Mdb") Then
    MsgBox "SISTEMA EN MANTENCIÓN 2," & vbCr & "INTÉNTELO MÁS TARDE", , "Inicio"
    Scp_Exit
    salir = True
    Exit Sub
End If

'Crea_Archivos

Fecha_Mask = "##/##/##"
Fecha_Format = "dd/mm/yy"
Fecha_Vacia = "__/__/__"
sql_Fecha_Formato = "yyyymmdd"

num_fmtgrl = "##########"
num_Format0 = "#,###,###,###"
' para planos
num_Formato = "########.0"


'///////////////////////////////////////
Dim sql As String, RsEmp As New ADODB.Recordset

'sql = "SELECT * FROM empresas WHERE rut='" & "89784800-7" & "'"
sql = "SELECT * FROM empresas" ' scp0
'sql = "SELECT * FROM tb_empresas" 'delgado1303
Rs_Abrir RsEmp, sql

With RsEmp
If Not .EOF Then
    
    'Set Db = OpenDatabase(Syst_file, False, True, ";pwd=eml")
    'Set Rs = Db.OpenRecordset("Empresa")
    
    '"89784800-7" 'por default esta en eml
    'With Rs
    'Do While Not .EOF
    '    If !Rut = Rut_Eml Then Exit Do
    '    .MoveNext
    'Loop
    
    Empresa.rut = RsEmp!rut
    'Empresa.Razon = RsEmp![Razón Social]
    Empresa.Razon = RsEmp![razon_social] ' scp0
    'Empresa.Razon = RsEmp![razonSocial] ' delgado1303
    'Empresa.Fantasia = RsEmp![Nombre Fantasia]
    Empresa.Fantasia = RsEmp![Fantasia]
    Empresa.Giro = NoNulo(RsEmp!Giro)
    Empresa.Direccion = NoNulo(RsEmp!Direccion)
    Empresa.Comuna = NoNulo(RsEmp!Comuna)
    Empresa.Ciudad = NoNulo(RsEmp!Ciudad)
    Empresa.Telefono1 = NoNulo(RsEmp![Telefono1]) ' central
    Empresa.Telefono2 = NoNulo(RsEmp![Telefono2]) ' fax
    Empresa.Telefono3 = NoNulo(RsEmp![Telefono3]) ' PAGO FACTURAS
    
    EmpOC = Empresa

End If

'.Close

End With

' 27/09/2003
'Set Rs = Db.OpenRecordset("Parametros")
'Parametro.Iva = Rs!Iva
'Rs.Close
sql = "SELECT * FROM parametros" ' tiene un solo registro ' sqlserver.scp0
'sql = "SELECT * FROM tb_parametros" ' tiene un solo registro sqlserver.delgado1303
Rs_Abrir RsEmp, sql

With RsEmp
If Not .EOF Then
    Parametro.Iva = RsEmp!Iva
End If
.Close
End With

'Db.Close

'Madq_file = Drive_Server & Path_Mdb & EmpOC.Fantasia & "\" & "ScpMadq" '03/02/1999
Madq_file = Drive_Server & Path_Mdb & "ScpMadq"  '03/02/1999

Usuario.ReadOnly = False

' 01/07/04
m_TextoIso = "FO713-01 Rev.0 30.06.04"

PrimeraEtiqueta = True

Tabla_PlanosCabecera = "[planos cabecera]"
Tabla_PlanosDetalle = "[planos detalle]"

Exit Sub

NoSql:

'Debug.Print "Error:|" & Err.Number & "|"

Select Case Err.Number

Case -2147467259
    ' si el ini esta cambido, tambien da este error
    MsgBox "Red o .ini|" & Err.Description & "|", , "Error !!!"
Case 3001
    MsgBox ".ini no esta configurado|" & Err.Description & "|", , "Error !!!"
Case 3706
    MsgBox "Sql Cliente NO Instalado o Win7-sqlncli10|" & Err.Description & "|", , "Error !!!"
Case Else
    MsgBox "Error|" & Err.Number & "|" & Err.Description & "|", , "Error !!!"
End Select

Scp_Exit
salir = True

End Sub
Private Function Version_Validar() As Boolean
' valida si la fecha del archivo en el servidor
' es mayor que la de la estación
'Version_Validar = True
'Exit Function
Dim s As String, l As String
s = Drive_Server & "\instalar\"  'Path_Server
l = Drive_Local & Path_Local
Version_Validar = False

' verifica si esta conectado con servidor
On Error GoTo NoConectado
If Dir(s & "scp.exe", vbArchive) = "" Then
    ' ok, conectado
End If
On Error GoTo 0

If Not Archivo_Existe(s, "SCP.EXE") Then GoTo Fin
'Debug.Print Archivo_Fecha(s, "SCP.EXE")
'Debug.Print Archivo_Fecha(l, "SCP.EXE")
If Archivo_Fecha(s, "SCP.EXE") > Archivo_Fecha(l, "SCP.EXE") Then Exit Function
'If Archivo_Fecha(s & "ScpCnfgSys.MDB") > Archivo_Fecha(l & "ScpCnfgSys.MDB") Then Exit Function
Fin:
Version_Validar = True
Exit Function
NoConectado:
MsgBox "NO Hay Conexión con la carpeta " & Path_Mdb & vbLf & " del Servidor" & Drive_Server
End

End Function
Private Sub Scp_Exit()
' sale completamente del sistema
Dim f As Form
For Each f In Forms
    Unload f
Next
End Sub
Public Sub BasePlano_Recalcula()
' recalcula valores de base de Datos Plano
' de acuerdo a OTf, ITOf, ITOpg y GD

If MsgBox("¿ Seguro que Recalcula ?", vbYesNo) <> vbYes Then Exit Sub

Dim Dbm As Database, RsPc As Recordset, RsPd As Recordset
Dim OTf As Recordset, ITOf As Recordset
Dim ITOpg As Recordset
Dim GDc As Recordset, GDd As Recordset, m_Tipo As String
Dim indice As String

Dim m_Nv As Double, m_Plano As String, m_Peso As Double, m_Superficie As Double

indice = "NV-Plano-Marca"

Set Dbm = OpenDatabase(mpro_file)

Set RsPc = Dbm.OpenRecordset("Planos Cabecera")
RsPc.Index = "NV-Plano"

Set RsPd = Dbm.OpenRecordset("Planos Detalle")
RsPd.Index = indice

Set OTf = Dbm.OpenRecordset("OT Fab Detalle")
Set ITOf = Dbm.OpenRecordset("ITO Fab Detalle")
Set ITOpg = Dbm.OpenRecordset("ITO pg Detalle")
Set GDc = Dbm.OpenRecordset("GD cabecera")
GDc.Index = "numero"
Set GDd = Dbm.OpenRecordset("GD Detalle")

GoTo LimpiaPlanos
'/////////////////
' pone Kg y m2 a cabecera
RsPd.MoveFirst
m_Nv = RsPd!Nv
m_Plano = RsPd!Plano
m_Peso = 0
m_Superficie = 0
Do While Not RsPd.EOF
    If m_Nv = RsPd!Nv And m_Plano = RsPd!Plano Then
        m_Peso = m_Peso + RsPd![Cantidad Total] * RsPd!Peso
        m_Superficie = m_Superficie + RsPd![Cantidad Total] * RsPd!Superficie
    Else
        Cabecera_Actualizar RsPc, m_Nv, m_Plano, m_Peso, m_Superficie
        m_Nv = RsPd!Nv
        m_Plano = RsPd!Plano
        m_Peso = RsPd![Cantidad Total] * RsPd!Peso
        m_Superficie = RsPd![Cantidad Total] * RsPd!Superficie
    End If
    RsPd.MoveNext
Loop
Cabecera_Actualizar RsPc, m_Nv, m_Plano, m_Peso, m_Superficie

' limpia cantidades de planos
LimpiaPlanos:
With RsPd
.MoveFirst
Do While Not .EOF
    .Edit
    ![OT fab] = 0
    ![ITO fab] = 0
    ![GD gal] = 0
    ![ITO pyg] = 0
    ![ito gr] = 0
    ![ito pp] = 0
    !GD = 0
    .Update
    .MoveNext
Loop
End With

OTf:
If OTf.RecordCount > 0 Then
    OTf.MoveFirst
    Do While Not OTf.EOF
        RsPd.Seek "=", OTf!Nv, 0, OTf!Plano, OTf!Marca
        If Not RsPd.NoMatch Then
            RsPd.Edit
            RsPd![OT fab] = RsPd![OT fab] + OTf!Cantidad
            RsPd.Update
        End If
        OTf.MoveNext
    Loop
End If

ITOf:
If ITOf.RecordCount > 0 Then
    ITOf.MoveFirst
    Do While Not ITOf.EOF
        RsPd.Seek "=", ITOf!Nv, 0, ITOf!Plano, ITOf!Marca
        If Not RsPd.NoMatch Then
            RsPd.Edit
            RsPd![ITO fab] = RsPd![ITO fab] + ITOf!Cantidad
            RsPd.Update
        End If
        ITOf.MoveNext
    Loop
End If

ITOpg:
If ITOpg.RecordCount > 0 Then
    ITOpg.MoveFirst
    Do While Not ITOpg.EOF
        RsPd.Seek "=", ITOpg!Nv, 0, ITOpg!Plano, ITOpg!Marca
        If Not RsPd.NoMatch Then
            RsPd.Edit
' P : pintura
' G : galvanizado
' R : granallado, Erwin
' S : granallado esp
' T : produccion pintura , Erwin
' U : pp esp
            Select Case ITOpg!Tipo
            Case "P", "G"
                RsPd![ITO pyg] = RsPd![ITO pyg] + ITOpg!Cantidad
            Case "R", "S"
                RsPd![ito gr] = RsPd![ito gr] + ITOpg!Cantidad
            Case "T", "U"
                RsPd![ito pp] = RsPd![ito pp] + ITOpg!Cantidad
            End Select
            RsPd.Update
        End If
        ITOpg.MoveNext
    Loop
End If

GD:
If GDd.RecordCount > 0 Then
    GDd.MoveFirst
    Do While Not GDd.EOF
        
        RsPd.Seek "=", GDd!Nv, 0, GDd!Plano, GDd!Marca
        If Not RsPd.NoMatch Then
            
            GDc.Seek "=", GDd!Numero
            If GDc.NoMatch Then
                m_Tipo = ""
            Else
                m_Tipo = GDc!Tipo
            End If
            
            RsPd.Edit
            Select Case m_Tipo
            Case "N"
                RsPd!GD = RsPd!GD + GDd!Cantidad
            Case "G"
                RsPd![GD gal] = RsPd![GD gal] + GDd!Cantidad
            End Select
            RsPd.Update
        End If
        GDd.MoveNext
    Loop
End If

MsgBox "LISTO"

End Sub
Private Sub Cabecera_Actualizar(RsC As Recordset, Nv As Double, Plano As String, Peso As Double, Superficie As Double)
RsC.Seek "=", Nv, Plano
If Not RsC.NoMatch Then
    RsC.Edit
    RsC![Peso Total] = Peso
    RsC![Superficie Total] = Superficie
    RsC.Update
End If
End Sub
Public Sub Login_Registrar_NO_VA(indice As Long, Optional Usuario_Nombre As String)
' creada el 01/07/98
Dim Db As Database, Rs As Recordset
Dim bm As Variant
Set Db = OpenDatabase(stat_file)
Set Rs = Db.OpenRecordset("Usuarios Login")
Rs.Index = "Indice"
If indice = 0 Then
    With Rs
        .AddNew
        !Estación = Left(Computador_Nombre(), 15)
        ![Usuario Win] = Left(Usuario_Win(), 15)
        ![Usuario Scp] = Left(Usuario_Nombre, 15)
        !Fecha = Format(Now, Fecha_Format)
        ![Hora Inicio] = Format(Now, "hh:mm:ss")
        .Update
        .Bookmark = .LastModified
        Usuario.indice = !indice
    End With
Else
    Rs.Seek "=", indice
    If Not Rs.NoMatch Then
        Rs.Edit
        Rs![Hora Final] = Format(Now, "hh:mm:ss")
        Rs.Update
    End If
End If
Rs.Close
Db.Close
End Sub
Public Function Movs_Path(RutEmpresa As String, ObrasTerminadas As Boolean) As String
'08/05/99
Dim Arch As String
Arch = IIf(ObrasTerminadas, "ScpHist", "ScpMovs")
If RutEmpresa = Empresa.rut Then   'eml
    Movs_Path = Drive_Server & Path_Mdb & Arch
Else
    Movs_Path = Drive_Server & Path_Mdb & "pyp\" & Arch
End If
End Function
Public Function GetNumDoc(Documento As String, RsDoc As Recordset, RsCorre As Recordset) As Double
' 08/05/99
' obtiene siguiente número disponible y bloquea registro
Dim SalirSinGrabar As Boolean, intento As Integer
SalirSinGrabar = False
intento = 1
GetNumDoc = 0

Do While intento < 2
    If RsCorre(Documento & " Bloqueada") Then
        If MsgBox("espere un instante..." & vbLf & "¿ ReIntenta Grabación ?", vbRetryCancel) = vbRetry Then
            intento = intento + 1
        Else
            SalirSinGrabar = True
            Exit Do
        End If
    Else
        Exit Do ' no bloqueado
    End If
Loop

If SalirSinGrabar Then
    ' sales
Else

    RsCorre.Edit
    RsCorre(Documento & " Bloqueada") = True
    RsCorre.Update
    
    On Error GoTo Error
    RsDoc.MoveLast
    GetNumDoc = RsDoc!Numero + 1
    
End If
Exit Function

Error:
GetNumDoc = 1
End Function
Public Sub DesBloqueo(Documento As String, RsCorre As Recordset)
'13/05/1999
RsCorre.Edit
RsCorre(Documento & " Bloqueada") = False
RsCorre.Update
End Sub
Public Sub Registrar_NO_VA(Accion As String)
Dim DbSt As Database, RsSt As Recordset 'status
Set DbSt = OpenDatabase(stat_file)
Set RsSt = DbSt.OpenRecordset("Mantención")
With RsSt
.AddNew
!Fecha = Now
!Acción = Left(Accion, 30)
!Estación = Left(Computador_Nombre(), 15)
![Usuario Win] = Left(Usuario_Win(), 15)
![Usuario Scp] = Left(Usuario.nombre, 15)
.Update
.Close
End With
DbSt.Close
End Sub
Public Sub Ini_Crear()
' creada el 08/02/05
' crea archivo scp.ini en carpeta system
Dim camino As String, Archivo As String, Texto As String, p As String, VAR As String, con As String
Dim arr(1, 99) As String, i As Integer, Fin As Integer, Contenido As String
camino = WindowsDirectory
Archivo = "scp.ini"
i = 0

If Not Archivo_Existe(camino & "\", Archivo) Then

    Open camino & "\" & Archivo For Output As #1
    
    ' este archivo tiene que ver con el PC mas que con el usuario
    Print #1, "[Ini]"
    Print #1, "VersionIni=1.00"
    Print #1, "VersionFecha=08/02/2005"
    Print #1, ""
    Print #1, "[Principal]"
    Print #1, "Programa=scp.exe"
    Print #1, "Servidor=\\acr166_dualpro\scp"
    Print #1, "Obras=En Proceso" ' usuario ??
    Print #1, "Impresora=IBM Proprinter II"
    Print #1, ""
    Print #1, "[Web]"
    Print #1, "FechaUltimaActualizacion="
    
    Close #1
    
Else

    ' lee variables desde archivo texto
    Open camino & "\" & Archivo For Input As #1
    Do While Not EOF(1)
        Input #1, Texto
        p = InStr(1, Texto, "=")
        If p <> 0 Then
            VAR = Left(Texto, p - 1) ' nombre de la variable
            con = Mid(Texto, p + 1)  ' contenido
            i = i + 1
            arr(0, i) = VAR
            arr(1, i) = con
            Debug.Print VAR & "|" & con
        End If
        Fin = i
    Loop
    Close #1
    
    ' busca variable y lee su contenido
    Contenido = ""
    VAR = "servidor"
    For i = 1 To Fin
        If UCase(arr(0, i)) = UCase(VAR) Then
            Contenido = arr(1, i)
            Debug.Print "º"
        End If
    Next
    Debug.Print "encontrado " & VAR & "|" & Contenido
    
End If

End Sub
Public Function Variable_Leer(Variable As String) As String

' lee el contenido de una variable desde archivo de texto *.ini
Dim camino As String, Archivo As String, Texto As String, p As Integer, i As Integer
Dim VAR As String, con As String

Variable_Leer = "Variable no Existe"

camino = WindowsDirectory
Archivo = "scp.ini"

If Archivo_Existe(camino & "\", Archivo) Then

    ' busca variable y lee su contenido
    Open camino & "\" & Archivo For Input As #1
    Do While Not EOF(1)
    
        Input #1, Texto
        p = InStr(1, Texto, "=")
        If p <> 0 Then
    
            VAR = Left(Texto, p - 1) ' nombre de la variable
    
            If UCase(VAR) = UCase(Variable) Then
            
                con = Mid(Texto, p + 1)  ' contenido
                Variable_Leer = con
                Exit Do
                
            End If
            
        End If
        
    Loop
    Close #1

Else

    Variable_Leer = "Archivo no Existe"

End If

End Function
Public Sub Variable_Escribir(Variable As String, Contenido As String)

' escribe el contenido de una variable a archivo de texto *.ini
' no verifica si orden esta correcto, solo reescribe archivo completo con nuevo contenido de variable
' si variable no existe, el archivo queda igual

Dim camino As String, Archivo As String, Texto As String, p As Integer, i As Integer
Dim VAR As String, con As String, arr(99) As String, Fin As Integer

'Variable_Leer = "Variable no Existe"

camino = WindowsDirectory
Archivo = "scp.ini"

i = 0
If Archivo_Existe(camino & "\", Archivo) Then

    ' lee archivo completo en arreglo
    Open camino & "\" & Archivo For Input As #1
    Do While Not EOF(1)
        
        Input #1, Texto
        
        i = i + 1
        arr(i) = Texto
        
        p = InStr(1, Texto, "=")
        
        If p <> 0 Then
    
            VAR = Left(Texto, p - 1) ' nombre de la variable
    
            If UCase(VAR) = UCase(Variable) Then
            
                arr(i) = VAR & "=" & Contenido
                            
            End If
            
        End If
        
    Loop
    Close #1

    Fin = i

    ' escribe archivo completo
    Open camino & "\" & Archivo For Output As #1

    For i = 1 To Fin
        Print #1, arr(i)
    Next

    Close #1

Else

'    Variable_Leer = "Archivo no Existe"

End If

End Sub
Public Function fecha2semana(P_Fecha As Date) As String
' trasforma fecha a semana del año, exclusivamente para checklist
' ej: 25/07/07 -> 5-04
' nota: fecha siempre debe ser miercoles
Dim m_Fecha As Date, dia As Integer, mes As Integer, semana As Integer ', primerdomingo As Integer
Dim diaespecial As Date
'mes = Month(P_Fecha)
'm_Fecha = "1/" & mes & "/" & Year(P_Fecha)
m_Fecha = P_Fecha + 4
dia = Day(m_Fecha)
mes = Month(m_Fecha)
semana = Int(dia / 7)
'primerdomingo = dia - semana * 7

diaespecial = P_Fecha + 6
If Day(diaespecial) = 2 Then
    semana = 0
    mes = Month(diaespecial)
End If

fecha2semana = semana + 1 & "-" & PadL(Trim(str(mes)), 2, "0") '& " pm " & primerdomingo

End Function
Public Sub Crea_Repo()

Dim Db As Database, td As TableDef, Campo(299) As Field

If testeo Then
    MsgBox "antes de scprepo"
End If

If Not Archivo_Existe(App.Path & "\", "ScpRepo.MDB") Then
    If testeo Then
        MsgBox "no existe antes"
    End If
    Set Db = CreateDatabase(App.Path & "\" & "ScpRepo", dbLangGeneral)
    If testeo Then
        MsgBox "no existe despues"
    End If
Else
    If testeo Then
        MsgBox "existe antes"
    End If
    Set Db = OpenDatabase(App.Path & "\" & "ScpRepo")
    If testeo Then
        MsgBox "existe despues"
    End If
End If

If testeo Then
    MsgBox "despues de scprepo"
End If

If Not Tabla_Existe(Db, "Planos") Then

    Set td = Db.CreateTableDef("Planos")
    Set Campo(0) = td.CreateField("NV", dbDouble)
    Set Campo(1) = td.CreateField("Obra", dbText, 30)
    Set Campo(2) = td.CreateField("Plano", dbText, 50)
    Set Campo(3) = td.CreateField("Rev", dbText, 10)
    
    Set Campo(4) = td.CreateField("Item", dbInteger)
    Set Campo(5) = td.CreateField("Cantidad Total", dbInteger)
    Set Campo(6) = td.CreateField("Descripcion", dbText, 30)
    Set Campo(7) = td.CreateField("Marca", dbText, 50)
    Set Campo(8) = td.CreateField("Peso Unitario", dbDouble)
    Set Campo(9) = td.CreateField("Peso Total", dbDouble)
    Set Campo(10) = td.CreateField("m2 Unitario", dbDouble)
    Set Campo(11) = td.CreateField("m2 Total", dbDouble)
    Set Campo(12) = td.CreateField("Observaciones", dbText, 50)
    Set Campo(13) = td.CreateField("densidad_n", dbInteger)
    Set Campo(14) = td.CreateField("densidad_s", dbText, 50)
    Set Campo(15) = td.CreateField("fecha", dbDate) ' fecha de ingreso del plano al sistema
    
    Campos_Append Db, td, Campo, 16
'    Campos_Append 13
'    Db.TableDefs.Append Td
    
End If

If Not Tabla_Existe(Db, "OT") Then
    Set td = Db.CreateTableDef("OT")
    Set Campo(0) = td.CreateField("NV", dbDouble)
    Set Campo(1) = td.CreateField("Obra", dbText, 30)
    Set Campo(2) = td.CreateField("Contratista", dbText, 50)
    Set Campo(3) = td.CreateField("Nº OT", dbDouble)
    Set Campo(4) = td.CreateField("Fecha", dbDate)
    Set Campo(5) = td.CreateField("Kg Total", dbDouble)
    Set Campo(6) = td.CreateField("$ Total", dbDouble)
    Set Campo(7) = td.CreateField("$ Promedio", dbDouble)
'    Set Campo(8) = td.CreateField("Montaje", dbBoolean)
    Campos_Append Db, td, Campo, 8
'    Db.TableDefs.Append Td
End If

' creada 17/08/04 para repo_general_ot
If Not Tabla_Existe(Db, "OTe") Then
    Set td = Db.CreateTableDef("OTe")
    Set Campo(0) = td.CreateField("NV", dbDouble)
    Set Campo(1) = td.CreateField("Obra", dbText, 30)
    Set Campo(2) = td.CreateField("Contratista", dbText, 50)
    Set Campo(3) = td.CreateField("Nº OT", dbDouble)
    Set Campo(4) = td.CreateField("Fecha", dbDate)
    Set Campo(5) = td.CreateField("Kg Total", dbDouble)
    Set Campo(6) = td.CreateField("$ Total", dbDouble)
    Set Campo(7) = td.CreateField("$ Promedio", dbDouble)
    Set Campo(8) = td.CreateField("Tipo", dbText, 1)
    Campos_Append Db, td, Campo, 9
'    Db.TableDefs.Append Td
End If

If Not Tabla_Existe(Db, "ITO") Then
    Set td = Db.CreateTableDef("ITO")
    Set Campo(0) = td.CreateField("NV", dbDouble)
    Set Campo(1) = td.CreateField("Obra", dbText, 30)
    Set Campo(2) = td.CreateField("Contratista", dbText, 50)
    Set Campo(3) = td.CreateField("Clasificacion", dbText, 10)
    Set Campo(4) = td.CreateField("Nº ITO", dbDouble)
    Set Campo(5) = td.CreateField("Fecha", dbDate)
    Set Campo(6) = td.CreateField("Kg Total", dbDouble)
    Set Campo(7) = td.CreateField("$ Total", dbDouble)
    Set Campo(8) = td.CreateField("m2 Total", dbDouble)
    Campos_Append Db, td, Campo, 9
'    Db.TableDefs.Append Td
End If

If Not Tabla_Existe(Db, "ITOpr_det") Then
    Set td = Db.CreateTableDef("ITOpr_det")
    Set Campo(0) = td.CreateField("Numero", dbDouble)
    Set Campo(1) = td.CreateField("Fecha", dbDate)
    Set Campo(2) = td.CreateField("NV", dbDouble)
    Set Campo(3) = td.CreateField("nvarea", dbInteger)
    Set Campo(4) = td.CreateField("Obra", dbText, 30)
    Set Campo(5) = td.CreateField("Plano", dbText, 50)
    Set Campo(6) = td.CreateField("Rev", dbText, 10)
    Set Campo(7) = td.CreateField("Marca", dbText, 50)
    Set Campo(8) = td.CreateField("descripcion", dbText, 40)
    Set Campo(9) = td.CreateField("cantidad_total", dbDouble)
    Set Campo(10) = td.CreateField("cantidad_pr", dbDouble)
    Set Campo(11) = td.CreateField("Peso Unitario", dbDouble)
    Set Campo(12) = td.CreateField("m2 Unitario", dbDouble)
    Set Campo(13) = td.CreateField("Precio Unitario", dbDouble)
    Set Campo(14) = td.CreateField("turno", dbInteger)
    Set Campo(15) = td.CreateField("maquina", dbText, 1)
    Set Campo(16) = td.CreateField("Rut Operador", dbText, 10)
    Set Campo(17) = td.CreateField("Nombre Operador", dbText, 30)
    Set Campo(18) = td.CreateField("tipo2", dbText, 4)
    Set Campo(19) = td.CreateField("manos1", dbInteger)
    Set Campo(20) = td.CreateField("manos2", dbInteger)

    Campos_Append Db, td, Campo, 21
'    Db.TableDefs.Append Td
End If


If Not Tabla_Existe(Db, "GD") Then
    Set td = Db.CreateTableDef("GD")
    Set Campo(0) = td.CreateField("NV", dbDouble)
    Set Campo(1) = td.CreateField("Obra", dbText, 30)
    Set Campo(2) = td.CreateField("GD", dbDouble)
    Set Campo(3) = td.CreateField("Tipo", dbText, 1)
    Set Campo(4) = td.CreateField("Fecha", dbDate)
    Set Campo(5) = td.CreateField("Cliente", dbText, 30)
    Set Campo(6) = td.CreateField("Kg Total", dbDouble)
    Set Campo(7) = td.CreateField("$ Total", dbDouble)
    Set Campo(8) = td.CreateField("Chofer", dbText, 30)
    Set Campo(9) = td.CreateField("Patente", dbText, 30)
    Set Campo(10) = td.CreateField("Contenido1", dbText, 50)
    Set Campo(11) = td.CreateField("Contenido2", dbText, 50)
    Campos_Append Db, td, Campo, 12
'    Db.TableDefs.Append Td
    Indice_Crear td, "GD", "GD", True
End If

' creada 25/03/2014
If Not Tabla_Existe(Db, "gd_det") Then

    Set td = Db.CreateTableDef("gd_det")
    Set Campo(0) = td.CreateField("nv", dbDouble)
    Set Campo(1) = td.CreateField("plano", dbText, 50)
    Set Campo(2) = td.CreateField("rev", dbText, 10)
    Set Campo(3) = td.CreateField("marca", dbText, 50)
    Set Campo(4) = td.CreateField("descripcion", dbText, 40)
    Set Campo(5) = td.CreateField("cantidadTotal", dbDouble)
    Set Campo(6) = td.CreateField("pesoUnitario", dbDouble)
    Set Campo(7) = td.CreateField("guia", dbDouble)
    Set Campo(8) = td.CreateField("fecha", dbDate)
    Set Campo(9) = td.CreateField("cantidadDespachada", dbDouble)
    Set Campo(10) = td.CreateField("bulto", dbDouble)

    Campos_Append Db, td, Campo, 11
    
End If

If Not Tabla_Existe(Db, "Piezas Pendientes") Then
    Set td = Db.CreateTableDef("Piezas Pendientes")
    Set Campo(0) = td.CreateField("NV", dbDouble)
    Set Campo(1) = td.CreateField("Obra", dbText, 30)
    Set Campo(2) = td.CreateField("Plano", dbText, 50)
    Set Campo(3) = td.CreateField("Rev", dbText, 10)
    Set Campo(4) = td.CreateField("Marca", dbText, 50)
    Set Campo(5) = td.CreateField("Descripcion", dbText, 30)
    Set Campo(6) = td.CreateField("Contratista", dbText, 10) ' solo diez
    Set Campo(7) = td.CreateField("Total", dbInteger)
    Set Campo(8) = td.CreateField("Peso Unitario", dbDouble)
    Set Campo(9) = td.CreateField("m2 Unitario", dbDouble)
    Set Campo(10) = td.CreateField("xFab", dbInteger) ' por asignar
    Set Campo(11) = td.CreateField("enFab", dbInteger) ' asignadas y no recibidas de ITOgr
    Set Campo(12) = td.CreateField("enGR", dbInteger) '
    Set Campo(13) = td.CreateField("enPP", dbInteger)
    Set Campo(14) = td.CreateField("enPIN", dbInteger) ' en pintura
    Set Campo(15) = td.CreateField("xDesp", dbInteger) ' por despachar
    Set Campo(16) = td.CreateField("Desp", dbInteger)
    Set Campo(17) = td.CreateField("densidad", dbInteger)
    Campos_Append Db, td, Campo, 18
'    Db.TableDefs.Append Td
End If

If Not Tabla_Existe(Db, "Avances de Pago") Then
    Set td = Db.CreateTableDef("Avances de Pago")
    Set Campo(0) = td.CreateField("Obra", dbText, 50)
    Set Campo(1) = td.CreateField("SubContratista", dbText, 50)
    Set Campo(2) = td.CreateField("Nº OT", dbDouble)
    Set Campo(3) = td.CreateField("Plano", dbText, 50)
    Set Campo(4) = td.CreateField("Rev", dbText, 10)
    Set Campo(5) = td.CreateField("Kg Total", dbDouble)
    Set Campo(6) = td.CreateField("$Total", dbDouble)
    Set Campo(7) = td.CreateField("Kg Recibidos", dbDouble)
    Set Campo(8) = td.CreateField("$a Pagar", dbDouble)
    Set Campo(9) = td.CreateField("%Avance", dbDouble)
    Campos_Append Db, td, Campo, 10
'    Db.TableDefs.Append Td
End If

If Not Tabla_Existe(Db, "Piezas x Asignar") Then
    Set td = Db.CreateTableDef("Piezas x Asignar")
    Set Campo(0) = td.CreateField("NV", dbDouble)
    Set Campo(1) = td.CreateField("Obra", dbText, 30)
    Set Campo(2) = td.CreateField("Plano", dbText, 50)
    Set Campo(3) = td.CreateField("Rev", dbText, 10)
    Set Campo(4) = td.CreateField("Marca", dbText, 50)
    Set Campo(5) = td.CreateField("Descripcion", dbText, 30)
    Set Campo(6) = td.CreateField("Total", dbDouble)
    Set Campo(7) = td.CreateField("x Asignar", dbDouble)
    Set Campo(8) = td.CreateField("Peso Unitario", dbDouble)
    Set Campo(9) = td.CreateField("Peso Total", dbDouble)
    Campos_Append Db, td, Campo, 10
'    Db.TableDefs.Append Td
End If

If Not Tabla_Existe(Db, "Piezas x Recibir") Then
    Set td = Db.CreateTableDef("Piezas x Recibir")
    Set Campo(0) = td.CreateField("NV", dbDouble)
    Set Campo(1) = td.CreateField("Obra", dbText, 30)
    Set Campo(2) = td.CreateField("Contratista", dbText, 50)
    Set Campo(3) = td.CreateField("Nº OT", dbDouble)
    Set Campo(4) = td.CreateField("Fecha", dbDate)
    Set Campo(5) = td.CreateField("Plano", dbText, 50)
    Set Campo(6) = td.CreateField("Rev", dbText, 10)
    Set Campo(7) = td.CreateField("Marca", dbText, 50)
    Set Campo(8) = td.CreateField("Descripcion", dbText, 30)
    Set Campo(9) = td.CreateField("x Recibir", dbDouble)
    Set Campo(10) = td.CreateField("Fecha Entrega", dbDate)
    Set Campo(11) = td.CreateField("Peso Unitario", dbDouble)
    Set Campo(12) = td.CreateField("Peso Total", dbDouble)
    Campos_Append Db, td, Campo, 13
'    Db.TableDefs.Append Td
End If

If Not Tabla_Existe(Db, "Piezas x Despachar") Then
    Set td = Db.CreateTableDef("Piezas x Despachar")
    Set Campo(0) = td.CreateField("NV", dbDouble)
    Set Campo(1) = td.CreateField("Obra", dbText, 30)
    Set Campo(2) = td.CreateField("Plano", dbText, 50)
    Set Campo(3) = td.CreateField("Rev", dbText, 10)
    Set Campo(4) = td.CreateField("Marca", dbText, 50)
    Set Campo(5) = td.CreateField("Descripcion", dbText, 30)
    Set Campo(6) = td.CreateField("Total", dbDouble)
    Set Campo(7) = td.CreateField("x Despachar", dbDouble)
    Set Campo(8) = td.CreateField("Peso Unitario", dbDouble)
    Set Campo(9) = td.CreateField("Peso Total", dbDouble)
    Campos_Append Db, td, Campo, 10
'    Db.TableDefs.Append Td
End If

If Not Tabla_Existe(Db, "ITOs de OTs") Then
    Set td = Db.CreateTableDef("ITOs de OTs")
    Set Campo(0) = td.CreateField("NV", dbDouble)
    Set Campo(1) = td.CreateField("Obra", dbText, 30)
    Set Campo(2) = td.CreateField("SubContratista", dbText, 50)
    Set Campo(3) = td.CreateField("OT Nº", dbDouble)
    Set Campo(4) = td.CreateField("OT Fecha", dbDate)
    Set Campo(5) = td.CreateField("OT Kg Total", dbDouble)
    Set Campo(6) = td.CreateField("OT $ Total", dbDouble)
    Set Campo(7) = td.CreateField("ITO Nº", dbDouble)
    Set Campo(8) = td.CreateField("ITO Fecha", dbDate)
    Set Campo(9) = td.CreateField("ITO Kg Total", dbDouble)
    Set Campo(10) = td.CreateField("ITO $ Total", dbDouble)
    Campos_Append Db, td, Campo, 11
'    Db.TableDefs.Append Td
End If

If Not Tabla_Existe(Db, "OTf x Plano") Then
    Set td = Db.CreateTableDef("OTf x Plano")
    Set Campo(0) = td.CreateField("NV", dbDouble)
    Set Campo(1) = td.CreateField("Obra", dbText, 30)
    Set Campo(2) = td.CreateField("Plano", dbText, 50)
    Set Campo(3) = td.CreateField("Rev", dbText, 10)
    Set Campo(4) = td.CreateField("Marca", dbText, 50)
    Set Campo(5) = td.CreateField("Nº OT", dbDouble)
    Set Campo(6) = td.CreateField("Fecha OT", dbDate)
    Set Campo(7) = td.CreateField("Fecha Entrega", dbDate)
    Set Campo(8) = td.CreateField("SubContratista", dbText, 50)
    Set Campo(9) = td.CreateField("Kg Total", dbDouble)
    Set Campo(10) = td.CreateField("Kg Recibidos", dbDouble)
    Set Campo(11) = td.CreateField("% Avance", dbDouble)
    Campos_Append Db, td, Campo, 12
'    Db.TableDefs.Append Td
End If

'/////////////////////////////////////////////
Dim nl As String, i As Integer, ent As Integer
If Not Tabla_Existe(Db, "OC Legal") Then
    Set td = Db.CreateTableDef("OC Legal")
    
    Set Campo(0) = td.CreateField("Numero", dbDouble)
    Set Campo(1) = td.CreateField("Emision", dbDate)
    Set Campo(2) = td.CreateField("Senores", dbText, 50) ' señores 18/02/08
    Set Campo(3) = td.CreateField("RUT", dbText, 10)
    Set Campo(4) = td.CreateField("Direccion", dbText, 50)
    Set Campo(5) = td.CreateField("Comuna", dbText, 50)
    Set Campo(6) = td.CreateField("Telefono", dbText, 10)
    Set Campo(7) = td.CreateField("Fax", dbText, 10)
    Set Campo(8) = td.CreateField("NV", dbText, 5)
    Set Campo(9) = td.CreateField("Obra", dbText, 50)
    Set Campo(10) = td.CreateField("At Sr", dbText, 50)
    Set Campo(11) = td.CreateField("Condiciones", dbText, 50)
    Set Campo(12) = td.CreateField("Entregar en", dbText, 50)
    Set Campo(13) = td.CreateField("Fecha Entrega", dbDate)
    Set Campo(14) = td.CreateField("Cotizacion", dbDouble)
    Set Campo(15) = td.CreateField("Guia de Despacho", dbDouble)
    
    Set Campo(16) = td.CreateField("SubTotal", dbDouble)
    Set Campo(17) = td.CreateField("% Descuento", dbDouble)
    Set Campo(18) = td.CreateField("Descuento", dbDouble)
    Set Campo(19) = td.CreateField("Otro Descuento", dbDouble)
    Set Campo(20) = td.CreateField("Neto", dbDouble)
    Set Campo(21) = td.CreateField("Iva", dbDouble)
    Set Campo(22) = td.CreateField("Total", dbDouble)
    
    Set Campo(23) = td.CreateField("Obs 1", dbText, 50)
    Set Campo(24) = td.CreateField("Obs 2", dbText, 50)
    Set Campo(25) = td.CreateField("Obs 3", dbText, 50)
    Set Campo(26) = td.CreateField("Obs 4", dbText, 50)
    
    'detalle
    For i = 1 To 20
        nl = str(i)
        ent = 26 + (i - 1) * 9
        Set Campo(ent + 1) = td.CreateField("Codigo Producto" & nl, dbText, 20)
        Set Campo(ent + 2) = td.CreateField("Cantidad" & nl, dbDouble)
        Set Campo(ent + 3) = td.CreateField("Unidad" & nl, dbText, 3)
        Set Campo(ent + 4) = td.CreateField("Descripcion" & nl, dbText, 50)
        Set Campo(ent + 5) = td.CreateField("Largo" & nl, dbDouble)  ' en mm
        Set Campo(ent + 6) = td.CreateField("Precio Unitario" & nl, dbDouble)
        Set Campo(ent + 7) = td.CreateField("Total" & nl, dbDouble)
        Set Campo(ent + 8) = td.CreateField("CuentaContable" & nl, dbText, 30)
        Set Campo(ent + 9) = td.CreateField("CentroCosto" & nl, dbText, 30)
    Next
    
    ' si existe [oc cabecera].fechaModificacion => debe decir "modificada el dd/mm/aaaa"
    Set Campo(198) = td.CreateField("fechaModificacion", dbText, 30)
    
    Campos_Append Db, td, Campo, 199
    'Campos_Append Db, td, Campo, 27
    
End If

'/////////////////////////////////////////////
If Not Tabla_Existe(Db, "OC Legal Detalle") Then

    Dim idxNombre As Index

    Set td = Db.CreateTableDef("OC Legal Detalle")
    'detalle
    Set Campo(0) = td.CreateField("Numero", dbDouble)
    Set Campo(1) = td.CreateField("Codigo Producto", dbText, 20)
    Set Campo(2) = td.CreateField("Cantidad", dbDouble)
    Set Campo(3) = td.CreateField("Unidad", dbText, 3)
    Set Campo(4) = td.CreateField("Descripcion", dbText, 50)
    Set Campo(5) = td.CreateField("Largo", dbDouble)    ' en mm
    Set Campo(6) = td.CreateField("Precio Unitario", dbDouble)
    Set Campo(7) = td.CreateField("Total", dbDouble)
    Set Campo(8) = td.CreateField("CuentaContable", dbText, 30)
    Set Campo(9) = td.CreateField("CentroCosto", dbText, 30)
    
    Campos_Append Db, td, Campo, 10
    
    Set idxNombre = td.CreateIndex
    With idxNombre
        .Name = "numero"
        .Fields.Append td.CreateField("Numero")
    End With
    td.Indexes.Append idxNombre
    
End If


If Not Tabla_Existe(Db, "bulto") Then
    Set td = Db.CreateTableDef("bulto")
    
    Set Campo(0) = td.CreateField("Numero", dbDouble)
    Set Campo(1) = td.CreateField("Emision", dbDate)
    Set Campo(2) = td.CreateField("Senores", dbText, 50) ' señores 18/02/08
    Set Campo(3) = td.CreateField("RUT", dbText, 10)
    Set Campo(4) = td.CreateField("NV", dbText, 5)
    Set Campo(5) = td.CreateField("Obra", dbText, 50)
    Set Campo(6) = td.CreateField("KgTotal", dbDouble)
    
    'detalle
    For i = 1 To 20
        nl = str(i)
        ent = 6 + (i - 1) * 7
        Set Campo(ent + 1) = td.CreateField("plano" & nl, dbText, 50)
        Set Campo(ent + 2) = td.CreateField("rev" & nl, dbText, 10)
        Set Campo(ent + 3) = td.CreateField("marca" & nl, dbText, 50)
        Set Campo(ent + 4) = td.CreateField("descripcion" & nl, dbText, 50)
        Set Campo(ent + 5) = td.CreateField("cantidad" & nl, dbDouble)
        Set Campo(ent + 6) = td.CreateField("pesounitario" & nl, dbDouble)
        Set Campo(ent + 7) = td.CreateField("pesototal" & nl, dbDouble)
    Next
    
    Campos_Append Db, td, Campo, 147
'    Db.TableDefs.Append Td
End If

'///////////////////////////////////////////////
If Not Tabla_Existe(Db, "Produccion Mensual") Then
    
    Set td = Db.CreateTableDef("Produccion Mensual")
    
    Set Campo(0) = td.CreateField("Fecha", dbDate)
    Set Campo(1) = td.CreateField("Mes Año", dbText, 8)
    Set Campo(2) = td.CreateField("Total", dbDouble)
    Set Campo(3) = td.CreateField("SubC 0", dbDouble)
    Set Campo(4) = td.CreateField("nombre1", dbText, 50)
    Set Campo(5) = td.CreateField("SubC 1", dbDouble)
    Set Campo(6) = td.CreateField("nombre2", dbText, 50)
    Set Campo(7) = td.CreateField("SubC 2", dbDouble)
    Set Campo(8) = td.CreateField("nombre3", dbText, 50)
    Set Campo(9) = td.CreateField("SubC 3", dbDouble)
    Set Campo(10) = td.CreateField("nombre4", dbText, 50)
    Set Campo(11) = td.CreateField("SubC 4", dbDouble)
    Set Campo(12) = td.CreateField("nombre5", dbText, 50)
    Set Campo(13) = td.CreateField("SubC 5", dbDouble)
    Set Campo(14) = td.CreateField("nombre6", dbText, 50)
    Set Campo(15) = td.CreateField("SubC 6", dbDouble)
    Set Campo(16) = td.CreateField("nombre7", dbText, 50)
    Set Campo(17) = td.CreateField("SubC 7", dbDouble)
    Set Campo(18) = td.CreateField("nombre8", dbText, 50)
    Set Campo(19) = td.CreateField("SubC 8", dbDouble)
    Set Campo(20) = td.CreateField("nombre9", dbText, 50)
    Set Campo(21) = td.CreateField("SubC 9", dbDouble)
    
    Campos_Append Db, td, Campo, 22
'    Db.TableDefs.Append Td
    Indice_Crear td, "Fecha", "Fecha", True

End If

'///////////////////////////////////////////////
If Not Tabla_Existe(Db, "Compras") Then
    Set td = Db.CreateTableDef("Compras")
    
    Set Campo(0) = td.CreateField("Nv", dbDouble)
    Set Campo(1) = td.CreateField("Tipo", dbText, 10)
    Set Campo(2) = td.CreateField("Orden", dbLong)
    Set Campo(3) = td.CreateField("Descripcion", dbText, 50)
    Set Campo(4) = td.CreateField("Neto", dbDouble)
    
    Campos_Append Db, td, Campo, 5
'    Db.TableDefs.Append Td
    Indice_Crear td, "Tipo", "Tipo", True

End If

'///////////////////////////////////////////////
If Not Tabla_Existe(Db, "General_Nv") Then
    Set td = Db.CreateTableDef("General_Nv")
    
    Set Campo(0) = td.CreateField("Orden", dbLong)
    Set Campo(1) = td.CreateField("Texto0", dbText, 50)
    Set Campo(2) = td.CreateField("Texto1", dbText, 50)
    Set Campo(3) = td.CreateField("Texto2", dbText, 50)
    Set Campo(4) = td.CreateField("Texto3", dbText, 50)
    Set Campo(5) = td.CreateField("Texto4", dbText, 50)
    Set Campo(6) = td.CreateField("v1_0", dbDouble)
    Set Campo(7) = td.CreateField("v1_1", dbDouble)
    Set Campo(8) = td.CreateField("v1_2", dbDouble)
    Set Campo(9) = td.CreateField("v1_t0", dbDouble)
    Set Campo(10) = td.CreateField("v1_t1", dbDouble)
    Set Campo(11) = td.CreateField("v1_t2", dbDouble)
    
    Campos_Append Db, td, Campo, 12
'    Db.TableDefs.Append Td
    Indice_Crear td, "Orden", "Orden", True

End If
'///////////////////////////////////////////////
If Not Tabla_Existe(Db, "Planos Kilos") Then

    Set td = Db.CreateTableDef("Planos Kilos")

    Set Campo(0) = td.CreateField("Tipo", dbText, 1) ' 1 x nv, 2 x contratista
    Set Campo(1) = td.CreateField("Nv", dbDouble)
    Set Campo(2) = td.CreateField("cc", dbDouble) ' centro de costo
    Set Campo(3) = td.CreateField("RUT", dbText, 10)
    Set Campo(4) = td.CreateField("Descripcion", dbText, 50)
    Set Campo(5) = td.CreateField("KgTotales", dbDouble)
    Set Campo(6) = td.CreateField("Esquema", dbText, 12) ' galvanizado o pintado
    Set Campo(7) = td.CreateField("ListadoPernos", dbText, 10)
    Set Campo(8) = td.CreateField("KgEntregados", dbDouble) ' nuevo 27/08/13
    Set Campo(9) = td.CreateField("KgxFab", dbDouble)
    Set Campo(10) = td.CreateField("KgenFab", dbDouble)

    Set Campo(11) = td.CreateField("KgenGR", dbDouble)
    Set Campo(12) = td.CreateField("KgenPP", dbDouble)
    Set Campo(13) = td.CreateField("KgenPin", dbDouble)

    Set Campo(14) = td.CreateField("KgparaDes1", dbDouble)
    Set Campo(15) = td.CreateField("KgparaDes2", dbDouble)
    Set Campo(16) = td.CreateField("KgDes", dbDouble)
    Set Campo(17) = td.CreateField("pFab1", dbText, 4)
    Set Campo(18) = td.CreateField("pFab2", dbText, 4)
    Set Campo(19) = td.CreateField("KgenFabxC", dbDouble) ' de cada contratista
    
    Campos_Append Db, td, Campo, 20
'    Db.TableDefs.Append Td
    Indice_Crear td, "Rut", "Rut", False

End If
'///////////////////////////////////////////////
If Not Tabla_Existe(Db, "ComponentesKilos") Then

    Set td = Db.CreateTableDef("ComponentesKilos")
    
    Set Campo(0) = td.CreateField("RUT", dbText, 10)
    Set Campo(1) = td.CreateField("Descripcion", dbText, 50)
    Set Campo(2) = td.CreateField("Nv", dbDouble)
    Set Campo(3) = td.CreateField("KgTotales", dbDouble)
    Set Campo(4) = td.CreateField("KgxCortar", dbDouble)
    Set Campo(5) = td.CreateField("KgCortados", dbDouble)
    Set Campo(6) = td.CreateField("KgPerforados", dbDouble)
    Set Campo(7) = td.CreateField("KgEntregados", dbDouble)
'aqui
    Campos_Append Db, td, Campo, 8
'    Db.TableDefs.Append td
    Indice_Crear td, "Rut", "Rut", False

End If
'///////////////////////////////////////////////
If Not Tabla_Existe(Db, "Kardex") Then
    ' KARDEX ///////////
    Set td = Db.CreateTableDef("Kardex")
    Set Campo(0) = td.CreateField("Fecha", dbDate)
    Set Campo(1) = td.CreateField("Tipo", dbText, 2)
    Set Campo(2) = td.CreateField("Numero", dbDouble)
    Set Campo(3) = td.CreateField("RUT", dbText, 10)
    Set Campo(4) = td.CreateField("Precio", dbDouble)
    Set Campo(5) = td.CreateField("Cant_Entra", dbDouble)
    Set Campo(6) = td.CreateField("Cant_Sale", dbDouble)
    Set Campo(7) = td.CreateField("Cant_Saldo", dbDouble)
    Set Campo(8) = td.CreateField("Val_Entra", dbDouble)
    Set Campo(9) = td.CreateField("Val_Sale", dbDouble)
    Set Campo(10) = td.CreateField("Val_Saldo", dbDouble)
    Set Campo(11) = td.CreateField("PPP", dbDouble) ' precio promedio ponderado
    
    Campos_Append Db, td, Campo, 12
'    Db.TableDefs.Append Td
    Indice_Crear td, "Fecha-Tipo-Numero", "Fecha;Tipo;Numero", True
    
End If
'///////////////////////////////////////////////
If Not Tabla_Existe(Db, "importa_planos") Then
    
    Set td = Db.CreateTableDef("importa_planos")
    Set Campo(0) = td.CreateField("Orden", dbInteger)
    Set Campo(1) = td.CreateField("Td", dbText, 10)
    Set Campo(2) = td.CreateField("Nd", dbDouble)
    Set Campo(3) = td.CreateField("Fecha", dbDate)
    Set Campo(4) = td.CreateField("Plano", dbText, 50)
    Set Campo(5) = td.CreateField("Marca", dbText, 50)
    
    Campos_Append Db, td, Campo, 6
    
'    Db.TableDefs.Append Td
'    Indice_Crear Td, "Fecha-Tipo-Numero", "Fecha;Tipo;Numero", True
    
End If
'///////////////////////////////////////////////
If Not Tabla_Existe(Db, "oce_genera") Then
    
    Set td = Db.CreateTableDef("oce_genera")
    Set Campo(0) = td.CreateField("numero", dbDouble)
    Set Campo(1) = td.CreateField("fecha", dbDate)
    Set Campo(2) = td.CreateField("nv", dbInteger)
    Set Campo(3) = td.CreateField("obra", dbText, 30)
    Set Campo(4) = td.CreateField("subtotal", dbDouble)
    
    Campos_Append Db, td, Campo, 5
'    Db.TableDefs.Append Td
    
End If

'///////////////////////////////////////////////

If Not Tabla_Existe(Db, "prodxdesc") Then
    Set td = Db.CreateTableDef("prodxdesc")
    
    Set Campo(0) = td.CreateField("descripcion", dbText, 4)
    Set Campo(1) = td.CreateField("kgs", dbDouble)
    Set Campo(2) = td.CreateField("numero", dbLong)
    
    Campos_Append Db, td, Campo, 3
'    Db.TableDefs.Append Td
    Indice_Crear td, "descripcion", "descripcion", True

End If

'///////////////////////////////////////////////

If False Then
'If Not Tabla_Existe(Db, "chklst_formato") Then
    Set td = Db.CreateTableDef("chklst_formato")
    
    Set Campo(0) = td.CreateField("titulo", dbText, 30)
    Set Campo(1) = td.CreateField("formato", dbInteger) ' codigo formato
    
    Campos_Append Db, td, Campo, 3
'    Db.TableDefs.Append Td
    Indice_Crear td, "descripcion", "descripcion", True

End If

'///////////////////////////////////////////////

If Not Tabla_Existe(Db, "chklst") Then

    Set td = Db.CreateTableDef("chklst")

    Set Campo(0) = td.CreateField("codigoarea", dbInteger)
    Set Campo(1) = td.CreateField("fecha", dbDate)
    Set Campo(2) = td.CreateField("area", dbText, 30)
    Set Campo(3) = td.CreateField("area_responsable", dbText, 50)
    Set Campo(4) = td.CreateField("evaluacion_responsable", dbText, 50)
    Set Campo(5) = td.CreateField("cargo", dbText, 30)
    Set Campo(6) = td.CreateField("semana_fecha", dbDate)
    Set Campo(7) = td.CreateField("semana_texto", dbText, 10)
    Set Campo(8) = td.CreateField("valor1", dbInteger)
    Set Campo(9) = td.CreateField("valor2", dbInteger)
    Set Campo(10) = td.CreateField("valor3", dbInteger)
    Set Campo(11) = td.CreateField("valor4", dbInteger)
    Set Campo(12) = td.CreateField("valor5", dbInteger)
    Set Campo(13) = td.CreateField("valor6", dbInteger)

    Campos_Append Db, td, Campo, 14
'    Db.TableDefs.Append Td
    Indice_Crear td, "area-fecha", "codigoarea;fecha", False ' True

End If

If Not Tabla_Existe(Db, "chklst_obs") Then

    Set td = Db.CreateTableDef("chklst_obs")

    Set Campo(0) = td.CreateField("codigoarea", dbInteger)
    Set Campo(1) = td.CreateField("area", dbText, 30)
    Set Campo(2) = td.CreateField("mes_fecha", dbDate)
    Set Campo(3) = td.CreateField("mes_nombre", dbText, 6) ' nombre del mes y año , SEP/07
    Set Campo(4) = td.CreateField("item", dbInteger)
    Set Campo(5) = td.CreateField("obs1", dbText, 255)
    Set Campo(6) = td.CreateField("obs2", dbText, 255)
    Set Campo(7) = td.CreateField("obs3", dbText, 255)
    Set Campo(8) = td.CreateField("obs4", dbText, 255)
    Set Campo(9) = td.CreateField("obs5", dbText, 255)

    Campos_Append Db, td, Campo, 10
'    Db.TableDefs.Append Td
    Indice_Crear td, "area-mes-item", "codigoarea;mes_fecha;item", False ' True

End If

If Not Tabla_Existe(Db, "as_piezasxturno") Then

    Set td = Db.CreateTableDef("as_piezasxturno")

    Set Campo(0) = td.CreateField("orden", dbInteger)
    Set Campo(1) = td.CreateField("tipo", dbText, 1)
    Set Campo(2) = td.CreateField("descripcion", dbText, 10)
    Set Campo(3) = td.CreateField("factor", dbInteger)
    Set Campo(4) = td.CreateField("dia_cantidad", dbInteger)
    Set Campo(5) = td.CreateField("dia_metros", dbDouble)
    Set Campo(6) = td.CreateField("dia_kilos", dbDouble)
    Set Campo(7) = td.CreateField("noche_cantidad", dbInteger)
    Set Campo(8) = td.CreateField("noche_metros", dbDouble)
    Set Campo(9) = td.CreateField("noche_kilos", dbDouble)
    Set Campo(10) = td.CreateField("total_cantidad", dbInteger)
    Set Campo(11) = td.CreateField("total_metros", dbDouble)
    Set Campo(12) = td.CreateField("total_kilos", dbDouble)

    Campos_Append Db, td, Campo, 13
    Indice_Crear td, "tipo", "tipo", True

End If

If Not Tabla_Existe(Db, "as_bono") Then ' bono arco sumergido

    Set td = Db.CreateTableDef("as_bono")

    Set Campo(0) = td.CreateField("tipoestructura", dbInteger)
    Set Campo(1) = td.CreateField("rut", dbText, 10)
    Set Campo(2) = td.CreateField("nombre", dbText, 50)
    Set Campo(3) = td.CreateField("n_v", dbInteger)
    Set Campo(4) = td.CreateField("n_t", dbInteger)
    Set Campo(5) = td.CreateField("n_s", dbInteger)
    Set Campo(6) = td.CreateField("cef_1", dbInteger)
    Set Campo(7) = td.CreateField("cef_2", dbInteger)
    Set Campo(8) = td.CreateField("cef_3", dbInteger)
    Set Campo(9) = td.CreateField("cef_4", dbInteger)
    
    Set Campo(10) = td.CreateField("pte_1", dbDouble)
    Set Campo(11) = td.CreateField("pte_2", dbDouble)
    Set Campo(12) = td.CreateField("pte_3", dbDouble)
    Set Campo(13) = td.CreateField("pte_4", dbDouble)
    
    Set Campo(14) = td.CreateField("totalbono", dbDouble)

    Campos_Append Db, td, Campo, 15
    Indice_Crear td, "te-rut", "tipoestructura;rut", True

End If

'////////////////////////////////////////////////////////////////////
If Not Tabla_Existe(Db, "GD packinglist") Then
'If False Then

    Set td = Db.CreateTableDef("GD packinglist")
    
    Set Campo(0) = td.CreateField("Numero", dbDouble)
    Set Campo(1) = td.CreateField("Emision", dbDate)
    Set Campo(2) = td.CreateField("Senores", dbText, 50)
    Set Campo(3) = td.CreateField("RUT", dbText, 10)
    Set Campo(4) = td.CreateField("Direccion", dbText, 50)
    Set Campo(5) = td.CreateField("Comuna", dbText, 50)
    Set Campo(6) = td.CreateField("Telefono", dbText, 10)
    Set Campo(7) = td.CreateField("Fax", dbText, 10)
    Set Campo(8) = td.CreateField("NV", dbText, 5)
    Set Campo(9) = td.CreateField("Obra", dbText, 50)
    
    Set Campo(10) = td.CreateField("Peso Total", dbDouble)
    Set Campo(11) = td.CreateField("Precio Total", dbDouble)
    
    Set Campo(12) = td.CreateField("Obs 1", dbText, 50)
    Set Campo(13) = td.CreateField("Obs 2", dbText, 50)
    Set Campo(14) = td.CreateField("Obs 3", dbText, 50)
    Set Campo(15) = td.CreateField("Obs 4", dbText, 50)
    
    'detalle
    For i = 1 To 21
        nl = str(i)
        ent = 15 + (i - 1) * 9
        Set Campo(ent + 1) = td.CreateField("plano" & nl, dbText, 50)
        Set Campo(ent + 2) = td.CreateField("marca" & nl, dbText, 50)
        Set Campo(ent + 3) = td.CreateField("cantidad" & nl, dbDouble)
        Set Campo(ent + 4) = td.CreateField("unidad" & nl, dbText, 3)
        Set Campo(ent + 5) = td.CreateField("descripcion" & nl, dbText, 50)
        'Set Campo(ent + 5) = td.CreateField("descripcion" & nl, dbText, 20) ' desde 20/01/15
        Set Campo(ent + 6) = td.CreateField("peso unitario" & nl, dbDouble)
        Set Campo(ent + 7) = td.CreateField("peso total" & nl, dbDouble)
        Set Campo(ent + 8) = td.CreateField("precio unitario" & nl, dbDouble)
        Set Campo(ent + 9) = td.CreateField("precio total" & nl, dbDouble)
    Next
    
    Campos_Append Db, td, Campo, 196
    
End If

'////////////////////////////////////////////////////////////////////
If Not Tabla_Existe(Db, "inc_xa") Then ' infome de no conformidad x area

    Set td = Db.CreateTableDef("inc_xa")
    
    Set Campo(0) = td.CreateField("gerencia_codigo", dbText, 10)
    Set Campo(1) = td.CreateField("gerencia_descripcion", dbText, 50)
    Set Campo(2) = td.CreateField("area_codigo", dbText, 10)
    Set Campo(3) = td.CreateField("area_descripcion", dbText, 50)
    
    ' campos para resumen
    Set Campo(4) = td.CreateField("total", dbInteger) ' total de nc
    Set Campo(5) = td.CreateField("abiertas", dbInteger) ' nc abiertas
    Set Campo(6) = td.CreateField("cerradas", dbInteger) ' nc cerradas
    
    ' campos para detalle
    Set Campo(7) = td.CreateField("numero", dbInteger) ' numero de la NO Conformidad
    Set Campo(8) = td.CreateField("fecha_emision", dbDate) ' fecha emision
    Set Campo(9) = td.CreateField("condicion", dbText, 10)
    Set Campo(10) = td.CreateField("fechaPrimeraRespuesta", dbDate)
    ' numero de dias entre la fecha de emision y la fecha de primera respuesta
    Set Campo(11) = td.CreateField("diasPrimeraRespuesta", dbInteger)
    Set Campo(12) = td.CreateField("costoestimado", dbDouble)
    Set Campo(13) = td.CreateField("fecha_cierre", dbDate) ' fecha cierre
    Set Campo(14) = td.CreateField("dias", dbInteger) ' numero de dias de gestion, (abiertas)
    Set Campo(15) = td.CreateField("descripcion", dbText, 50)
    
    Campos_Append Db, td, Campo, 16
    Indice_Crear td, "gerencia-area-numero", "gerencia_codigo;area_codigo;numero", True
    
End If

If Not Tabla_Existe(Db, "generico") Then ' tabla generica para cualquier infome
    ' originalmente creada para nv
    ' 05/07/10

    Set td = Db.CreateTableDef("generico")
    
    Set Campo(0) = td.CreateField("texto10_0", dbText, 10)
    Set Campo(1) = td.CreateField("texto10_1", dbText, 10)
    Set Campo(2) = td.CreateField("texto10_2", dbText, 10)
    Set Campo(3) = td.CreateField("texto10_3", dbText, 10)
    Set Campo(4) = td.CreateField("texto10_4", dbText, 10)
    Set Campo(5) = td.CreateField("texto10_5", dbText, 10)
    Set Campo(6) = td.CreateField("texto10_6", dbText, 10)
    Set Campo(7) = td.CreateField("texto10_7", dbText, 10)
    Set Campo(8) = td.CreateField("texto10_8", dbText, 10)
    Set Campo(9) = td.CreateField("texto10_9", dbText, 10)
    Set Campo(10) = td.CreateField("texto50_0", dbText, 50)
    Set Campo(11) = td.CreateField("texto50_1", dbText, 50)
    Set Campo(12) = td.CreateField("valori_0", dbInteger)
    Set Campo(13) = td.CreateField("valord_0", dbDouble)
    
    Campos_Append Db, td, Campo, 14
    Indice_Crear td, "texto10_0", "texto10_0", True
    
End If

'/////////////////////////////////////////////
' 28/05/13
If Not Tabla_Existe(Db, "protocoloPintura") Then
    
    Set td = Db.CreateTableDef("protocoloPintura")
    
    Set Campo(0) = td.CreateField("numero", dbLong) ' numero de itop, es indice
    Set Campo(1) = td.CreateField("numeroProtocolo", dbLong) ' numero de protocolo
    Set Campo(2) = td.CreateField("pagina", dbText, 10)
    Set Campo(3) = td.CreateField("fecha", dbDate)
    Set Campo(4) = td.CreateField("responsable", dbText, 50)
    Set Campo(5) = td.CreateField("proyecto", dbText, 50)
    Set Campo(6) = td.CreateField("cliente", dbText, 50)
    Set Campo(7) = td.CreateField("esquema", dbText, 10)
    Set Campo(8) = td.CreateField("nv", dbInteger)
    Set Campo(9) = td.CreateField("probetaPrevia", dbText, 2)
    Set Campo(10) = td.CreateField("granallaMezclada", dbText, 2)
    Set Campo(11) = td.CreateField("calibre", dbText, 50)
    
    'detalle
    For i = 1 To 30
        nl = str(i)
        ent = 11 + (i - 1) * 3
        Set Campo(ent + 1) = td.CreateField("cantidad" & nl, dbDouble)
        Set Campo(ent + 2) = td.CreateField("descripcion" & nl, dbText, 50)
        Set Campo(ent + 3) = td.CreateField("marca" & nl, dbText, 25)
    Next
    
    ' 11 + 30 * 3 = 101
    Campos_Append Db, td, Campo, 102
'    Db.TableDefs.Append Td
End If

' 14/10/13
If Not Tabla_Existe(Db, "maestros") Then

    Set td = Db.CreateTableDef("maestros")

    Set Campo(0) = td.CreateField("codigo", dbText, 10)
    Set Campo(1) = td.CreateField("descripcion", dbText, 50)
    Set Campo(2) = td.CreateField("dato1", dbText, 20)
    Set Campo(3) = td.CreateField("orden", dbInteger)

    Campos_Append Db, td, Campo, 4
End If

'////////////////////////////////////////////////////////////////////
' cerado 18/07/2014
If Not Tabla_Existe(Db, "nc") Then
    Set td = Db.CreateTableDef("nc")
    
    Set Campo(0) = td.CreateField("numero", dbDouble)
    Set Campo(1) = td.CreateField("e_fecha", dbDate)
    Set Campo(2) = td.CreateField("e_nombre", dbText, 100)
    Set Campo(3) = td.CreateField("e_gerencia", dbText, 50)
    Set Campo(4) = td.CreateField("e_tipo", dbText, 50)
    Set Campo(5) = td.CreateField("e_descripcion", dbMemo)
    Set Campo(6) = td.CreateField("e_evidencia", dbText, 255)
    Set Campo(7) = td.CreateField("r_investigacion", dbMemo)
    Set Campo(8) = td.CreateField("r_accionCorrectivaOpciones", dbText, 255)
    Set Campo(9) = td.CreateField("r_accionCorrectiva", dbMemo)
    Set Campo(10) = td.CreateField("r_accionPreventiva", dbMemo)
    Set Campo(11) = td.CreateField("r_evidencia", dbText, 255)
    Set Campo(12) = td.CreateField("r_comentarios", dbMemo)
    Set Campo(13) = td.CreateField("r_encargadoNombre", dbText, 100)
    Set Campo(14) = td.CreateField("r_gerenciaNombre", dbText, 100)
    Set Campo(15) = td.CreateField("r_areaNombre", dbText, 100)
    Set Campo(16) = td.CreateField("e_fechaCierre", dbDate)

    Campos_Append Db, td, Campo, 17
    
End If

Db.Close

End Sub
Public Function Documento_Numero_Nuevo_PG(TipoDoc As String, RsITOpgc As Recordset) As Long
' busca nuevo correlativo para galvanizado o pintura (que usan el mismo archivo)
Documento_Numero_Nuevo_PG = 0
On Error GoTo Sigue
If TipoDoc = "G" Then ' ito galvanizado
    RsITOpgc.Seek "<", "P"
    Documento_Numero_Nuevo_PG = RsITOpgc!Numero
End If
If TipoDoc = "P" Then ' ito pintura
    RsITOpgc.Seek "<", "R"
    Documento_Numero_Nuevo_PG = RsITOpgc!Numero
End If
If TipoDoc = "R" Then ' ito granallado
    RsITOpgc.Seek "<", "S"
    Documento_Numero_Nuevo_PG = RsITOpgc!Numero
End If
If TipoDoc = "S" Then ' ito granallado especial
    RsITOpgc.Seek "<", "T"
    If RsITOpgc!Tipo = "S" Then
        Documento_Numero_Nuevo_PG = RsITOpgc!Numero
    End If
End If
If TipoDoc = "T" Then ' ito produccion pintura
    RsITOpgc.Seek "<", "U"
    Documento_Numero_Nuevo_PG = RsITOpgc!Numero
End If
If TipoDoc = "U" Then ' ito produccion pintura especial
    RsITOpgc.Seek "<", "Z" ' Z dummy
    If RsITOpgc!Tipo = "U" Then
        Documento_Numero_Nuevo_PG = RsITOpgc!Numero
    End If
End If
Sigue:
On Error GoTo 0
Documento_Numero_Nuevo_PG = Documento_Numero_Nuevo_PG + 1
End Function
Public Sub Doc_Actualizar_FaltaContratista(Nv As Double, Dbm As Database) ', RsPd As Recordset, RsOTfd As Recordset, RsITOfd As Recordset)
' acutaliza documentos en base a movimientos
' 1) actualiza planos
' 2) doc: puede ser OT, ITO, GD, etc
Dim RsOTfd As Recordset, RsITOfd As Recordset
Dim m_nOT As Double, m_Recibido As Integer

Set RsOTfd = Dbm.OpenRecordset("SELECT * FROM [ot fab detalle] WHERE nv=" & Nv)

Set RsITOfd = Dbm.OpenRecordset("ito fab detalle")
'RsITOfd.Index = "nv-plano-marca"
RsITOfd.Index = "ot"

' actualiza cantidad recibida en ot fab detalle
Dbm.Execute "UPDATE [ot fab detalle] SET [cantidad recibida]=0 WHERE nv=" & Nv

With RsOTfd
Do While Not .EOF
   m_nOT = !Numero
   m_Recibido = 0
   RsITOfd.Seek "=", m_nOT
   If Not RsITOfd.NoMatch Then
      Do While Not RsITOfd.EOF
         If RsITOfd![Numero OT] <> m_nOT Then Exit Do
         
         If RsITOfd!Plano = RsOTfd!Plano And RsITOfd!Marca = RsOTfd!Marca Then
            m_Recibido = m_Recibido + RsITOfd!Cantidad
         End If
         
         RsITOfd.MoveNext
         
      Loop
   End If
   
   .Edit
   ![Cantidad Recibida] = m_Recibido
   .Update
   
   .MoveNext
   
Loop
End With

End Sub
Public Sub PlanoDetalle_Actualizar(Dbm As Database, Nv As Double, NvArea As Integer, RsDoc As Recordset, indice As String, Campo As String, Optional TipoDoc As String)

' actualiza tabla "planos detalle" de acuerdo a detalle de documento, solo lo hace con un campo

Dim RsPd As Recordset, m_Plano As String, m_Marca As String, m_Cant As Integer, Indice_Actual As String

Indice_Actual = RsDoc.Index
RsDoc.Index = indice

Set RsPd = Dbm.OpenRecordset("SELECT * FROM [planos detalle] WHERE nv=" & Nv)
With RsPd
Do While Not .EOF

    m_Plano = !Plano
    m_Marca = !Marca
    m_Cant = 0
    
'If m_Marca = "V34" Then
'MsgBox ""
'End If
    Select Case Campo
    Case "ot fab", "ito pyg"
        RsDoc.Seek "=", Nv, 0, m_Plano, m_Marca 'ot
    Case "ito gr", "ito pp"
        RsDoc.Seek "=", Nv, NvArea, m_Plano, m_Marca
    Case "ito fab"
        RsDoc.Seek "=", Nv, NvArea, m_Plano, m_Marca 'itof
    End Select
    
    If Not RsDoc.NoMatch Then
    
        Do While Not RsDoc.EOF
        
            If Nv <> RsDoc!Nv Or m_Plano <> RsDoc!Plano Or m_Marca <> RsDoc!Marca Then Exit Do
            
            If IsMissing(TipoDoc) Or TipoDoc = "" Then
                m_Cant = m_Cant + RsDoc!Cantidad
            Else
            
                If TipoDoc = RsDoc!Tipo Then
                    m_Cant = m_Cant + RsDoc!Cantidad
                End If
            
            End If
            
            RsDoc.MoveNext
            
        Loop
    End If
    
    .Edit
    RsPd(Campo) = m_Cant
    .Update
    
    .MoveNext
    
Loop
End With

RsDoc.Index = Indice_Actual

End Sub
Public Sub Track_Registrar(Doc_Tipo As String, Doc_Numero As Double, Operacion As String)
' graba un registro en la tabla "track"

'Exit Sub

Dim ac(7, 1) As String
Dim av(7) As String

ac(1, 0) = "fechahora"
ac(1, 1) = "'"
ac(2, 0) = "estacion"
ac(2, 1) = "'"
ac(3, 0) = "usuario_win"
ac(3, 1) = "'"
ac(4, 0) = "usuario_scp"
ac(4, 1) = "'"
ac(5, 0) = "documento_tipo"
ac(5, 1) = "'"
ac(6, 0) = "documento_numero"
ac(6, 1) = ""
ac(7, 0) = "operacion"
ac(7, 1) = "'"
'    av(1) = "20100301.5" ' formato ansi' ok para tipo de datos datetime
    av(1) = Format(Date, "yyyymmdd") & " " & Format(time, "hh:mm:ss") ' ok
'av(1) = Format(Date, "yyyymmdd") & " " & Format(Time, "h:m:s") ' ok
'    av(1) = "20100301 12:34:56:00" ' no funciona
'    av(1) = "20100301 12:34" ' ok
'    av(1) = "20100301 12:34:56" ' ok
av(2) = Computador_Nombre()
av(3) = Usuario_Win()
av(4) = Usuario.nombre
av(5) = Doc_Tipo ' varchar(4)
av(6) = Doc_Numero
av(7) = Operacion ' varchar(3)

Registro_Agregar CnxSqlServer_scp0, "track", ac, av, 7
    
End Sub
Public Function Arreglo_DescripcionBuscar(arreglo, Codigo As String) As String
' busca descripcion en arreglo de dos dmensiones
' ejemplo:
' arreglo(1,0)="codigo1"
' arreglo(1,1)="descripcion1"
' arreglo(2,0)="codigo2"
' arreglo(2,1)="descripcion2"
' se busca codigo y se devuelve descripcion
Dim i As Integer
Arreglo_DescripcionBuscar = "CODIGO NO ENCONTRADO"
For i = 1 To 999
    If arreglo(i, 0) = Codigo Then
        Arreglo_DescripcionBuscar = arreglo(i, 1)
        Exit For
    End If
Next
End Function
Public Sub OT_Detalle_Recalcular_NEW(RsOTd As Recordset, RsITOfd As Recordset, m_Nv As Double, m_NvArea As Integer, N_Ot As Double)
' recalcula piezas recibidas en una sola OT
Dim m_recibidas As Integer
RsOTd.Seek "=", N_Ot, 1
If Not RsOTd.NoMatch Then
    Do While Not RsOTd.EOF
        If RsOTd!Numero <> N_Ot Then Exit Do
        ' busca itos
        m_recibidas = 0
        RsITOfd.Seek "=", m_Nv, m_NvArea, RsOTd!Plano, RsOTd!Marca
        If Not RsITOfd.NoMatch Then
            Do While Not RsITOfd.EOF
                If RsITOfd!Nv <> m_Nv Or RsITOfd!Plano <> RsOTd!Plano Or RsITOfd!Marca <> RsOTd!Marca Then Exit Do
                    If RsITOfd![Numero OT] = N_Ot Then
                        m_recibidas = m_recibidas + RsITOfd!Cantidad
                    End If
                RsITOfd.MoveNext
            Loop
        End If
        
        RsOTd.Edit
        RsOTd![Cantidad Recibida] = m_recibidas
        RsOTd.Update
        '
        RsOTd.MoveNext
        
    Loop
End If
End Sub
Public Sub ProdxDesp()
' muestra productos por despachar
Dim Dbm As Database, RsPd As Recordset
Set Dbm = OpenDatabase(mpro_file)
Set RsPd = Dbm.OpenRecordset("SELECT * FROM [planos detalle] WHERE [ito pyg]>[GD]ORDER BY nv")
With RsPd
Do While Not .EOF
    Debug.Print !Nv, !Marca
    .MoveNext
Loop
End With
End Sub
Public Sub Pd_Actualizar(Nv As Double)

' actualiza cantidades en plano detalle de acuerdo a movimiento de documentos de la nv

Dim m_Total As Integer, i As Integer

Dim t_OTf As Integer, t_ITOf As Integer, t_GDg As Integer, t_ITOp As Integer, t_ITOr As Integer, t_ITOt As Integer, t_GD As Integer
Dim Suma_Cant_Recib As Integer

Dim Dbm As Database, RsNVc As Recordset, RsNvPla As Recordset, RsPd As Recordset, RsOTfd As Recordset, RsITOfd As Recordset, RsITOpgd As Recordset, RsGDc As Recordset, RsGDd As Recordset

Dim m_Nv As Double, m_NvArea As Integer
m_Nv = Nv

'Set DbD = OpenDatabase(data_file)
'Set RsSc = DbD.OpenRecordset("Contratistas")
'RsSc.Index = "RUT"
'Set RsCl = DbD.OpenRecordset("Clientes")
'RsCl.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)

Dbm.Execute "UPDATE [planos detalle] SET [OT fab] = 0, [ITO fab] = 0, [GD gal] = 0, [ITO pyg] = 0, [ito gr] = 0, [ITO pp] = 0, GD = 0 WHERE nv=" & m_Nv

'Set RsNVc = DbM.OpenRecordset("NV Cabecera")
'RsNVc.Index = "Numero"

'Set RsNvPla = DbM.OpenRecordset("Planos Cabecera")
'RsNvPla.Index = "NV-Plano"

'Set RsPd = DbM.OpenRecordset("Planos Detalle")
'RsPd.Index = "NV-Plano-Marca"

Set RsPd = Dbm.OpenRecordset("SELECT * FROM [Planos Detalle] WHERE nv=" & m_Nv & " ORDER BY [plano],[marca]")

Set RsOTfd = Dbm.OpenRecordset("OT Fab Detalle")
RsOTfd.Index = "NV-Plano-Marca"

Set RsITOfd = Dbm.OpenRecordset("ITO Fab Detalle")
RsITOfd.Index = "NV-Plano-Marca"

Set RsITOpgd = Dbm.OpenRecordset("ITO pg Detalle")
RsITOpgd.Index = "NV-Plano-Marca"

Set RsGDc = Dbm.OpenRecordset("GD Cabecera")
RsGDc.Index = "Numero"
Set RsGDd = Dbm.OpenRecordset("GD Detalle")
RsGDd.Index = "NV-Plano-Marca"

t_OTf = 0
t_ITOf = 0
t_GDg = 0
t_ITOp = 0
t_ITOr = 0
t_ITOt = 0
t_GD = 0


Do While Not RsPd.EOF

Debug.Print RsPd!Nv, RsPd!Plano, RsPd!Marca

If m_Nv = RsPd!Nv Then

' OTf
m_Total = 0
RsITOfd.Index = "ot"
With RsOTfd
.Seek "=", m_Nv, m_NvArea, RsPd!Plano, RsPd!Marca
If Not .NoMatch Then
    Do While Not .EOF
        If !Nv <> m_Nv Then Exit Do
        If RsPd!Plano = !Plano And RsPd!Marca = !Marca Then
        
            ' actualiza cantidad recibida
            Suma_Cant_Recib = 0
            RsITOfd.Seek "=", !Numero
            If Not RsITOfd.NoMatch Then
               Do While Not RsITOfd.EOF
                  If !Numero <> RsITOfd![Numero OT] Then Exit Do
'                  If !marca <> RsITOfd![marca] Then
                  If !Marca = RsITOfd![Marca] Then
                     Suma_Cant_Recib = Suma_Cant_Recib + RsITOfd!Cantidad
                  End If
                  RsITOfd.MoveNext
               Loop
            End If
            If Suma_Cant_Recib > 0 Then
               .Edit
               ![Cantidad Recibida] = Suma_Cant_Recib
               .Update
            End If
 
            m_Total = m_Total + !Cantidad
            
        End If
        RsOTfd.MoveNext
    Loop
    
RsITOfd.Index = "nv-plano-marca"

End If
End With
t_OTf = m_Total

' ITOf
m_Total = 0


With RsITOfd
.Index = "nv-plano-marca"
'.Seek ">=", NVnumero.Caption, m_Plano, Marca.Caption
'.Seek "=", m_Nv, m_NvArea, m_Plano, Marca.Caption
.Seek "=", m_Nv, 0, RsPd!Plano, RsPd!Marca

If Not .NoMatch Then

    Do While Not .EOF
    
        If !Nv <> m_Nv Then Exit Do
        
        If RsPd!Plano = !Plano And RsPd!Marca = !Marca Then
        
            m_Total = m_Total + !Cantidad
            
        End If
        .MoveNext
    Loop
    
End If
End With
t_ITOf = m_Total

GDg:
' GD, galvanizado
m_Total = 0
With RsGDd
.Seek "=", m_Nv, m_NvArea, RsPd!Plano, RsPd!Marca
If Not .NoMatch Then
    Do While Not .EOF
        If !Nv <> m_Nv Then Exit Do
        
        RsGDc.Seek "=", RsGDd!Numero
        If Not RsGDc.NoMatch Then
        If RsGDc!Tipo = "G" Then
        
            If RsPd!Plano = !Plano And RsPd!Marca = !Marca Then
            
                m_Total = m_Total + !Cantidad
                                        
            End If
        
        End If
        End If
        
        .MoveNext
        
    Loop
    
End If
End With
t_GDg = m_Total

ITOpg:
' ITO pintura, galvanizado, granallado y produccion pintura
Dim td(1, 4) As String

td(0, 1) = "P": td(1, 1) = "ItoP"
td(0, 2) = "G": td(1, 2) = "ItoGa"
td(0, 3) = "R": td(1, 3) = "ItoGr"
td(0, 4) = "T": td(1, 4) = "ItoPP"

With RsITOpgd

For i = 1 To 4

    m_Total = 0
    .Seek "=", m_Nv, m_NvArea, RsPd!Plano, RsPd!Marca
    
    If Not .NoMatch Then
    
        Do While Not .EOF
        
            If !Nv <> m_Nv Then Exit Do
            
            If td(0, i) = !Tipo Then
            
                If RsPd!Plano = !Plano And RsPd!Marca = !Marca Then
                
                    m_Total = m_Total + !Cantidad
        
                End If
                
             End If
             
            .MoveNext
            
        Loop
           
    End If
    
    Select Case i
    Case 1
        t_ITOp = m_Total
    Case 2
        t_ITOp = t_ITOp + m_Total ' porque itop e itog son "la misma"
    Case 3
        t_ITOr = m_Total
    Case 4
        t_ITOt = m_Total
    End Select
    
Next
End With

GD:
' GD, NO galvanizado
m_Total = 0
With RsGDd
.Seek "=", m_Nv, m_NvArea, RsPd!Plano, RsPd!Marca
If Not .NoMatch Then
    Do While Not .EOF
        If !Nv <> m_Nv Then Exit Do
        
        RsGDc.Seek "=", RsGDd!Numero
        
        If Not RsGDc.NoMatch Then
        
        If RsGDc!Tipo = "N" Then
        
            If RsPd!Plano = !Plano And RsPd!Marca = !Marca Then
                
                m_Total = m_Total + !Cantidad
                
            End If
        
        End If
        End If
        
        .MoveNext
    Loop
    
End If
End With
t_GD = m_Total

    ' actualiza cantidades en planos detalle
    With RsPd
    .Edit
    ![OT fab] = t_OTf
    ![ITO fab] = t_ITOf
    ![GD gal] = t_GDg
    ![ITO pyg] = t_ITOp
    ![ito gr] = t_ITOr
    ![ito pp] = t_ITOt
    !GD = t_GD
    .Update
    End With
    
    End If

    RsPd.MoveNext
    
Loop

End Sub
Public Sub GDDescuadradas(Nv As Double)
' busca guias de despacho descuadradas entre cabecera y detalle
' solo kilos
Dim Dbm As Database, RsGDc As Recordset, RsGDd As Recordset
Dim sql As String, pesoCabecera As Double, pesoDetalle As Double, pesoCabeceraTotal As Double, pesoDetalleTotal As Double, numeroGD As Double, primero As Boolean

Set Dbm = OpenDatabase(mpro_file)
Set RsGDc = Dbm.OpenRecordset("GD Cabecera")
RsGDc.Index = "Numero"
'sql = "SELECT numero,SUM(cantidad) * SUM([peso unitario]) AS pesoDetalle FROM [GD Detalle] WHERE nv=" & nv & " GROUP BY numero"
sql = "SELECT * FROM [GD Detalle] WHERE nv=" & Nv & " ORDER BY numero"
Debug.Print sql
Set RsGDd = Dbm.OpenRecordset(sql)
'RsGDd.Index = "NV-Plano-Marca"

primero = True
pesoDetalle = 0
pesoCabeceraTotal = 0
pesoDetalleTotal = 0

With RsGDd

Do While Not .EOF

    If primero Then
        numeroGD = !Numero
        primero = False
    End If
    
    pesoDetalleTotal = pesoDetalleTotal + !Cantidad * ![Peso Unitario]

    If numeroGD = !Numero Then
    
        pesoDetalle = pesoDetalle + !Cantidad * ![Peso Unitario]
        
    Else
        
        RsGDc.Seek "=", numeroGD
        If RsGDc.NoMatch Then
            Debug.Print numeroGD & " no enconrada en cabecera"
        Else
        
            pesoCabecera = RsGDc![Peso Total]
            pesoCabeceraTotal = pesoCabeceraTotal + pesoCabecera
            
            If pesoCabecera = pesoDetalle Then
                Debug.Print numeroGD, pesoDetalle, pesoCabecera
            Else
                Debug.Print "ERROR ", numeroGD, pesoDetalle, pesoCabecera
            End If
        End If
        
        numeroGD = !Numero
        pesoDetalle = !Cantidad * ![Peso Unitario]
        
    End If

    .MoveNext
    
Loop

RsGDc.Seek "=", numeroGD
If RsGDc.NoMatch Then
    Debug.Print numeroGD & " no enconrada en cabecera"
Else
    pesoCabecera = RsGDc![Peso Total]
    pesoCabeceraTotal = pesoCabeceraTotal + pesoCabecera
    
    If pesoCabecera = pesoDetalle Then
        Debug.Print numeroGD, pesoDetalle, pesoCabecera
    Else
        Debug.Print "ERROR ", numeroGD, pesoDetalle, pesoCabecera
    End If
End If

End With

RsGDc.Close

Debug.Print pesoCabeceraTotal
Debug.Print pesoDetalleTotal

End Sub
Public Sub OCcertificadosRecibidos()
' mayo 2014
' certificados recibidos ahora van en oc detalle
Dim Dba As Database
Set Dba = OpenDatabase(Madq_file)
Dba.Execute "UPDATE documentos SET certificadoRecibido=1 WHERE fecha<=#04/01/2014#" ' 01 abril 2014
Dba.Close
End Sub
Public Sub certificadosNoRecibidos()

Dim Db As Database, sql As String, Oc As String
Set Db = OpenDatabase(Madq_file)
Open "c:\scp_1308\ocpendientes.txt" For Input As #1
Do While Not EOF(1)

    Line Input #1, Oc
    sql = "UPDATE documentos SET certificadoRecibido=0 WHERE oc=" & Oc
    Debug.Print sql
    Db.Execute sql
    
Loop
End Sub
