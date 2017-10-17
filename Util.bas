Attribute VB_Name = "Util"
' Referencias: "DAO 3.5"
' Componentes: "MS Common Controls 5.0" NO DEBE IR
' Componentes: "MS Windows Common Controls 6.0 (SP6)" desde 24/03/11

' Rutinas Útiles para Visual Basic
Option Explicit

Public Type EnterPrice
    Rut As String
    Razon As String
    Fantasia As String
    Giro As String
    Direccion As String
    Comuna As String
    Ciudad As String
    CodigoPostal As String
    Telefono1 As String
    Telefono2 As String
    Telefono3 As String
    Telefono4 As String
End Type

Public Type Usr
    nombre As String
    clave As String
    Descripcion As String ' 20/11/2012
    ReadOnly As Boolean
    AccesoTotal As Boolean
    indice As Long
    ObrasTerminadas As Boolean
    Adquis_Actual As Boolean
    Nv_Activas As Boolean ' 20/07/2004
    Nv_Orden As String
    tipo As String ' "C", contratista, "N" normal
    Rut As String '  rut del contratista
    Scc_Mod As Boolean ' true,S o false,N, indica si usuario puede modificar datos scc
End Type

'Public miColección As New Collection 'ok
Public Fecha_Format As String, Fecha_Vacia As String, Fecha_Mask As String
Public Hora_Format As String, Hora_Vacia As String, Hora_Mask As String
Public num_fmtgrl As String
Public myPC_file As String
'Public Drive_Server As String, path_mdb As String

'privilegios de usuarios, para opciones de menu
Public Privi(12, 25)
Public Menu_Columnas As Integer, Menu_Filas As Integer
'Public suma(9) As Double

'funciones en Win32api.txt
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Uni As String, Dec As String, Cen As String

' para dejar impresora predeterminada
Public cSetPrinter As New cSetDfltPrinter
' modulo de clase: cSetDfltPrinter

' funciones para locale, ej: formato fecha corto windows//////////
Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Declare Function GetUserDefaultLCID% Lib "kernel32" ()
' ejemplo
'Formato = Get_locale(LOCALE_SSHORTDATE)
Public Const LOCALE_SSHORTDATE = &H1F

Public Const testeo As Boolean = False

'Public aNv(999, 1) As String ' arreglo de notas de venta
Public aNv(2999) As NotaVenta
Public nvTotal As Integer ' total de nv en el arreglo
' segundo parametro
' 0 : numero de nota de venta
' 1 : obra

' variable que indica que tipo de acceso se usa para Notas de Venta
' Access: acceso a tabla y rutinas en Access
' Sql: acceso a tabla y rutinas en Sql Server
Public Const NvCnx As String = "Access"
'Public Const NvCnx As String = "Sql"
'/////////////////////////////////////////////////////////////////
Public Function Get_Locale(Valor) As String  ' Retrieve the regional setting

Dim Symbol As String
Dim iRet1 As Long
Dim iRet2 As Long
Dim lpLCDataVar As String
Dim pos As Integer
Dim Locale As Long
      
Get_Locale = ""
      
Locale = GetUserDefaultLCID()

'LOCALE_SDATE is the constant for the date separator
'as stated in declarations
'for any other locale setting just change the constant

'Function can also be re-written to take the
'locale symbol being requested as a parameter

'iRet1 = GetLocaleInfo(Locale, LOCALE_SDATE, lpLCDataVar, 0)
iRet1 = GetLocaleInfo(Locale, Valor, lpLCDataVar, 0)
      
Symbol = String$(iRet1, 0)
'iRet2 = GetLocaleInfo(Locale, LOCALE_SDATE, Symbol, iRet1)
iRet2 = GetLocaleInfo(Locale, Valor, Symbol, iRet1)
pos = InStr(Symbol, Chr$(0))
If pos > 0 Then
    Symbol = Left$(Symbol, pos - 1)
'    MsgBox "Regional Setting = " + Symbol
    Get_Locale = Symbol
End If

End Function
Public Function ValLet(num)
If num = 0 Then ValLet = "": Exit Function
Dim nn As String, LPaso As String, Monto As String
Dim n210 As String, n987 As String, n654 As String, n321 As String
Uni = "UN     DOS    TRES   CUATRO CINCO  SEIS   SIETE  OCHO   NUEVE  DIEZ   ONCE   DOCE   TRECE  CATORCEQUINCE "
Dec = "DIECI    VEINTI   TREINTA  CUARENTA CINCUENTASESENTA  SETENTA  OCHENTA  NOVENTA  "
Cen = "CIENTO       DOSCIENTOS   TRESCIENTOS  CUATROCIENTOSQUINIENTOS   SEISCIENTOS  SETECIENTOS  OCHOCIENTOS  NOVECIENTOS  "
nn = Format(num, "000000000000")
n210 = Left(nn, 3): n987 = Mid(nn, 4, 3): n654 = Mid(nn, 7, 3): n321 = Right(nn, 3)

LPaso = r321(n210)
If LPaso <> "" Then
    If Val(n210) = 1 Then
        Monto = "MIL "
    Else
        Monto = LPaso & " MIL "
    End If
End If

Millones:
If LPaso <> "" Then 'HAY MILES DE MILLONES
    LPaso = r321(n987)
    Monto = Monto & LPaso & " MILLONES "
Else            ' NO HAY MILES DE MILLONES
    LPaso = r321(n987)
    If LPaso = "" Then GoTo Miles
    If Val(n987) = 1 Then
        Monto = LPaso & " MILLÓN "
    Else
        Monto = LPaso & " MILLONES "
    End If
End If

Miles:
LPaso = r321(n654)
If LPaso = "" Then GoTo Cientos
If Val(n654) = 1 Then
    Monto = Monto & " MIL "
Else
    Monto = Monto & LPaso & " MIL "
End If

Cientos:
LPaso = r321(n321)
If LPaso <> "" Then Let Monto = Monto + LPaso + " "

ValLet = Monto

End Function
Private Function r21(N21)
Dim n1 As Integer, n2 As Integer
r21 = ""
If N21 = 0 Then: Exit Function
If N21 < 16 Then r21 = RTrim(Mid$(Uni, (N21 - 1) * 7 + 1, 7)): Exit Function
If N21 = 20 Then r21 = "VEINTE": Exit Function
n2 = Val(Left(N21, 1)): n1 = Val(Right(N21, 1))
If N21 < 30 Then r21 = RTrim(Mid(Dec, (n2 - 1) * 9 + 1, 9)) + RTrim(Mid(Uni, (n1 - 1) * 7 + 1, 7)): Exit Function
If n1 = 0 Then r21 = RTrim(Mid(Dec, (n2 - 1) * 9 + 1, 9)): Exit Function
r21 = RTrim(Mid(Dec, (n2 - 1) * 9 + 1, 9)) + " Y " + RTrim(Mid(Uni, (n1 - 1) * 7 + 1, 7))
End Function
Private Function r321(n321)
Dim S21 As String, N21 As Integer, LPaso As String
Dim n3 As Integer
r321 = ""
If Val(n321) = 0 Then Exit Function
If n321 = "100" Then r321 = "CIEN": Exit Function
S21 = Right(n321, 2): N21 = Val(S21)
LPaso = r21(S21): r321 = LPaso
n3 = Val(Left(n321, 1))
If n3 = 0 Then Exit Function
If N21 = 0 Then r321 = RTrim(Mid(Cen, (n3 - 1) * 13 + 1, 13)): Exit Function
r321 = RTrim(Mid(Cen, (n3 - 1) * 13 + 1, 13)) + " " + LPaso
End Function
Public Function Replace(Texto As String, CaracterViejo As String, Optional CaracterNuevo As String)
' cambia caracter
' Ejemplo : replace("zapato","a","E") -> "zEpEto"
Dim nt As String, pos As Integer
nt = Texto
pos = 1

If IsMissing(CaracterNuevo) Then
    CaracterNuevo = ""
End If

Do While True
    pos = InStr(pos, nt, CaracterViejo)
    If pos = 0 Then Exit Do
    nt = Left(nt, pos - 1) & CaracterNuevo & Right(nt, Len(nt) - pos)
    pos = pos - 1
    If pos = 0 Then pos = 1
Loop
Replace = nt
End Function
Public Function IsObjBlanco(Objeto As Object, Descripcion As String, btnGrabar As Button) As Boolean
' valida que objeto no esté en blanco
IsObjBlanco = False
If Trim(Objeto) = "" Then
    btnGrabar.Value = tbrUnpressed
    IsObjBlanco = True
    Beep
    MsgBox Descripcion & " NO DEBE ESTAR EN BLANCO"
    Objeto.SetFocus
End If
End Function
Public Function Corchetes_Pone(Texto As String) As String
' pone corchetes a nombres de campos que lo necesiten
Corchetes_Pone = Texto
If Left(Texto, 1) = "[" Then Exit Function
If InStr(1, Texto, " ") <> 0 Then Corchetes_Pone = "[" & Texto & "]"
End Function
Public Function Form_CentraX(f As Object) As Integer
Form_CentraX = Int((Screen.Width - f.Width) / 2)
End Function
Public Function Form_CentraY(f As Object) As Integer
Form_CentraY = Int((Screen.Height - f.Height) / 2)
End Function
Public Function Fecha_Valida(Fecha As Object, Optional Asume) As Boolean
' Fecha : Objeto MaskEdit
' Asume : Fecha que se asume si se presiona ENTER, si Asume es vacío  entonces, fecha="__/__/__"
Dim s As String, f As Date

Fecha_Valida = True
s = Replace(Fecha.Text, "_", "")
If s = "//" Then
    If Not IsMissing(Asume) Then Fecha = Format(Asume, Fecha_Format)
    Exit Function
End If
On Error GoTo Error
f = CDate(s)
Fecha = Format(f, Fecha_Format)
On Error GoTo 0
Exit Function
Error:
Fecha_Valida = False
MsgBox "FECHA NO VÁLIDA"
Fecha.visible = True ' para FlexGrid
Fecha.SetFocus
End Function
Public Function EsFechaVacia(Fecha As String) As Boolean
If Fecha = Fecha_Vacia Then
    EsFechaVacia = True
Else
    EsFechaVacia = False
End If
End Function
Public Function Hora_Obj_Valida(Hora As Object, Optional Asume) As Boolean
' Fecha : Objeto MaskEdit
' Asume : Fecha que se asume si se presiona ENTER, si Asume es vacío  entonces, hora="__:__"
Dim s As String, f As Date

Hora_Obj_Valida = True
s = Replace(Hora.Text, "_", "")
If s = ":" Then
    If Not IsMissing(Asume) Then Hora = Format(Asume, "hh:mm")
    Exit Function
End If
On Error GoTo Error
f = CDate(s)
Hora = Format(f, "hh:mm")
On Error GoTo 0
Exit Function
Error:
Hora_Obj_Valida = False
MsgBox "HORA NO VÁLIDA"
Hora.visible = True
Hora.SetFocus
End Function
Public Function Hora_Txt_Valida(Hora As String) As Boolean

'If Len(txt) > 0 Then
'    p = InStr(1, txt, ":")
'    If p = 0 Then
'        txt = Detalle.TextMatrix(i, 2) & " " & Detalle.TextMatrix(i, 3) & " " & Detalle.TextMatrix(i, 4)
''                MsgBox "Formato de Horas Extra: h:mm" & vbLf & Txt
'        MsgBox "Formato de Horas: h:mm" & vbLf & txt
'        Detalle.Row = i
'        Detalle.col = 5
'        Detalle.SetFocus
'        Exit Function
'    Else
'        Dim HH As Integer, mm As Integer
'        HH = Val(Left(txt, p - 1))
'        mm = Val(Mid(txt, p + 1))
'
'        If 0 <= HH And HH < 100 Then
'            ' ok horas
'        Else
'            txt = Detalle.TextMatrix(i, 2) & " " & Detalle.TextMatrix(i, 3) & " " & Detalle.TextMatrix(i, 4)
'            MsgBox "Horas entre 0 y 99" & vbLf & txt
'            Detalle.Row = i
'            Detalle.col = 5
'            Detalle.SetFocus
'            Exit Function
'        End If
'
'        If 0 <= mm And mm < 60 Then
'            ' ok minutos
'        Else
'            txt = Detalle.TextMatrix(i, 2) & " " & Detalle.TextMatrix(i, 3) & " " & Detalle.TextMatrix(i, 4)
'            MsgBox "Minutos entre 0 y 59" & vbLf & txt
'            Detalle.Row = i
'            Detalle.col = 5
'            Detalle.SetFocus
'            Exit Function
'        End If
'
'    End If
'End If
End Function
Public Function Hora2Decimal(Hora As String) As Double
' trasforma hora (en formato hh:mm) a numero decimal
' ejemplo:  3:45 -> 3,75
' ojo: hora debe estar validada
Dim pos As String, HH As Double, mm As Double

pos = InStr(1, Hora, ":")
If pos = 0 Then
    Hora2Decimal = 0
Else
    HH = Val(Left(Hora, pos - 1))
    mm = Val(Mid(Hora, pos + 1))
    Hora2Decimal = HH + mm / 60
End If
End Function
Public Function Decimal2Hora(Hora As Double) As String
' convierte numero decimal en formato hora
Dim m_HH As Integer, m_MM As Integer, m_double As Double
m_HH = Int(Hora)
m_double = Hora - m_HH
m_MM = m_double * 60
Decimal2Hora = m_HH & ":" & m_MM
End Function
Public Function NoNulo(Texto As Variant) As String
' entrega texto no nulo
NoNulo = ""
On Error Resume Next
If IsNull(Texto) Then
    NoNulo = ""
Else
    NoNulo = Texto
End If
End Function
Public Function NoNulo_Double(Numero As Variant) As Double
' entrega texto no nulo
If IsNull(Numero) Then
    NoNulo_Double = 0
Else
    NoNulo_Double = Numero
End If
End Function
Public Function Rut_Verifica(Rut) As Boolean
Dim val_num As Integer, sum As Integer, largo As Integer, dv As String
Rut_Verifica = False
Rut = Trim(Rut)
largo = Len(Rut)

If largo < 3 Then Exit Function

If Mid(Rut, largo - 1, 1) <> "-" Then Exit Function

dv = Right(Rut, 1)
Rut = "00000000" & Left(Rut, largo - 2)
Rut = Right(Rut, 8)

For val_num = 1 To 8
    If Asc(Mid(Rut, val_num, 1)) < 48 Or Asc(Mid(Rut, val_num, 1)) > 57 Then
        Rut_Verifica = False
        Exit Function
    End If
Next val_num

sum = 0
sum = sum + (Mid(Rut, 1, 1) * 3)
sum = sum + (Mid(Rut, 2, 1) * 2)
sum = sum + (Mid(Rut, 3, 1) * 7)
sum = sum + (Mid(Rut, 4, 1) * 6)
sum = sum + (Mid(Rut, 5, 1) * 5)
sum = sum + (Mid(Rut, 6, 1) * 4)
sum = sum + (Mid(Rut, 7, 1) * 3)
sum = sum + (Mid(Rut, 8, 1) * 2)

If Mid("0K987654321", (sum Mod 11) + 1, 1) = dv Then
    Rut_Verifica = True
Else
    Rut_Verifica = False
End If
End Function
Public Function Rut_Formato(Rut As String) As String
Dim largo As Integer
Rut = Trim(Rut)
'Rut_Formato = Rut
'Exit Function ' 09/06/10
largo = Len(Rut)
'rut = Space(10) & Left(rut, largo - 1) & "-" & Right(rut, 1)
Rut = Space(10) & Rut
Rut = Right(Rut, 10)
Rut_Formato = Rut
End Function
Public Function m_CDbl(txt As String) As Double
' mi propio "convert double" 28/02/98
Dim num As Double
'Dim num As Single
txt = Replace(txt, ",", ".") ' correccion 11/04/98
num = Val(txt)
If num <> 0 Then num = CDbl(num)
m_CDbl = num
End Function
Public Sub Printer_Set(tipo As String)

Dim ImpresoraNombre As String, Path_Local As String
Path_Local = App.Path & "\" ' ojo 16/04/08
'Dim FontNombre As String
'Dim prt As Printer, i As Integer

Select Case tipo
Case "Documentos"
    ImpresoraNombre = ReadIniValue(Path_Local & "scp.ini", "Printer", "Docs")
    ' predetermina impresora
    Impresora_Predeterminada ImpresoraNombre
End Select

'Printer.PaperSize = vbPRPSLetter ' 29/07/98 por problema en impresora de pastoreli OJO 20/02/06

' setea font
Printer.Font.Name = "Courier New"

' truco para que tome el Font
Printer.Font.Bold = True
Printer.Print "";
Printer.Font.Bold = False

End Sub
Public Sub Printer_Set_OLD(tipo As String)
Dim Db As Database, Rs As Recordset
Dim ImpresoraNombre As String
Dim FontNombre As String
Dim prt As Printer, i As Integer

Set Db = OpenDatabase(myPC_file)
Set Rs = Db.OpenRecordset("Configuración Impresoras")
ImpresoraNombre = NoNulo(Rs(tipo & " Printer"))
FontNombre = NoNulo(Rs(tipo & " Font"))
Db.Close

' setea impresora
For Each prt In Printers
    If prt.DeviceName = ImpresoraNombre Then
        ' deja la impresora como predeterminada para VB
        Set Printer = prt
    End If
Next

Printer.PaperSize = vbPRPSLetter ' 29/07/98 por problema en impresora de pastoreli

' setea font
For i = 0 To Printer.FontCount - 1
    If Printer.Fonts(i) = FontNombre Then
        Printer.Font.Name = FontNombre
        Exit For
    End If
Next

' truco para que tome el Font
Printer.Font.Bold = True
Printer.Print "";
Printer.Font.Bold = False

End Sub
Public Sub Printer_Set_vb50(Empresa_RUT As String, doc As String)
Dim Db As Database, Rs As Recordset
Dim ImpresoraNombre As String
Dim FontNombre As String
Dim FontSize As Long
Dim prt As Printer, i As Integer

Set Db = OpenDatabase(myPC_file)
'Set Db = OpenDatabase(cnfg_file)
Set Rs = Db.OpenRecordset("Configuración Impresoras")
'debe ir con Indice por rut empresa para cuando existen varias empresas
Rs.Index = "RUT"
Rs.Seek "=", Empresa_RUT
If Not Rs.NoMatch Then
    
    ImpresoraNombre = Rs(doc & " Impresora")
    FontNombre = Rs(doc & " Font Name")
    FontSize = Rs(doc & " Font Size")
    
    ' setea impresora
    For Each prt In Printers
        If prt.DeviceName = ImpresoraNombre Then
            ' deja la impresora como predeterminada para VB
            Set Printer = prt
        End If
    Next
    
    Printer.PaperSize = vbPRPSLetter ' 29/07/98 por problema en impresora de pastorelli
    
    ' setea font
    For i = 0 To Printer.FontCount - 1
        If Printer.Fonts(i) = FontNombre Then
            Printer.Font.Name = FontNombre
            Printer.Font.Size = FontSize 'ojo probrar 17/06/1999
            Exit For
        End If
    Next
    
    ' truco para que tome el Font
    Printer.Font.Bold = True
    Printer.Print "";
    Printer.Font.Bold = False
    
End If

Rs.Close
Db.Close

End Sub
Public Function Archivo_Existe(Path As String, File As String) As Boolean
' 19/03/98
' verifica si el archivo FILE existe en la ruta PATH (path con \ al final)
If UCase(File) = UCase(Dir(Path & File, vbArchive)) Then
    Archivo_Existe = True
Else
    Archivo_Existe = False
End If
End Function
Public Function Archivo_Fecha(Ruta As String, Archivo As String) As Date
' 19/03/98
' entrega fecha de modificación de ARCHIVO
If Archivo_Existe(Ruta, Archivo) Then
    Archivo_Fecha = FileDateTime(Ruta & Archivo)
End If
End Function
Public Function Documento_Numero_Nuevo(Rs As Recordset, CampoNumero As String) As Double
Documento_Numero_Nuevo = 1
On Error GoTo Error
Rs.MoveLast
Documento_Numero_Nuevo = Rs(CampoNumero) + 1
Exit Function
Error:
End Function
Public Function Documento_Numero_Nuevo2(Rs As Recordset, TipoDoc As String, TipoPieza, CampoNumero As String) As Double
Documento_Numero_Nuevo2 = 1
'On Error GoTo error
Rs.Seek ">=", TipoDoc, TipoPieza, 0 ' Rs(CampoNumero), 0
If Not Rs.EOF Then
Do While Not Rs.EOF
    If Rs("doc_tipo") <> TipoDoc Then
        Exit Do
    End If
    Documento_Numero_Nuevo2 = Rs(CampoNumero) + 1
    Rs.MoveNext
Loop
End If
'Documento_Numero_Nuevo2 = Rs(CampoNumero) + 1
Exit Function
Error:
End Function
Public Function m_Format(Numero As Variant, Formato As String) As String
' formatea número (entero?) con espacios en blanco delante
' 17/04/98, rev 18/07/98
m_Format = Format(Format(Numero, Formato), String(Len(Formato), "@"))
End Function
Public Function PadR(Texto As String, largo As Integer, Optional caracter) As String
' rellena con caracteres por la derecha
If IsMissing(caracter) Then
    caracter = " "
End If
Texto = Texto & String(largo, caracter)
Texto = Left(Texto, largo)
PadR = Format(Texto, String(largo, "@"))
End Function
Public Function PadL(Texto As String, largo As Integer, Optional caracter) As String
' rellena con caracteres por la izquierda
Texto = Trim(Texto)
If IsMissing(caracter) Then
    caracter = " "
End If
Texto = String(largo, caracter) & Texto
Texto = Right(Texto, largo)
PadL = Format(Texto, String(largo, "@"))
End Function
Public Function HtoD(Hnum As String) As Integer
' convierte un numero hexadecimal en decimal
' HtoD("A") -> "10"
Dim Valor As String
HtoD = 0
On Error GoTo Error
Valor = "&H" & UCase(Hnum)
HtoD = CInt(Valor)
Error:
End Function
Public Function Mi_Hex(Dnum As Integer, ancho As Integer) As String
' mi propio conversor a hexadecimal
' ej:  Mi_Hex(10,2) -> "0A"   . Mi_Hex(11,3) -> "00B"
Mi_Hex = Hex(Dnum)
Mi_Hex = String(ancho, "0") & Mi_Hex
Mi_Hex = Right(Mi_Hex, ancho)
End Function
Public Sub Recordsets_Mostrar()
Dim Ws As Workspace
Dim n As Integer, d As Integer
Set Ws = Workspaces(0)
n = Ws.Databases.Count
For d = 0 To n - 1
    Debug.Print d, Ws.Databases(d).Name
Next
End Sub
Public Sub DataBases_Cerrar()
Dim Ws As Workspace
Set Ws = Workspaces(0)
Do While Ws.Databases.Count > 0
    Ws.Databases(0).Close
Loop
End Sub
Function Usuario_Win() As String
' creada el 03/07/98
' Entrega el nombre del usuario que entró a windows
' ej : Administrador
Dim Usr As String, Temp
Usr = String(145, Chr(0))
Temp = GetUserName(Usr, 145)
Usuario_Win = Left(Usr, InStr(Usr, Chr(0)) - 1)
End Function
Function Computador_Nombre() As String
' creada el 06/07/98
' Entrega el nombre del computador
' ej : ACR166_WS15
Dim Cptr As String, Temp
Cptr = String(145, Chr(0))
Temp = GetComputerName(Cptr, 145)
Computador_Nombre = Left(Cptr, InStr(Cptr, Chr(0)) - 1)
End Function
Function SystemDirectory() As String
' creada el 07/01/05
' Entrega el nombre de la carpeta de windows\system
' ej : c:\winnt\system32
Dim Cptr As String, Temp
Cptr = String(145, Chr(0))
Temp = GetSystemDirectory(Cptr, 145)
SystemDirectory = Left(Cptr, InStr(Cptr, Chr(0)) - 1)
End Function
Function WindowsDirectory() As String
' creada el 07/01/05
' Entrega el nombre del directorio de windows
' ej : c:\winnt
Dim Cptr As String, Temp
Cptr = String(145, Chr(0))
Temp = GetWindowsDirectory(Cptr, 145)
WindowsDirectory = Left(Cptr, InStr(Cptr, Chr(0)) - 1)
End Function
Public Sub Campos_Append(Db As Database, td As TableDef, Campo() As Field, nc As Integer)
Dim c As Integer
For c = 0 To nc - 1
    Select Case Campo(c).Type
    Case dbText
        Campo(c).AllowZeroLength = True
    Case dbInteger, dbDouble
        Campo(c).DefaultValue = 0
    Case dbDate
'        Campo(c).DefaultValue = CDate("__/__/__")
    End Select
    td.Fields.Append Campo(c)
Next
Db.TableDefs.Append td
End Sub
Public Sub Indice_Crear(Tabla As TableDef, nombre As String, Campos As String, Primario As Boolean)
Dim indice As Index
Set indice = Tabla.CreateIndex
With indice
    .Name = nombre
    .Fields = Campos
    .Primary = Primario
End With
Tabla.Indexes.Append indice
End Sub
Public Function IsNumValido(Objeto As Object, btnGrabar As Button) As Boolean
' valida que objeto no esté en blanco
IsNumValido = True
If Not Numero_Valido(Objeto.Text) Then
    btnGrabar.Value = tbrUnpressed
    IsNumValido = False
    Beep
    MsgBox "NÚMERO NO VÁLIDO"
    Objeto.SetFocus
End If

End Function
Public Function Numero_Valido(Numero As String) As Boolean
Dim n As Double
Numero_Valido = False
On Error GoTo Error
n = Val(Numero)
Numero_Valido = True
Exit Function
Error:
End Function
Public Function n2mes(mes As Integer) As String
n2mes = ""
If mes < 1 Or 12 < mes Then Exit Function
Dim m As String
m = "Enero     Febrero   Marzo     Abril     Mayo      Junio     Julio     Agosto    SeptiembreOctubre   Noviembre Diciembre "
n2mes = Trim(Mid(m, (mes - 1) * 10 + 1, 10))
End Function
Public Function Tabla_Existe(Db As Database, Tabla As String) As Boolean
Dim t As Integer, td As TableDef
Tabla_Existe = False
For t = 0 To Db.TableDefs.Count - 1
    Set td = Db.TableDefs(t)
    If td.Attributes = 0 Then
        If td.Name = Tabla Then
            Tabla_Existe = True
            Exit For
        End If
    End If
Next
End Function
Public Function n(num As String) As Double
Dim i As Integer
num = Replace(num, "_", "") 'NEW 12/12/96
i = InStr(num, ",")
If i = 0 Then
    n = Val(num)
Else
    n = Val(Left(num, i - 1) & "." & Right(num, Len(num) - i))
End If
End Function
 Public Function p_fn(Numero As Double, Formato As String) As String
' FORMATEA NÚMERO PARA SER IMPRESO
p_fn = Format(Format(Numero, Formato), ">" & String(Len(Formato), "@"))
End Function
Public Function Hora_Actual()
'05/07/1999
Hora_Actual = Format(Now, "HH:MM")
End Function
Public Function StrMirror(Texto As String) As String
' invierte texto:  "hola" -> "aloh"
Dim t As String, i As Integer
t = ""
For i = Len(Texto) To 1 Step -1
    t = t + Mid(Texto, i, 1)
Next
StrMirror = t
End Function
Public Function InStrLast(Texto As String, caracter As String) As Integer
' busca la posicion del ultimo "caracter" encontrado en "texto"
Dim i As Integer, pos As Integer
pos = 0
For i = Len(Texto) To 1 Step -1
    If Mid(Texto, i, 1) = caracter Then pos = i: Exit For
Next
InStrLast = pos
End Function
Public Sub Mdb_Compactar(Path As String, Archivo_Mdb As String)
Dim Db_Old As String, Db_New As String
Db_Old = Archivo_Mdb & ".Bak"
Db_New = Archivo_Mdb & ".Mdb"
If Archivo_Existe(Path, Db_Old) Then
    Kill Path & Db_Old
End If
Name Path & Db_New As Path & Db_Old
DBEngine.CompactDatabase Path & Db_Old, Path & Db_New
End Sub
Public Sub Enter(KeyAscii As Integer)
'If Not Win7 Then
'    If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{Tab}"
'End If
If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeysA vbKeyTab, True
End Sub
Public Function Directorio(Carpeta As String) As String
' verifica su existe carpeta, si no, va recortando path hasta llegar a "C:"
' parametro: carpeta: viene con disco y ruta, ej: c:\scp\txt\2005\etc
' creada 31/03/05
Dim i As Integer, largo As Integer
Directorio = "C:"

'Carpeta = "c:"

largo = Len(Carpeta)
For i = largo To 1 Step -1
    If Mid(Carpeta, i, 1) = "\" Then
    
        ' recorta ruta
        Carpeta = Left(Carpeta, i)
        
'        Debug.Print Dir(Carpeta, vbDirectory)
        If Dir(Carpeta, vbDirectory) <> "" Then
            ' ok
            Directorio = Carpeta
            Exit Function
            
        End If
        
    End If
Next

End Function
Public Sub Impresora_Predeterminada(NombreImpresora As String)

    If UCase(NombreImpresora) = "DEFAULT" Then
        Set cSetPrinter = Nothing
        Exit Sub
    End If

' establece impresora predeterminada
    Dim sMsg As String
    Dim NombreAparato As String
    Dim p As Printer, largo As Integer
    
    largo = Len(NombreImpresora)
    
    For Each p In Printers
        NombreAparato = p.DeviceName
'        If UCase(Left(DeviceName, 5)) = "ZEBRA" Then
        If UCase(Left(NombreAparato, largo)) = UCase(NombreImpresora) Then
            If cSetPrinter.SetPrinterAsDefault(NombreAparato) Then
                sMsg = NombreAparato & " a sido configurada como predeterminada."
            Else
                sMsg = NombreAparato & " a fallado al ser configurada como predeterminada."
            End If
'            MsgBox sMsg, vbExclamation, App.Title
'            Debug.Print sMsg
'            MsgBox sMsg & vbLf & "|" & NombreImpresora & "|" & vbLf & "|" & NombreAparato & "|==|" & p.DeviceName & "|"
            Exit For
        Else
'            MsgBox "|" & NombreAparato & "|<>|" & NombreImpresora & "|"
        End If
    Next
    
End Sub
'Public Sub Font_Setear(Impresora As Printer, Optional NombreFont As String, Optional ConPaperSize As Boolean)
Public Sub Font_Setear(Impresora As Printer, Optional NombreFont As String)

' setea tamaño del papel
'On Error Resume Next
'If ConPaperSize Then
'Impresora.PaperSize = vbPRPSLetter  ' 29/07/98 por problema en impresora de pastoreli OJO 20/02/06
'End If

' setea font
If NoNulo(NombreFont) = "" Then
    NombreFont = "Courier New"
End If
Impresora.Font.Name = NombreFont

' truco para que tome el Font
Impresora.Font.Bold = True
Impresora.Print "";
Impresora.Font.Bold = False

End Sub
Public Function CharCount(OrigString As String, Chars As String, Optional CaseSensitive As Boolean = False) As Long

'**********************************************
'PURPOSE: Returns Number of occurrences of a character or
'or a character sequencence within a string

'PARAMETERS:
    'OrigString: String to Search in
    'Chars: Character(s) to search for
    'CaseSensitive (Optional): Do a case sensitive search
    'Defaults to false

'RETURNS:
    'Number of Occurrences of Chars in OrigString

'EXAMPLES:
'Debug.Print CharCount("FreeVBCode.com", "E") -- returns 3
'Debug.Print CharCount("FreeVBCode.com", "E", True) -- returns 0
'Debug.Print CharCount("FreeVBCode.com", "co") -- returns 2
''**********************************************

Dim lLen As Long
Dim lCharLen As Long
Dim lAns As Long
Dim sInput As String
Dim sChar As String
Dim lCtr As Long
Dim lEndOfLoop As Long
Dim bytCompareType As Byte

sInput = OrigString
If sInput = "" Then Exit Function
lLen = Len(sInput)
lCharLen = Len(Chars)
lEndOfLoop = (lLen - lCharLen) + 1
bytCompareType = IIf(CaseSensitive, vbBinaryCompare, vbTextCompare)

    For lCtr = 1 To lEndOfLoop
        sChar = Mid(sInput, lCtr, lCharLen)
        If StrComp(sChar, Chars, bytCompareType) = 0 Then lAns = lAns + 1
    Next

CharCount = lAns

End Function
Public Sub split(Texto As String, separador As String, aCampos, NumeroCampos As Integer)
' un texto largo lo divide en un arreglo
' debes declarar aCampos(9) as string
' ejemplo:
'    split("pieza1 pieza2 pieza3"," ")
' aCampos[1]="pieza1"
' aCampos[2]="pieza2"
' aCampos[3]="pieza3"

'     1       2    3      4        5        6   7      8          9        10    11
' proveedor/mina/sector/marca/descripcion/peso/nv/maestranza/nombreSector/xxxx/itemOC
' xxx: no se que es

Dim largo As Integer, indice As Integer, Desde As Integer, i As Integer
Texto = Trim(Texto)

largo = Len(Texto)

indice = 0
Desde = 0

For i = 1 To NumeroCampos
    aCampos(i) = ""
Next

For i = 1 To largo
    If Mid(Texto, i, 1) = separador Then
        indice = indice + 1
        aCampos(indice) = Mid(Texto, Desde + 1, i - Desde - 1)
        Desde = i
    End If
Next

' verifica ultimo caracter no es separador
If Mid(Texto, largo, 1) <> separador Then
    indice = indice + 1
    aCampos(indice) = Mid(Texto, Desde + 1, largo - Desde)
End If

'For i = 1 To numeroCampos
'    Debug.Print i & "|" & aCampos(i) & "|"
'Next

End Sub
