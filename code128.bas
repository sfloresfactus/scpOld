Attribute VB_Name = "code128"
' modulo de rutinas de codigo de barra code 128
' debe estar instalado el font code128.ttf en c:\windows\fonts
' este archivo de fonts tiene el codigo de barra correspondiente a cada codigo ascii
' instrucciones de uso:
'   en el modulo ppal debe poner linea:
'   c128_Poblar
Option Explicit
Public c128(107) As String
Public Sub c128_Poblar()
' puebla arreglo de codigos de barra
Dim i As Integer

For i = 0 To 94
    c128(i) = Chr(i + 32)
Next
For i = 95 To 102
    c128(i) = Chr(i + 100)
Next

' ejemplo de impresion de codigo de barras
If False Then
    Dim p As Printer
    Set p = Printer
    p.Font.Name = "code 128"
    'p.Font.Size = 24
    p.Font.Size = 40 ' 48 '36
    'p.Print txt2code128("768/001-F07/A/VB4")
    p.Print txt2code128("ST-302-2122B-Z-ST-3020/2122B")
    'p.Print txt2code128("A00")
    p.EndDoc
    Set p = Nothing
End If

End Sub
Public Function txt2code128(datos As String) As String
Dim largo As Integer, i As Integer, suma As Long, n As Integer, resto As Integer, checksum As String
' transforma string en codigo 128
' usando tabla B
' el codigo 128 se arma con:
' caracter de comienzo - datos - checksum - caracter de parada
' caracter de comienzo es chr(204) para start B (tabla B)
' caracter de parada   es chr(206)

' el checksum se calcula de la siguiente manera:
' ejemplo: "ZB65"
' 104 + ( 1 x 58 ) + ( 2 x 34 ) + ( 3 x 22 ) + ( 4 x 21 ) = 380
' 104       58           68           66           84     = 380
' 380/103 = 3    resto = 71 -> "g"

largo = Len(datos)
suma = 104 ' start B
For i = 1 To largo
    n = Asc(Mid(datos, i, 1)) - 32
    suma = suma + n * i
'    Debug.Print i, n
Next

n = Int(suma / 103)
resto = suma - (n * 103)

'Debug.Print n
'Debug.Print suma
'Debug.Print resto

checksum = c128(resto)
'Debug.Print checksum

txt2code128 = Chr(204) & datos & checksum & Chr(206)

End Function
'////////////////////////////////////////////////////////////////////
' valor | Tabla | ascii | patron
'       |   B   | code  |
'-------|-------|-------|-------------
'   0   | space |  032  | 11011001100
'   1   |   !   |  033  | 11001101100
'   2   |   "   |  034  |
'   3   |   #   |  035  |
'   4   |   $   |  036  |
' etc
