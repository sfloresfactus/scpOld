Attribute VB_Name = "Emd_Util"
Option Explicit
Dim Db As Database, Rs As Recordset
Dim item As Integer
Public Sub Ascii_Traspasa()
GoTo Sigue
Dim Reg As String
Set Db = OpenDatabase("D:\emd\planos.mdb")
Set Rs = Db.OpenRecordset("Plano")
item = 0
' LEE ARCHIVO ASCII
Open "D:\EMD\ROM.TXT" For Input As #1
Do While Not EOF(1)
    Line Input #1, Reg
    'Parentesis_Show Reg
    Conjuntos Reg
    
    ' crea linea en blanco
    item = item + 1
    Rs.AddNew
    Rs("Item") = item
    Rs.Update
    
Loop
Close #1
Sigue:
End Sub
Private Sub Parentesis_Show(Reg As String)
' muestra posiciones de los paréntesis
Dim largo As Integer, c As Integer, posi As String
Dim p_ini As Integer, ant As Integer
Dim p As String

largo = Len(Reg)
p_ini = 0: ant = 0
p = ""

For c = 1 To largo

    If Mid(Reg, c, 1) = "(" Then
        p_ini = p_ini + 1
        posi = posi & Space(c - ant - 1) & p_ini
        ant = c
        p = p & p_ini
        
    End If
    
    If Mid(Reg, c, 1) = ")" Then
        posi = posi & Space(c - ant - 1) & p_ini
        p = p & p_ini
        p_ini = p_ini - 1
        ant = c
    End If
Next
'Debug.Print Reg
'Debug.Print posi
'Debug.Print p
End Sub
Private Sub Conjuntos(Reg As String)
' Reg : registro
' busca conjuntos entre paréntesis
Dim largo As Integer, pos_pare As Integer, c As Integer, Conj As String

largo = Len(Reg)
pos_pare = 0
'Debug.Print Reg
For c = 1 To largo

    If Mid(Reg, c, 1) = "(" Or Mid(Reg, c, 1) = ")" Then
        If pos_pare + 1 <> c Then
            Conj = Mid(Reg, pos_pare, c - pos_pare + 1)
            If Left(Conj, 1) = "(" Then
'                Debug.Print Conj
                Traspasa Conj
            End If
        End If
        pos_pare = c
    End If
    
Next
End Sub
Private Function Comillas_Cuenta(paren As String)
' cuenta las comillas "
Dim largo As String, c As Integer, cc As Integer
largo = Len(paren)
cc = 0
For c = 1 To largo
    If Mid(paren, c, 1) = Chr(34) Then cc = cc + 1
Next
Comillas_Cuenta = cc
End Function
Private Function Parametros_Cuenta(Parentesis As String)
Dim Nc As Integer, Np As Integer
Nc = Comillas_Cuenta(Parentesis)
Np = -1
Select Case Nc
Case 0
    ' (0 0 0 ... 0)
    Np = 0
Case 2
    ' 1 parámetro
    ' 0 : material
    Np = 1
Case 8

    ' 7 parámetros
    
    ' 0 : cantidad total
    ' 1 : material
    ' 2 : peso unitario
    ' 3 : peso total
    ' 4 : obs plano-marca
    ' 5 : marca
    ' 6 : observaciones
    
    Np = 7
Case 12

    ' 14 parámetros
    
    ' 0 : cantidad de material
    ' 1 : cantidad parcial
    ' 2 : sección
    ' 3 : largo
    ' 4 : marca
    ' 5 : peso unitario
    ' 6 : peso total
    ' 7 : observación
    ' 8 : icha          *
    ' 9 : cintac        *
    '10 : especial      *
    '11 : sección
    '12 : designación
    '13 : solo/para     *
    
    Np = 14
Case 14

    ' 17 parámetros
    
    ' 0 : código material
    ' 1 : cantidad total
    ' 2 : cantidad parcial
    ' 3 : sección
    ' 4 : largo
    ' 5 : marca
    ' 6 : peso unitario
    ' 7 : peso total
    ' 8 : observación plano+marca
    ' 9 : útimo         *
    '10 : icha          *
    '11 : cintac        *
    '12 : especial      *
    '13 : sección       *
    '14 : observación
    '15 : designación
    '16 : solo/para     *

    Np = 17
End Select
Parametros_Cuenta = Np
End Function
Private Sub Traspasa(Conj As String)
' Conj Viene con parentesis, ej: (cvbcxcx) ó (dfgsdfsdf(
Dim pi As Integer  ' posición inicial del campo
Dim pf As Integer  ' posición final del delimitador, ya sea, chr(32) o chr(34)
Dim Np As Integer, doble As Double
Np = Parametros_Cuenta(Conj)
If Np <> 0 Then item = item + 1
Select Case Np
Case 1

    Rs.AddNew
    Rs("Item") = item
    
    ' material
    pi = InStr(1, Conj, Chr(34)) + 1
    pf = InStr(pi, Conj, Chr(34))
    Rs("Material") = Mid(Conj, pi, pf - pi)
    
    Rs.Update
    
Case 7

    Rs.AddNew
    ' item
    Rs("Item") = item
    ' cantidad total
    pi = 2
    pf = InStr(pi, Conj, " ")
    Rs("Cantidad Total") = Mid(Conj, pi, pf - pi)
    ' material
    pi = InStr(1, Conj, Chr(34)) + 1
    pf = InStr(pi, Conj, Chr(34))
    Rs("Material") = Mid(Conj, pi, pf - pi)
    ' peso unitario
    pi = pf + 2
    pf = InStr(pi, Conj, " ")
    doble = CDbl(Val(Mid(Conj, pi, pf - pi)))
    Rs("Peso Unitario") = doble
    ' peso total
    pi = pf + 1
    pf = InStr(pi, Conj, " ")
    doble = CDbl(Val(Mid(Conj, pi, pf - pi)))
    Rs("Peso Total") = doble
    ' observaciones
    pi = pf + 2
    pf = InStr(pi, Conj, Chr(34))
    Rs("Observaciones") = Mid(Conj, pi, pf - pi)
    
    Rs.Update
    
Case 14

    Rs.AddNew
    ' item
    Rs("Item") = item
    ' cantidad parcial
    pi = InStr(3, Conj, Chr(34)) + 2
    pf = InStr(pi, Conj, " ")
    doble = CDbl(Val(Mid(Conj, pi, pf - pi)))
    Rs("Cantidad Parcial") = doble
    ' material
    pi = pf + 2
    pf = InStr(pi, Conj, Chr(34))
    Rs("Material") = Mid(Conj, pi, pf - pi)
    ' largo
    pi = pf + 2
    pf = InStr(pi, Conj, " ")
    doble = CDbl(Val(Mid(Conj, pi, pf - pi)))
    Rs("Largo") = doble
    ' marca
    pi = pf + 2
    pf = InStr(pi, Conj, Chr(34))
    Rs("Marca") = Mid(Conj, pi, pf - pi)
    ' peso unitario
    pi = pf + 2
    pf = InStr(pi, Conj, " ")
    doble = CDbl(Val(Mid(Conj, pi, pf - pi)))
    Rs("Peso Unitario") = doble
    ' peso total
    pi = pf + 1
    pf = InStr(pi, Conj, " ")
    doble = CDbl(Val(Mid(Conj, pi, pf - pi)))
    Rs("Peso Total") = doble
    ' observaciones
    pi = pf + 2
    pf = InStr(pi, Conj, Chr(34))
    Rs("Observaciones") = Mid(Conj, pi, pf - pi)
    
    Rs.Update
    
Case 17

    If Mid(Conj, 2, 1) <> Chr(34) Then Exit Sub
    Rs.AddNew
    ' item
    Rs("Item") = item
    ' cantidad total
    pi = InStr(1, Conj, " ") + 1
    pf = InStr(pi, Conj, " ")
    doble = CDbl(Val(Mid(Conj, pi, pf - pi)))
    Rs("Cantidad Total") = doble
    ' cantidad parcial
    pi = pf + 1
    pf = InStr(pi, Conj, " ")
    doble = CDbl(Val(Mid(Conj, pi, pf - pi)))
    Rs("Cantidad Parcial") = doble
    ' material
    pi = pf + 2
    pf = InStr(pi, Conj, Chr(34))
    Rs("Material") = Mid(Conj, pi, pf - pi)
    ' largo
    pi = pf + 2
    pf = InStr(pi, Conj, " ")
    doble = CDbl(Val(Mid(Conj, pi, pf - pi)))
    Rs("Largo") = doble
    ' marca ??
    pi = pf + 2
    pf = InStr(pi, Conj, Chr(34))
'    Rs("Marca") = Mid(Conj, pi, pf - pi)
    ' peso unitario
    pi = pf + 2
    pf = InStr(pi, Conj, " ")
    doble = CDbl(Val(Mid(Conj, pi, pf - pi)))
    Rs("Peso Unitario") = doble
    ' peso total
    pi = pf + 1
    pf = InStr(pi, Conj, " ")
    doble = CDbl(Val(Mid(Conj, pi, pf - pi)))
    Rs("Peso Total") = doble
    ' observaciones
    pi = pf + 2
    pf = InStr(pi, Conj, Chr(34))
    Rs("Observaciones") = Mid(Conj, pi, pf - pi)
    
    Rs.Update
End Select
End Sub

