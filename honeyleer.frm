VERSION 5.00
Begin VB.Form HoneyLeer 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "HoneyLeer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
' LEE ARCHIVO ASCII
Dim Reg As String, i As Integer
Open "F:\honeywell\FMC(2).pwi" For Input As #1
'Open "F:\honeywell\test.txt" For Input As #1
i = 0
Do While Not EOF(1)
    i = i + 1
    Line Input #1, Reg
    If False Then
        Reg = Replace(Reg, Chr(0), "") ' nul
        Reg = Replace(Reg, Chr(1), "") ' soh
        Reg = Replace(Reg, Chr(2), "") ' stx
        Reg = Replace(Reg, Chr(4), "") ' eot
        Reg = Replace(Reg, Chr(6), "") ' ack
        Reg = Replace(Reg, Chr(7), "") ' bell
        Reg = Replace(Reg, Chr(8), "") ' bs
        Reg = Replace(Reg, Chr(15), "") ' si
        Reg = Replace(Reg, Chr(21), "") ' nak
        Reg = Replace(Reg, Chr(23), "") ' etb
        Reg = Replace(Reg, Chr(29), "") ' gs
    End If
    For c = 0 To 31
        Reg = Replace(Reg, Chr(c), "")
    Next
    Debug.Print i & "|" & Reg
Loop
Close #1
End Sub
Private Function Replace(Texto As String, CaracterViejo As String, Optional CaracterNuevo As String)
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
