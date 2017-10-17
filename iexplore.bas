Attribute VB_Name = "iexplore"
Option Explicit
'///////////////
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Sub GoURL(Destination As Variant)
Dim hWnd As Long
Dim msg As String
'    On Error GoTo ErrHandler
    'check and see if there is a valid link
    If Destination = "" Then
        'The programmer did not enter a link
        Err.Raise 100
    End If
        'execute the link
        ShellExecute hWnd, "open", Destination, vbNullString, vbNullString, 0 ' conSwNormal
'        ShellExecute hWnd, "open", "file:///d:/aros/help/wth.chm", vbNullString, vbNullString, 0 ' conSwNormal
    Exit Sub
ErrHandler:
    msg = "Error" & Chr(10) & Destination ' no se puede abrir URL
    MsgBox msg, vbCritical
End Sub
