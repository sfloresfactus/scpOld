Attribute VB_Name = "email"
Option Explicit
' requiere formulario dummy "email_frm" donde esta el control inet
Public Sub Email_Generar(destinatario As String, Nv As String, pla As String, Rev As String, mar As String, des As String, can As String, pun As String)
' genera email

Dim ServidorHTTP As String, intranet

Dim pagina As String ' la que realiza envio
Dim Parametros As String, textoWeb As String, txt As String

email_frm.Inet.Protocol = icHTTP

intranet = ReadIniValue(Path_Local & "scp.ini", "Path", "intranet_server")
'ServidorHTTP = "HTTP://acr3006-dualpro/intranet/"
ServidorHTTP = intranet & "intranet/"

pagina = "emailSend.asp"
Parametros = "?to=" & destinatario
Parametros = Parametros & "&nv=" & Nv
Parametros = Parametros & "&pla=" & pla
Parametros = Parametros & "&rev=" & Rev
Parametros = Parametros & "&mar=" & mar
Parametros = Parametros & "&des=" & des
Parametros = Parametros & "&can=" & can
Parametros = Parametros & "&pun=" & pun
textoWeb = ServidorHTTP & pagina & Parametros
Debug.Print textoWeb
txt = email_frm.Inet.OpenURL(textoWeb)
Debug.Print txt
If Left(txt, 2) = "OK" Then
'    MsgBox "Correo enviado Exitosamente"
Else
    MsgBox "NO se pudo enviar correo a :" & destinatario
End If

End Sub
