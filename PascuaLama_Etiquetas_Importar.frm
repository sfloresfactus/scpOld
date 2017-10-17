VERSION 5.00
Begin VB.Form PascuaLama_Etiquetas_Importar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Etiquetas"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "PascuaLama_Etiquetas_Importar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Planilla As Object, Hoja As String
Private SqlRsSc As New ADODB.Recordset
Private ac(21, 1) As String, av(21) As String, asn(21) As Boolean
Private Sub Form_Load()
' importa desde planilla excel a sql server
' para impresion de eqtiquetas proyecto pascua lama

ac(1, 0) = "cliente"
ac(1, 1) = "'"
ac(2, 0) = "proyecto"
ac(2, 1) = "'"
ac(3, 0) = "proyecto_numero"
ac(3, 1) = "'"
ac(4, 0) = "odc_numero"
ac(4, 1) = "'"
ac(5, 0) = "mk_numero"
ac(5, 1) = "'"
ac(6, 0) = "mk_descripcion"
ac(6, 1) = "'"
ac(7, 0) = "nv_nombre"
ac(7, 1) = "'"
ac(8, 0) = "kg_unitario"
ac(8, 1) = "'"
ac(9, 0) = "cantidad"
ac(9, 1) = ""

'Excel_Leer "e:\scp-02-sql\pascualama\", "planilla.xls", "etiquetas"
'Excel_Leer2 "c:\scp-02-sql\pascualama\", "2725-2726.xls", "etiquetas"
'Excel_Leer2 "c:\scp-02-sql\pascualama\", "pascualama_2725_2726_v2.xls", "etiquetas"
'Excel_Leer2 "c:\scp-02-sql\pascualama\", "pascualama_2725_2726_v3.xls", "etiquetas"

'CnxSqlServer.Execute "delete from etiquetas_barrick"
' 03/01/2012
'Excel_Leer2 "c:\scp-02-sql\pascualama\", "pascualama_2725_2726_v4.xls", "etiquetas"
' 06/01/2012
'Excel_Leer2 "c:\scp-02-sql\pascualama\", "pascualama_2726.xls", "etiquetas"
' 13/01/2012
'Excel_Leer2 "c:\scp-02-sql\pascualama\", "pascualama_2725_2726_v5.xls", "etiquetas"
' 06/03/2012
'Excel_Leer2 "c:\scp-02-sql\pascualama\", "pascualama_120306.xls", "etiquetas"
' 06/03/2012
'Excel_Leer2 "H:\notebook_lenovo\scp-02-sql\pascualama\", "ANEXO ETIQUETAS PASCUALAMA TRIPPERCAR 2717.xls", "etiquetas"
' 23/06/2013
Excel_Leer2 "f:\scp-02-sql\pascualama\", "pascualama_130523.xls", "alexpizarro"

End Sub
Private Sub Excel_Leer(Path As String, Archivo As String, Hoja As String)
' importa datos desde planilla excel
Dim fi As Integer, co As Integer
Dim filas_vacias As Integer

Dim m_Cli As String ' cliente
Dim m_Pro As String ' proyecto
Dim m_ProNum As String ' proyecto numero
Dim m_OdcNum As String ' odc numero
Dim m_MKN As String ' MK Nº
Dim m_MKD As String ' MK Descripcion
Dim m_NVM As String ' NV + Nombre
Dim m_KgU As String ' kg unitario
Dim m_Can As String ' cantidad


If Not Archivo_Existe(Path, Archivo) Then
    MsgBox "No existe archivo " & vbLf & Path & Archivo
    Exit Sub
End If

'On Error GoTo NoExcel
Set Planilla = GetObject(Path & Archivo, "Excel.Sheet.8") 'assign sheet object as an OLE excel
'On Error GoTo 0

With Planilla.Worksheets(Hoja)

fi = 1
filas_vacias = 0

Do While True

    fi = fi + 1
    
'    m_cli = Val(Trim(.cells(fi, 1).Value))
    m_Cli = Trim(.cells(fi, 1).Value)
    
    ' si linea esta en blanco
    If Len(m_Cli) = 0 Then
        Exit Do
    End If
    
    m_Pro = Trim(.cells(fi, 2).Value)
    m_ProNum = Trim(.cells(fi, 3).Value)
    m_OdcNum = Trim(.cells(fi, 4).Value)
    m_MKN = Trim(.cells(fi, 5).Value)
    m_MKD = Trim(.cells(fi, 6).Value)
    m_NVM = Trim(.cells(fi, 7).Value)
    m_KgU = Trim(.cells(fi, 8).Value)
    m_Can = Trim(.cells(fi, 9).Value)
    
    Debug.Print m_Cli, m_Pro, m_ProNum, m_OdcNum, m_MKN, m_MKD, m_NVM, m_KgU, m_Can
    
    av(1) = m_Cli
    av(2) = m_Pro
    av(3) = m_ProNum
    av(4) = m_OdcNum
    av(5) = m_MKN
    av(6) = m_MKD
    av(7) = m_NVM
    av(8) = m_KgU
    av(9) = m_Can
    
    'Registro_Agregar CnxSqlServer, "etiquetas_barrick", ac, av, 9
        
Loop

End With

Set Planilla = Nothing

Exit Sub

NoExcel:
'MsgBox "No Tiene Instalado Microsoft Excel"
' o archivo esta abierto (en uso por ejemplo por excel)
MsgBox "Nombre de Hoja NO Válido"

End Sub
Private Sub Excel_Leer2(Path As String, Archivo As String, Hoja As String)
' importa datos desde planilla excel

' para planilla enviada poe alex pizarro, 5 agosto 2011

Dim fi As Integer, co As Integer
Dim filas_vacias As Integer

Dim m_Cli As String ' cliente
Dim m_Pro As String ' proyecto
Dim m_ProNum As String ' proyecto numero
Dim m_OdcNum As String ' odc numero
Dim m_MKN As String ' MK Nº
Dim m_MKD As String ' MK Descripcion
Dim m_NVM As String ' NV + Nombre
Dim m_KgU As String ' kg unitario
Dim m_Can As String ' cantidad


If Not Archivo_Existe(Path, Archivo) Then
    MsgBox "No existe archivo " & vbLf & Path & Archivo
    Exit Sub
End If

'On Error GoTo NoExcel
Set Planilla = GetObject(Path & Archivo, "Excel.Sheet.8") 'assign sheet object as an OLE excel
'On Error GoTo 0

With Planilla.Worksheets(Hoja)

fi = 1
filas_vacias = 0

Do While True

    fi = fi + 1
    
    m_Cli = Trim(.cells(fi, 1).Value)
    ' si linea esta en blanco
    If Len(m_Cli) = 0 Then
        Exit Do
    End If

    m_Pro = Trim(.cells(fi, 2).Value)
    m_ProNum = Trim(.cells(fi, 3).Value)
    m_OdcNum = Trim(.cells(fi, 4).Value)
    m_MKN = Trim(.cells(fi, 5).Value)
    m_MKD = Trim(.cells(fi, 6).Value)
    m_NVM = Trim(.cells(fi, 7).Value)
    m_KgU = Trim(.cells(fi, 8).Value)
    m_Can = Trim(.cells(fi, 9).Value)
    
    Debug.Print m_Cli, m_Pro, m_ProNum, m_OdcNum, m_MKN, m_MKD, m_NVM, m_KgU, m_Can
    
    av(1) = m_Cli
    av(2) = m_Pro
    av(3) = m_ProNum
    av(4) = m_OdcNum
    av(5) = m_MKN
    av(6) = m_MKD
    av(7) = m_NVM
    av(8) = m_KgU
    av(9) = m_Can
    
    Registro_Agregar CnxSqlServer_scp0, "etiquetas_barrick", ac, av, 9
        
Loop

End With

Set Planilla = Nothing

Exit Sub

NoExcel:
'MsgBox "No Tiene Instalado Microsoft Excel"
' o archivo esta abierto (en uso por ejemplo por excel)
MsgBox "Nombre de Hoja NO Válido"

End Sub
