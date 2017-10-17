VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form oce_generar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Genera OC Especial desde planilla Excel"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ComboHoja 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton btnHelp 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   1920
      Width           =   375
   End
   Begin Crystal.CrystalReport cr 
      Left            =   4080
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton btnGenerar 
      Caption         =   "&Generar"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton btnAbrir 
      Caption         =   "&Abrir Archivo"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   345
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      _Version        =   327680
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblHoja 
      Caption         =   "Hoja Excel"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Emisión"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   300
      Width           =   1095
   End
   Begin VB.Label lblEstado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label lblCarpeta 
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   4455
   End
End
Attribute VB_Name = "oce_generar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'####################################################################
'####################################################################
Option Explicit
Private Dbm As Database, RsNVc As Recordset ', RsPc As Recordset, RsPd As Recordset

' para oc esp
Private Dba As Database, RsOcc As Recordset, RsOCd As Recordset, RsCorre As Recordset

' para reporte
Private DbR As Database, RsR As Recordset

Private m_Path As String, m_Archivo As String
Private i As Integer, j As Integer, existe As Boolean
Private Planilla As Object, Hoja As String
Private m_Nv As Integer, m_Numero As Double, st As Double, oc_linea As Integer
Private m_Rut As String, m_Fecha As Date, m_Condiciones As String
Private m_FechaaRecibir As Date, m_Atencion As String, m_EntregarEn As String
Private m_Cotizacion As Double, m_Obs1 As String, m_Obs2 As String
Private m_Obs3 As String, m_Obs4 As String
Private m_TotaldeOCs As Integer
Private m_Titulo As String, m_NvArea As Integer

Private Sub btnHelp_Click()
' muestra pantalla de ayuda:
'   planilla de abrirse en forma exclusiva
'   hoja debe llamarse "hoja1"
'   primer fila con encabezados, datos a partir de segunda fila
'   cada oc debe estar separada por una linea en blanco
'   maximo 19 lineas de detalle por oc especial
'   columna B: NV
'   columnas D, E, F, G, H, I, J y O, son el detalle
'   columna L: cantidad, con un decimal
'   columna M: precio unitario
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    MsgBox "Ayuda"
End If
End Sub

' 0: numero,  1: nombre obra
Private Sub Form_Load()

Set Dbm = OpenDatabase(mpro_file)

Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
'RsNVc.Index = Nv_Index ' "Número"
RsNVc.Index = "Numero" ' debe ser x numero para buscar

Set Dba = OpenDatabase(Madq_file)
Set RsOcc = Dba.OpenRecordset("OC Cabecera")
RsOcc.Index = "Numero"

Set RsOCd = Dba.OpenRecordset("OC Detalle")
RsOCd.Index = "Numero-Linea"
Set RsCorre = Dba.OpenRecordset("Correlativo")

Set DbR = OpenDatabase(repo_file)
Set RsR = DbR.OpenRecordset("oce_genera")

m_Rut = "96637830-1" ' masteri s.a.
m_Fecha = Format(CStr(Date), "dd/mm/yy")
m_FechaaRecibir = m_Fecha

m_Condiciones = "30 DIAS"
m_EntregarEn = "LAS ACACIAS"

m_TotaldeOCs = 0

m_NvArea = 0

ComboHoja.Clear
ComboHoja.AddItem "--- Elegir Hoja ---"

End Sub
Private Sub btnAbrir_Click()
' abre carpeta

Dim m_PathArchivo As String, p As Integer, m_nHojas As Integer, m_Hoja As String

cd.DialogTitle = "Buscar Carpeta"
cd.Filter = "Microsoft Excel (*.xls)|*.xls|Todos los Archivos (*.*)|*.*"

' busca ultima ruta
m_Path = GetSetting("scp", "oce", "ruta")
If m_Path = "" Then m_Path = "C:"

' si m_path no existe, muestra directorio actual
'cd.InitDir = m_Path

'm_Path = "C:" ' por mientras

m_Path = Directorio(m_Path)

cd.InitDir = m_Path
cd.ShowOpen

m_PathArchivo = cd.filename

If m_PathArchivo = "" Then
'    MsgBox "Debe escoger Archivo"
    Exit Sub
End If

'MsgBox Cd.filename ' viene con ruta completa
m_PathArchivo = cd.filename

' separa path y archivo
p = InStrLast(m_PathArchivo, "\")
If p > 0 Then

    m_Path = Left(m_PathArchivo, p)
    m_Archivo = Mid(m_PathArchivo, p + 1)
    
    lblCarpeta.Caption = m_Path
    
Set Planilla = GetObject(m_Path & m_Archivo, "Excel.Sheet.8") 'assign sheet object as an OLE excel

' busca hojas
'Debug.Print Planilla.Worksheets.Count
m_nHojas = Planilla.Worksheets.Count
For i = 1 To m_nHojas
   m_Hoja = Planilla.Worksheets(i).Name
'      Debug.Print m_Hoja
   ComboHoja.AddItem m_Hoja
Next



If m_nHojas = 1 Then
   ComboHoja.ListIndex = 1
Else
   ComboHoja.ListIndex = 0
End If
   
Hoja = ComboHoja.Text
'   On Error GoTo 0
    
    ' guarda ultima ruta usada
    SaveSetting "scp", "planos", "oce", m_Path
    
End If

End Sub
Private Sub Fecha_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Fecha_LostFocus()
Dim d As Variant
d = Fecha_Valida(Fecha, Now)
End Sub
Private Sub btnGenerar_Click()

Dim NombreArchTXT As String, p As Integer

If lblCarpeta.Caption = "" Then
    MsgBox "Debe Escoger ARCHIVO de ORIGEN"
    btnAbrir.SetFocus
    Exit Sub
End If

If Fecha.Text = "__/__/__" Then
    MsgBox "Debe Digitar Fecha Emisión"
    Fecha.SetFocus
    Exit Sub
End If

If ComboHoja.ListIndex = 0 Then
    MsgBox "Debe Escoger Hoja"
    ComboHoja.SetFocus
    Exit Sub
End If

' importa excel

DbR.Execute "DELETE * FROM oce_genera"

m_Fecha = Format(Fecha.Text, "dd/mm/yy")
m_FechaaRecibir = m_Fecha

'Hoja = "hoja1"
Hoja = ComboHoja.Text
Excel_Leer m_Path, m_Archivo, Hoja

m_Titulo = "Ordenes de Compra Generadas"

cr.WindowTitle = m_Titulo

cr.WindowMaxButton = False
cr.WindowMinButton = False
cr.WindowState = crptMaximized

cr.DataFiles(0) = repo_file & ".MDB"
cr.Formulas(0) = "RAZON SOCIAL=""" & Empresa.Razon & """"
cr.Formulas(1) = "TITULO=""" & m_Titulo & """"
cr.Formulas(2) = "archivo=""" & "Archivo: " & m_Path & m_Archivo & """"
cr.ReportSource = crptReport
cr.ReportFileName = Drive_Server & Path_Rpt & "oce_genera.rpt"
cr.Action = 1

'MsgBox "ok"

End Sub

Private Sub Excel_Leer(Path As String, Archivo As String, Hoja As String)
' importa datos desde planilla excel
Dim fi As Integer, co As Integer
Dim filas_vacias As Integer, m_NvAnterior As Double, PrimeraLineaOc As Boolean

If Not Archivo_Existe(Path, Archivo) Then
    MsgBox "No existe archivo " & vbLf & Path & Archivo
    Exit Sub
End If

'On Error GoTo NoExcel
'Set Planilla = GetObject(Path & Archivo, "Excel.Sheet.8") 'assign sheet object as an OLE excel
'On Error GoTo 0

With Planilla.Worksheets(Hoja)

fi = 1
filas_vacias = 0
st = 0
oc_linea = 0
'oce_nueva = True
m_NvAnterior = Val(Trim(.cells(2, 2).Value))  ' nv
PrimeraLineaOc = True

Do While True

    fi = fi + 1
    
    m_Nv = Val(Trim(.cells(fi, 2).Value)) ' nv
    
    ' si linea esta en blanco
    If m_Nv = 0 Then
    
        ' fila está vacia
        filas_vacias = filas_vacias + 1
        If filas_vacias > 1 Then
            ' hay mas de una fila vacia => fin archivo excel
            Exit Do
        End If
        
        m_Nv = m_NvAnterior
        
        Cabecera_Grabar
        
        DesBloqueo "OC", RsCorre
        
        ' graba en reporte
        RsR.AddNew
        RsR!Numero = m_Numero
        RsR!Fecha = m_Fecha
        RsR!Nv = m_Nv
        RsNVc.Seek "=", m_Nv, m_NvArea
        If Not RsNVc.NoMatch Then
            RsR!obra = RsNVc!obra
        End If
        RsR!SubTotal = st
        RsR.Update
        ' ////////////////
        
        st = 0
        oc_linea = 0
        PrimeraLineaOc = True
        
    Else
    
        If PrimeraLineaOc Then
            m_Numero = GetNumDoc("OC", RsOcc, RsCorre)
            PrimeraLineaOc = False
        End If
        
        linea_grabar fi
        filas_vacias = 0
        m_NvAnterior = m_Nv
        
    End If
    
Loop

End With

Set Planilla = Nothing

Exit Sub

NoExcel:
'MsgBox "No Tiene Instalado Microsoft Excel"
' o archivo esta abierto (en uso por ejemplo por excel)
MsgBox "Nombre de Hoja NO Válido"
   
    
End Sub
Private Sub Cabecera_Grabar()
'Private Function Doc_Grabar(Nueva As Boolean) As Double

Dim m_cantidad As String, m_pr As Double
'Doc_Grabar = 0

save:
' CABECERA
With RsOcc

If m_Numero = 0 Then
    MsgBox "OC NO SE GRABÓ!"
'    Doc_Grabar = 0
    MousePointer = vbDefault
    Exit Sub
End If

'Numero.Text = m_Numero
'Doc_Grabar = m_Numero

.AddNew
!Numero = m_Numero
!Tipo = "E"
!Fecha = m_Fecha
!Nv = m_Nv
![RUT Proveedor] = m_Rut
'![Codigo Direccion] = Direccion.ListIndex
![Condiciones de Pago] = m_Condiciones
![Fecha a Recibir] = m_FechaaRecibir
!atencion = m_Atencion
![Entregar en] = m_EntregarEn
!Cotizacion = m_Cotizacion
![Observacion 1] = m_Obs1
![Observacion 2] = m_Obs2
![Observacion 3] = m_Obs3
![Observacion 4] = m_Obs4
!SubTotal = st
![% Descuento] = 0
!Descuento = 0
!Neto = st
!Iva = Int(st * Parametro.Iva / 100 + 0.5)
!Total = st + !Iva
!Pendiente = True
!Nula = False

!certificado = False ' ?

.Update
End With

m_TotaldeOCs = m_TotaldeOCs + 1
lblEstado.Caption = "Total de OC Generada(s) = " & m_TotaldeOCs
Me.Refresh

End Sub
Private Sub linea_grabar_old(fi As Integer)

Dim m_Can As Double, m_Des As String, m_PrU As Double
Dim s_Paso As String, d_Paso As Double

oc_linea = oc_linea + 1

With Planilla.Worksheets(Hoja)

m_Can = m_CDbl(Trim(.cells(fi, 10).Value))
m_Des = Trim(.cells(fi, 4).Value) & " " & Trim(.cells(fi, 5).Value)
m_Des = m_Des & " " & Trim(.cells(fi, 6).Value) & " " & Trim(.cells(fi, 7).Value)
m_Des = m_Des & " " & Trim(.cells(fi, 8).Value) & " " & Trim(.cells(fi, 13).Value)

' precio unitario
d_Paso = m_CDbl(Trim(.cells(fi, 11).Value))

If d_Paso = 0 Then
    ' error kg no puede venir con cero o alfanumerico
    MsgBox "Formato de Planilla No Válido" & vbLf & _
    "Columna K de planilla debe venir con Precio Unitario"
    Exit Sub
End If

m_PrU = d_Paso

'Debug.Print m_Can, m_Des, m_PrU
    
st = st + Int(m_Can * m_PrU + 0.5)
    
' graba linea
With RsOCd
.AddNew
!Numero = m_Numero
!linea = oc_linea
!Tipo = "E"
!Fecha = m_Fecha
!Nv = m_Nv
![RUT Proveedor] = m_Rut ' masteri
!unidad = "KGS" ' pedido por felix 24/08/05
!Descripcion = m_Des
!Cantidad = m_Can
![Precio Unitario] = m_PrU
!Total = m_Can * m_PrU
.Update
End With

End With

End Sub
Private Sub linea_grabar(fi As Integer)
' version 30/09/05
Dim m_Can As Double, m_Des As String, m_PrU As Double
Dim s_Paso As String, d_Paso As Double

oc_linea = oc_linea + 1

With Planilla.Worksheets(Hoja)

m_Can = m_CDbl(Trim(.cells(fi, 12).Value)) ' col L
'If m_Can > 0 Then m_Can = Format(m_Can, "#.#")
If m_Can > 0 Then m_Can = Format(m_Can, "#.##")

m_Des = Trim(.cells(fi, 4).Value) & " " & Trim(.cells(fi, 5).Value)                 ' col D  col E
m_Des = m_Des & " " & Trim(.cells(fi, 6).Value) & " " & Trim(.cells(fi, 7).Value)   ' col F  col G
m_Des = m_Des & " " & Trim(.cells(fi, 8).Value) & " " & Trim(.cells(fi, 9).Value)   ' col H  col I
m_Des = m_Des & " " & Trim(.cells(fi, 10).Value) & " " & Trim(.cells(fi, 15).Value) ' col J  col O

'd_Paso = m_CDbl(Trim(.cells(fi, 12).Value))  ' col K
'If d_Paso > 0 Then d_Paso = Format(d_Paso, "#.#")
'm_Des = m_Des & " " & d_Paso & " " & Trim(.cells(fi, 15).Value) ' col K  col O

' precio unitario
d_Paso = m_CDbl(Trim(.cells(fi, 13).Value)) ' col M

If d_Paso = 0 Then
    ' puede haber precio unitario vacio
    ' se trata de una linea de descripcion
End If

m_PrU = d_Paso

'Debug.Print m_Can, m_Des, m_PrU

st = st + Int(m_Can * m_PrU + 0.5)

' graba linea
With RsOCd
.AddNew
!Numero = m_Numero
!linea = oc_linea
!Tipo = "E"
!Fecha = m_Fecha
!Nv = m_Nv
![RUT Proveedor] = m_Rut ' masteri
!unidad = "KGS" ' pedido por felix 24/08/05

![Fecha a Recibir] = m_FechaaRecibir ' agregado el 13/03/06
!Pendiente = True                    ' agregado el 13/03/06

!Descripcion = Left(m_Des, 50)

!Cantidad = m_Can
![Precio Unitario] = m_PrU
!Total = m_Can * m_PrU

.Update
End With
    
End With
    
End Sub
