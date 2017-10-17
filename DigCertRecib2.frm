VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form DigCertRecib2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificados Recibidos"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnImprimir 
      Height          =   400
      Left            =   3360
      Picture         =   "DigCertRecib2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   400
   End
   Begin VB.CheckBox Check_Todos 
      Caption         =   "Todas las OC"
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton btnGrabar 
      Caption         =   "&Grabar"
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   3960
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid fg 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5953
      _Version        =   327680
   End
   Begin VB.Label lbl 
      Caption         =   "Doble Click en ""Recib"" para cambiar estado"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "DigCertRecib2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private n_columnas As Integer, n_filas As Integer, i As Integer
Private Dba As Database, RsOcc As Recordset
Private Dbm As Database, RsNVc As Recordset, RsRmC As Recordset
Private DbD As Database, RsProv As Recordset

Private m_Nv As Double, m_NvArea As String, m_Rut As String, m_FechaInicio As String, m_FechaTermino As String

Public Property Let Numero_NotaVenta(ByVal New_Opcion As Double)
m_Nv = New_Opcion
End Property
Public Property Let Rut_Proveedor(ByVal New_Opcion As String)
m_Rut = New_Opcion
End Property
Public Property Let Fecha_Inicio(ByVal New_Opcion As String)
m_FechaInicio = New_Opcion
End Property
Public Property Let Fecha_Termino(ByVal New_Opcion As String)
m_FechaTermino = New_Opcion
End Property

Private Sub Form_Load()

'Debug.Print "|" & m_Nv & "|"
'Debug.Print "|" & m_RUT & "|"
'Debug.Print "|" & m_FechaInicio & "|"
'Debug.Print "|" & m_FechaTermino & "|"

n_columnas = 8
n_filas = 10

Set DbD = OpenDatabase(data_file)
Set RsProv = DbD.OpenRecordset("Proveedores")
RsProv.Index = "RUT"

Set Dba = OpenDatabase(Madq_file)

Set RsRmC = Dba.OpenRecordset("RM Cabecera")
RsRmC.Index = "Oc-Rm"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Fg_Inicializa

Leer

m_NvArea = 0

End Sub
Private Sub Fg_Inicializa()
Dim ancho As Integer

fg.Cols = n_columnas + 1
fg.Rows = n_filas + 1

fg.ColWidth(0) = 0
fg.ColWidth(1) = 800
fg.ColWidth(2) = 900
fg.ColWidth(3) = 1600
fg.ColWidth(4) = 450
fg.ColWidth(5) = 900
fg.ColWidth(6) = 500
fg.ColWidth(7) = 1600
fg.ColWidth(8) = 600

ancho = 350
For i = 0 To fg.Cols - 1
    ancho = ancho + fg.ColWidth(i)
Next

fg.Width = ancho
Me.Width = ancho + fg.Left * 2

i = 0
fg.TextMatrix(i, 0) = ""
fg.TextMatrix(i, 1) = "OC"
fg.TextMatrix(i, 2) = "Fecha"
fg.TextMatrix(i, 3) = "Proveedor"
fg.TextMatrix(i, 4) = "Doc"
fg.TextMatrix(i, 5) = "Nº"
fg.TextMatrix(i, 6) = "NV"
fg.TextMatrix(i, 7) = "Obra"
fg.TextMatrix(i, 8) = "Recib"

End Sub
Private Sub Leer()
' lee archivo y le trae al FlexGrid
Dim sql As String
Dim Descripcion As String, m_doble As Double

sql = "SELECT "
sql = sql & "[OC Cabecera].NV as NV,"
sql = sql & "[OC Cabecera].Numero as numero,"
sql = sql & "[OC Cabecera].fecha as fecha,"
sql = sql & "[OC Cabecera].[RUT Proveedor] as rut,"
sql = sql & "[OC Cabecera].[Certificado Recibido] as recibido"
sql = sql & " FROM [Oc Cabecera]"
sql = sql & " WHERE Certificado AND not nula"

If m_Nv > 0 Then sql = sql & " AND NV=" & m_Nv

If m_Rut <> "" Then sql = sql & " AND [rut proveedor]='" & m_Rut & "'"

If m_FechaInicio <> "__/__/__" Then sql = sql & " AND [fecha]>=CDate('" & m_FechaInicio & "')"

If m_FechaTermino <> "__/__/__" Then sql = sql & " AND [fecha]<=CDate('" & m_FechaTermino & "')"

If Check_Todos.Value = 1 Then
    'todos
Else
    ' solo pendientes
    sql = sql & " AND NOT [Certificado Recibido]"
End If

Set RsOcc = Dba.OpenRecordset(sql)

n_filas = 10
fg.Rows = n_filas + 1

i = 0
With RsOcc
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    i = i + 1
    If i > n_filas Then
        n_filas = n_filas + 1
        fg.Rows = n_filas + 1
    End If
    
    fg.TextMatrix(i, 1) = !Numero
    fg.TextMatrix(i, 2) = !Fecha
    
    Descripcion = ""
    RsProv.Seek "=", !Rut
    If Not RsProv.NoMatch Then
        Descripcion = RsProv![Razon Social]
    End If
    fg.TextMatrix(i, 3) = Descripcion
    
'    If !Numero = 47142 Then
'    MsgBox "ok"
'    End If
    
    Descripcion = ""
    m_doble = 0
    RsRmC.Seek ">=", !Numero, 0
    If Not RsRmC.NoMatch Then
        If RsRmC!Oc = !Numero Then
            Descripcion = RsRmC![Tipo Documento]
            m_doble = RsRmC![Numero Documento]
        End If
    End If
    fg.TextMatrix(i, 4) = Descripcion
    fg.TextMatrix(i, 5) = m_doble
    
    Descripcion = ""
    RsNVc.Seek "=", !Nv, m_NvArea
    If Not RsNVc.NoMatch Then
        Descripcion = RsNVc![obra]
    End If
    fg.TextMatrix(i, 6) = !Nv
    fg.TextMatrix(i, 7) = Descripcion
    
    fg.TextMatrix(i, 8) = IIf(!recibido, "SI", "NO")
    
    .MoveNext
    
Loop
End If
End With
End Sub
Private Sub Check_Todos_Click()
' pregunta si realmente quiere todos
'If MsgBox("¿ Todos ?", vbYesNo) = vbYes Then
'    Check_Todos.Value = 1
'Else
'    Check_Todos.Value = 0
'End If
Leer
'Todos_Leer
End Sub
Private Sub fg_DblClick()

If fg.col = 8 Then
    Estado_Cambiar
End If

End Sub
Private Sub fg_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    Estado_Cambiar
End If
End Sub
Private Sub Estado_Cambiar()
    If Trim(fg.TextMatrix(fg.Row, 8)) = "SI" Then
        fg.TextMatrix(fg.Row, 8) = "NO"
    Else
        fg.TextMatrix(fg.Row, 8) = "SI"
    End If
End Sub
Private Sub btnGrabar_Click()
Dim m_Numero As Double

RsOcc.Close

Set RsOcc = Dba.OpenRecordset("OC Cabecera")
With RsOcc
.Index = "Numero"

For i = 1 To n_filas

    m_Numero = Val(fg.TextMatrix(i, 1))
    If m_Numero > 0 Then
    .Seek "=", m_Numero
    If Not .NoMatch Then
        .Edit
        ![Certificado Recibido] = IIf(fg.TextMatrix(i, 8) = "SI", True, False)
        .Update
    End If
    End If
    
Next
.Close
End With
Unload Me
End Sub
Private Sub btnImprimir_Click()
Imprimir
End Sub
Private Sub Imprimir()
Dim j As Integer
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer, tab4 As Integer
Dim tab5 As Integer, tab6 As Integer, tab7 As Integer, prt As Printer, li As Integer, linea As String

linea = String(76, "-")
tab0 = 4 'margen izquierdo
tab1 = tab0 + 0  'oc
tab2 = tab1 + 10 'fecha
tab3 = tab2 + 10 'proveedor
tab4 = tab3 + 20 'doc
tab5 = tab4 + 5  'numero
tab6 = tab5 + 11 'nv
tab7 = tab6 + 5  'obra

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt
'prt.Font.Name = "Courier New"

For j = 1 To 1 'n_Copias

    prt.Font.Size = 12
    prt.Print Tab(tab0); Empresa.Razon
    prt.Print Tab(tab0); " "
    prt.Print Tab(tab0); "Certificados de Calidad Pendientes"
    prt.Print Tab(tab0); "FECHA "; Format(Now, Fecha_Format)
    prt.Print Tab(tab0); " "
'    prt.Print Tab(tab0); "Proveedor : "; UCase(fg.TextMatrix(1, 3))
'    prt.Print Tab(tab0); " "

    prt.Font.Bold = True
    prt.Print Tab(tab1); PadL("OC", 6);
    prt.Print Tab(tab2); PadL("FECHA", 8);
    prt.Print Tab(tab3); "Proveedor";
    prt.Print Tab(tab4); "Doc";
    prt.Print Tab(tab5); PadL("Nº", 10);
    prt.Print Tab(tab6); " NV";
    prt.Print Tab(tab7); "OBRA"
    prt.Font.Bold = False
    
    prt.Print Tab(tab0); linea
    
    For li = 1 To n_filas
        If Trim(fg.TextMatrix(li, 1)) <> "" Then
            prt.Print Tab(tab1); PadL(fg.TextMatrix(li, 1), 6);
            prt.Print Tab(tab2); Format(fg.TextMatrix(li, 2), Fecha_Format);
            prt.Print Tab(tab3); PadR(fg.TextMatrix(li, 3), 18);
            prt.Print Tab(tab4); PadL(fg.TextMatrix(li, 4), 3);
            prt.Print Tab(tab5); PadL(fg.TextMatrix(li, 5), 10);
            prt.Print Tab(tab6); PadL(fg.TextMatrix(li, 6), 4);
            prt.Print Tab(tab7); PadR(fg.TextMatrix(li, 7), 16)
        End If
    Next
    
    prt.Print Tab(tab0); linea
    prt.Print Tab(tab0); "G: Guías de Despacho"
    prt.Print Tab(tab0); "F: Factura de Venta"
    
    prt.EndDoc
    
Next

Impresora_Predeterminada "default"

End Sub
