VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PagoContratistas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago a Contratistas"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox FechaPago 
      Height          =   300
      Left            =   3600
      TabIndex        =   9
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   327680
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton btnMarcar 
      Caption         =   "Desmarcar Todas"
      Height          =   495
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton btnMarcar 
      Caption         =   "Marcar Todas"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton btnImprimir 
      Height          =   400
      Left            =   5520
      Picture         =   "PagoContratistas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   400
   End
   Begin VB.CommandButton btnGrabar 
      Caption         =   "&Grabar"
      Height          =   315
      Left            =   4800
      TabIndex        =   2
      Top             =   5280
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid fg 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7011
      _Version        =   327680
   End
   Begin VB.Label Contratista 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label lbl 
      Caption         =   "Fecha de Pago"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   8
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label PrecioTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   300
      Left            =   4560
      TabIndex        =   7
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label KgTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   300
      Left            =   3600
      TabIndex        =   6
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Doble Click en ""Pagada"" para cambiar estado"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
End
Attribute VB_Name = "PagoContratistas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private n_columnas As Integer, n_filas As Integer, i As Integer
Private DbD As Database, RsSc As Recordset
Private Dbm As Database, RsNVc As Recordset, RsITOfc As Recordset

Private m_Nv As Double, m_NvArea As Integer, m_Rut As String, m_Razon As String, m_FechaInicio As String, m_FechaTermino As String
Private m_PesoTotal As Double, m_PrecioTotal As Double

Public Property Let Numero_NotaVenta(ByVal New_Opcion As Double)
m_Nv = New_Opcion
End Property
Public Property Let Contratista_Rut(ByVal New_Opcion As String)
m_Rut = New_Opcion
End Property
Public Property Let Contratista_Nombre(ByVal New_Opcion As String)
m_Razon = New_Opcion
End Property
Public Property Let Fecha_Inicio(ByVal New_Opcion As String)
m_FechaInicio = New_Opcion
End Property
Public Property Let Fecha_Termino(ByVal New_Opcion As String)
m_FechaTermino = New_Opcion
End Property

Private Sub btnMarcar_Click(Index As Integer)
Dim m_Pagada As String

m_PesoTotal = 0
m_PrecioTotal = 0

If Index = 0 Then
    m_Pagada = "SI"
Else
    m_Pagada = "NO"
End If

For i = 1 To n_filas
    If fg.TextMatrix(i, 7) <> "" Then
        fg.TextMatrix(i, 7) = m_Pagada
        If m_Pagada = "SI" Then
            m_PesoTotal = m_PesoTotal + m_CDbl(fg.TextMatrix(i, 5))
            m_PrecioTotal = m_PrecioTotal + m_CDbl(fg.TextMatrix(i, 6))
        End If
    End If
Next

KgTotal.Caption = m_PesoTotal
PrecioTotal.Caption = m_PrecioTotal

End Sub
Private Sub Sumar()

m_PesoTotal = 0
m_PrecioTotal = 0

For i = 1 To n_filas
    If fg.TextMatrix(i, 7) = "SI" Then
        m_PesoTotal = m_PesoTotal + m_CDbl(fg.TextMatrix(i, 5))
        m_PrecioTotal = m_PrecioTotal + m_CDbl(fg.TextMatrix(i, 6))
    End If
Next

KgTotal.Caption = m_PesoTotal
PrecioTotal.Caption = m_PrecioTotal

End Sub
Private Sub Form_Load()

Contratista.Caption = m_Razon

n_columnas = 7
n_filas = 10

Set DbD = OpenDatabase(data_file)
Set RsSc = DbD.OpenRecordset("contratistas")
RsSc.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index

Fg_Inicializa

Leer

m_NvArea = 0

btnImprimir.visible = False

End Sub
Private Sub Fg_Inicializa()
Dim ancho As Integer

fg.Cols = n_columnas + 1
fg.Rows = n_filas + 1

fg.ColWidth(0) = 0
fg.ColWidth(1) = 600
fg.ColWidth(2) = 1300
fg.ColWidth(3) = 700
fg.ColWidth(4) = 900
fg.ColWidth(5) = 900
fg.ColWidth(6) = 900
fg.ColWidth(7) = 700

ancho = 350
For i = 0 To fg.Cols - 1
    ancho = ancho + fg.ColWidth(i)
Next

fg.Width = ancho
Me.Width = ancho + fg.Left * 2

i = 0
fg.TextMatrix(i, 0) = ""
fg.TextMatrix(i, 1) = "NV"
fg.TextMatrix(i, 2) = "Obra"
fg.TextMatrix(i, 3) = "ITOf"
fg.TextMatrix(i, 4) = "Fecha"
fg.TextMatrix(i, 5) = "Kg.Total"
fg.TextMatrix(i, 6) = "$ Total"
fg.TextMatrix(i, 7) = "Pagada"

End Sub
Private Sub Leer()
' lee archivo y le trae al FlexGrid
Dim sql As String
Dim descripcion As String, m_doble As Double

sql = "SELECT *"
sql = sql & " FROM [ito fab Cabecera]"
sql = sql & " WHERE not pagada"

If m_Nv > 0 Then sql = sql & " AND NV=" & m_Nv

If m_Rut <> "" Then sql = sql & " AND [rut contratista]='" & m_Rut & "'"

If m_FechaInicio <> "__/__/__" And m_FechaInicio <> "" Then
    sql = sql & " AND [fecha]>=CDate('" & m_FechaInicio & "')"
End If
If m_FechaTermino <> "__/__/__" And m_FechaTermino <> "" Then
    sql = sql & " AND [fecha]<=CDate('" & m_FechaTermino & "')"
End If

sql = sql & " ORDER BY nv,numero"

Set RsITOfc = Dbm.OpenRecordset(sql)

n_filas = 10
fg.Rows = n_filas + 1

i = 0
With RsITOfc
If .RecordCount > 0 Then

.MoveFirst
Do While Not .EOF

    i = i + 1
    If i > n_filas Then
        n_filas = n_filas + 1
        fg.Rows = n_filas + 1
    End If
    
    descripcion = "-"
    RsNVc.Seek "=", !Nv, m_NvArea
    If Not RsNVc.NoMatch Then
        descripcion = RsNVc![obra]
    End If
    
    fg.TextMatrix(i, 1) = !Nv
    fg.TextMatrix(i, 2) = descripcion
    
    fg.TextMatrix(i, 3) = !Numero
    fg.TextMatrix(i, 4) = !Fecha
    
    fg.TextMatrix(i, 5) = ![Peso Total]
    fg.TextMatrix(i, 6) = ![Precio Total]
        
    fg.TextMatrix(i, 7) = IIf(!pagada, "SI", "NO")
    
    .MoveNext
    
Loop
End If
End With
End Sub
Private Sub fg_DblClick()

If fg.col = 7 Then
    Estado_Cambiar
End If

End Sub
Private Sub fg_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    Estado_Cambiar
End If
End Sub
Private Sub Estado_Cambiar()

If Trim(fg.TextMatrix(fg.Row, 7)) = "SI" Then
    fg.TextMatrix(fg.Row, 7) = "NO"
Else
    fg.TextMatrix(fg.Row, 7) = "SI"
End If
    
Sumar
    
End Sub
Private Sub FechaPago_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub FechaPAgo_LostFocus()
Dim d As Variant
d = Fecha_Valida(FechaPago, Now)
End Sub

Private Sub btnGrabar_Click()
Dim m_Numero As Double

If FechaPago.Text = "__/__/__" Then
    MsgBox "Debe digitar Fecha de Pago"
    FechaPago.SetFocus
    Exit Sub
End If

RsITOfc.Close

Set RsITOfc = Dbm.OpenRecordset("ITO fab Cabecera")

With RsITOfc
.Index = "Numero"

For i = 1 To n_filas

    If fg.TextMatrix(i, 7) = "SI" Then
    
        m_Numero = Val(fg.TextMatrix(i, 3))
        
        If m_Numero > 0 Then
        
            .Seek "=", m_Numero
            If Not .NoMatch Then
                .Edit
                ![FechaPago] = FechaPago.Text
                ![pagada] = IIf(fg.TextMatrix(i, 7) = "SI", True, False)
                .Update
            End If
            
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
