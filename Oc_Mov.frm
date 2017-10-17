VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Oc_Mov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial de una OC"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.TextBox N_Oc 
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton btnImprimir 
         Height          =   400
         Left            =   6840
         Picture         =   "Oc_Mov.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton btnLimpiar 
         Caption         =   "&Limpiar"
         Height          =   255
         Left            =   5400
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton btnBuscar 
         Caption         =   "&Buscar"
         Height          =   255
         Left            =   3840
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label impresa 
         Caption         =   "NO"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6840
         TabIndex        =   12
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "IMPRESA"
         Height          =   255
         Left            =   5880
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.Label obra 
         Caption         =   "Obra"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   900
         Width           =   3615
      End
      Begin VB.Label lbl 
         Caption         =   "OBRA"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Proveedor 
         Caption         =   "Proveedor"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label lbl 
         Caption         =   "PROVEEDOR"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Especial 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lbl 
         Caption         =   "Nº Orden Compra"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Detalle 
      Height          =   3495
      Left            =   240
      TabIndex        =   13
      Top             =   1560
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6165
      _Version        =   327680
      FixedCols       =   0
   End
End
Attribute VB_Name = "Oc_Mov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DbD As Database, RsProd As Recordset, RsProv As Recordset
Private Dbm As Database, RsNVc As Recordset
Private DbAdq As Database
Private RsOcc As Recordset, RsOCd As Recordset
Private RsRmC As Recordset, RsRMd As Recordset
Private indice As Integer, n_filas As Integer, n_columnas As Integer
Private m_Fecha As Date, m_Texto As String
Private m_NvArea As Integer
Private Sub Form_Load()

Set DbD = OpenDatabase(data_file)
Set RsProd = DbD.OpenRecordset("Productos")
RsProd.Index = "Codigo"
Set RsProv = DbD.OpenRecordset("Proveedores")
RsProv.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set DbAdq = OpenDatabase(Madq_file)

Set RsOcc = DbAdq.OpenRecordset("OC Cabecera")
RsOcc.Index = "Numero"

Set RsOCd = DbAdq.OpenRecordset("OC Detalle")
RsOCd.Index = "Numero-Linea"

Set RsRmC = DbAdq.OpenRecordset("RM Cabecera")
RsRmC.Index = "Numero"

Set RsRMd = DbAdq.OpenRecordset("Documentos")
RsRMd.Index = "OC-Linea-Rm"

n_filas = 20
n_columnas = 12 '11 '10
Resultado_Config

'btnBuscar.Enabled = False

btnLimpiar_Click

'btnImprimir.Visible = False ' ojo

m_NvArea = 0

End Sub
Private Sub Resultado_Config()

Dim i As Integer, ancho As Integer

Detalle.WordWrap = True
'Detalle.RowHeight(0) = 450
Detalle.Rows = n_filas + 1
Detalle.Cols = n_columnas

Detalle.ColAlignment(3) = 0 ' izquierda
Detalle.ColAlignment(4) = 1 ' izquierda
'Detalle.ColAlignment(5) = 0 ' izquierda
Detalle.ColAlignment(6) = 1 ' izquierda
Detalle.ColAlignment(7) = 1 ' izquierda

Detalle.TextMatrix(0, 0) = "Doc"
Detalle.TextMatrix(0, 1) = "Número"
Detalle.TextMatrix(0, 2) = "Doc"
Detalle.TextMatrix(0, 3) = "Número"
Detalle.TextMatrix(0, 4) = "Fecha"
Detalle.TextMatrix(0, 5) = "L" ' nueva linea
Detalle.TextMatrix(0, 6) = "Codigo"
Detalle.TextMatrix(0, 7) = "Descripción"
Detalle.TextMatrix(0, 8) = "Largo"
Detalle.TextMatrix(0, 9) = "Cant"
Detalle.TextMatrix(0, 10) = "Recib"
Detalle.TextMatrix(0, 11) = "Cert.Rec"

Detalle.ColWidth(0) = 400
Detalle.ColWidth(1) = 700
Detalle.ColWidth(2) = 400
Detalle.ColWidth(3) = 700
Detalle.ColWidth(4) = 900
Detalle.ColWidth(5) = 300
Detalle.ColWidth(6) = 1500
Detalle.ColWidth(7) = 3000
Detalle.ColWidth(8) = 1000
Detalle.ColWidth(9) = 1000
Detalle.ColWidth(10) = 1000
Detalle.ColWidth(11) = 800

ancho = 360
For i = 0 To n_columnas - 1
    ancho = ancho + Detalle.ColWidth(i)
    If i > 0 Then
        Detalle.Row = i
        Detalle.col = 11
        Detalle.CellAlignment = flexAlignLeftCenter
        Detalle.CellForeColor = vbRed
    End If
Next

Me.Width = ancho + 400
Detalle.Width = ancho

End Sub
Private Sub btnBuscar_Click()
' busca oc
Dim primeralinea As Boolean, m_CantidadRecibida As Double

primeralinea = True

MousePointer = vbHourglass

Especial.Caption = ""
Proveedor.Caption = ""
obra.Caption = ""
impresa.Caption = ""

Detalle_Limpiar

Detalle.TopRow = 1

btnLimpiar.Enabled = True

'Marca.Caption = UCase(Marca.Caption)

indice = 0
RsOCd.Seek ">=", N_Oc.Text

If RsOCd.NoMatch Then

    MsgBox "OC NO existe"
    N_Oc.SetFocus
    
Else
    
    Do While Not RsOCd.EOF
    
        If RsOCd![Numero] <> N_Oc.Text Then Exit Do
        
        If primeralinea Then
        
            RsOcc.Seek "=", N_Oc.Text
            If Not RsOcc.NoMatch Then
            
                If RsOcc!Certificado Then
                    'If RsOcc![Certificado Recibido] Then
                    '    certificado.Caption = "RECIBIDO"
                    'Else
                    '    certificado.Caption = "NO RECIBIDO"
                    'End If
                Else
                    'certificado.Caption = "NO VA CON CERTIFICADO"
                End If
                
                If RsOcc!impresa Then
                    impresa.Caption = "SI"
                Else
                    impresa.Caption = "NO"
                End If
                
            End If
        
            RsProv.Seek "=", RsOCd![RUT Proveedor]
            If Not RsProv.NoMatch Then
                Proveedor.Caption = RsProv![Razon Social]
            End If
            
            RsNVc.Seek "=", RsOCd!Nv, m_NvArea ' RsOCd!nvarea
            If Not RsNVc.NoMatch Then
                obra.Caption = RsOCd!Nv & " " & RsNVc!obra
            End If
            
            primeralinea = False
            
        End If
        
        Especial.Caption = IIf(RsOCd!tipo = "E", "E", "")
        
        indice = indice + 1
        If indice >= n_filas Then
            n_filas = indice
            Detalle.Rows = n_filas + 1
        End If
        Detalle.TextMatrix(indice, 0) = "OC"
        Detalle.TextMatrix(indice, 1) = N_Oc.Text
        Detalle.TextMatrix(indice, 4) = RsOCd!Fecha ' m_Fecha
        
        Detalle.TextMatrix(indice, 5) = RsOCd!linea
        
        m_Texto = NoNulo(RsOCd![codigo producto])
        
        If Len(m_Texto) = 0 Then
        
            ' si codigo producto esta vacio es porque es oc especial
            Detalle.TextMatrix(indice, 7) = NoNulo(RsOCd!Descripcion)
            
        Else
        
            ' oc normal
            Detalle.TextMatrix(indice, 6) = NoNulo(RsOCd![codigo producto])
            
            m_Texto = "Codigo NO Existe"
            RsProd.Seek "=", RsOCd![codigo producto]
            If Not RsProd.NoMatch Then
                m_Texto = RsProd!Descripcion
            End If
            Detalle.TextMatrix(indice, 7) = m_Texto
            Detalle.TextMatrix(indice, 8) = RsOCd!largo
            
        End If
        
        Detalle.TextMatrix(indice, 9) = RsOCd!Cantidad
        
        ' busca recepcion de material
        m_CantidadRecibida = 0
        RsRMd.Seek ">=", N_Oc.Text, RsOCd!linea
        If Not RsRMd.NoMatch Then
            Do While Not RsRMd.EOF
                If RsRMd!Oc <> N_Oc.Text Then Exit Do ', RsOCd!Línea
                
                If RsRMd![Linea Oc] = RsOCd!linea Then
                
                    ' puebla linea (puede haber mas de una linea de recepcion)
                    
                    indice = indice + 1
                    If indice >= n_filas Then
                        n_filas = indice
                        Detalle.Rows = n_filas + 1
                    End If
                    Detalle.TextMatrix(indice, 0) = "RM"
                    Detalle.TextMatrix(indice, 1) = RsRMd!Numero
                    
                    ' busca documento proveedor
                    RsRmC.Seek "=", RsRMd!Numero
                    If Not RsRmC.NoMatch Then
                        Detalle.TextMatrix(indice, 2) = RsRmC![Tipo Documento]
                        Detalle.TextMatrix(indice, 3) = RsRmC![Numero Documento]
                    End If
                    
                    Detalle.TextMatrix(indice, 4) = RsRMd!Fecha
                    Detalle.TextMatrix(indice, 5) = RsRMd!linea
                    Detalle.TextMatrix(indice, 6) = RsRMd![codigo producto]
                    
                    'm_Texto = "Codigo NO Existe"
                    'RsProd.Seek "=", RsOCd![Código Producto]
                    'If Not RsProd.NoMatch Then
                    '    m_Texto = RsProd!descripción
                    'End If
                    Detalle.TextMatrix(indice, 7) = m_Texto
                    
                    Detalle.TextMatrix(indice, 10) = RsRMd!Cant_Entra
                    
                    Detalle.TextMatrix(indice, 11) = IIf(RsRMd!certificadoRecibido, "S", "")
                    
                    m_CantidadRecibida = m_CantidadRecibida + RsRMd!Cant_Entra
                    
                End If
                
                RsRMd.MoveNext
            Loop
        End If
        
'        Detalle.Row = indice
'        Detalle.col = 5
'        Detalle.CellForeColor = vbRed
            
        ' aqui debo actualizar cantidad recibida en oc detalle
        RsOCd.Edit
'        ojo aqui voy 13/07/06
        RsOCd![Cantidad Recibida] = m_CantidadRecibida
        RsOCd.Update
        
        RsOCd.MoveNext
        
    Loop
    
End If

MousePointer = vbDefault

End Sub
Private Sub Fila_Crear(NuevaFila As Integer)
If NuevaFila >= n_filas Then 'antes era = y no >=
    n_filas = NuevaFila + 1
    Detalle.Rows = n_filas + 1
End If
End Sub
Private Sub btnLimpiar_Click()
'limpia variables
N_Oc.Text = ""
Especial.Caption = ""
Proveedor.Caption = ""
obra.Caption = ""
'certificado.Caption = ""
impresa.Caption = ""

Detalle_Limpiar

Detalle.TopRow = 1

btnBuscar.Enabled = True
btnLimpiar.Enabled = True

End Sub
Private Sub Detalle_Limpiar()
Dim f As Integer, c As Integer
For f = 1 To n_filas '- 1
    For c = 0 To n_columnas - 1
        Detalle.TextMatrix(f, c) = ""
    Next
Next
End Sub
Private Sub btnImprimir_Click()
' imprimir
'PrinterNCopias.Numero_Copias = 1
'PrinterNCopias.Show 1
'If PrinterNCopias.Numero_Copias > 0 Then
'    Imprimir PrinterNCopias.Numero_Copias
'End If

Imprimir 1

End Sub
Private Sub Form_Unload(Cancel As Integer)
DataBases_Cerrar
End Sub
Private Sub Imprimir(n_Copias)
Dim j As Integer
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer, tab4 As Integer, tab5 As Integer
Dim tab6 As Integer, tab7 As Integer, tab8 As Integer, tab9 As Integer, tab10 As Integer, tab11 As Integer
Dim prt As Printer, li As Integer, linea As String

linea = String(101, "-")

tab0 = 5 'margen izquierdo
tab1 = tab0 + 0  ' doc
tab2 = tab1 + 4  ' numero
tab3 = tab2 + 7  ' doc
tab4 = tab3 + 4  ' numero
tab5 = tab4 + 7  ' fecha
tab6 = tab5 + 9  ' L
tab7 = tab6 + 3  ' cod
tab8 = tab7 + 16  ' des
tab9 = tab8 + 20 '15  ' largo
tab10 = tab9 + 10 ' cant
tab11 = tab10 + 10  ' recib

'Printer_Set "Documentos"

prt_escoger.ImpresoraNombre = ""
prt_escoger.Show 1
'm_ImpresoraNombre = prt_escoger.ImpresoraNombre

Set prt = Printer '??
prt.Font.Name = "Courier New"
'Font_Setear prt

For j = 1 To n_Copias

    prt.Font.Size = 8
    prt.Print Tab(tab0); Empresa.Razon
    prt.Print Tab(tab0); " "
    prt.Print Tab(tab0); "HISTORICO DE OC Nº " & N_Oc.Text & " " & Especial.Caption
    prt.Print Tab(tab0); "FECHA REPORTE "; Format(Now, Fecha_Format); " "; time
    prt.Print Tab(tab0); " "
    prt.Print Tab(tab0); "PROVEEDOR  : " & Proveedor.Caption
    prt.Print Tab(tab0); "OBRA  : " & obra.Caption
    'prt.Print Tab(tab0); "CERTIFICADO  : " & certificado.Caption
    prt.Print Tab(tab0); "IMPRESA : " & impresa.Caption
    prt.Print Tab(tab0); " "

    prt.Font.Bold = True
    
    prt.Print Tab(tab1); "DOC";
    prt.Print Tab(tab2); PadL("Nº", 6);
    prt.Print Tab(tab3); "DOC";
    prt.Print Tab(tab4); PadL("Nº", 6);
    prt.Print Tab(tab5); " FECHA";
    prt.Print Tab(tab6); PadL("L", 2);
    prt.Print Tab(tab7); "COD";
    prt.Print Tab(tab8); "DESCRIPCION";
    prt.Print Tab(tab9); PadL("LARGO", 10);
    prt.Print Tab(tab10); PadL("CANT", 10);
    prt.Print Tab(tab11); PadL("RECIB", 10)
    
    prt.Font.Bold = False
    
    prt.Print Tab(tab0); linea
    
    For li = 1 To n_filas '- 1
    
        If Trim(Detalle.TextMatrix(li, 0)) <> "" Then
        
            prt.Print Tab(tab1); Detalle.TextMatrix(li, 0);            ' doc
            prt.Print Tab(tab2); PadL(Detalle.TextMatrix(li, 1), 6);   ' nº
            prt.Print Tab(tab3); Detalle.TextMatrix(li, 2);            ' doc
            prt.Print Tab(tab4); PadL(Detalle.TextMatrix(li, 3), 6);   ' nº
            prt.Print Tab(tab5); Format(Detalle.TextMatrix(li, 4), Fecha_Format);
            prt.Print Tab(tab6); PadL(Detalle.TextMatrix(li, 5), 2);   ' L
            prt.Print Tab(tab7); PadR(Detalle.TextMatrix(li, 6), 15);  ' codigo
            prt.Print Tab(tab8); PadR(Detalle.TextMatrix(li, 7), 19);  ' descirpcion
            prt.Print Tab(tab9); PadL(Detalle.TextMatrix(li, 8), 10);  ' largo
            prt.Print Tab(tab10); PadL(Detalle.TextMatrix(li, 9), 10); ' cant
            prt.Print Tab(tab11); PadL(Detalle.TextMatrix(li, 10), 10) ' recib
            
        End If
        
    Next
    
    prt.Print Tab(tab0); linea
    
    If j < n_Copias Then
        prt.NewPage
    Else
        prt.EndDoc
    End If
    
Next

Impresora_Predeterminada "default"

End Sub
Private Sub N_Oc_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
