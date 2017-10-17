VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form OCe_Buscar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busca OC Especial"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.TextBox Descripcion 
         Height          =   300
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox Nv 
         Height          =   300
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton btnImprimir 
         Height          =   400
         Left            =   6600
         Picture         =   "OCe_Buscar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
         Width           =   400
      End
      Begin VB.CommandButton btnLimpiar 
         Caption         =   "&Limpiar"
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton btnBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   4920
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox ComboNV 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lbl 
         Caption         =   "&Descripcion"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "&Obra"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Detalle 
      Height          =   3615
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6376
      _Version        =   327680
      FixedCols       =   0
   End
End
Attribute VB_Name = "OCe_Buscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DbD As Database, RsPrv As Recordset
Private Dbm As Database, RsNVc As Recordset
Private Dba As Database, RsOCd As Recordset, RsRMd As Recordset
Private indice As Integer, n_filas As Integer, n_columnas As Integer
Private m_Fecha As Date ', m_Plano As String, m_SubC As String
Private a_Nv(2999, 1) As String, i As Integer, m_Nv As Double
Private qry As String
Private Sub Form_Load()

Set DbD = OpenDatabase(data_file)

Set RsPrv = DbD.OpenRecordset("Proveedores")
RsPrv.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)

Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set Dba = OpenDatabase(Madq_file)
'Set RsOCd = DbA.OpenRecordset("OC Detalle")
'RsOCd.Index = "Número"

Nv.MaxLength = 5

' Combo Obra
i = 0
ComboNV.AddItem " "
Do While Not RsNVc.EOF

    If Usuario.Nv_Activas = False Then ' todas
        GoTo IncluirNV
    Else
        If Usuario.Nv_Activas And RsNVc!Activa Then
            GoTo IncluirNV
        End If
    End If
    
    If False Then
IncluirNV:
        i = i + 1
        a_Nv(i, 0) = RsNVc!Numero
        a_Nv(i, 1) = RsNVc!Obra
        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!Obra
    End If
    
    RsNVc.MoveNext
    
Loop

n_filas = 20
n_columnas = 9 ' 7  '6
Resultado_Config

'btnBuscar.Enabled = False

btnLimpiar_Click

End Sub
Private Sub Resultado_Config()

Dim i As Integer, ancho As Integer

Detalle.WordWrap = True
Detalle.RowHeight(0) = 450
Detalle.Rows = n_filas + 1
Detalle.Cols = n_columnas

Detalle.TextMatrix(0, 0) = "OC Nº"
Detalle.TextMatrix(0, 1) = "Fecha"
Detalle.TextMatrix(0, 2) = "Proveedor"
Detalle.TextMatrix(0, 3) = "Cant"
Detalle.TextMatrix(0, 4) = "Uni"
Detalle.TextMatrix(0, 5) = "Descripcion"
Detalle.TextMatrix(0, 6) = "$ Uni"

Detalle.ColWidth(0) = 600
Detalle.ColWidth(1) = 900
Detalle.ColWidth(2) = 1100
Detalle.ColWidth(3) = 700
Detalle.ColWidth(4) = 400
Detalle.ColWidth(5) = 2600
Detalle.ColWidth(6) = 900
Detalle.ColWidth(7) = 0
Detalle.ColWidth(8) = 0

End Sub
Private Sub Nv_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Nv_LostFocus()

m_Nv = m_CDbl(Nv.Text)
If m_Nv = 0 Then Exit Sub
' busca nv en combo
i = 1
Do Until a_Nv(i, 0) = ""
    If Val(a_Nv(i, 0)) = m_Nv Then
        ComboNV.ListIndex = i
        Exit Sub
    End If
    i = i + 1
Loop

MsgBox "NV no existe"
Nv.SetFocus
'SendKeys "{Home}+{End}"

End Sub
Private Sub ComboNV_Click()

MousePointer = vbHourglass

m_Nv = Val(Left(ComboNV.Text, 6))
If m_Nv = 0 Then
    Nv.Text = ""
Else
    Nv.Text = m_Nv
End If

btnBuscar.Enabled = True

MousePointer = vbDefault

End Sub
Private Sub btnBuscar_Click()

If Len(Descripcion.Text) < 2 Then
    MsgBox "Debe digitar al menos 2 caracteres"
    Descripcion.SetFocus
    Exit Sub
End If
If InStr(1, Descripcion.Text, "'") > 0 Then
    MsgBox "caracter NO Valido '"
    Descripcion.SetFocus
    Exit Sub
End If

MousePointer = vbHourglass

Detalle_Limpiar

'ComboNV.Enabled = False
'Marca.Enabled = False
'btnBuscar.Enabled = False

btnLimpiar.Enabled = True

qry = "SELECT * FROM [OC Detalle] "
qry = qry & " WHERE descripcion LIKE '*" & Descripcion.Text & "*'"

If Nv.Text <> "" Then
    qry = qry & " AND Nv=" & Nv.Text
End If

Set RsOCd = Dba.OpenRecordset(qry)

indice = 0
With RsOCd
Do While Not .EOF

    indice = indice + 1
    If indice = n_filas Then
        n_filas = indice + 1
        Detalle.Rows = n_filas + 1
    End If
    Detalle.TextMatrix(indice, 0) = !Numero
    Detalle.TextMatrix(indice, 1) = !Fecha
    
    RsPrv.Seek "=", ![RUT Proveedor]
    If Not RsPrv.NoMatch Then
        Detalle.TextMatrix(indice, 2) = RsPrv![Razon Social]
    End If
    
    Detalle.TextMatrix(indice, 3) = !Cantidad
    Detalle.TextMatrix(indice, 4) = !unidad
    Detalle.TextMatrix(indice, 5) = !Descripcion
    Detalle.TextMatrix(indice, 6) = ![Precio Unitario]
    
    Set RsRMd = Dba.OpenRecordset("SELECT * FROM [rm cabecera] WHERE oc=" & !Numero)
    If RsRMd.RecordCount > 0 Then
        Detalle.TextMatrix(indice, 7) = RsRMd![Tipo Documento]
        Detalle.TextMatrix(indice, 8) = RsRMd![Numero Documento]
    End If
    
'    Debug.Print !Numero, !Fecha, !descripcion
    
    .MoveNext
    
Loop

.Close

End With

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
Nv.Text = ""
ComboNV.Text = " "
Descripcion.Text = ""
Detalle_Limpiar

ComboNV.Enabled = True

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
PrinterNCopias.Numero_Copias = 1
PrinterNCopias.Show 1
If PrinterNCopias.Numero_Copias > 0 Then
    Imprimir PrinterNCopias.Numero_Copias
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
DataBases_Cerrar
End Sub
Private Sub Imprimir(n_Copias)
Dim j As Integer
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer, tab4 As Integer
Dim tab5 As Integer, tab6 As Integer, tab7 As Integer
Dim prt As Printer, li As Integer, linea As String

Dim sep As String
sep = ","

linea = String(76, "-")
tab0 = 5 'margen izquierdo
tab1 = tab0 + 0  'num
tab2 = tab1 + 6  'fecha
tab3 = tab2 + 10 'prov
tab4 = tab3 + 12 'cant
tab5 = tab4 + 10 'uni
tab6 = tab5 + 5  'descrip
tab7 = tab6 + 20 'largo

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt

For j = 1 To n_Copias

    prt.Font.Size = 12
    prt.Print Tab(tab0); Empresa.Razon
    prt.Print Tab(tab0); " "
    prt.Print Tab(tab0); "Busca OT Especial"
    prt.Print Tab(tab0); "FECHA "; Format(Now, Fecha_Format)
    prt.Print Tab(tab0); " "
    prt.Print Tab(tab0); "OBRA  : "; ComboNV.Text
    prt.Print Tab(tab0); " "

    prt.Font.Bold = True
    prt.Print Tab(tab1); "NÚMERO";
    prt.Print Tab(tab2); " FECHA";
    prt.Print Tab(tab3); "PROVEEDOR";
    prt.Print Tab(tab4); "  CANT";
    prt.Print Tab(tab5); "UNI";
    prt.Print Tab(tab6); "DESCRIPCION";
    prt.Print Tab(tab7); "$ UNI"
    prt.Font.Bold = False
    
    prt.Print Tab(tab0); linea
    
    For li = 1 To n_filas - 1
        If Trim(Detalle.TextMatrix(li, 0)) <> "" Then
        
            prt.Print Tab(tab1); PadL(Detalle.TextMatrix(li, 0), 5);
            prt.Print Tab(tab2); Format(Detalle.TextMatrix(li, 1), Fecha_Format);
            prt.Print Tab(tab3); PadR(Detalle.TextMatrix(li, 2), 10); 'proveedor
            prt.Print Tab(tab4); PadL(Detalle.TextMatrix(li, 3), 6); ' cant
            prt.Print Tab(tab5); PadR(Detalle.TextMatrix(li, 4), 3); ' uni
            prt.Print Tab(tab6); Detalle.TextMatrix(li, 5); ' descr
            prt.Print Tab(tab7); Detalle.TextMatrix(li, 6) ' precio unitario
            
            Debug.Print PadL(Detalle.TextMatrix(li, 0), 5); sep; _
             Format(Detalle.TextMatrix(li, 1), Fecha_Format); sep; _
             Detalle.TextMatrix(li, 2); sep; _
             PadL(Detalle.TextMatrix(li, 3), 6); sep; _
             PadR(Detalle.TextMatrix(li, 4), 18); sep; _
             Detalle.TextMatrix(li, 5); sep; _
             Detalle.TextMatrix(li, 6); sep; _
             Detalle.TextMatrix(li, 7); sep; _
             Detalle.TextMatrix(li, 8)
            
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
