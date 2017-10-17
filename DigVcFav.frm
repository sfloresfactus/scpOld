VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form DigVcFav 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vales de Consumo Facturados"
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
   Begin VB.TextBox txtEdit 
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton btnImprimir 
      Height          =   400
      Left            =   3360
      Picture         =   "DigVcFav.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   400
   End
   Begin VB.CheckBox Check_Todos 
      Caption         =   "Todos los Vales"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton btnGrabar 
      Caption         =   "&Grabar"
      Height          =   315
      Left            =   1680
      TabIndex        =   1
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
End
Attribute VB_Name = "DigVcFav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private n_columnas As Integer, n_filas As Integer, i As Integer
Private DbD As Database, RsCl As Recordset
Private Dbm As Database, RsNVc As Recordset, RsVc As Recordset, Dba As Database

Private m_Nv As Double, m_Rut As String, m_FechaInicio As String, m_FechaTermino As String
Private m_NvArea As Integer
Private m_Total As String
Public Property Let Numero_NotaVenta(ByVal New_Opcion As Double)
m_Nv = New_Opcion
End Property
Public Property Let Rut_Contratista(ByVal New_Opcion As String)
m_Rut = New_Opcion
End Property
Public Property Let Fecha_Inicio(ByVal New_Opcion As String)
m_FechaInicio = New_Opcion
End Property
Public Property Let Fecha_Termino(ByVal New_Opcion As String)
m_FechaTermino = New_Opcion
End Property

Private Sub Form_Load()

btnImprimir.visible = False

txtEdit.visible = False

'Debug.Print "|" & m_NV & "|"
'Debug.Print "|" & m_RUT & "|"
'Debug.Print "|" & m_FechaInicio & "|"
'Debug.Print "|" & m_FechaTermino & "|"

n_columnas = 6 '7
n_filas = 10

Set DbD = OpenDatabase(data_file)
Set RsCl = DbD.OpenRecordset("Contratistas")
RsCl.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)

Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"


Set Dba = OpenDatabase(Madq_file)
Set RsVc = Dba.OpenRecordset("documentos") ' ojo
'RsGDc.Index = "Numero"

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
fg.ColWidth(4) = 1300
fg.ColWidth(5) = 900
fg.ColWidth(6) = 800

ancho = 350
For i = 0 To fg.Cols - 1
    ancho = ancho + fg.ColWidth(i)
Next

fg.Width = ancho
Me.Width = ancho + fg.Left * 2

i = 0
fg.TextMatrix(i, 0) = ""
fg.TextMatrix(i, 1) = "VC"
fg.TextMatrix(i, 2) = "Fecha"
fg.TextMatrix(i, 3) = "Contratista"
fg.TextMatrix(i, 4) = "Obra"
fg.TextMatrix(i, 5) = "Total"
fg.TextMatrix(i, 6) = "Factura"

End Sub
Private Sub Leer()
' lee archivo y le trae al FlexGrid
Dim sql As String, p As Integer, j As Integer
Dim descripcion As String, m_doble As Double, m_Num As Double, m_Fec As Date, m_Oc As Double

sql = "SELECT "
sql = sql & "[documentos].NV as NV,"
sql = sql & "[documentos].tipo as tipo,"
sql = sql & "[documentos].Numero as numero,"
sql = sql & "[documentos].fecha as fecha,"
sql = sql & "[documentos].cant_sale as cantidad,"
sql = sql & "[documentos].[RUT] as rut,"
sql = sql & "[documentos].[Precio unitario] as precio,"
sql = sql & "[documentos].oc as oc"
sql = sql & " FROM [documentos]"


p = 0
If m_Nv > 0 Then sql = sql & " WHERE NV=" & m_Nv: p = p + 1

If m_Rut <> "" Then
    m_Rut = SqlRutPadL(m_Rut)
    sql = sql & " " & IIf(p = 0, "WHERE", "AND")
    sql = sql & " [rut]='" & m_Rut & "'"
    p = p + 1
End If

If m_FechaInicio <> "__/__/__" Then
    sql = sql & " " & IIf(p = 0, "WHERE", "AND")
    sql = sql & " [fecha]>=CDate('" & m_FechaInicio & "')"
    p = p + 1
End If

If m_FechaTermino <> "__/__/__" Then
    sql = sql & " " & IIf(p = 0, "WHERE", "AND")
    sql = sql & " [fecha]<=CDate('" & m_FechaTermino & "')"
    p = p + 1
End If

If True Then
If Check_Todos.Value = 1 Then
    'todos
Else
    ' solo pendientes
    sql = sql & " " & IIf(p = 0, "WHERE", "AND")
    sql = sql & " oc=0"
    p = p + 1
End If
End If

'sql = sql & " GROUP BY [documentos].numero"


sql = sql & " " & IIf(p = 0, "WHERE", "AND")
sql = sql & " tipo='VC'" ' ojo

sql = sql & " ORDER BY numero"

Set RsVc = Dba.OpenRecordset(sql)

n_filas = 10
fg.Rows = n_filas + 1

' limpia flex
For i = 1 To n_filas
    For j = 1 To n_columnas
        fg.TextMatrix(i, j) = ""
    Next
Next

i = 0
With RsVc
If .RecordCount > 0 Then

    .MoveFirst
    m_Num = !Numero
    m_Total = 0
    Do While Not .EOF
        
        If m_Num = !Numero Then
        
            m_Nv = !Nv
            m_Total = m_Total + !Cantidad * !precio
    
        Else
        
            Linea_Mostrar m_Num, m_Fec, m_Oc, False
            
        End If
        
        m_Fec = !Fecha
        m_Rut = NoNulo(!rut)
        m_Oc = !Oc
        
        .MoveNext
        
    Loop
    
    Linea_Mostrar m_Num, m_Fec, m_Oc, True

End If

End With
End Sub
Private Sub Linea_Mostrar(m_Num As Double, m_Fec As Date, m_Oc As Double, FinArchivo As Boolean)

Dim descripcion As String

i = i + 1
If i > n_filas Then
    n_filas = n_filas + 1
    fg.Rows = n_filas + 1
End If

fg.TextMatrix(i, 1) = m_Num
fg.TextMatrix(i, 2) = m_Fec

descripcion = ""
RsCl.Seek "=", m_Rut
If Not RsCl.NoMatch Then
    descripcion = RsCl![Razon Social]
End If
fg.TextMatrix(i, 3) = descripcion

descripcion = ""
'        m_doble = 0

descripcion = ""
RsNVc.Seek "=", m_Nv, m_NvArea
If Not RsNVc.NoMatch Then
    descripcion = RsNVc![obra]
End If
fg.TextMatrix(i, 4) = descripcion
fg.TextMatrix(i, 5) = m_Total
fg.TextMatrix(i, 6) = m_Oc 'factura

If Not FinArchivo Then
    m_Num = RsVc!Numero
    m_Total = RsVc!Cantidad * RsVc!precio
End If

End Sub
Private Sub Check_Todos_Click()
' pregunta si realmente quiere todos
If Check_Todos.Value = 1 Then
    If MsgBox("¿ Todos ?", vbYesNo) = vbYes Then
        Check_Todos.Value = 1
    Else
        Check_Todos.Value = 0
    End If
End If
Leer
'Todos_Leer
End Sub
Private Sub btnGrabar_Click()
Dim m_Numero As Double

RsVc.Close

If False Then
    Set RsVc = Dba.OpenRecordset("documentos")
    With RsVc
    .Index = "tipo-numero-linea"
    
    For i = 1 To n_filas - 1
    
        m_Numero = Val(fg.TextMatrix(i, 1))
        If m_Numero > 0 Then
            .Seek "=", "VC", m_Numero, 1
            If Not .NoMatch Then
                Do While Not .EOF
                    If !Tipo <> "VC" Or !Numero <> m_Numero Then Exit Do
                    .Edit
                    !Oc = Val(fg.TextMatrix(i, 6))
                    .Update
                    .MoveNext
                Loop
            End If
        End If
        
    Next
    .Close
    End With
Else
    For i = 1 To n_filas - 1
        m_Numero = Val(fg.TextMatrix(i, 1))
        If m_Numero > 0 Then
            Dba.Execute "UPDATE documentos SET oc=" & Val(fg.TextMatrix(i, 6)) & " WHERE tipo='VC' AND numero=" & m_Numero
        End If
    Next
End If

Unload Me
End Sub
Private Sub btnImprimir_Click()
Imprimir
End Sub
Private Sub Imprimir()
Dim j As Integer
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer, tab4 As Integer
Dim tab5 As Integer, tab6 As Integer, prt As Printer, li As Integer, linea As String

linea = String(76, "-")
tab0 = 5 'margen izquierdo
tab1 = tab0 + 0  'oc
tab2 = tab1 + 10 'fecha
tab3 = tab2 + 10 'doc
tab4 = tab3 + 5  'numero
tab5 = tab4 + 12 'obra

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt

For j = 1 To 1 'n_Copias

    prt.Font.Size = 12
    prt.Print Tab(tab0); Empresa.Razon
    prt.Print Tab(tab0); " "
    prt.Print Tab(tab0); "Certificados de Calidad Pendientes"
    prt.Print Tab(tab0); "FECHA "; Format(Now, Fecha_Format)
    prt.Print Tab(tab0); " "
    prt.Print Tab(tab0); "Proveedor : "; UCase(fg.TextMatrix(1, 3))
    prt.Print Tab(tab0); " "

    prt.Font.Bold = True
    prt.Print Tab(tab1); PadL("OC", 6);
    prt.Print Tab(tab2); PadL("FECHA", 8);
    prt.Print Tab(tab3); "Doc";
    prt.Print Tab(tab4); PadL("Nº", 10);
    prt.Print Tab(tab5); "OBRA"
    prt.Font.Bold = False
    
    prt.Print Tab(tab0); linea
    
    For li = 1 To n_filas - 1
        If Trim(fg.TextMatrix(li, 1)) <> "" Then
            prt.Print Tab(tab1); PadL(fg.TextMatrix(li, 1), 6);
            prt.Print Tab(tab2); Format(fg.TextMatrix(li, 2), Fecha_Format);
            prt.Print Tab(tab3); PadL(fg.TextMatrix(li, 4), 3);
            prt.Print Tab(tab4); PadL(fg.TextMatrix(li, 5), 10);
            prt.Print Tab(tab5); PadR(fg.TextMatrix(li, 6), 20);
'            prt.Print Tab(tab6); fg.TextMatrix(li, 5)
        End If
    Next
    
    prt.Print Tab(tab0); linea
    prt.Print Tab(tab0); "G: Guías de Despacho"
    prt.Print Tab(tab0); "F: Factura de Venta"
    
    prt.EndDoc
    
Next

Impresora_Predeterminada "default"

End Sub
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'
' RUTINAS PARA EL FLEXGRID
'
'
Private Sub fg_DblClick()
' simula un espacio
'If Imprimiendo Then Exit Sub
MSFlexGridEdit fg, txtEdit, 32
End Sub
Private Sub fg_GotFocus()
'If Imprimiendo Then Exit Sub
If txtEdit.visible = False Then Exit Sub
fg = txtEdit
txtEdit.visible = False
End Sub
Private Sub fg_LeaveCell()
If txtEdit.visible = False Then Exit Sub
fg = txtEdit
txtEdit.visible = False
End Sub
Private Sub fg_KeyPress(KeyAscii As Integer)
'If Imprimiendo Then Exit Sub
MSFlexGridEdit fg, txtEdit, KeyAscii
End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode fg, txtEdit, KeyCode, Shift
End Sub
Private Sub txtEdit_LostFocus()
Cursor_NoMueve
End Sub
Sub EditKeyCode(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape ' Esc
    Edt.visible = False
    MSFlexGrid.SetFocus
Case vbKeyReturn ' Enter
    MSFlexGrid.SetFocus
    DoEvents
'    Actualiza
    Cursor_Mueve MSFlexGrid
Case 38 ' Flecha Arriba
    MSFlexGrid.SetFocus
    DoEvents
'    Actualiza
    If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
        MSFlexGrid.Row = MSFlexGrid.Row - 1
    End If
Case 40 ' Flecha Abajo
    MSFlexGrid.SetFocus
    DoEvents
'    Actualiza
    If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
        MSFlexGrid.Row = MSFlexGrid.Row + 1
    End If
End Select
End Sub
Private Sub txtEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)

If MSFlexGrid.col = 6 Then
    Select Case KeyAscii
    Case 0 To 32
        Edt = MSFlexGrid
        Edt.SelStart = 1000
    Case Else
        Edt = Chr(KeyAscii)
        Edt.SelStart = 1
    End Select
    'Edt.Move SSTab.Left + MSFlexGrid.CellLeft + MSFlexGrid.Left, SSTab.Top + MSFlexGrid.CellTop + MSFlexGrid.Top, MSFlexGrid.CellWidth, MSFlexGrid.CellHeight
    Edt.Move MSFlexGrid.CellLeft + MSFlexGrid.Left, MSFlexGrid.CellTop + MSFlexGrid.Top, MSFlexGrid.CellWidth, MSFlexGrid.CellHeight
    Edt.visible = True
    Edt.SetFocus
End If

End Sub
Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)
'If Imprimiendo Then Exit Sub
If KeyCode = vbKeyF2 Then
    MSFlexGridEdit fg, txtEdit, 32
End If
End Sub
Private Sub Cursor_Mueve(MSFlexGrid As Control)
'MIA
Select Case MSFlexGrid.col
Case 6
    If MSFlexGrid.Row + 1 < MSFlexGrid.Rows Then
        MSFlexGrid.col = 1
        MSFlexGrid.Row = MSFlexGrid.Row + 1
    End If
Case Else
    MSFlexGrid.col = MSFlexGrid.col + 1
End Select
End Sub
'
' FIN RUTINAS PARA FLEXGRID
'
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
Private Sub Cursor_NoMueve()
i = fg.col
fg.col = IIf(i = 1, 2, 1)
fg.col = i
End Sub
