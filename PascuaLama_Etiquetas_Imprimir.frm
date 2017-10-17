VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PascuaLama_Etiquetas_Imprimir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprime Etiquitas por Lotes"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEdit 
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton btnImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   5760
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid fg 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9763
      _Version        =   327680
   End
End
Attribute VB_Name = "PascuaLama_Etiquetas_Imprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' rutina especifica
' para impresion de etiquetas autoadhesivas
' formato 4" x 4" (10cm x 10cm)
' exclusivo para cliente Barrick, proyecto Pacual Lama
' marzo 2011
' solicitado por Jose Luis Gonzalez y Rodrigo Nuñez
Option Explicit
Private i As Integer, n_filas As Integer
Private prt As Printer

Dim m_Cli As String ' cliente
Dim m_Pro As String ' proyecto
Dim m_ProNum As String ' proyecto numero
Dim m_OdcNum As String ' odc numero
Dim m_MKN As String ' MK Nº
Dim m_MKD As String ' MK Descripcion
Dim m_NVN As String ' NV + Nombre
Dim m_KgU As String ' kg unitario
Dim m_Can As String ' cantidad

Private AjusteX As Double, AjusteY As Double
'Private Dbm As Database, RsPd As Recordset
Private RsEtiq As New ADODB.Recordset
'Private a_Nv(3, 1) As String, Total_Obras As Integer
Private sql As String
Private Sub btnImprimir_Click()
Dim Cant As Integer
'Prt_Ini
For i = 1 To n_filas
    If fg.TextMatrix(i, 7) = "SI" Then
        Cant = Val(fg.TextMatrix(i, 6))
        Etiquetas_Imprimir Cant, i
    End If
Next
prt.EndDoc

End Sub

Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    MSFlexGridEdit fg, txtEdit, 32
End If
End Sub

' 0: nv  1: obra
Private Sub Form_Load()

'Total_Obras = 1

'Set Dbm = OpenDatabase(mpro_file)

'sql = "SELECT * FROM [planos detalle] WHERE nv=2717 ORDER BY nv,plano,marca"
'sql = "SELECT * FROM [etiquetas_barrick] ORDER BY mk_numero"
sql = "SELECT * FROM [etiquetas_barrick] ORDER BY nv_nombre,mk_descripcion,mk_numero"
'Set RsPd = Dbm.OpenRecordset(sql)
'RsPd.Index = "nv-plano-marca"

'a_Nv(1, 0) = 2717
'a_Nv(1, 1) = "PASCUALAMA TRIPPER CAR"

' lee archivo excel con marcas a imprimir
n_filas = 2999
fg.Rows = n_filas + 1
fg.Cols = 8

fg.ColWidth(0) = 500
fg.ColWidth(1) = 2200 ' ODC Numero
fg.ColWidth(2) = 1800 ' 3100 ' MK Numero
fg.ColWidth(3) = 1600 ' MK Descripcion
fg.ColWidth(4) = 2000 ' nv + nombre
fg.ColWidth(5) = 650  ' peso
fg.ColWidth(6) = 600  ' cantidad total
fg.ColWidth(7) = 700  ' si/no indica si se imprime o no

fg.TextMatrix(0, 0) = ""
fg.TextMatrix(0, 1) = "ODC Numero"
fg.TextMatrix(0, 2) = "MK Numero"
fg.TextMatrix(0, 3) = "MK Descripcion"
fg.TextMatrix(0, 4) = "NV Nombre"
fg.TextMatrix(0, 5) = "Kg Uni"
fg.TextMatrix(0, 6) = "Cantidad"
fg.TextMatrix(0, 7) = "Imprimir"

For i = 1 To n_filas
    fg.TextMatrix(i, 0) = i
    fg.Row = i
    fg.col = 7
    fg.CellForeColor = vbRed
Next

fg.Width = fg.ColWidth(0) + fg.ColWidth(1) + fg.ColWidth(2) + fg.ColWidth(3) + fg.ColWidth(4) + fg.ColWidth(5) + fg.ColWidth(6) + fg.ColWidth(7) + 400
Me.Width = fg.Width + 400

'm_Nv = 2315
'm_obra = "Placas Desgaste Sulfuros RT"

Piezas_Leer

End Sub
Private Sub fg_DblClick()

If fg.col = 6 Then ' cantidad
    MSFlexGridEdit fg, txtEdit, vbKeySpace
End If

If fg.col = 7 Then ' si/no
    Estado_Cambiar
End If

End Sub
Private Sub fg_KeyPress(KeyAscii As Integer)

Dim Cant As Integer

If fg.col = 6 Then ' cantidad
    MSFlexGridEdit fg, txtEdit, KeyAscii
End If

If fg.col = 7 Then ' si/no
    If KeyAscii = vbKeySpace Then ' 32
        Estado_Cambiar
    End If
End If

'If KeyAscii = 13 Then
'    Cant = Val(fg.TextMatrix(fg.Row, 6))
'    Etiquetas_Imprimir Cant, fg.Row
'    fg.Row = fg.Row + 1
'End If


End Sub
Private Sub Estado_Cambiar()

If Trim(fg.TextMatrix(fg.Row, 7)) = "SI" Then
    fg.TextMatrix(fg.Row, 7) = ""
Else
    fg.TextMatrix(fg.Row, 7) = "SI"
End If
    
End Sub
Private Sub Etiquetas_Imprimir(Cantidad As Integer, linea As Integer)

Dim li As Integer

Dim li1 As Double, li2 As Double, li3 As Double, li4 As Double, li5 As Double, li6 As Double, li7 As Double, li8 As Double, li9 As Double
Dim dif_linea As Double, Copia As Integer
Dim Margen_Izquierdo As Double, m_TamanoCodeBar As Integer

dif_linea = 0.7 ' 0.57

Margen_Izquierdo = 0.5

m_TamanoCodeBar = Val(ReadIniValue(Path_Local & "scp.ini", "Printer", "LabelBarCodeSize"))
If m_TamanoCodeBar = 0 Then m_TamanoCodeBar = 30

m_TamanoCodeBar = 42
' 40 -> 1,25
' 42
' 50 -> 1,6

li1 = 1.4
li2 = li1 + dif_linea
li3 = li2 + dif_linea
li4 = li3 + dif_linea
li5 = li4 + dif_linea
li6 = li5 + dif_linea
li7 = li6 + dif_linea
li8 = li7 + dif_linea
li9 = li8 + dif_linea ' codigo de barras

If Cantidad > 0 Then

    For Copia = 1 To Cantidad
    
'            Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Etiquetas") ' puesta el 16/03/06
'            MsgBox Printer.DeviceName & vbLf & prt.DeviceName
        
        Prt_Ini
    
        m_OdcNum = fg.TextMatrix(linea, 1)
        m_MKN = fg.TextMatrix(linea, 2)
        m_MKD = fg.TextMatrix(linea, 3)
        m_NVN = fg.TextMatrix(linea, 4)
        m_KgU = fg.TextMatrix(linea, 5)
        m_Can = fg.TextMatrix(linea, 6)
        
'            m_Peso = m_CDbl(Detalle.TextMatrix(li, 16)) ' old 12
    
        ' font para logo
        Printer.Font.Name = "delgado"
        Printer.Font.Size = 32
'            Printer.Font.Bold = True
        
        SetpYX -0.15, Margen_Izquierdo
        Printer.Print "Delgado";
        '//////////////////
        
        Printer.Font.Name = "Arial"
        Printer.Font.Bold = False
        Printer.Font.Size = 15
                  
        SetpYX 0.5, 6
        prt.Print Format(Date, "dd/mm/yyyy")
                  
        SetpYX li1, Margen_Izquierdo
        prt.Print "Cliente: BARRICK"
                
        SetpYX li2, Margen_Izquierdo
        prt.Print "Proyecto: PASCUALAMA"
        
        SetpYX li3, Margen_Izquierdo
        prt.Print "Proyecto Número: P5SL"
        
        SetpYX li4, Margen_Izquierdo
        prt.Print "ODC Número: "; m_OdcNum
        
        SetpYX li5, Margen_Izquierdo
        prt.Print "MK Número: "; m_MKN
        
        SetpYX li6, Margen_Izquierdo
        prt.Print "MK Descripción: "; m_MKD
        
        SetpYX li7, Margen_Izquierdo
        prt.Print m_NVN
        
        SetpYX li8, Margen_Izquierdo
        prt.Print "Peso(Kg): "; m_KgU
        
        prt.Font.Name = "code 128"
        prt.Font.Size = m_TamanoCodeBar ' 32 oficial
        
        SetpYX li9, 0.4 'Margen_Izquierdo
        prt.Print txt2code128(m_MKN)
        
'        Printer.Font.Name = "Arial"
'        Printer.Font.Size = 16
        
'        SetpYX li7 + 1.9, Margen_Izquierdo
'        prt.Print Format(Date, "dd/mm/yyyy")
        
        ' punto para saltodeetiqueta
            prt.Font.Name = "Arial"
            prt.Font.Size = 8
        If False Then
            SetpYX 4.91, 9
        Else
            SetpYX 9.91, 9
        End If
        
        prt.Print "."
        
    
    Next
    
End If

'prt.EndDoc
    
End Sub
Private Sub Prt_Ini()
Set prt = Printer
Printer.ScaleMode = 1 ' twips : 576 twips x cm
Printer.ScaleMode = 7 ' centimetros
End Sub
Private Sub SetpYX(Y As Double, x As Double)
Printer.CurrentY = AjusteY + Y
Printer.CurrentX = AjusteX + x
End Sub
Private Sub Piezas_Leer()

Dim fi As Integer
'Dim filas_vacias As Integer, m_NvAnterior As Double, PrimeraLineaOc As Boolean

Rs_Abrir RsEtiq, sql

fi = 0
With RsEtiq
Do While Not .EOF

    m_Can = ![Cantidad]
    
    If m_Can > 0 Then
    
        fi = fi + 1
               
        fg.TextMatrix(fi, 1) = !odc_numero
        fg.TextMatrix(fi, 2) = !mk_numero
        fg.TextMatrix(fi, 3) = !mk_descripcion
        fg.TextMatrix(fi, 4) = !NV_nombre
        fg.TextMatrix(fi, 5) = !kg_unitario
        fg.TextMatrix(fi, 6) = !Cantidad
'        fg.TextMatrix(fi, 7) = !descripcion
    
    End If
    
    .MoveNext
        
Loop

End With

End Sub
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)
Select Case MSFlexGrid.col
Case 6

    Edt = Chr(KeyAscii)
    Edt.SelStart = 1

'    Edt = MSFlexGrid

    Edt.Move MSFlexGrid.CellLeft + MSFlexGrid.Left, MSFlexGrid.CellTop + MSFlexGrid.Top, MSFlexGrid.CellWidth, MSFlexGrid.CellHeight + 50
    Edt.visible = True
    Edt.SetFocus
End Select
End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode fg, txtEdit, KeyCode, Shift
End Sub
Private Sub txtEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Private Sub EditKeyCode(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
' rutina que se ejecuta con los keydown de Edt
Dim m_col As Integer, Cantidad As Integer

m_col = MSFlexGrid.col

'Cantidad = Val(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 6))

'Exit Sub

'Edt.visible = False

Select Case KeyCode
Case vbKeyEscape ' Esc
    Edt.visible = False
    MSFlexGrid.SetFocus
Case vbKeyReturn ' Enter
    
    Cantidad = Val(Edt.Text)
    
    If Cantidad <= 0 Then
        MsgBox "Cantidad No Válida"
        Exit Sub
    End If
    
    MSFlexGrid = Cantidad
    
    MSFlexGrid.TextMatrix(MSFlexGrid.Row, 7) = "SI"
    
    MSFlexGrid.Row = MSFlexGrid.Row + 1
    
    Edt.visible = False
    MSFlexGrid.SetFocus
    DoEvents
    
End Select

End Sub
