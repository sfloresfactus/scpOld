VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form EtiquetasContratistasImprimir 
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
   Begin VB.PictureBox PictureQrCode 
      Height          =   1170
      Left            =   0
      Picture         =   "EtiquetasContratistasImprimir.frx":0000
      ScaleHeight     =   74
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   50
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   810
   End
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
      Left            =   7200
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
Attribute VB_Name = "EtiquetasContratistasImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' rutina especifica
' para impresion de etiquetas autoadhesivas
' formato 4" x 2" (10cm x 5cm)
' exclusivo para Contratistas, todas las NV
' noviembre 2012
Option Explicit
Private i As Integer, n_filas As Integer
Private prt As Printer

Dim m_KgU As String ' kg unitario
Dim m_Can As String ' cantidad

Private AjusteX As Double, AjusteY As Double
'Private Dbm As Database, RsPd As Recordset
'Private a_Nv(3, 1) As String, Total_Obras As Integer
Private sql As String
Private Const Columna As Integer = 9
Private cQrCode As ClsQrCode
Private Sub btnImprimir_Click()
Dim Cant As Integer
'Prt_Ini
For i = 1 To n_filas
    If fg.TextMatrix(i, Columna) = "SI" Then
'        Cant = Val(fg.TextMatrix(i, 8))
        etiquetaImprimirV1211 i
    End If
Next
prt.EndDoc

End Sub

Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    MSFlexGridEdit fg, txtEdit, 32
End If
End Sub
Private Sub Form_Load()

'Debug.Print Usuario.nombre

n_filas = 3999
fg.Rows = n_filas + 1
fg.Cols = Columna + 1

fg.ColWidth(0) = 500
fg.ColWidth(1) = 500  ' NV
fg.ColWidth(2) = 1500 ' Obra
fg.ColWidth(3) = 1600 ' plano
fg.ColWidth(4) = 300  ' rev
fg.ColWidth(5) = 1500 ' marca
fg.ColWidth(6) = 0 ' descripcion
fg.ColWidth(7) = 650  ' kg uni
fg.ColWidth(8) = 600  ' cant total
fg.ColWidth(9) = 700  ' si/no indica si se imprime o no

fg.TextMatrix(0, 0) = ""
fg.TextMatrix(0, 1) = "NV"
fg.TextMatrix(0, 2) = "Obra"
fg.TextMatrix(0, 3) = "Plano"
fg.TextMatrix(0, 4) = "Rev"
fg.TextMatrix(0, 5) = "Marca"
fg.TextMatrix(0, 6) = "Marca"
fg.TextMatrix(0, 7) = "Kg Uni"
fg.TextMatrix(0, 8) = "Cantidad"
fg.TextMatrix(0, 9) = "Imprimir"

For i = 1 To n_filas
    fg.TextMatrix(i, 0) = i
    fg.Row = i
    fg.col = Columna
    fg.CellForeColor = vbRed
Next

fg.Width = 0
For i = 0 To Columna
    fg.Width = fg.Width + fg.ColWidth(i)
Next
fg.Width = fg.Width + 400

Me.Width = fg.Width + 400

'm_Nv = 2315
'm_obra = "Placas Desgaste Sulfuros RT"

Piezas_Leer

End Sub
Private Sub fg_DblClick()

If fg.col = 8 Then ' cantidad
    MSFlexGridEdit fg, txtEdit, vbKeySpace
End If

If fg.col = Columna Then ' si/no
    Estado_Cambiar
End If

End Sub
Private Sub fg_KeyPress(KeyAscii As Integer)

Dim Cant As Integer

If fg.col = 8 Then ' cantidad
    MSFlexGridEdit fg, txtEdit, KeyAscii
End If

If fg.col = Columna Then ' si/no
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

If Trim(fg.TextMatrix(fg.Row, Columna)) = "SI" Then
    fg.TextMatrix(fg.Row, Columna) = ""
Else
    fg.TextMatrix(fg.Row, Columna) = "SI"
End If
    
End Sub
Private Sub Etiquetas_Imprimir_Old(Cantidad As Integer, linea As Integer)

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
    
        m_KgU = fg.TextMatrix(linea, 5)
        m_Can = fg.TextMatrix(linea, 6)
        
'            m_Peso = m_CDbl(Detalle.TextMatrix(li, 16)) ' old 12
    
        ' font para logo
        Printer.Font.Name = "delgado"
        Printer.Font.Size = 30
'            Printer.Font.Bold = True
        
        SetpYX -0.15, Margen_Izquierdo
        Printer.Print "Delgado";
        '//////////////////
        
        Printer.Font.Name = "Arial"
        Printer.Font.Bold = False
        Printer.Font.Size = 12
                  
        SetpYX 0.5, 6
        prt.Print Format(Date, "dd/mm/yyyy")
                  
        SetpYX li1, Margen_Izquierdo
        prt.Print "Cliente: BARRICK"
                
        SetpYX li2, Margen_Izquierdo
        prt.Print "Proyecto: PASCUALAMA"
        
        SetpYX li3, Margen_Izquierdo
        prt.Print "Proyecto Número: P5SL"
        
        SetpYX li4, Margen_Izquierdo
'        prt.Print "ODC Número: "; m_OdcNum
        
        SetpYX li5, Margen_Izquierdo
'        prt.Print "MK Número: "; m_MKN
        
        SetpYX li6, Margen_Izquierdo
'        prt.Print "MK Descripción: "; m_MKD
        
        SetpYX li7, Margen_Izquierdo
'        prt.Print m_NVN
        
        SetpYX li8, Margen_Izquierdo
        prt.Print "Peso(Kg): "; m_KgU
        
        prt.Font.Name = "code 128"
        prt.Font.Size = m_TamanoCodeBar ' 32 oficial
        
        SetpYX li9, 0.4 'Margen_Izquierdo
'        prt.Print txt2code128(m_MKN)
        
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
Private Sub etiquetaImprimirV1211(li As Integer)

' imprime etiquetas version noviembre 2012
' QR Code (codigo bidimensional)

'Dim li As Integer
Dim nCopias As Integer, Copia As Integer
Dim tab0 As Double, tab1 As Double, mDes As String, QrCodeText As String
Dim m_Nv As String, m_obra As String, m_Plano As String, m_Rev As String, m_Marca As String, m_Peso As Double
Dim mProyecto As String, mTag As String, Imprime As Boolean
mProyecto = ""
mTag = ""

tab0 = 0.5
tab1 = 4.5

Set prt = Printer
Set cQrCode = New ClsQrCode

'For li = 1 To n_filas

'    If n_Copias > 0 Then n_Copias = 1
'    Imprime = IIf(Trim(fg.TextMatrix(li, Columna)) = "SI", True, False) ' SI o NO imprime
    
'    If Imprime Then
'    If nCopias > 0 Then
    
        nCopias = Val(Trim(fg.TextMatrix(li, 8))) ' Cant a imprimir
        nCopias = 1
        
        For Copia = 1 To nCopias

            Prt_Ini
            
            m_Nv = fg.TextMatrix(li, 1)
            m_obra = fg.TextMatrix(li, 2)
            m_Plano = fg.TextMatrix(li, 3)
            m_Rev = fg.TextMatrix(li, 4)
            m_Marca = fg.TextMatrix(li, 5)
            mDes = fg.TextMatrix(li, 6)
            m_Peso = m_CDbl(fg.TextMatrix(li, 7))
            
            '//////////////////////////////////////////////////
            '
            ' CODIGO QR
            '
            ' carga imagen generada desde la web
            ' NombreCliente/Proyecto/Tag/Plano/Descripcion/PesoUnitario/NV
            'QrCodeText = cliNombreFantasia & "/" & mProyecto & "/" & mTag & "/" & m_Plano & "/" & mDes & "/" & m_Peso & "/" & Nv.Text
            
            'QrCodeText = Usuario.descripcion & "/" & m_Plano & "/" & m_Marca & "/" & mDes & "/" & m_Peso & "/" & m_Nv
            QrCodeText = Usuario.Descripcion & "/" & m_Plano & "/" & m_Marca & "/" & "" & "/" & m_Peso & "/" & m_Nv
            
            ' no esta establecida
            PictureQrCode.Picture = cQrCode.GetPictureQrCode(QrCodeText, PictureQrCode.ScaleWidth, PictureQrCode.ScaleHeight)
            If PictureQrCode.Picture Is Nothing Then MsgBox "Error!"
            
            '////////////////////////////////////////////////////
            ' para servidor qrcode.com (que no estaba disponible los dias 05 y 06/09/12
            'Printer.PaintPicture PictureQrCode.Picture, tab0, 1, 3.7, 3.7 ' en cms
            '////////////////////////////////////////////////////
            ' para api by Goolge
            ' sAPI = "http://chart.apis.google.com/chart?cht=qr&chs=" & Width & "x" & Height & "&chl=" & GetSafeURL(Unicode2UTF8(sText)) & "&choe=" & Encoding & "&chld=" & ErrCorrectionLevel
            Printer.PaintPicture PictureQrCode.Picture, -1, -1.8, 6.4, 9.5 ' ajustados segun lo que se imprime
            '////////////////////////////////////////////////////
            
            
            '////////////////////////
            '
            ' LOGO DELGADO
            '
            prt.Font.Name = "delgado"
            prt.Font.Size = 18
            prt.Font.Bold = True
            SetpYX -0.1, tab0
            prt.Print "Delgado"
            '////////////////////////
            '
            ' TEXTOS
            '
            prt.Font.Name = "Arial Black"
            prt.Font.Size = 12 '15
            prt.Font.Bold = False
            
            SetpYX 0.1, 7.1
            Printer.Print Format(Date, "dd/mm/yy")
            
            SetpYX 1, tab1
            prt.Print UCase(Usuario.Descripcion)
            
            SetpYX 1.6, tab1
            prt.Print m_Plano
            
            SetpYX 2.2, tab1
            prt.Print m_Marca
            
            SetpYX 2.8, tab1
            prt.Print mDes
            
            SetpYX 3.4, tab1
            prt.Print m_Peso; " Kgs"
            
            SetpYX 5, tab0
'            prt.Print mProyecto
            prt.Print "NV " & m_Nv & " " & m_obra
            
            ' Get the picture's dimensions in the printer's scale
            ' mode.
            'wid = ScaleX(Picture1.ScaleWidth, Picture1.ScaleMode, Printer.ScaleMode)
            'hgt = ScaleY(Picture1.ScaleHeight, Picture1.ScaleMode, Printer.ScaleMode)
            
            ' Draw the box.
            'Printer.Line (1440, 1440)-Step(wid, hgt), , B
            
            prt.EndDoc

        Next
'    End If
'Next

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

Dim fi As Integer, Dbi As Database, RsRepo As Recordset

PiezasxRecibir 0, Usuario.rut, "", "__/__/__", "__/__/__"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset("SELECT * FROM [piezas x recibir] ORDER BY nv")

fi = 0
With RsRepo
Do While Not .EOF

    m_Can = ![x Recibir]
    
    If m_Can > 0 Then
    
        fi = fi + 1
               
        fg.TextMatrix(fi, 1) = !Nv
        fg.TextMatrix(fi, 2) = !obra
        fg.TextMatrix(fi, 3) = !Plano
        fg.TextMatrix(fi, 4) = !Rev
        fg.TextMatrix(fi, 5) = !Marca
        fg.TextMatrix(fi, 6) = !Descripcion
        fg.TextMatrix(fi, 7) = ![Peso Unitario]
        fg.TextMatrix(fi, 8) = m_Can
    
    End If
    
    .MoveNext
        
Loop

Debug.Print fi

.Close
End With
Dbi.Close

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
    
    MSFlexGrid.TextMatrix(MSFlexGrid.Row, Columna) = "SI"
    
    MSFlexGrid.Row = MSFlexGrid.Row + 1
    
    Edt.visible = False
    MSFlexGrid.SetFocus
    DoEvents
    
End Select

End Sub
