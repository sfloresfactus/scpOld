VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Marca_Mov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos de una Marca"
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
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.TextBox Nv 
         Height          =   300
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton btnImprimir 
         Height          =   400
         Left            =   5880
         Picture         =   "Marca_Mov.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1320
         Width           =   400
      End
      Begin VB.ComboBox ComboMarca 
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Text            =   "ComboMarca"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.ComboBox ComboPlano 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   3735
      End
      Begin VB.CommandButton btnLimpiar 
         Caption         =   "&Limpiar"
         Height          =   375
         Left            =   5400
         TabIndex        =   9
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton btnBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   5400
         TabIndex        =   8
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
         Caption         =   "&Plano"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lbl 
         Caption         =   "&Marca"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   615
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
      Height          =   3135
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5530
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.Label Marca 
      Caption         =   "Marca"
      Height          =   255
      Left            =   7320
      TabIndex        =   12
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Marca_Mov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DbD As Database, RsCl As Recordset
', RsSc As Recordset
'Private SqlRsSc As New ADODB.Recordset
Private Dbm As Database, RsNVc As Recordset, RsNvPla As Recordset, RsPd As Recordset
Private RsOTfd As Recordset, RsITOfd As Recordset
Private RsITOpgd As Recordset
Private RsBulto As Recordset, RsGDc As Recordset, RsGDd As Recordset
Private indice As Integer, n_filas As Integer, n_columnas As Integer
Private m_Fecha As Date, m_Plano As String, m_SubC As String
Private a_Nv(2999, 1) As String, i As Integer, m_Nv As Double, m_NvArea As Integer
Private Sub Form_Load()

Set DbD = OpenDatabase(data_file)
'Set RsSc = DbD.OpenRecordset("Contratistas")
'RsSc.Index = "RUT"

Set RsCl = DbD.OpenRecordset("Clientes")
RsCl.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)

Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set RsNvPla = Dbm.OpenRecordset("Planos Cabecera")
RsNvPla.Index = "NV-Plano"

Set RsPd = Dbm.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Marca"

Set RsOTfd = Dbm.OpenRecordset("OT Fab Detalle")
RsOTfd.Index = "NV-Plano-Marca"

Set RsITOfd = Dbm.OpenRecordset("ITO Fab Detalle")
RsITOfd.Index = "NV-Plano-Marca"

Set RsITOpgd = Dbm.OpenRecordset("ITO pg Detalle")
RsITOpgd.Index = "NV-Plano-Marca"

'Set RsBulto = Dbm.OpenRecordset("bultos")

Set RsGDc = Dbm.OpenRecordset("GD Cabecera")
RsGDc.Index = "Numero"
Set RsGDd = Dbm.OpenRecordset("GD Detalle")
RsGDd.Index = "NV-Plano-Marca"

Nv.MaxLength = 5

' Combo obra
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
IncluirNV:        i = i + 1
        a_Nv(i, 0) = RsNVc!Numero
        a_Nv(i, 1) = RsNVc!obra
        ComboNV.AddItem Format(RsNVc!Numero, "0000") & " - " & RsNVc!obra
    End If
    
    RsNVc.MoveNext
    
Loop

n_filas = 20
n_columnas = 6
Resultado_Config

'btnBuscar.Enabled = False

'btnLimpiar_Click

m_NvArea = 0

ComboMarca.Clear

End Sub
Private Sub Resultado_Config()

Dim i As Integer, ancho As Integer

Detalle.WordWrap = True
Detalle.RowHeight(0) = 450
Detalle.Rows = n_filas + 1
Detalle.Cols = n_columnas

Detalle.TextMatrix(0, 0) = "Doc"
Detalle.TextMatrix(0, 1) = "Número"
Detalle.TextMatrix(0, 2) = "Fecha"
Detalle.TextMatrix(0, 3) = "Piezas"
Detalle.TextMatrix(0, 4) = "Kg ó m2"
Detalle.TextMatrix(0, 5) = "Descripción"

Detalle.ColWidth(0) = 600
Detalle.ColWidth(1) = 2700
Detalle.ColWidth(2) = 900 '800
Detalle.ColWidth(3) = 600
Detalle.ColWidth(4) = 1700 ' 1800
Detalle.ColWidth(5) = 2200 '1600

Detalle.Width = Detalle.ColWidth(0) + Detalle.ColWidth(1) + Detalle.ColWidth(2) + Detalle.ColWidth(3) + Detalle.ColWidth(4) + Detalle.ColWidth(5) + 350
Me.Width = Detalle.Width + 450

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

ComboPlano.Clear

ComboPlano.AddItem " " ' para "borrar" linea

RsNvPla.Seek ">=", m_Nv, ""
If Not RsNvPla.NoMatch Then
    Do While Not RsNvPla.EOF
        If RsNvPla!Nv = m_Nv Then
            ComboPlano.AddItem RsNvPla!Plano
        Else
            Exit Do
        End If
        RsNvPla.MoveNext
    Loop
End If

' limpia combos de marca
ComboMarca.Clear
MousePointer = vbDefault

End Sub
Private Sub ComboPlano_Click()
Dim np As String
np = ComboPlano.Text
ComboMarca.Clear
i = 0
RsPd.Seek ">=", m_Nv, m_NvArea, np, ""
If Not RsPd.NoMatch Then
    Do While Not RsPd.EOF
        If RsPd!Nv <> m_Nv Then Exit Do
        If RsPd!Plano = np Then
            i = i + 1
            ComboMarca.AddItem RsPd!Marca
        Else
            Exit Do
        End If
        RsPd.MoveNext
    Loop
End If

If i = 1 Then

    ' solo hay una marca en el plano
    ' click en combomarca
    ComboMarca.ListIndex = 0
    ' y busca inmediatamente
    btnBuscar_Click
    Detalle.SetFocus
    
End If

End Sub
'Private Sub Marca_KeyPress(KeyAscii As Integer)
'If KeyAscii = vbKeyReturn Then btnBuscar.SetFocus
'End Sub
Private Sub btnBuscar_Click()
' busca marca
' MARCA NO ES ÚNICA EN OBRA
MousePointer = vbHourglass

Limpiar_Detalle
Marca.Caption = ComboMarca.Text

'ComboNV.Enabled = False
'Marca.Enabled = False
'btnBuscar.Enabled = False
btnLimpiar.Enabled = True

m_Plano = ""
Marca.Caption = UCase(Marca.Caption)

indice = 0
' planos detalle
RsPd.Seek ">=", m_Nv, m_NvArea, "", ""
If Not RsPd.NoMatch Then
    Do While Not RsPd.EOF
    
        If RsPd!Nv <> m_Nv Then Exit Do
        If RsPd!Plano = ComboPlano.Text Then
        If RsPd!Marca = Marca.Caption Then
        
'            m_Plano = Trim(RsPd!Plano) 'le puse trim por "DUCTOS " 14/08/98
            m_Plano = RsPd!Plano 'le puse trim por "DUCTOS " 14/08/98
            
            m_Fecha = CDate(0)
            RsNvPla.Seek "=", m_Nv, m_NvArea, m_Plano
            If Not RsNvPla.NoMatch Then
                m_Fecha = RsNvPla![Fecha Modificacion]
            End If
            
            indice = indice + 1
            If indice = n_filas Then
                n_filas = indice
                Detalle.Rows = n_filas + 1
            End If
            Detalle.TextMatrix(indice, 0) = "Plano"
            Detalle.TextMatrix(indice, 1) = m_Plano & ", " & RsPd!Rev
            Detalle.TextMatrix(indice, 2) = m_Fecha
            Detalle.TextMatrix(indice, 3) = RsPd![Cantidad Total]
            Detalle.TextMatrix(indice, 4) = "x " & Format(RsPd![Peso]) & " Kg= " & Format(RsPd![Cantidad Total] * RsPd![Peso]) & " Kg"
            Detalle.Row = indice
            Detalle.col = 5
'            Detalle.CellAlignment = flexAlignLeftCenter
            Detalle.CellForeColor = vbRed
            Detalle.TextMatrix(indice, 5) = RsPd!Descripcion
            
            Buscar_en_Docs RsPd![Peso], RsPd![Superficie]
            
        End If
        End If
        RsPd.MoveNext
    Loop
End If

MousePointer = vbDefault

End Sub
Private Sub Buscar_en_Docs(PesoUnitario As Double, m2Unitario As Double)

Dim m_Total As Integer
Dim t_OTf As Integer, t_ITOf As Integer, t_GDg As Integer, t_ITOp As Integer, t_ITOr As Integer, t_ITOt As Integer, t_GD As Integer
Dim Suma_Cant_Recib As Integer

t_OTf = 0
t_ITOf = 0
t_GDg = 0
t_ITOr = 0
t_ITOt = 0
t_ITOp = 0
t_GD = 0

' OTf
m_SubC = ""
m_Total = 0
RsITOfd.Index = "ot"
With RsOTfd
'.Seek ">=", NVnumero.Caption, m_Plano, Marca.Caption
.Seek "=", m_Nv, m_NvArea, m_Plano, Marca.Caption
If Not .NoMatch Then
    Do While Not .EOF
        If !Nv <> m_Nv Then Exit Do
        If m_Plano = !Plano And !Marca = Marca.Caption Then
        
            indice = indice + 1
            Fila_Crear indice
            Detalle.TextMatrix(indice, 0) = "OTf"
            Detalle.TextMatrix(indice, 1) = !Numero
            Detalle.TextMatrix(indice, 2) = !Fecha
            Detalle.TextMatrix(indice, 3) = !Cantidad
            Detalle.TextMatrix(indice, 4) = !Cantidad * PesoUnitario & " Kg"
            
            m_SubC = ""
'            RsSc.Seek "=", ![RUT Contratista]
'            If Not RsSc.NoMatch Then
'                m_SubC = RsSc![Razon Social]
'            End If
            m_SubC = Contratista_Lee(SqlRsSc, ![Rut contratista])
            Detalle.TextMatrix(indice, 5) = m_SubC
            
            ' actualiza cantidad recibida
            Suma_Cant_Recib = 0
            RsITOfd.Seek "=", !Numero
            If Not RsITOfd.NoMatch Then
               Do While Not RsITOfd.EOF
                  If !Numero <> RsITOfd![Numero OT] Then Exit Do
                  If !Plano = RsITOfd![Plano] Then
                  If !Marca = RsITOfd![Marca] Then
                     Suma_Cant_Recib = Suma_Cant_Recib + RsITOfd!Cantidad
                  End If
                  End If
                  RsITOfd.MoveNext
               Loop
            End If
            If Suma_Cant_Recib >= 0 Then
               .Edit
               ![Cantidad Recibida] = Suma_Cant_Recib
               .Update
            End If
 
            m_Total = m_Total + !Cantidad
            
        End If
        RsOTfd.MoveNext
    Loop
    
    RsITOfd.Index = "nv-plano-marca"

    If m_Total <> Detalle.TextMatrix(indice, 3) Then
        indice = indice + 1
        Fila_Crear indice
        Detalle.TextMatrix(indice, 0) = "OTf"
        Detalle.TextMatrix(indice, 1) = "Total"
        Detalle.TextMatrix(indice, 3) = m_Total
    End If
End If
End With
t_OTf = m_Total

' ITOf
m_Total = 0


With RsITOfd
.Index = "nv-plano-marca"
'.Seek ">=", NVnumero.Caption, m_Plano, Marca.Caption
'.Seek "=", m_Nv, m_NvArea, m_Plano, Marca.Caption
.Seek "=", m_Nv, m_NvArea, m_Plano, Marca.Caption

If Not .NoMatch Then

    Do While Not .EOF
    
        If !Nv <> m_Nv Then Exit Do
        
        If m_Plano = !Plano And !Marca = Marca.Caption Then
        
            indice = indice + 1
            
            Fila_Crear indice
            Detalle.TextMatrix(indice, 0) = "ITOf"
            Detalle.TextMatrix(indice, 1) = !Numero
            Detalle.TextMatrix(indice, 2) = !Fecha
            Detalle.TextMatrix(indice, 3) = !Cantidad
            Detalle.TextMatrix(indice, 4) = !Cantidad * PesoUnitario & " Kg"
'            Detalle.TextMatrix(indice, 4) = ![numero ot]
            
            m_SubC = ""
'            RsSc.Seek "=", ![RUT Contratista]
'            If Not RsSc.NoMatch Then
'                m_SubC = RsSc![Razon Social]
'            End If
            m_SubC = Contratista_Lee(SqlRsSc, ![Rut contratista])
            Detalle.TextMatrix(indice, 5) = m_SubC
            
            m_Total = m_Total + !Cantidad
            
        End If
        .MoveNext
    Loop
    
    If m_Total <> Detalle.TextMatrix(indice, 3) Then
        indice = indice + 1
        Fila_Crear indice
        Detalle.TextMatrix(indice, 0) = "ITOf"
        Detalle.TextMatrix(indice, 1) = "Total"
        Detalle.TextMatrix(indice, 3) = m_Total
    End If

End If
End With
t_ITOf = m_Total

GDg:
' GD, galvanizado
m_Total = 0
With RsGDd
'.Seek ">=", NVnumero.Caption, m_Plano, Marca.Caption
.Seek "=", m_Nv, m_NvArea, m_Plano, Marca.Caption
If Not .NoMatch Then
    Do While Not .EOF
        If !Nv <> m_Nv Then Exit Do
        If m_Plano = !Plano And !Marca = Marca.Caption Then
        
            RsGDc.Seek "=", !Numero
            
            If Not RsGDc.NoMatch Then
            
                If RsGDc!Tipo = "G" Then
                
                    indice = indice + 1
                    Fila_Crear indice
                    Detalle.TextMatrix(indice, 0) = "GDg"
                    Detalle.TextMatrix(indice, 1) = !Numero
                    Detalle.TextMatrix(indice, 2) = !Fecha
                    Detalle.TextMatrix(indice, 3) = !Cantidad
                    Detalle.TextMatrix(indice, 4) = !Cantidad * PesoUnitario & " Kg"
                    
                    m_SubC = ""
                    RsCl.Seek "=", ![RUT CLiente]
                    m_SubC = ""
                    If Not RsCl.NoMatch Then
                        m_SubC = RsCl![Razon Social]
                    End If
                    Detalle.TextMatrix(indice, 5) = m_SubC
        
                    m_Total = m_Total + !Cantidad
                    
                End If
                
            End If
            
        End If
        .MoveNext
    Loop
    
    If m_Total > 0 Then
    If m_Total <> Detalle.TextMatrix(indice, 3) Then
        indice = indice + 1
        Fila_Crear indice
        Detalle.TextMatrix(indice, 0) = "GDg"
        Detalle.TextMatrix(indice, 1) = "Total"
        Detalle.TextMatrix(indice, 3) = m_Total
    End If
    End If

End If
End With
t_GDg = m_Total

ITOpg:
' ITO pintura, galvanizado, granallado y produccion pintura
Dim td(1, 4) As String

'td(0, 1) = "P": td(1, 1) = "ITO Pin"
'td(0, 2) = "G": td(1, 2) = "ITO Gal"
'td(0, 3) = "R": td(1, 3) = "ITO Gra"
'td(0, 4) = "T": td(1, 4) = "ITO PPin"

'td(0, 1) = "P": td(1, 1) = "ItoP"
'td(0, 2) = "G": td(1, 2) = "ItoGa"
'td(0, 3) = "R": td(1, 3) = "ItoGr"
'td(0, 4) = "T": td(1, 4) = "ItoPP"

td(0, 1) = "G": td(1, 1) = "ItoGa"
td(0, 2) = "R": td(1, 2) = "ItoGr"
td(0, 3) = "T": td(1, 3) = "ItoPP"
td(0, 4) = "P": td(1, 4) = "ItoP"

With RsITOpgd
'For i = 1 To 4
For i = 1 To 3 ' modif 10/06/2017
    m_Total = 0
    .Seek "=", m_Nv, m_NvArea, m_Plano, Marca.Caption
    
    If Not .NoMatch Then
    
        Do While Not .EOF
        
            If !Nv <> m_Nv Then Exit Do
            
            If td(0, i) = !Tipo Then
            
                If m_Plano = !Plano And !Marca = Marca.Caption Then
                
                    indice = indice + 1
                    Fila_Crear indice
                    
                    Detalle.TextMatrix(indice, 0) = td(1, i)
                    
                    Detalle.TextMatrix(indice, 1) = !Numero
                    Detalle.TextMatrix(indice, 2) = !Fecha
                    Detalle.TextMatrix(indice, 3) = !Cantidad
                    Detalle.TextMatrix(indice, 4) = !Cantidad * m2Unitario & " m2"
        '            RsSc.Seek "=", ![RUT SubContratista]
                    m_SubC = ""
'                    RsSc.Seek "=", ![RUT Contratista]
'                    If Not RsSc.NoMatch Then
'                        m_SubC = RsSc![Razon Social]
'                    End If
                    m_SubC = Contratista_Lee(SqlRsSc, ![Rut contratista])
                    Detalle.TextMatrix(indice, 5) = m_SubC
                    m_Total = m_Total + !Cantidad
        
                End If
                
             End If
             
            .MoveNext
            
        Loop
        
        If m_Total > 0 Then
        
            If m_Total <> Detalle.TextMatrix(indice, 3) Then
            
                indice = indice + 1
                Fila_Crear indice
                
                Detalle.TextMatrix(indice, 0) = td(1, i)
                
                Detalle.TextMatrix(indice, 1) = "Total"
                Detalle.TextMatrix(indice, 3) = m_Total
                
            End If
            
        End If
    
    End If
    
    Select Case i
    Case 1
        t_ITOp = t_ITOp + m_Total ' porque itop e itog son "la misma"
    Case 2
        t_ITOr = m_Total
    Case 3
        t_ITOt = m_Total
    Case 4
        t_ITOp = m_Total
    End Select
    
Next
End With

Bultos:
' bultos
m_Total = 0
Set RsBulto = Dbm.OpenRecordset("SELECT * FROM bultos WHERE nv=" & m_Nv & " AND plano='" & m_Plano & "' AND marca='" & Marca.Caption & "'")
With RsBulto

Do While Not .EOF
        
    indice = indice + 1
    Fila_Crear indice
    Detalle.TextMatrix(indice, 0) = "Bulto"
    Detalle.TextMatrix(indice, 1) = !Numero
    Detalle.TextMatrix(indice, 2) = !Fecha
    Detalle.TextMatrix(indice, 3) = !Cantidad
    Detalle.TextMatrix(indice, 4) = !Cantidad * PesoUnitario & " Kg"
    

    m_Total = m_Total + !Cantidad
    .MoveNext
Loop
    
If m_Total > 0 Then
    If m_Total <> Detalle.TextMatrix(indice, 3) Then
        indice = indice + 1
        Fila_Crear indice
        Detalle.TextMatrix(indice, 0) = "Bultos"
        Detalle.TextMatrix(indice, 1) = "Total"
        Detalle.TextMatrix(indice, 3) = m_Total
    End If
End If

End With
't_GD = m_Total

GD:
' GD, NO galvanizado
m_Total = 0
With RsGDd
'.Seek ">=", NVnumero.Caption, m_Plano, Marca.Caption
.Seek "=", m_Nv, m_NvArea, m_Plano, Marca.Caption
If Not .NoMatch Then
    Do While Not .EOF
        If !Nv <> m_Nv Then Exit Do
        If m_Plano = !Plano And !Marca = Marca.Caption Then
            
            RsGDc.Seek "=", !Numero
            
            If Not RsGDc.NoMatch Then
            
                If RsGDc!Tipo <> "G" Then
            
                    indice = indice + 1
                    Fila_Crear indice
                    Detalle.TextMatrix(indice, 0) = "GD"
                    Detalle.TextMatrix(indice, 1) = !Numero
                    Detalle.TextMatrix(indice, 2) = !Fecha
                    Detalle.TextMatrix(indice, 3) = !Cantidad
                    Detalle.TextMatrix(indice, 4) = !Cantidad * PesoUnitario & " Kg"
                    
                    RsCl.Seek "=", ![RUT CLiente]
                    m_SubC = ""
                    If Not RsCl.NoMatch Then
                        m_SubC = RsCl![Razon Social]
                    End If
                    Detalle.TextMatrix(indice, 5) = m_SubC
        
                    m_Total = m_Total + !Cantidad
                    
                End If
                
            End If
            
        End If
        .MoveNext
    Loop
    
    If m_Total > 0 Then
    If m_Total <> Detalle.TextMatrix(indice, 3) Then
        indice = indice + 1
        Fila_Crear indice
        Detalle.TextMatrix(indice, 0) = "GD"
        Detalle.TextMatrix(indice, 1) = "Total"
        Detalle.TextMatrix(indice, 3) = m_Total
    End If
    End If

End If
End With
t_GD = m_Total

' actualiza cantidades en planos detalle
With RsPd
.Seek "=", m_Nv, m_NvArea, m_Plano, Marca.Caption
If Not .NoMatch Then
    .Edit
    ![OT fab] = t_OTf
    ![ITO fab] = t_ITOf
    ![GD gal] = t_GDg
    ![ITO pyg] = t_ITOp
    ![ito gr] = t_ITOr
    ![ito pp] = t_ITOt
    !GD = t_GD
    .Update
End If
End With

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
ComboPlano.Clear
ComboMarca.Clear
Marca.Caption = ""

Limpiar_Detalle

ComboNV.Enabled = True
ComboPlano.Enabled = True
ComboMarca.Enabled = True
btnBuscar.Enabled = True
btnLimpiar.Enabled = True

End Sub
Private Sub Limpiar_Detalle()
Dim f As Integer, c As Integer

Detalle.SetFocus
'SendKeys "^{HOME}", True

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
Dim tab5 As Integer, tab6 As Integer, prt As Printer, li As Integer, linea As String

linea = String(76, "-")
tab0 = 5 'margen izquierdo
tab1 = tab0 + 0  'doc
tab2 = tab1 + 6  'numero
tab3 = tab2 + 10 'fecha
tab4 = tab3 + 10 'piezas
tab5 = tab4 + 8  'kg ó m2
tab6 = tab5 + 20 'contratista

Impresora_Predeterminada ReadIniValue(Path_Local & "scp.ini", "Printer", "Docs")
'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt
'prt.Font.Name = "Courier New"
prt.Font.Size = 12

For j = 1 To n_Copias

    prt.Font.Size = 12
    prt.Print Tab(tab0); Empresa.Razon
    prt.Print Tab(tab0); " "
    prt.Print Tab(tab0); "MOVIMIENTOS DE MARCA"
    prt.Print Tab(tab0); "FECHA "; Format(Now, Fecha_Format)
    prt.Print Tab(tab0); " "
    prt.Print Tab(tab0); "OBRA  : "; ComboNV.Text
    prt.Print Tab(tab0); "MARCA : "; UCase(ComboMarca.Text)
    prt.Print Tab(tab0); " "

    prt.Font.Bold = True
    prt.Print Tab(tab1); "DOC";
    prt.Print Tab(tab2); "   NÚMERO";
    prt.Print Tab(tab3); " FECHA";
    prt.Print Tab(tab4); "PIEZAS";
    prt.Print Tab(tab5); "           KG ó m2";
    prt.Print Tab(tab6); "CONTRATISTA"
    prt.Font.Bold = False
    
    prt.Print Tab(tab0); linea
    
    For li = 1 To n_filas - 1
        If Trim(Detalle.TextMatrix(li, 0)) <> "" Then
            prt.Print Tab(tab1); Detalle.TextMatrix(li, 0);
            prt.Print Tab(tab2); PadL(Detalle.TextMatrix(li, 1), 9);
            prt.Print Tab(tab3); Format(Detalle.TextMatrix(li, 2), Fecha_Format);
            prt.Print Tab(tab4); PadL(Detalle.TextMatrix(li, 3), 6);
            prt.Print Tab(tab5); PadL(Detalle.TextMatrix(li, 4), 18);
            prt.Print Tab(tab6); Detalle.TextMatrix(li, 5)
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
