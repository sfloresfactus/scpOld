VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ITOfaAnular 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ITO Fabricación a Anular"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnImprimir 
      Height          =   300
      Left            =   120
      Picture         =   "ITOfaAnular.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   300
   End
   Begin MSFlexGridLib.MSFlexGrid Detalle 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5318
      _Version        =   327680
      BackColorBkg    =   12632256
   End
   Begin VB.Label Plano 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lbl 
      Caption         =   "PLANO"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "ITOfaAnular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private i As Integer, j As Integer
Private fin As Integer
Private n_filas As Integer, n_columnas As Integer
Private DbD As Database, RsSc As Recordset, RsPd As Recordset
Private Dbm As Database, RsNVc As Recordset, RsITOc As Recordset, RsITOd As Recordset
Private prt As Printer
Private m_Nv As Double, m_Plano As String, m_Rev As String, n_marcas As Integer, m_NvArea As Integer
'////////////////////////////////////////////////////////////////////
Public Property Let Nv(ByVal New_Value As Double)
m_Nv = New_Value
End Property
Public Property Let PlanoNombre(ByVal New_Value As String)
m_Plano = New_Value
End Property
Public Property Let Rev(ByVal New_Value As String)
m_Rev = New_Value
End Property
Public Property Let NumerodeMarcas(ByVal New_Value As Integer)
n_marcas = New_Value
End Property
'////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
Dim qry As String, m_desc As String, m_sub As String, m_obra As String

n_filas = 100
n_columnas = 8

Detalle_Config

Set DbD = OpenDatabase(data_file)
Set RsSc = DbD.OpenRecordset("Contratistas")
RsSc.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)

Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set RsPd = Dbm.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Marca"

Set RsITOc = Dbm.OpenRecordset("ITO Fab Cabecera")
RsITOc.Index = "Numero"

Set RsITOd = Dbm.OpenRecordset("ITO Fab Detalle")
RsITOd.Index = "NV-Plano-Marca"

Plano.Caption = m_Plano & ", " & m_Rev

j = 0
For i = 1 To n_marcas
    RsITOd.Seek "=", m_Nv, m_NvArea, m_Plano, Marcas(i, 0)
    If Not RsITOd.NoMatch Then
        Do While Not RsITOd.EOF
            If RsITOd!Plano <> m_Plano Or RsITOd!Marca <> Marcas(i, 0) Then Exit Do
            
            RsITOc.Seek "=", RsITOd!Numero
            If Not RsITOc.NoMatch Then
            
                RsPd.Seek "=", m_Nv, m_NvArea, m_Plano, Marcas(i, 0)
                If RsPd.NoMatch Then
                    m_desc = ""
                Else
                    m_desc = RsPd!Descripcion
                End If
                
                RsNVc.Seek "=", RsITOc!Nv, RsITOc!NvArea
                If RsNVc.NoMatch Then
                    m_obra = ""
                Else
                    m_obra = RsNVc!obra
                End If
                
                RsSc.Seek "=", RsITOc![RUT Contratista]
                If RsSc.NoMatch Then
                    m_sub = ""
                Else
                    m_sub = RsSc![Razon Social]
                End If
                
                j = j + 1
                Detalle.TextMatrix(j, 0) = RsITOd!Marca
                Detalle.TextMatrix(j, 1) = m_desc
                Detalle.TextMatrix(j, 2) = RsITOc!Numero
                Detalle.TextMatrix(j, 3) = RsITOc![Fecha]
                Detalle.TextMatrix(j, 4) = m_sub
                Detalle.TextMatrix(j, 5) = RsITOc![Nv]
                Detalle.TextMatrix(j, 6) = m_obra
                Detalle.TextMatrix(j, 7) = RsITOd![Numero OT]
                
            End If
            
            RsITOd.MoveNext
            
        Loop
    End If
Next

Detalle.Rows = j + 1

DbD.Close
Dbm.Close

m_NvArea = 0

End Sub
Private Sub Detalle_Config()
Dim ancho As Integer

Detalle.FixedCols = 0
Detalle.Cols = n_columnas
Detalle.FixedRows = 1
Detalle.Rows = n_filas + 1

Detalle.TextMatrix(0, 0) = "Marca"
Detalle.TextMatrix(0, 1) = "Descripción"
Detalle.TextMatrix(0, 2) = "Nº ITO"
Detalle.TextMatrix(0, 3) = "Fecha"
Detalle.TextMatrix(0, 4) = "Contratista"
Detalle.TextMatrix(0, 5) = "Nº NV"
Detalle.TextMatrix(0, 6) = "Obra"
Detalle.TextMatrix(0, 7) = "OT Nº"

Detalle.ColWidth(0) = 800
Detalle.ColWidth(1) = 1000
Detalle.ColWidth(2) = 900
Detalle.ColWidth(3) = 900
Detalle.ColWidth(4) = 1200
Detalle.ColWidth(5) = 900
Detalle.ColWidth(6) = 1200
Detalle.ColWidth(7) = 900

ancho = 350
For i = 0 To n_columnas - 1
    ancho = ancho + Detalle.ColWidth(i)
Next
Detalle.Width = ancho
ITOfaAnular.Width = ancho + 200

Detalle.Row = 1
Detalle.col = 0

End Sub
Private Sub btnImprimir_Click()
MousePointer = vbHourglass
' imprime OTs a Anular o Modificar
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer, tab4 As Integer
Dim tab5 As Integer, tab6 As Integer, tab7 As Integer, tab8 As Integer, tab9 As Integer
Dim linea As String
linea = String(110, "-")

tab0 = 7 'margen izquierdo
tab1 = tab0 + 0
tab2 = tab1 + 10 'ancho marca
tab3 = tab2 + 12 'desc
tab4 = tab3 + 6  'nºito
tab5 = tab4 + 9  'fecha
tab6 = tab5 + 15 'subc
tab7 = tab6 + 5  'nºnv
tab8 = tab7 + 10 'obra
tab9 = tab8 + 5  'nºot

Dim can_valor As String, can_col As Integer

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt

prt.Print Tab(tab0); "ITO(S) FABRICACIÓN A ANULAR o MODIFICAR"
prt.Print ""
prt.Print ""

' cabecera
prt.Print Tab(tab0); Empresa.Razon
prt.Print Tab(tab0); "GIRO: " & Empresa.Giro
prt.Print Tab(tab0); Empresa.Direccion
prt.Print Tab(tab0); "Teléfono: " & Empresa.Telefono1 & " - " & Empresa.Comuna
prt.Print ""
prt.Print Tab(tab0); "FECHA     : " & Format(Now, Fecha_Format)
prt.Print Tab(tab0); "PLANO     : " & Plano.Caption
prt.Print ""

' detalle
prt.Print Tab(tab1); "MARCA";
prt.Print Tab(tab2); "DESCRIPCIÓN";
prt.Print Tab(tab3); "NºITO";
prt.Print Tab(tab4); "FECHA";
prt.Print Tab(tab5); "CONTRATISTA";
prt.Print Tab(tab6); "NºNV";
prt.Print Tab(tab7); "O B R A";
prt.Print Tab(tab8); "NºOT"

prt.Print Tab(tab1); linea

For i = 1 To Detalle.Rows - 1 'n_filas

    ' MARCA
    prt.Print Tab(tab1); Detalle.TextMatrix(i, 0);
    
    ' DESCRIPCION
    prt.Print Tab(tab2); Detalle.TextMatrix(i, 1);
    
    ' Nº
    prt.Print Tab(tab3); Detalle.TextMatrix(i, 2);
    
    ' FECHA
    prt.Print Tab(tab4); Format(Detalle.TextMatrix(i, 3), Fecha_Format);
    
    ' CONTRATISTA
    prt.Print Tab(tab5); Left(Detalle.TextMatrix(i, 4), 14);
    
    ' Nº NV
    prt.Print Tab(tab6); Detalle.TextMatrix(i, 5);
    
    ' OBRA
    prt.Print Tab(tab7); Left(Detalle.TextMatrix(i, 6), 9);
    
    ' Nº OT
    prt.Print Tab(tab8); Detalle.TextMatrix(i, 7)
        
Next

prt.Print Tab(tab1); linea

prt.EndDoc

Impresora_Predeterminada "default"

MousePointer = vbDefault

End Sub
