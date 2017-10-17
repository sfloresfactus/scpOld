VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Plano_Dig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digitacion de Planos"
   ClientHeight    =   5370
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   9390
   Icon            =   "Plano_Dig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo Plano"
            ImageIndex      =   1
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar Plano"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar Plano"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Plano"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Deshacer"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Grabar Plano"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox txtEditP 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Detalle 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7223
      _Version        =   393216
      BackColorBkg    =   12632256
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   8640
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Plano_Dig.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Plano_Dig.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Plano_Dig.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Plano_Dig.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Plano_Dig.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Plano_Dig.frx":0864
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label porcentaje 
      Alignment       =   1  'Right Justify
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   6
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label SuperficieTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6600
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label PesoTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   4800
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Posicion 
      Caption         =   "Lín 1, Col 1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.Menu MenuPop 
      Caption         =   "MenuPop"
      Visible         =   0   'False
      Begin VB.Menu FilaInsertar 
         Caption         =   "&Insertar Fila"
      End
      Begin VB.Menu FilaEliminar 
         Caption         =   "&Eliminar Fila"
      End
      Begin VB.Menu FilaBorrarContenido 
         Caption         =   "&Borrar Contenido"
      End
   End
End
Attribute VB_Name = "Plano_Dig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Dbm As Database
'Private RsNVc As Recordset 'nota venta cabecera
Private RsPc As Recordset  'plano cabecera
Private RsPd As Recordset  'plano detalle
Private n_filas As Integer, n_columnas As Integer
Private Titulo As String
Private m_PesoTot As Double, m_SuperTot As Double
Private i As Integer, j As Integer
Private RsOTc As Recordset, RsOTd As Recordset
Private RsITOd As Recordset, RsGDd As Recordset
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button, btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button
Private n_marcas As Integer
Private Imprimiendo As Boolean, Imprimir_Abrir As Boolean
Private m_Plano As String, m_Nv As Double, m_obra As String, m_Rev As String, m_Nuevo As Boolean, m_Obs As String, m_RevSN As Boolean
Private m_Path As String, m_Archivo As String
Private largo As Integer, m_NombreHoja As String, m_NvArea As Integer

Private Sub Form_Load()

Inicializa

Set Dbm = OpenDatabase(mpro_file)

'Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
'RsNVc.Index = Nv_Index ' "Número"

Set RsPc = Dbm.OpenRecordset("Planos Cabecera")
RsPc.Index = "NV-Plano"

Set RsPd = Dbm.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Item"

Set RsOTd = Dbm.OpenRecordset("OT Fab Detalle")
RsOTd.Index = "NV-Plano-Marca"

Set RsITOd = Dbm.OpenRecordset("ITO Fab Detalle")
RsITOd.Index = "NV-Plano-Marca"

Set RsGDd = Dbm.OpenRecordset("GD Detalle")
RsGDd.Index = "NV-Plano-Marca"

Imprimiendo = False
Imprimir_Abrir = True

m_NvArea = 0

End Sub
Private Sub Inicializa()
Titulo = "Digitación de Planos"
Plano_Dig.Caption = Titulo

opGrabar False

Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnDesHacer = Toolbar.Buttons(6)
Set btnGrabar = Toolbar.Buttons(7)

If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 0, 0
End If

Variables_Limpiar

n_filas = 130
n_columnas = 9

Detalle_Config
Detalle.Enabled = False

End Sub
Private Sub Variables_Limpiar()
m_Nv = 0
m_obra = ""
m_Plano = ""
m_Rev = ""
m_Nuevo = False
m_RevSN = False
PesoTotal.Caption = "0"
SuperficieTotal.Caption = "0"
porcentaje.Caption = ""
End Sub
Private Sub Botones_Enabled(btn_Agregar As Boolean, btn_Modificar As Boolean, btn_Eliminar, btn_Imprimir As Boolean, btn_DesHacer As Boolean, btn_Grabar As Boolean)

btnAgregar.Enabled = btn_Agregar
btnModificar.Enabled = btn_Modificar
btnEliminar.Enabled = btn_Eliminar
btnImprimir.Enabled = btn_Imprimir
btnDesHacer.Enabled = btn_DesHacer
btnGrabar.Enabled = btn_Grabar

btnAgregar.Value = tbrUnpressed
btnModificar.Value = tbrUnpressed
btnEliminar.Value = tbrUnpressed
btnImprimir.Value = tbrUnpressed
btnDesHacer.Value = tbrUnpressed
btnGrabar.Value = tbrUnpressed

End Sub
Private Sub Detalle_Config()
Dim i As Integer, ancho As Integer

Detalle.Cols = n_columnas + 1
Detalle.Rows = n_filas + 1
Detalle.Left = 100
Detalle.WordWrap = True
Detalle.RowHeight(0) = 450

' solo seleccion filas
'Detalle.SelectionMode = flexSelectionByRow

Detalle.TextMatrix(0, 0) = ""
Detalle.TextMatrix(0, 1) = "Marca"
Detalle.TextMatrix(0, 2) = "Descripción"
Detalle.TextMatrix(0, 3) = "Cant.  Total"
Detalle.TextMatrix(0, 4) = "Cant.   Asignad." '*
Detalle.TextMatrix(0, 5) = "  Kg Unitario"
Detalle.TextMatrix(0, 6) = "  Kg  Total"          '*
Detalle.TextMatrix(0, 7) = "m2 Unitario"
Detalle.TextMatrix(0, 8) = "  m2  Total"          '*
Detalle.TextMatrix(0, 9) = "Observaciones"

Detalle.ColWidth(0) = 350 '250
Detalle.ColWidth(1) = 2000
Detalle.ColWidth(2) = 1200
Detalle.ColWidth(3) = 700
Detalle.ColWidth(4) = 700
Detalle.ColWidth(5) = 800
Detalle.ColWidth(6) = 800
Detalle.ColWidth(7) = 800
Detalle.ColWidth(8) = 800
Detalle.ColWidth(9) = 1200

ancho = 350
For i = 0 To n_columnas
    If i = 6 Then
        PesoTotal.Left = ancho - 200
        PesoTotal.Width = Detalle.ColWidth(i)
    End If
    If i = 8 Then
        SuperficieTotal.Left = ancho - 200
        SuperficieTotal.Width = Detalle.ColWidth(i)
    End If
    ancho = ancho + Detalle.ColWidth(i)
Next

porcentaje.Left = SuperficieTotal.Left + 800

Detalle.Width = ancho
Me.Width = ancho + Detalle.Left * 2

For i = 1 To n_filas
    Detalle.TextMatrix(i, 0) = i
    
    Detalle.Row = i
    ' establece colores a columnas
    Detalle.col = 4
    Detalle.CellForeColor = vbRed
    Detalle.col = 6
    Detalle.CellForeColor = vbRed
    Detalle.col = 8
    Detalle.CellForeColor = vbRed
Next

Detalle.Row = 1
Detalle.col = 1

txtEditP = ""

End Sub
Private Sub opGrabar(Habilitada As Boolean)
Toolbar.Buttons(7).Enabled = Habilitada
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Imprimiendo = False
Select Case Button.Index
Case 1
    nuevo
Case 2
    modificar
Case 3
    Eliminar
Case 4
    Imprimiendo = True
    Imprimir
Case 6
    DesHacer
Case 7
    If Datos_Valida Then Plano_Grabar
End Select
End Sub
Private Sub nuevo()
' PIDE NOMBRE PARA NUEVO PLANO

'PlanoDialog.nuevo RsNVc, RsPc
PlanoDialog.nuevo RsPc

m_Nv = PlanoDialog.Nv
m_obra = PlanoDialog.obra
m_Plano = PlanoDialog.NombrePlano
m_Rev = PlanoDialog.Rev
m_Obs = PlanoDialog.Obs
m_Path = PlanoDialog.Path
m_Archivo = PlanoDialog.Archivo
m_Nuevo = True

If m_Plano <> "" Then

    Titulo_Poner
    
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Detalle.Enabled = True
    Detalle_Limpiar
    
    If m_Archivo <> "" Then
        ' quiere importar plano
        largo = Len(m_Archivo)
        If largo > 4 Then
        
            m_NombreHoja = Left(m_Archivo, Len(m_Archivo) - 4)
            
            Excel_Leer m_Path, m_Archivo, m_NombreHoja
'            Txt_Leer m_Path, m_Archivo
            
            btnGrabar.Enabled = True
            
        Else
            MsgBox "Nombre de Archivo NO Válido" & vbLf & m_Archivo
        End If
        
    End If
    
End If

End Sub
Private Sub Excel_Leer(Path As String, Archivo As String, Hoja As String)
Dim Hay_Datos As Boolean
' importa datos desde planilla excel
Dim Planilla As Object, fi As Integer, co As Integer
Dim m_Can As Integer, m_Mar As String, m_Des As String, m_KgU As Double, m_KgT As Double, m_m2U As Double, m_m2T As Double
Dim s_Paso As String, d_Paso As Double

If Not Archivo_Existe(Path, Archivo) Then
    MsgBox "No existe archivo " & vbLf & Path & Archivo
    Exit Sub
End If

On Error GoTo NoExcel
Set Planilla = GetObject(Path & Archivo, "Excel.Sheet.8") 'assign sheet object as an OLE excel
On Error GoTo 0

With Planilla.Worksheets(Hoja)

Hay_Datos = True
fi = 0

m_PesoTot = 0
m_SuperTot = 0

Do While Hay_Datos
'For fi = 1 To 20 ' fila excel

    fi = fi + 1
    
    m_Can = Val(Trim(.cells(fi, 3).Value))
'    m_Can = Val(Trim(.cells(fi, 2).Value))
    
    If m_Can = 0 Then
        ' fila está vacia
        ' fin de lectura de planilla
'        Hay_Datos = False
        Exit Do
    End If
    
    m_Mar = Trim(.cells(fi, 4).Value)
    m_Des = Trim(.cells(fi, 5).Value)
'    m_KgU = replace(Trim(.cells(fi, 6).Value), ",", ".")

    ' Kg Unitarios
'    d_Paso = Val(Trim(.cells(fi, 6).Value))
    d_Paso = m_CDbl(Trim(.cells(fi, 6).Value))
    If d_Paso = 0 Then
        ' error kg no puede venir con cero o alfanumerico
        MsgBox "Formato de Planilla No Válido" & vbLf & _
        "Columna F de planilla debe venir con Kg Unitario"
        Exit Do
    End If
    m_KgU = Trim(.cells(fi, 6).Value)
    
    ' Kg Totales
'    d_Paso = Val(Trim(.cells(fi, 7).Value))
    d_Paso = m_CDbl(Trim(.cells(fi, 7).Value))
    If d_Paso = 0 Then
        ' error kg no puede venir con cero o alfanumerico
        MsgBox "Formato de Planilla No Válido" & vbLf & _
        "Columna G de planilla debe venir con Kg Totales"
        Exit Do
    End If
    
    m_KgT = Trim(.cells(fi, 7).Value)
    
    m_m2U = Trim(.cells(fi, 8).Value)
    m_m2T = Trim(.cells(fi, 9).Value)
    
'    m_Mar = Trim(.cells(fi, 3).Value)
'    m_Des = Trim(.cells(fi, 4).Value)
'    m_KgU = Trim(.cells(fi, 5).Value)
'    m_KgT = Trim(.cells(fi, 6).Value)
'    m_m2U = Trim(.cells(fi, 7).Value)
'    m_m2T = Trim(.cells(fi, 8).Value)

'    Debug.Print m_can, m_Mar, m_des, m_KgU, m_KgT, m_m2U, m_m2T
    
    '///// traspasa valores leidos de excel a flexgrid
    
    Detalle.TextMatrix(fi, 1) = m_Mar
    Detalle.TextMatrix(fi, 2) = m_Des
    Detalle.TextMatrix(fi, 3) = m_Can
'    Detalle.TextMatrix(fi, 4) = RsPd![OT fab]

    s_Paso = m_KgU
    Detalle.TextMatrix(fi, 5) = Replace(s_Paso, ",", ".")
    
'    Detalle.TextMatrix(fi, 5) = replace(m_KgU, ",", ".")
    Detalle.TextMatrix(fi, 6) = m_KgT
    
    s_Paso = m_m2U
    Detalle.TextMatrix(fi, 7) = Replace(s_Paso, ",", ".")
    
    Detalle.TextMatrix(fi, 8) = m_m2T
'    Detalle.TextMatrix(fi, 9) = NoNulo(RsPd!Observaciones)
    
    m_PesoTot = m_PesoTot + CDbl(Detalle.TextMatrix(fi, 6))
    m_SuperTot = m_SuperTot + CDbl(Detalle.TextMatrix(fi, 8))
    
    ' traspaso de variables de memoria a pantalla (flexgrid?)
    
Loop
'Next
End With

PesoTotal.Caption = m_PesoTot
SuperficieTotal.Caption = m_SuperTot

If m_PesoTot <> 0 Then
    porcentaje.Caption = Format(m_SuperTot * 100 / m_PesoTot, "##0.##") & "%"
End If

Set Planilla = Nothing

Exit Sub

NoExcel:
'MsgBox "No Tiene Instalado Microsoft Excel"
' o archivo esta abierto (en uso por ejemplo por excel)
MsgBox "Nombre de Hoja NO Válido"

End Sub
Private Sub Txt_Leer(Path As String, Archivo As String)
'Dim Hay_Datos As Boolean
' importa datos desde archivo plano
Dim m_Can As Integer, m_Mar As String, m_Des As String, m_KgU As Double, m_KgT As Double, m_m2U As Double, m_m2T As Double
Dim s_Paso As String, d_Paso As Double, Reg As String
Dim largoreg As Integer, m_col As Integer, largocont As Integer
Dim p1 As Integer, fi As Integer, m_Contenido As String

If Not Archivo_Existe(Path, Archivo) Then
    MsgBox "No existe archivo " & vbLf & Path & Archivo
    Exit Sub
End If

m_PesoTot = 0
m_SuperTot = 0

' abre archivo
Open Path & Archivo For Input As #1

fi = 0

Do While Not EOF(1)

    Line Input #1, Reg
    
    largoreg = Len(Reg)
    
    p1 = 1
    m_col = 0
    
    For i = 1 To largoreg
    
'        Debug.Print "|" & Mid(Reg, i, 1) & "|" & Asc(Mid(Reg, i, 1)) & "|"
        
        If Asc(Mid(Reg, i, 1)) = 9 Then
        
            m_col = m_col + 1
            largocont = i - p1
            m_Contenido = Mid(Reg, p1, largocont)
            
            Select Case m_col
            Case 1 ' plano
                'Debug.Print "plano|" & m_Contenido & "|"
            Case 2 ' rev
                'Debug.Print "rev|" & m_Contenido & "|"
            Case 3 ' cantidad
                m_Can = m_Contenido
                'Debug.Print "cant|" & m_Contenido & "|"
            Case 4 ' marca
                m_Mar = m_Contenido
                'Debug.Print "marca|" & m_Contenido & "|"
            Case 5 ' descr
                m_Des = m_Contenido
                'Debug.Print "descr|" & m_Contenido & "|"
            Case 6 ' peso uni
            
                ' Kg Unitarios
                d_Paso = m_CDbl(m_Contenido)
                If d_Paso = 0 Then
                    ' error kg no puede venir con cero o alfanumerico
                    MsgBox "Formato de Planilla No Válido" & vbLf & _
                    "Columna F de planilla debe venir con Kg Unitario"
                    Exit Do
                End If
                
                m_KgU = d_Paso ' Trim(m_Contenido)

                'Debug.Print "pesouni|" & m_Contenido & "|"
                
            Case 7 ' peso tot
'                m_KgT = m_Contenido
                
                ' Kg Totales
                d_Paso = m_CDbl(m_Contenido)
                If d_Paso = 0 Then
                    ' error kg no puede venir con cero o alfanumerico
                    MsgBox "Formato de Planilla No Válido" & vbLf & _
                    "Columna G de planilla debe venir con Kg Totales"
                    Exit Do
                End If
                
                m_KgT = d_Paso 'Trim(m_Contenido)
                
                'Debug.Print "pesotot|" & m_Contenido & "|"
                
            Case 8 ' m2 uni
                m_m2U = m_Contenido
                'Debug.Print "m2uni|" & m_Contenido & "|"
            End Select
                
            p1 = i + 1
            
        End If
    
    Next
    
    m_col = m_col + 1
    largocont = i - p1
    m_Contenido = Mid(Reg, p1, largocont)
    m_m2T = m_Contenido
'    Debug.Print "m2tot|" & m_Contenido & "|"
    
'    m_m2U = Trim(.cells(fi, 8).Value)
'    m_m2T = Trim(.cells(fi, 9).Value)
    
    '///// traspasa valores leidos de excel a flexgrid
    
    fi = fi + 1
    
    Detalle.TextMatrix(fi, 1) = m_Mar
    Detalle.TextMatrix(fi, 2) = m_Des
    Detalle.TextMatrix(fi, 3) = m_Can
'    Detalle.TextMatrix(fi, 4) = RsPd![OT fab]

    s_Paso = m_KgU
    Detalle.TextMatrix(fi, 5) = Replace(s_Paso, ",", ".")
    
'    Detalle.TextMatrix(fi, 5) = replace(m_KgU, ",", ".")
    Detalle.TextMatrix(fi, 6) = m_KgT
    
    s_Paso = m_m2U
    Detalle.TextMatrix(fi, 7) = Replace(s_Paso, ",", ".")
    
    Detalle.TextMatrix(fi, 8) = m_m2T
'    Detalle.TextMatrix(fi, 9) = NoNulo(RsPd!Observaciones)
    
    m_PesoTot = m_PesoTot + CDbl(Detalle.TextMatrix(fi, 6))
    m_SuperTot = m_SuperTot + CDbl(Detalle.TextMatrix(fi, 8))
   
Loop

Close #1

PesoTotal.Caption = m_PesoTot
SuperficieTotal.Caption = m_SuperTot

If m_PesoTot <> 0 Then
    porcentaje.Caption = Format(m_SuperTot * 100 / m_PesoTot, "##0.##") & "%"
End If

End Sub
Private Sub modificar()
' PIDE NOMBRE del PLANO (editable) a modificar

'PlanoDialog.Abrir RsNVc, RsPc, "Modificar"
PlanoDialog.Abrir RsPc, "Modificar"
m_Nv = PlanoDialog.Nv
m_obra = PlanoDialog.obra
m_Plano = PlanoDialog.NombrePlano
m_Rev = PlanoDialog.Rev
m_Obs = PlanoDialog.Obs
m_Path = PlanoDialog.Path
m_Archivo = PlanoDialog.Archivo

'm_Nuevo = True
If m_Plano <> "" Then
    Titulo_Poner
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Detalle.Enabled = True
    
    If m_Archivo = "" Then
    
        ' quiere digitar manualmente
        Plano_Leer ' leer tradicional
        
    Else
    
        ' quiere importar plano desde excel
        largo = Len(m_Archivo)
        If largo > 4 Then
        
            m_NombreHoja = Left(m_Archivo, Len(m_Archivo) - 4)
            
            Excel_Leer m_Path, m_Archivo, m_NombreHoja
'            Txt_Leer m_Path, m_Archivo

            Plano_BuscarOT

            btnGrabar.Enabled = True

        Else
        
            MsgBox "Nombre de Archivo NO Válido" & vbLf & m_Archivo
            
        End If

    End If
    
End If
End Sub
Private Sub Eliminar()

' PIDE NOMBRE del PLANO (editable) a eliminar

PlanoDialog.Abrir RsPc, "Eliminar"
m_Nv = PlanoDialog.Nv
m_obra = PlanoDialog.obra
m_Plano = PlanoDialog.NombrePlano
m_Rev = PlanoDialog.Rev
m_Obs = PlanoDialog.Obs
'm_Nuevo = True
If m_Plano <> "" Then

    Titulo_Poner
    Detalle.Enabled = True
    Plano_Leer
    
    If Plano_Borrable Then
        ' no tiene OTs
    
        If MsgBox("¿ ELIMINA PLANO ?", vbYesNo) = vbYes Then
            Plano_Eliminar
        End If
        
    Else
    
        ' si tiene OTs
        
        RsPd.Seek ">=", m_Nv, 0, m_Plano, ""
        If Not RsPd.NoMatch Then
            Do While Not RsPd.EOF
                If m_Nv <> RsPd!Nv Or m_Plano <> RsPd!Plano Then Exit Do
                Marcas_Agregar n_marcas, RsPd!Marca, RsPd![Cantidad Total], RsPd![Peso]
                RsPd.MoveNext
            Loop
        End If

        OTfaAnular.Nv = m_Nv
        OTfaAnular.PlanoNombre = m_Plano
        OTfaAnular.Rev = m_Rev
        OTfaAnular.NumerodeMarcas = n_marcas
        OTfaAnular.Show 1
        
        ITOfaAnular.Nv = m_Nv
        ITOfaAnular.PlanoNombre = m_Plano
        ITOfaAnular.Rev = m_Rev
        ITOfaAnular.NumerodeMarcas = n_marcas
        ITOfaAnular.Show 1
        
        If MsgBox("¿ ELIMINA PLANO ?", vbYesNo) = vbYes Then
            Plano_Eliminar
        End If
    
    End If
    
    Variables_Limpiar
    Titulo_Poner
    Detalle_Limpiar
    Detalle.Enabled = False

End If

End Sub
Private Function Plano_Borrable() As Boolean
Plano_Borrable = True
RsPd.Seek ">=", m_Nv, 0, m_Plano, 0
If Not RsPd.NoMatch Then
    Do While Not RsPd.EOF
        If RsPd!Nv <> m_Nv Or RsPd!Plano <> m_Plano Then Exit Do
        
        'ojo, sería mejor buscar en OTs
        If RsPd![OT fab] <> 0 Then
            Plano_Borrable = False
            Exit Function
        End If
        
        RsPd.MoveNext
        
    Loop
End If
End Function
Private Sub Plano_Eliminar()
Dim qry As String

' borra de nota venta planos
RsPc.Index = "NV-Plano"
RsPc.Seek "=", m_Nv, m_NvArea, m_Plano
If Not RsPc.NoMatch Then
    RsPc.Delete
End If
'rspc.Index = "Nota-Línea"

' borra plano detalle
qry = "DELETE * FROM [Planos Detalle] WHERE NV=CDbl(" & m_Nv & ")"
qry = qry & " AND Plano='" & m_Plano & "'"
Dbm.Execute qry

End Sub
Private Sub Imprimir()
' PIDE NOMBRE del PLANO (Editable) a imprimir
If Imprimir_Abrir Then

    PlanoDialog.Abrir RsPc, "Imprimir"
    m_Nv = PlanoDialog.Nv
    m_obra = PlanoDialog.obra
    m_Plano = PlanoDialog.NombrePlano
    m_Rev = PlanoDialog.Rev
    m_Obs = PlanoDialog.Obs
'    m_Nuevo = True
    If m_Plano <> "" Then
    
        Titulo_Poner
        Botones_Enabled 0, 0, 0, 1, 1, 0
        Detalle.Enabled = True
        Plano_Leer
        Imprimir_Abrir = False
        
    Else
        Imprimir_Abrir = True
    End If
    
Else

    If MsgBox("¿ IMPRIME PLANO ?", vbYesNo) = vbYes Then
        MousePointer = vbHourglass
        Plano_Imprimir
        MousePointer = vbDefault
    End If
    
    Variables_Limpiar
    Titulo_Poner
    Detalle_Limpiar
    Detalle.Enabled = False
    
    Imprimir_Abrir = True
    
End If
End Sub
Private Sub DesHacer()
If btnGrabar.Enabled = True Then
    If MsgBox("¿ ABANDONA PLANO SIN GRABAR CAMBIOS ?", vbYesNo) = vbYes Then
        DesHacer1
    End If
Else
    DesHacer1
End If
Imprimir_Abrir = True
End Sub
Private Sub DesHacer1()
Variables_Limpiar
Titulo_Poner
Detalle_Limpiar
Detalle.Enabled = False
Botones_Enabled 1, 1, 1, 1, 0, 0
End Sub
Private Sub Titulo_Poner()
If m_obra = "" Then
    Plano_Dig.Caption = Titulo
Else
    Plano_Dig.Caption = Titulo & " [ " & StrConv(m_obra, vbProperCase) & "," & m_Plano & "," & m_Rev & " ]"
End If
End Sub
Private Sub Actualiza()
Dim fi As Integer, co As Integer, num As Double
fi = Detalle.Row
co = Detalle.col
If co = 3 Or co = 5 Or co = 7 Then
    Detalle.TextMatrix(fi, 6) = Format(Val(Detalle.TextMatrix(fi, 3)) * Val(Detalle.TextMatrix(fi, 5)), num_Formato)
    Detalle.TextMatrix(fi, 8) = Format(Val(Detalle.TextMatrix(fi, 3)) * Val(Detalle.TextMatrix(fi, 7)), num_Formato)
End If

m_PesoTot = 0
m_SuperTot = 0
For i = 1 To n_filas
    m_PesoTot = m_PesoTot + m_CDbl(Detalle.TextMatrix(i, 6))
    m_SuperTot = m_SuperTot + m_CDbl(Detalle.TextMatrix(i, 8))
Next

PesoTotal.Caption = Format(m_PesoTot, num_Formato)
SuperficieTotal.Caption = Format(m_SuperTot, num_Formato)

End Sub
Private Function Datos_Valida()
Datos_Valida = False
For i = 1 To n_filas

    'MARCA
    If Not LargoString_Valida(Detalle.TextMatrix(i, 1), 25, i, 1) Then Exit Function
    
    If Trim(Detalle.TextMatrix(i, 1)) <> "" Then
        If Trim(Detalle.TextMatrix(i, 2)) = "" Then
            Beep
            MsgBox "Descripción Obligatoria"
            Detalle.Row = i
            Detalle.col = 2
            Detalle.SetFocus
            Exit Function
        End If
    End If
    
    'DESCRIPCION
    If Not LargoString_Valida(Detalle.TextMatrix(i, 2), 30, i, 2) Then Exit Function
    
    'CANTIDAD
    If Not Numero_Valida(Detalle.TextMatrix(i, 3), i, 3) Then Exit Function
    
    'KILOS UNITARIOS
    If Not Numero_Valida(Detalle.TextMatrix(i, 5), i, 5) Then Exit Function
    
    'Superficie
    If Not Numero_Valida(Detalle.TextMatrix(i, 7), i, 7) Then Exit Function
    
    'OBSERVACION
    If Not LargoString_Valida(Detalle.TextMatrix(i, 9), 30, i, 9) Then Exit Function
    
Next
Datos_Valida = True
End Function
Private Function LargoString_Valida(txt As String, max As Integer, fil As Integer, col As Integer) As Boolean
If Len(Trim(txt)) > max Then
    LargoString_Valida = False
    Beep
    MsgBox "Largo Máximo es " & max & " caracteres"
    Detalle.Row = fil
    Detalle.col = col
    Detalle.SetFocus
Else
    LargoString_Valida = True
End If
End Function
Private Function Numero_Valida(txt As String, fil As Integer, col As Integer) As Boolean
Dim num As String
Numero_Valida = False
num = txt
If Not IsNumeric(num) Then
    If num <> "" Then
        GoTo Sigue
    End If
Else
    If Val(num) < 0 Then
Sigue:
        Beep
        MsgBox "Número no Válido"
        Detalle.Row = fil
        Detalle.col = col
        Detalle.SetFocus
        Exit Function
    End If
End If
Numero_Valida = True
End Function

Private Sub Plano_Grabar()

If m_Nuevo Then
    ' graba plano nuevo
    PlanoNuevo_Grabar
Else
'    If Plano_RevSN Then
        ' graba nueva revision
'        PlanoRevision_Grabar
'    Else
        ' graba plano editable (ya esta en archivo)
        PlanoModificar_Grabar
'    End If
End If

End Sub
Private Sub PlanoNuevo_Grabar()
Dim linea As Integer

save:

' graba cabecera
RsPc.AddNew
RsPc!Nv = m_Nv
'rspc!Línea = linea + 1
RsPc!Editable = True
RsPc!Plano = UCase(m_Plano)
RsPc!Rev = UCase(m_Rev)
RsPc![Peso Total] = PesoTotal.Caption
RsPc![Superficie Total] = SuperficieTotal.Caption
RsPc![Fecha Modificacion] = Format(Now, Fecha_Format)
RsPc!Observacion = m_Obs

RsPc.Update

Plano_Detalle_Grabar False

Botones_Enabled 1, 1, 1, 1, 0, 0
Variables_Limpiar
Titulo_Poner
Detalle_Limpiar
Detalle.Enabled = False
opGrabar False

End Sub
Private Sub PlanoModificar_Grabar()

' igual consulta si plano es "editable"

If Plano_Borrable Then
    ' no tiene OTs
    Plano_ModGrabar
    
Else
    ' si tiene OTs
    Plano_RevisionGrabar

End If

Botones_Enabled 1, 1, 1, 1, 0, 0
Variables_Limpiar
Titulo_Poner
Detalle_Limpiar
Detalle.Enabled = False
opGrabar False

End Sub
Private Sub Plano_ModGrabar()

' cabecera
RsPc.Seek "=", m_Nv, m_NvArea, m_Plano
If Not RsPc.NoMatch Then
    RsPc.Edit
    RsPc![Peso Total] = PesoTotal.Caption
    RsPc![Superficie Total] = SuperficieTotal.Caption
    RsPc!Rev = UCase(m_Rev)
    RsPc!Observacion = m_Obs
    RsPc![Fecha Modificacion] = Format(Now, Fecha_Format)
    RsPc.Update
End If

' borra y graba detalle
Plano_Detalle_Grabar True

End Sub
Private Sub Plano_RevisionGrabar()

Marcas_Contar

' informe de OTs a anular
OTfaAnular.Nv = m_Nv
OTfaAnular.PlanoNombre = m_Plano
OTfaAnular.Rev = m_Rev
OTfaAnular.NumerodeMarcas = n_marcas
OTfaAnular.Show 1

ITOfaAnular.Nv = m_Nv
ITOfaAnular.PlanoNombre = m_Plano
ITOfaAnular.Rev = m_Rev
ITOfaAnular.NumerodeMarcas = n_marcas
ITOfaAnular.Show 1

' cabecera
RsPc.Seek "=", m_Nv, m_NvArea, m_Plano
If Not RsPc.NoMatch Then
    RsPc.Edit
    RsPc![Peso Total] = PesoTotal.Caption
    RsPc![Superficie Total] = SuperficieTotal.Caption
    RsPc!Rev = UCase(m_Rev)
    RsPc!Observacion = m_Obs
    RsPc![Fecha Modificacion] = Format(Now, Fecha_Format)
    RsPc.Update
End If

' borra y graba detalle
Plano_Detalle_Grabar True

Lineas_Actualiza

End Sub
Private Sub Marcas_Contar()
' cuenta marcas y llena arreglo de marcas
Dim m_Marca As String

RsPd.Index = "NV-Plano-Marca"

n_marcas = 0
' recorre detalle
For i = 1 To n_filas

    m_Marca = Detalle.TextMatrix(i, 1)

    If m_Marca <> "" Then
    
        Marcas_Modificadas m_Marca, Detalle.TextMatrix(i, 3), m_CDbl(Detalle.TextMatrix(i, 5))
        
    End If
Next

Marcas_Eliminadas

RsPd.Index = "NV-Plano-Item"

End Sub
Private Sub Marcas_Modificadas(Marca As String, Cantidad As Integer, Peso As Double)
' busca y "graba" marcas modificadas
' marca,cantidad y peso son del plano nuevo

' busca marca(linea) en plano antiguo
RsPd.Seek "=", m_Nv, m_NvArea, m_Plano, Marca
If RsPd.NoMatch Then

    ' marca nueva
    ' no hay drama con OTs ni ITOs
    
Else

    ' marca ya existe
    RsPd.Edit
    RsPd!Chequeada = True
    RsPd.Update
    
    ' revisa variacion de peso
    If Peso = RsPd![Peso] Then
        ' no hay drama con OTs ni ITOs (con los kilos)
    Else
        ' marca con peso modificado
        Marcas_Agregar n_marcas, Marca, Cantidad, Peso
        Exit Sub
    End If
    
    ' revisa variacion de cantidad
    If Cantidad - RsPd![Cantidad Total] >= 0 Then
        ' no hay drama con OTs
    Else
        ' marca disminuyó cantidad
        Marcas_Agregar n_marcas, Marca, Cantidad, Peso
    End If
    
End If

End Sub
Private Sub Marcas_Eliminadas()
' busca marcas eliminadas
RsPd.Seek ">=", m_Nv, m_NvArea, m_Plano, ""
If Not RsPd.NoMatch Then
    Do While Not RsPd.EOF
        If m_Nv <> RsPd!Nv Or m_Plano <> RsPd!Plano Then Exit Do
        If Not RsPd!Chequeada Then
            Marcas_Agregar n_marcas, RsPd!Marca, 0, 0
        End If
        RsPd.MoveNext
    Loop
End If
End Sub
Public Sub Marcas_Agregar(n_marcas As Integer, Marca As String, Cantidad As Integer, Peso As Double)
' agrega marcas al arreglo de marcas
Dim i As Integer
For i = 0 To n_marcas
    If Marcas(i, 0) = Marca Then
        Exit Sub
    End If
Next
n_marcas = n_marcas + 1
Marcas(n_marcas, 0) = Marca
Marcas(n_marcas, 1) = Cantidad
Marcas(n_marcas, 2) = Peso
End Sub
Private Sub Lineas_Actualiza()
Dim m_Marca As String, Can_OT As Integer, Can_ITOf As Integer, Can_ITOp As Integer, Can_GD As Integer
'actualiza los campos: cantidades y peso
'OT fab
'Cantidad ITO Fabricacion
'ITO pin
'GD
' y Rev en Docs 03/06/98

RsPd.Seek "=", m_Nv, m_NvArea, m_Plano, 1
If Not RsPd.NoMatch Then
    Do While Not RsPd.EOF
        If RsPd!Plano <> m_Plano Then Exit Do
        
            m_Marca = RsPd!Marca
            Can_OT = 0
            Can_ITOf = 0
            Can_ITOp = 0
            Can_GD = 0
            
            ' OT
            RsOTd.Seek "=", m_Nv, m_NvArea, m_Plano, m_Marca
            If Not RsOTd.NoMatch Then
                Do While Not RsOTd.EOF
                    If RsOTd!Plano <> m_Plano Or RsOTd!Marca <> m_Marca Then Exit Do
                        RsOTd.Edit
                        RsOTd!Rev = m_Rev
                        RsOTd![Peso Unitario] = RsPd![Peso]
                        RsOTd.Update
                        Can_OT = Can_OT + RsOTd!Cantidad
                    RsOTd.MoveNext
                Loop
            End If
            
            ' ITO
            RsITOd.Seek "=", m_Nv, m_NvArea, m_Plano, m_Marca
            If Not RsITOd.NoMatch Then
                Do While Not RsITOd.EOF
                    If RsITOd!Plano <> m_Plano Or RsITOd!Marca <> m_Marca Then Exit Do
                        RsITOd.Edit
                        RsITOd!Rev = m_Rev
                        RsITOd![Peso Unitario] = RsPd![Peso]
                        RsITOd.Update
                        Can_ITOf = Can_ITOf + RsITOd!Cantidad
                    RsITOd.MoveNext
                Loop
            End If
            
            ' GD
            RsGDd.Seek "=", m_Nv, m_NvArea, m_Plano, m_Marca
            If Not RsGDd.NoMatch Then
                Do While Not RsGDd.EOF
                    If RsGDd!Plano <> m_Plano Or RsGDd!Marca <> m_Marca Then Exit Do
                        RsGDd.Edit
                        RsGDd!Rev = m_Rev
                        RsGDd![Peso Unitario] = RsPd![Peso]
                        RsGDd.Update
                        Can_GD = Can_GD + RsGDd!Cantidad
                    RsGDd.MoveNext
                Loop
            End If
    
            RsPd.Edit
            RsPd![OT fab] = Can_OT
            RsPd![ITO fab] = Can_ITOf
            RsPd![ITO pyg] = Can_ITOp
            RsPd![GD] = Can_GD
            RsPd.Update
    
        RsPd.MoveNext
    Loop
End If
End Sub
Private Sub Plano_Detalle_Grabar(Borra As Boolean)

If Borra Then
    ' borra detalle
    Dim qry  As String
    qry = "DELETE * FROM [Planos Detalle] WHERE NV=Cdbl(" & m_Nv & ")"
    qry = qry & " AND Plano='" & m_Plano & "'"
    Dbm.Execute qry
End If

' graba detalle
j = 0
For i = 1 To n_filas
    ' verifica marca y cantidad
    If Detalle.TextMatrix(i, 1) <> "" And Val(Detalle.TextMatrix(i, 3)) <> 0 Then
        RsPd.AddNew
        RsPd!Nv = m_Nv
        RsPd!Plano = UCase(m_Plano)
        RsPd!Rev = Left(UCase(m_Rev), 10)
        j = j + 1
        RsPd!item = j
        RsPd!Marca = UCase(Detalle.TextMatrix(i, 1))
        RsPd!Descripcion = Detalle.TextMatrix(i, 2)

        RsPd![Cantidad Total] = Detalle.TextMatrix(i, 3)
        
        RsPd![Peso] = Val(Detalle.TextMatrix(i, 5))
        RsPd![Superficie] = Val(Detalle.TextMatrix(i, 7))
        RsPd![Observaciones] = Detalle.TextMatrix(i, 9)
        RsPd![OT fab] = 0
        RsPd![ITO fab] = 0
        RsPd![ITO pyg] = 0
        RsPd![GD] = 0
        RsPd!Chequeada = False
        RsPd.Update
    End If
Next

End Sub
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'
' RUTINAS PARA EL FLEXGRID
'
'
Private Sub Detalle_DblClick()
' simula un espacio
If Imprimiendo Then Exit Sub
MSFlexGridEdit Detalle, txtEditP, 32
End Sub
Private Sub Detalle_GotFocus()
'If Imprimiendo Then Exit Sub
If txtEditP.visible = False Then Exit Sub
Detalle = txtEditP
txtEditP.visible = False
End Sub
Private Sub Detalle_LeaveCell()
If txtEditP.visible = False Then Exit Sub
Detalle = txtEditP
txtEditP.visible = False
End Sub
Private Sub Detalle_KeyPress(KeyAscii As Integer)
If Imprimiendo Then Exit Sub
MSFlexGridEdit Detalle, txtEditP, KeyAscii
End Sub
Private Sub txtEditP_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCodeP Detalle, txtEditP, KeyCode, Shift
End Sub
Sub EditKeyCodeP(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape ' Esc
    Edt.visible = False
    MSFlexGrid.SetFocus
Case vbKeyReturn ' Enter
    MSFlexGrid.SetFocus
    DoEvents
    Actualiza
    Cursor_Mueve MSFlexGrid
Case 38 ' Flecha Arriba
    MSFlexGrid.SetFocus
    DoEvents
    Actualiza
    If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
        MSFlexGrid.Row = MSFlexGrid.Row - 1
    End If
Case 40 ' Flecha Abajo
    MSFlexGrid.SetFocus
    DoEvents
    Actualiza
    If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
        MSFlexGrid.Row = MSFlexGrid.Row + 1
    End If
End Select
End Sub
Private Sub txtEditP_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)


Select Case MSFlexGrid.col
Case 4, 6, 8 ' cant OT, total kg , total m2
    Exit Sub
Case 1
    Edt.MaxLength = 25
End Select

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
opGrabar True
End Sub
Private Sub Detalle_KeyDown(KeyCode As Integer, Shift As Integer)
If Imprimiendo Then Exit Sub
If KeyCode = vbKeyF2 Then
    MSFlexGridEdit Detalle, txtEditP, 32
End If
End Sub
Private Sub Detalle_RowColChange()
'MIA
Posicion = "Lín " & Detalle.Row & ", Col " & Detalle.col
End Sub
Private Sub Cursor_Mueve(MSFlexGrid As Control)
'MIA
Select Case MSFlexGrid.col
Case 3
    MSFlexGrid.col = 5
Case 5
    MSFlexGrid.col = 7
Case 7
    MSFlexGrid.col = 9
Case 9
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
Private Sub Plano_Leer()

Detalle_Limpiar

m_PesoTot = 0
m_SuperTot = 0
For i = 1 To n_filas
    RsPd.Seek "=", m_Nv, m_NvArea, m_Plano, i
    If Not RsPd.NoMatch Then
        Detalle.TextMatrix(i, 1) = NoNulo(RsPd!Marca)
        Detalle.TextMatrix(i, 2) = NoNulo(RsPd!Descripcion)
        Detalle.TextMatrix(i, 3) = RsPd![Cantidad Total]
        Detalle.TextMatrix(i, 4) = RsPd![OT fab]
        Detalle.TextMatrix(i, 5) = Replace(RsPd!Peso, ",", ".")
        Detalle.TextMatrix(i, 6) = Format(RsPd![Cantidad Total] * RsPd!Peso, num_Formato)
'        Detalle.TextMatrix(i, 7) = RsPd![Superficie]
        Detalle.TextMatrix(i, 7) = Replace(RsPd![Superficie], ",", ".")
        Detalle.TextMatrix(i, 8) = RsPd![Cantidad Total] * RsPd![Superficie]
        Detalle.TextMatrix(i, 9) = NoNulo(RsPd!Observaciones)
        m_PesoTot = m_PesoTot + CDbl(Detalle.TextMatrix(i, 6))
        m_SuperTot = m_SuperTot + CDbl(Detalle.TextMatrix(i, 8))
    End If
Next

PesoTotal.Caption = m_PesoTot
SuperficieTotal.Caption = m_SuperTot

Detalle.Enabled = True
opGrabar (False)
End Sub
Private Sub Plano_BuscarOT()
Dim m_Marca As String

RsPd.Index = "Nv-Plano-Marca"

m_PesoTot = 0
m_SuperTot = 0

For i = 1 To n_filas

    m_Marca = Detalle.TextMatrix(i, 1)
    
    If m_Marca = "" Then Exit For
    
    RsPd.Seek "=", m_Nv, m_Plano, m_Marca
    If Not RsPd.NoMatch Then
    
        Detalle.TextMatrix(i, 4) = RsPd![OT fab]
'        m_PesoTot = m_PesoTot + CDbl(Detalle.TextMatrix(i, 6))
'        m_SuperTot = m_SuperTot + CDbl(Detalle.TextMatrix(i, 8))
        
    End If
Next

'PesoTotal.Caption = m_PesoTot
'SuperficieTotal.Caption = m_SuperTot

Detalle.Enabled = True
opGrabar (False)

RsPd.Index = "Nv-Plano-Item"

End Sub
Private Sub Detalle_Limpiar()
Dim j As Integer
For i = 1 To n_filas
    For j = 1 To n_columnas
        Detalle.TextMatrix(i, j) = ""
    Next
Next
'Detalle.Row = 1
'Detalle.col = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim Rs As Recordset
If btnGrabar.Enabled = True Then
    If MsgBox("¿ ABANDONA PLANO SIN GRABAR CAMBIOS ?", vbYesNo) = vbNo Then
        ' NO sale
        Cancel = True
    End If
    ' abandona
Else
    ' sale sin problemas, pues no modificó
    DataBases_Cerrar
End If
End Sub
Private Sub Detalle_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Imprimiendo Then Exit Sub
If Button = 2 Then
'    MsgBox Detalle.ColSel & vbCr & Detalle.col
    If Detalle.ColSel = 9 And Detalle.col = 1 Then
        PopupMenu MenuPop
    End If
End If
End Sub
Private Sub FilaInsertar_Click()
' inserta fila en FlexGrid
Dim fi As Integer, co As Integer, fi_ini As Integer
fi_ini = Detalle.Row
For fi = n_filas To fi_ini + 1 Step -1
    For co = 1 To n_columnas
        Detalle.TextMatrix(fi, co) = Detalle.TextMatrix(fi - 1, co)
    Next
Next
' fila nueva
For co = 1 To n_columnas
    Detalle.TextMatrix(fi_ini, co) = ""
Next
Detalle.col = 1
Detalle.Row = fi_ini
End Sub
Private Sub FilaEliminar_Click()
' elimina fila en FlexGrid
Dim fi As Integer, co As Integer, fi_ini As Integer
fi_ini = Detalle.Row
For fi = fi_ini To n_filas - 1
    For co = 1 To n_columnas
        Detalle.TextMatrix(fi, co) = Detalle.TextMatrix(fi + 1, co)
    Next
Next
' última fila
For co = 1 To n_columnas
    Detalle.TextMatrix(fi, co) = ""
Next
Detalle.col = 1
Detalle.Row = fi_ini
End Sub
Private Sub FilaBorrarContenido_Click()
Dim fi As Integer, co As Integer
fi = Detalle.Row
For co = 1 To n_columnas
    Detalle.TextMatrix(fi, co) = ""
Next
End Sub
Private Sub Detalle_SelChange()
' se produce con cada click
End Sub
Private Sub Plano_Imprimir()
Dim prt As Printer, linea As String
linea = String(99, "-")
' imprime OT
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer, tab4 As Integer
Dim tab5 As Integer, tab6 As Integer, tab7 As Integer, tab8 As Integer, tab9 As Integer
tab0 = 5 'margen izquierdo
tab1 = tab0 + 0
tab2 = tab0 + 6
tab3 = tab0 + 19
tab4 = tab0 + 34
tab5 = tab0 + 44
tab6 = tab0 + 55
tab7 = tab0 + 66
tab8 = tab0 + 79
tab9 = tab0 + 88

Dim can_valor As String, can_col As Integer

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt

'prt.Font.Size = 15
prt.Print Tab(tab0); "PLANO"
prt.Font.Size = 10
prt.Print ""
prt.Print ""

' cabecera
RsPc.Seek "=", m_Nv, m_NvArea, m_Plano
If RsPc.NoMatch Then Exit Sub
prt.Print Tab(tab0); Empresa.Razon
prt.Print Tab(tab0); "GIRO: " & Empresa.Giro
prt.Print Tab(tab0); Empresa.Direccion
prt.Print Tab(tab0); "Teléfono: " & Empresa.Telefono1 & " - " & Empresa.Comuna
prt.Print Tab(tab8); Format(Now, Fecha_Format)
prt.Print Tab(tab0); "NÚMERO NV   : " & RsPc!Nv
prt.Print Tab(tab0); "OBRA        : " & m_obra
prt.Print Tab(tab0); "PLANO       : " & RsPc!Plano
prt.Print Tab(tab0); "REV         : " & RsPc!Rev
prt.Print Tab(tab0); "FECHA MOD.  : " & Format(RsPc![Fecha Modificacion], Fecha_Format)
prt.Print Tab(tab0); "OBSERVACIÓN : " & NoNulo(RsPc!Observacion)

prt.Print ""
' detalle
prt.Print Tab(tab1); "ITEM";
prt.Print Tab(tab2); "MARCA";
prt.Print Tab(tab3); "DESCRIPCIÓN";
prt.Print Tab(tab4); "CANTIDAD";
prt.Print Tab(tab5); "Kg UNITARIO";
prt.Print Tab(tab6); "   Kg TOTAL";
prt.Print Tab(tab7); "  m2 UNITARIO";
prt.Print Tab(tab8); "     m2 TOTAL"

prt.Print Tab(tab1); linea
'j = -1
For i = 1 To n_filas

    can_valor = Detalle.TextMatrix(i, 3)
    
    If Val(can_valor) = 0 Then
    
    '    j = j + 1
    '    prt.Print Tab(tab1 + j * 5); "----\"

    Else
    
        ' ITEM
        prt.Print Tab(tab1); Format(i, "#### ");
        
        ' MARCA
        prt.Print Tab(tab2); Detalle.TextMatrix(i, 1);
        
        ' DESCRIPCION
        prt.Print Tab(tab3); Detalle.TextMatrix(i, 2);
        
        ' CANTIDAD
        can_valor = Trim(Format(can_valor, "####,###"))
        can_col = 8 - Len(can_valor)
        prt.Print Tab(tab4 + can_col); can_valor;
        
        ' KG UNITARIO
        can_valor = Trim(Format(m_CDbl(Detalle.TextMatrix(i, 5)), "#,###,###.0")) ' 18/05/98
        can_col = 11 - Len(can_valor)
        prt.Print Tab(tab5 + can_col); can_valor;
        
        ' KG TOTAL
        can_valor = Trim(Format(m_CDbl(Detalle.TextMatrix(i, 6)), "#,###,###.0")) '18/05/98
        can_col = 11 - Len(can_valor)
        prt.Print Tab(tab6 + can_col); can_valor;
        
        ' m2 UNITARIO
        can_valor = Trim(Format(m_CDbl(Detalle.TextMatrix(i, 7)), "###,###,###.0"))
        can_col = 13 - Len(can_valor)
        prt.Print Tab(tab7 + can_col); can_valor;
        
        ' m2 TOTAL
        can_valor = Trim(Format(m_CDbl(Detalle.TextMatrix(i, 8)), "###,###,###.0"))
        can_col = 13 - Len(can_valor)
        prt.Print Tab(tab8 + can_col); can_valor
        
    End If
    
Next

prt.Print Tab(tab1); linea
prt.Print ""
prt.Print Tab(tab0 + 40); "TOTAL KILOS : " & Format(PesoTotal, "#,###,###.0");
prt.Print ""

For i = 1 To 5
    prt.Print ""
Next

prt.Print Tab(tab0); Tab(14), "__________________", Tab(56), "__________________"
prt.Print Tab(tab0); Tab(14), "       VºBº       ", Tab(56), "       VºBº       "

prt.EndDoc

Impresora_Predeterminada "default"

End Sub
