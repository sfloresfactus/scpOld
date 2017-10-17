VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form RecepcionMateriales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Materiales"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ingresar Nueva Recepción"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar Recepción"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar Recepción"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Recepción"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Grabar Recepción"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Deshacer"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CheckBox RecepcionTotal 
      Caption         =   "Recibir Completamente"
      Height          =   375
      Left            =   5160
      TabIndex        =   31
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton btnBuscarOC 
      Height          =   350
      Left            =   2280
      Picture         =   "RecepcionMateriales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Busca OC de Proveedor"
      Top             =   1560
      Width           =   350
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   30
      Top             =   6780
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox NumeroDoc 
      Height          =   300
      Left            =   3480
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "##########"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox ComboTipo 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1920
      Width           =   1575
   End
   Begin MSMask.MaskEdBox Oc 
      Height          =   300
      Left            =   840
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "##########"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   3
      Left            =   240
      MaxLength       =   30
      TabIndex        =   18
      Top             =   6405
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   2
      Left            =   240
      MaxLength       =   30
      TabIndex        =   17
      Top             =   6105
      Width           =   5000
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   1
      Left            =   240
      MaxLength       =   30
      TabIndex        =   16
      Top             =   5805
      Width           =   5000
   End
   Begin VB.Frame Frame 
      Caption         =   "PROVEEDOR"
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   2280
      TabIndex        =   23
      Top             =   400
      Width           =   6375
      Begin VB.CommandButton btnBuscarProveedor 
         Height          =   350
         Left            =   2040
         Picture         =   "RecepcionMateriales.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Busca Proveedor"
         Top             =   190
         Width           =   350
      End
      Begin MSMask.MaskEdBox Rut 
         Height          =   300
         Left            =   600
         TabIndex        =   19
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Comuna 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4080
         TabIndex        =   29
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Direccion 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   480
         TabIndex        =   28
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Razon 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2520
         TabIndex        =   27
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label lbl 
         Caption         =   "DIR"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lbl 
         Caption         =   "RUT"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox txtEditGD 
      Height          =   285
      Left            =   9360
      TabIndex        =   20
      Text            =   "txtEditGD"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Obs 
      Height          =   300
      Index           =   0
      Left            =   240
      MaxLength       =   30
      TabIndex        =   15
      Top             =   5505
      Width           =   5000
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   327680
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Numero 
      Height          =   300
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "##########"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid Detalle 
      Height          =   2925
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5159
      _Version        =   327680
      Enabled         =   -1  'True
      ScrollBars      =   2
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   9600
      Top             =   2160
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
            Picture         =   "RecepcionMateriales.frx":0204
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RecepcionMateriales.frx":0316
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RecepcionMateriales.frx":0428
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RecepcionMateriales.frx":053A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RecepcionMateriales.frx":064C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RecepcionMateriales.frx":075E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "&OC Nº"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Obra 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lbl 
      Caption         =   "OBRA"
      Height          =   255
      Index           =   6
      Left            =   2880
      TabIndex        =   8
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "&Nº Doc"
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   12
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label NVnumero 
      Caption         =   "NVnumero"
      Height          =   255
      Left            =   9480
      TabIndex        =   26
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   22
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "&Tipo Doc"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "&FECHA"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "Nº"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   375
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
Attribute VB_Name = "RecepcionMateriales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnGrabar As Button, btnDesHacer As Button
Private Obj As String, Objs As String, Accion As String
Private i As Integer, j As Integer, d As Variant

Private DbD As Database, RsPrv As Recordset, RsPrd As Recordset
Private Dbm As Database, RsNVc As Recordset
Private DbAdq As Database, RsOcc As Recordset, RsOCd As Recordset, RsCorre As Recordset
Private RsRmC As Recordset, RsRMd As Recordset

Private n_filas As Integer, n_columnas As Integer
Private prt As Printer
Private m_TipoRM As String, TipoDoc As String
Private Const m_Saldada As String = "<S>" 'para productos que ya no se van a recibir
Private m_NvArea As Integer
Private Sub btnBuscarOC_Click()

OCpendientes.Proveedor_RUT = Rut.Text
OCpendientes.Proveedor_Razon = Razon.Caption
OCpendientes.Show 1

If OCpendientes.Oc <> 0 Then
    Oc.Text = OCpendientes.Oc
    Oc_KeyPress 13
End If

End Sub
Private Sub Form_Load()

' abre archivos
Set DbD = OpenDatabase(data_file)
Set RsPrv = DbD.OpenRecordset("Proveedores")
RsPrv.Index = "RUT"
Set RsPrd = DbD.OpenRecordset("Productos")
RsPrd.Index = "Codigo"

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

'Set RsRmD = DbAdq.OpenRecordset("RM Detalle")
Set RsRMd = DbAdq.OpenRecordset("Documentos")
RsRMd.Index = "Tipo-Numero-Linea"

Set RsCorre = DbAdq.OpenRecordset("Correlativo")

Inicializa
Detalle_Config
Botones_Enabled 1, 1, 1, 1, 0, 0

StatusBar.Panels(1) = EmpOC.Razon

m_NvArea = 0

End Sub
Private Sub Form_Unload(Cancel As Integer)
DataBases_Cerrar
End Sub
Private Sub Inicializa()
Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnGrabar = Toolbar.Buttons(6)
Set btnDesHacer = Toolbar.Buttons(7)

Obj = "RECEPCIÓN"
Objs = "RECEPCIONES"

Accion = ""

Campos_Enabled False

ComboTipo.AddItem " "
ComboTipo.AddItem "Factura"
ComboTipo.AddItem "Guía"

TipoDoc = "RM"

End Sub
Private Sub Detalle_Config()
Dim i As Integer, ancho As Integer, c_totales As Integer

n_filas = 20
n_columnas = 11 '10 '9 '8

Detalle.Left = 100
Detalle.WordWrap = True
Detalle.RowHeight(0) = 450
Detalle.Rows = n_filas + 1
Detalle.Cols = n_columnas + 1

'Detalle.ColIsVisible(9) = False '04/11/1999 para guardar [Línea Oc]

Detalle.TextMatrix(0, 0) = ""
Detalle.TextMatrix(0, 1) = "Código"
Detalle.TextMatrix(0, 2) = "Cantidad"
Detalle.TextMatrix(0, 3) = "Uni"               '*
Detalle.TextMatrix(0, 4) = "Descripción"       '*
Detalle.TextMatrix(0, 5) = "Largo(mm)"         '*
Detalle.TextMatrix(0, 6) = "Precio Unitario"   '*
Detalle.TextMatrix(0, 7) = "Cantidad Recibida" '*
Detalle.TextMatrix(0, 8) = "Cantidad a Recibir"
Detalle.TextMatrix(0, 9) = "Cerrar"
Detalle.TextMatrix(0, 10) = "Cert.Recib"

Detalle.ColWidth(0) = 250
Detalle.ColWidth(1) = 1350
Detalle.ColWidth(2) = 850
Detalle.ColWidth(3) = 450
Detalle.ColWidth(4) = 2750
Detalle.ColWidth(5) = 850
Detalle.ColWidth(6) = 850
Detalle.ColWidth(7) = 850
Detalle.ColWidth(8) = 850
Detalle.ColWidth(9) = 550
Detalle.ColWidth(10) = 550
Detalle.ColWidth(11) = 0

'Detalle.CellAlignment = 0
Detalle.ColAlignment(4) = 0 'justificado a la izquierda

ancho = 350 ' con scroll vertical

For i = 0 To n_columnas
    ancho = ancho + Detalle.ColWidth(i)
Next

Detalle.Width = ancho
Me.Width = ancho + Detalle.Left * 2.5

For i = 1 To n_filas
    Detalle.TextMatrix(i, 0) = i
Next

For i = 1 To n_filas
    Detalle.Row = i
    Detalle.col = 1
    Detalle.CellForeColor = vbRed
    Detalle.col = 2
    Detalle.CellForeColor = vbRed
    Detalle.col = 3
    Detalle.CellForeColor = vbRed
    Detalle.col = 4
    Detalle.CellForeColor = vbRed
    Detalle.col = 5
    Detalle.CellForeColor = vbRed
    Detalle.col = 6
    Detalle.CellForeColor = vbRed
    Detalle.col = 7
    Detalle.CellForeColor = vbRed
Next

txtEditGD.Text = ""

'Top = Form_CentraY(Me)
'Left = Form_CentraX(Me)

'Detalle.Enabled = False

End Sub
Private Sub Numero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    After_Enter
End If
End Sub
Private Sub After_Enter()
Select Case Accion
Case "Agregando"
    If Val(Numero.Text) <= 0 Then
        MsgBox "Número NO Válido"
        Exit Sub
    End If
    RsRmC.Seek "=", Numero.Text
    If RsRmC.NoMatch Then
    
        Campos_Enabled True
        Numero.Enabled = False
'        Detalle.Enabled = False
        
        Fecha.SetFocus
        btnGrabar.Enabled = True
        
    Else
    
        Doc_Leer
        MsgBox Obj & " YA EXISTE"
'        Detalle.Enabled = False
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
        
    End If
    
Case "Modificando"
    RsRmC.Seek "=", Numero.Text
    If RsRmC.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
    
        Doc_Leer
        Campos_Enabled True
        Numero.Enabled = False
        Fecha.SetFocus
        btnGrabar.Enabled = True

    End If
Case "Eliminando"
    RsRmC.Seek "=", Numero.Text
    If RsRmC.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Doc_Leer
        Numero.Enabled = False
        If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo) = vbYes Then
            Doc_Eliminar
        End If
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
    End If
Case "Imprimiendo"
    
    RsRmC.Seek "=", Numero.Text
    If RsRmC.NoMatch Then
        MsgBox Obj & " NO EXISTE"
    Else
        Doc_Leer
        Numero.Enabled = False
        Detalle.Enabled = True
    End If
    
End Select

End Sub
Private Sub Doc_Leer()
Dim m_Unidad As String, m_descri As String, Lin_Oc As Integer
' CABECERA
Fecha.Text = Format(RsRmC!Fecha, Fecha_Format)
NvNumero.Caption = RsRmC!Nv
Rut.Text = RsRmC![RUT Proveedor]

Proveedor_Lee Rut.Text
Oc.Text = RsRmC!Oc

m_TipoRM = RsRmC!tipo

RsNVc.Seek "=", NvNumero.Caption, m_NvArea
If Not RsNVc.NoMatch Then
    obra.Caption = NvNumero & " - " & RsNVc!obra
End If

If RsRmC![Tipo Documento] = "F" Then ComboTipo.Text = "Factura"
If RsRmC![Tipo Documento] = "G" Then ComboTipo.Text = "Guía"

NumeroDoc.Text = RsRmC![Numero Documento]

Obs(0).Text = NoNulo(RsRmC![Observacion 1])
Obs(1).Text = NoNulo(RsRmC![Observacion 2])
Obs(2).Text = NoNulo(RsRmC![Observacion 3])
Obs(3).Text = NoNulo(RsRmC![Observacion 4])

'DETALLE

With RsRMd
.Seek ">=", TipoDoc, Val(Numero.Text), 0
i = 0
j = 0
If Not .NoMatch Then
    Do While Not .EOF
        If TipoDoc = !tipo And !Numero = Val(Numero.Text) Then
        
'            i = !Línea
            Lin_Oc = ![Linea Oc]
            
            If !TipoNE = "N" Then 'normal
            
                i = i + 1
                
                Detalle.TextMatrix(i, 1) = ![codigo producto]
                
                m_Unidad = "": m_descri = ""
                RsPrd.Seek "=", ![codigo producto]
                If Not RsPrd.NoMatch Then
                    m_Unidad = RsPrd![unidad de medida]
                    m_descri = RsPrd!Descripcion
                End If
                Detalle.TextMatrix(i, 3) = m_Unidad
                Detalle.TextMatrix(i, 4) = m_descri
                Detalle.TextMatrix(i, 10) = IIf(!certificadoRecibido, m_Saldada, "") ' 29/04/2014
                
                RsOCd.Seek "=", Oc.Text, Lin_Oc
                If Not RsOCd.NoMatch Then
                    Do While Not RsOCd.EOF
                        If RsOCd!Numero <> Oc.Text Then Exit Do
                        If RsOCd![codigo producto] = ![codigo producto] And Lin_Oc = RsOCd!linea Then
                    
                            Detalle.TextMatrix(i, 2) = RsOCd!Cantidad
                            Detalle.TextMatrix(i, 5) = RsOCd![largo]
                            Detalle.TextMatrix(i, 6) = RsOCd![Precio Unitario]
                            Detalle.TextMatrix(i, 7) = Format(RsOCd![Cantidad Recibida], "0.000")
                            Detalle.TextMatrix(i, 10) = IIf(!certificadoRecibido, m_Saldada, "") ' 29/04/2014
                            Detalle.TextMatrix(i, 11) = RsOCd!linea ' 11/04/1999
                            
                            If Accion = "Modificando" Then Detalle.TextMatrix(i, 7) = Format(RsOCd![Cantidad Recibida] - RsRMd!Cant_Entra, "0.000")
                            
                        End If
                        RsOCd.MoveNext
                    Loop
                End If
                Detalle.TextMatrix(i, 8) = !Cant_Entra
                Detalle.TextMatrix(i, 9) = IIf(!Pendiente, "", m_Saldada)
                
            Else 'especial
            
                RsOCd.Seek "=", Oc.Text, Lin_Oc
                If Not RsOCd.NoMatch Then
                    j = j + 1
                    Detalle.TextMatrix(j, 2) = RsOCd!Cantidad
                    Detalle.TextMatrix(j, 3) = NoNulo(RsOCd!unidad)
                    Detalle.TextMatrix(j, 4) = NoNulo(RsOCd!Descripcion)
                    
                    Detalle.TextMatrix(j, 6) = RsOCd![Precio Unitario]
                    Detalle.TextMatrix(j, 7) = Format(RsOCd![Cantidad Recibida], "0.000")
                    Detalle.TextMatrix(j, 8) = !Cant_Entra
                    Detalle.TextMatrix(j, 9) = IIf(!Pendiente, "", m_Saldada)
                    Detalle.TextMatrix(j, 10) = IIf(!certificadoRecibido, m_Saldada, "") ' 29/04/2014
                    Detalle.TextMatrix(j, 11) = RsOCd!linea ' 11/04/1999
                End If
                'Detalle.TextMatrix(j, 10) = IIf(!certificadoRecibido, m_Saldada, "") ' 29/04/2014
                
            End If
        Else
            Exit Do
        End If
        .MoveNext
    Loop
End If
End With

Detalle.Row = 1 ' para q' actualice la primera fila del detalle

End Sub
Private Sub Proveedor_Lee(Rut)
RsPrv.Seek "=", Rut
If Not RsPrv.NoMatch Then
    Razon.Caption = RsPrv![Razon Social]
    Direccion.Caption = RsPrv!Direccion
    Comuna.Caption = NoNulo(RsPrv!Comuna)
End If
End Sub
Private Sub Fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Oc.SetFocus
End Sub
Private Sub Fecha_LostFocus()
d = Fecha_Valida(Fecha, Now)
End Sub
Private Sub ComboTipo_Click()
If NumeroDoc.Enabled Then NumeroDoc.SetFocus
End Sub
Private Sub NumeroDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Detalle.Row = 1
    Detalle.col = 8
    Detalle.SetFocus
End If
End Sub
Private Function Doc_Validar() As Boolean
Dim porDespachar As Integer
Doc_Validar = False

For i = 1 To n_filas

'    If Trim(Detalle.TextMatrix(i, 1)) <> "" Then
    If Trim(Detalle.TextMatrix(i, 2)) <> "" Then

        ' cantidad a recibir
        If Not Numero_Valida(Detalle.TextMatrix(i, 8), i, 8) Then Exit Function
        
        Dim s2 As String, s7 As String, s8 As String
        Dim n2 As Currency, n7 As Currency, n8 As Currency
        
        ' cantidad a recibir
        s2 = m_CDbl(Detalle.TextMatrix(i, 2))
        s7 = m_CDbl(Detalle.TextMatrix(i, 7))
        s8 = m_CDbl(Detalle.TextMatrix(i, 8))
        
        n2 = CCur(s2)
        n7 = CCur(s7)
        n8 = CCur(s8)
        
'        If n2 - n7 <= n8 Then
        If n2 >= (n7 + n8) Then
            ' ok
        Else
            Beep
            MsgBox "Cantidad a Recibir:" & vbLf & "Línea " & i & vbLf & n2 & " < " & n7 & "+" & n8, , "Error"
            Detalle.Row = i
            Detalle.col = 8
            Detalle.SetFocus
            Exit Function
        End If
        
    End If
    
Next

Doc_Validar = True

End Function
Private Function CampoReq_Valida(txt As String, fil As Integer, col As Integer) As Boolean
' valida si campo requerido
If Len(Trim(txt)) = 0 Then
    CampoReq_Valida = False
    Beep
    MsgBox "CAMPO OBLIGATORIO"
    Detalle.Row = fil
    Detalle.col = col
    Detalle.SetFocus
Else
    CampoReq_Valida = True
End If
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
num = "0" & txt
If Not IsNumeric(num) Then
    If num <> "" Then
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
Private Function Fecha_Req(Fecha As String, fil As Integer, col As Integer) As Boolean
Fecha_Req = False
Fecha = Replace(Fecha, "_")
If Fecha = "//" Or Fecha = "" Then
    Beep
    MsgBox "DEBE digitar Fecha"
    Detalle.Row = fil
    Detalle.col = col
    Detalle.SetFocus
    Exit Function
End If
Fecha_Req = True
End Function
Private Function Doc_Grabar(Nueva As Boolean) As Boolean
MousePointer = vbHourglass
Dim m_cantidad As Double, m_Numero As Double, m_nLinea As Integer, m_PrecioUnitario As Double
Doc_Grabar = True
save:
' Rm Cabecera
With RsRmC

If Nueva Then
    
    m_Numero = GetNumDoc("RM", RsRmC, RsCorre)
    If m_Numero = 0 Then
        MsgBox "RM NO SE GRABÓ!"
        Doc_Grabar = False
        MousePointer = vbDefault
        Exit Function
    End If
    
    Numero.Text = m_Numero
    
    .AddNew
    !Numero = Numero.Text
    
Else

    Doc_Detalle_Eliminar
    .Edit
    
End If

!Fecha = Fecha.Text
!Nv = NvNumero.Caption
![RUT Proveedor] = Rut.Text
!Oc = Oc.Text
!tipo = m_TipoRM
![Tipo Documento] = Left(ComboTipo.Text, 1)
![Numero Documento] = Val(NumeroDoc.Text)
![Observacion 1] = Obs(0).Text
![Observacion 2] = Obs(1).Text
![Observacion 3] = Obs(2).Text
![Observacion 4] = Obs(3).Text
.Update

End With

' Rm Detalle
j = 0
For m_nLinea = 1 To n_filas

'    m_cantidad = Val(Detalle.TextMatrix(m_nLinea, 8))
    m_cantidad = m_CDbl(Detalle.TextMatrix(m_nLinea, 8))
    
    If m_cantidad <> 0 Or Detalle.TextMatrix(m_nLinea, 9) = m_Saldada Then
        
        m_PrecioUnitario = 0 '??
        
        With RsOCd
'        .Seek "=", Oc.Text, m_nLinea
        .Seek "=", Oc.Text, Detalle.TextMatrix(m_nLinea, 11) '11/11/1999
        If Not .NoMatch Then
        
            m_PrecioUnitario = ![Precio Unitario]
        
            If !tipo = "N" Then
        
'                Do While Not .EOF
'                    If !Número <> Oc.Text Then Exit Do
                    
                    If ![codigo producto] = Detalle.TextMatrix(m_nLinea, 1) Then
                        .Edit
                        ![Cantidad Recibida] = ![Cantidad Recibida] + m_cantidad
                        ![Fecha Recepcion] = Fecha.Text
                        !Pendiente = IIf(Detalle.TextMatrix(m_nLinea, 9) = m_Saldada, False, True)
                        .Update
'                        Exit Do
                    End If
'                    .MoveNext
'                Loop
                
            Else 'especial
            
                .Edit
                ![Cantidad Recibida] = ![Cantidad Recibida] + m_cantidad
                ![Fecha Recepcion] = Fecha.Text
                !Pendiente = IIf(Detalle.TextMatrix(m_nLinea, 9) = m_Saldada, False, True)
                .Update
            
            End If
            
        End If
        End With
        
        With RsRMd
        .AddNew
        !tipo = TipoDoc
        !Numero = Numero.Text
        j = j + 1
        !linea = j 'i ' j 'i
        !TipoNE = m_TipoRM
        !Fecha = Fecha.Text
        !Nv = NvNumero.Caption
        ![Rut] = Rut.Text
        ![codigo producto] = Detalle.TextMatrix(m_nLinea, 1)
        ![Precio Unitario] = m_PrecioUnitario
        !Cant_Entra = m_cantidad
        !Pendiente = IIf(Detalle.TextMatrix(m_nLinea, 9) = m_Saldada, False, True)
        !certificadoRecibido = IIf(Detalle.TextMatrix(m_nLinea, 10) = m_Saldada, True, False)
        
        ' nuevo 28/10/1999
        !Oc = Oc.Text
        ![Linea Oc] = Val(Detalle.TextMatrix(m_nLinea, 11)) ' <- 04/11/1999
        
        .Update
        
        End With
        
    End If
Next

DesBloqueo "RM", RsCorre

Oc_Marca_Pendiente Oc.Text

MousePointer = vbDefault

End Function
Private Sub Doc_Eliminar()

' elimina cabecera
RsRmC.Seek "=", Numero.Text
If Not RsRmC.NoMatch Then

    RsRmC.Delete
   
End If

' elimina detalle
Doc_Detalle_Eliminar

End Sub
Private Sub old_Doc_Detalle_Eliminar()
Dim m_Tipo As String
RsRMd.Seek ">=", TipoDoc, Numero.Text, 0
If Not RsRMd.NoMatch Then
    Do While Not RsRMd.EOF
        If RsRMd!tipo <> TipoDoc Or RsRMd!Número <> Numero.Text Then Exit Do
    
        ' actualiza "cantidad recibida" en oc
        With RsOCd
        If RsRMd!tipo = "N" Then
'            .Seek ">=", Oc.Text, 1
            .Seek "=", Oc.Text, RsRMd![Línea Oc]
            If Not .NoMatch Then
'               Do While Not .EOF
'                    If !Número <> Oc.Text Then Exit Do
                    If ![Código Producto] = RsRMd![Código Producto] Then
                        .Edit
                        ![Cantidad Recibida] = ![Cantidad Recibida] - RsRMd!Cantidad
                        .Update
'                        Exit Do
                    End If
'                   .MoveNext
'               Loop
            End If
        Else 'rm especial
'            .Seek ">=", Oc.Text, 1 ', RsRmD!Línea
            .Seek "=", Oc.Text, RsRMd![Línea Oc]
'            Debug.Print RsRmD![Línea Oc]
            If Not .NoMatch Then
'                Do While Not .EOF
'                    If !Número <> Oc.Text Then Exit Do
'                    If !Línea = RsRmD![Línea Oc] Then
                        .Edit
                        ![Cantidad Recibida] = ![Cantidad Recibida] - RsRMd!Cantidad
                        .Update
'                        Exit Do
'                    End If
'                    .MoveNext
'                Loop
            End If
        End If
        End With
        
        ' borra detalle
        RsRMd.Delete
    
        RsRMd.MoveNext
        
    Loop
End If
End Sub
Private Sub Doc_Detalle_Eliminar()

RsRMd.Seek ">=", TipoDoc, Numero.Text, 0
If RsRMd.NoMatch Then Exit Sub

If RsRMd!TipoNE = "N" Then

    ' elimina detalle normal
    Do While Not RsRMd.EOF
    
        If RsRMd!tipo <> TipoDoc Or RsRMd!Numero <> Numero.Text Then Exit Do
        
        ' actualiza "cantidad recibida" en oc
        With RsOCd
        .Seek "=", Oc.Text, RsRMd![Linea Oc]
        If Not .NoMatch Then
            If ![codigo producto] = RsRMd![codigo producto] Then
                .Edit
                ![Cantidad Recibida] = ![Cantidad Recibida] - RsRMd!Cant_Entra
                .Update
            End If
        End If
        End With
        RsRMd.Delete
        RsRMd.MoveNext
    Loop

Else
    ' detalle especial
    Do While Not RsRMd.EOF
    
        If RsRMd!tipo <> TipoDoc Or RsRMd!Numero <> Numero.Text Then Exit Do
        
        ' actualiza "cantidad recibida" en oc
        With RsOCd
        .Seek "=", Oc.Text, RsRMd![Linea Oc]
        If Not .NoMatch Then
            .Edit
            ![Cantidad Recibida] = ![Cantidad Recibida] - RsRMd!Cant_Entra
            .Update
        End If
        End With
        RsRMd.Delete
        RsRMd.MoveNext
    Loop

End If

End Sub
Private Sub Campos_Limpiar()
Numero.Text = ""
Fecha.Text = Fecha_Vacia
Rut.Text = ""
Razon.Caption = ""
Direccion.Caption = ""
Comuna.Caption = ""

Oc.Text = ""
obra.Caption = ""
ComboTipo.Text = " "
NumeroDoc.Text = ""

RecepcionTotal.Value = 0

Detalle_Limpiar

Obs(0).Text = ""
Obs(1).Text = ""
Obs(2).Text = ""
Obs(3).Text = ""

End Sub
Private Sub Detalle_Limpiar()
Dim fi As Integer, co As Integer
For fi = 1 To n_filas
    For co = 1 To n_columnas
        Detalle.TextMatrix(fi, co) = ""
    Next
Next
End Sub
Private Sub Obs_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If Index = 3 Then
        Obs(0).SetFocus
    Else
        Obs(Index + 1).SetFocus
    End If
End If
End Sub
Private Sub Oc_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Dim m_cod As String
    Detalle_Limpiar
    RsOcc.Seek "=", Oc.Text
    If RsOcc.NoMatch Then
        MsgBox "ORDEN DE COMPRA NO EXISTE"
        Oc.SetFocus
    Else
    
        If RsOcc!Nula Then
            MsgBox "OC Nula"
            Oc.Text = ""
            Oc.SetFocus
            Exit Sub
        End If
    
        Dim m_Unidad As String, m_descri As String
        ' lee orden de compra
        NvNumero.Caption = RsOcc!Nv
        
        RsNVc.Seek "=", NvNumero.Caption, m_NvArea
        If Not RsNVc.NoMatch Then
            obra.Caption = NvNumero & " - " & RsNVc!obra
        End If

        m_TipoRM = RsOcc!tipo ' 04/11/1999

        Rut.Text = RsOcc![RUT Proveedor]
        Proveedor_Lee Rut.Text
        
        ' lee detalle oc
        RsOCd.Seek "=", Oc.Text, 1
        If Not RsOCd.EOF Then
            Do While Not RsOCd.EOF
                If RsOCd!Numero <> Oc.Text Then Exit Do
                i = RsOCd!linea
                
                Detalle.TextMatrix(i, 2) = RsOCd!Cantidad
                
                m_cod = NoNulo(RsOCd![codigo producto])
                m_Unidad = "": m_descri = ""
                If m_cod = "" Then
                    m_Unidad = RsOCd!unidad
                    m_descri = RsOCd!Descripcion
                Else
                    Detalle.TextMatrix(i, 1) = m_cod
                    RsPrd.Seek "=", m_cod
                    If Not RsPrd.NoMatch Then
                        m_Unidad = RsPrd![unidad de medida]
                        m_descri = RsPrd!Descripcion
                    End If
                End If
                Detalle.TextMatrix(i, 3) = m_Unidad
                Detalle.TextMatrix(i, 4) = m_descri
                Detalle.TextMatrix(i, 5) = RsOCd![largo]
                Detalle.TextMatrix(i, 6) = RsOCd![Precio Unitario]
                Detalle.TextMatrix(i, 7) = Format(RsOCd![Cantidad Recibida], "0.000")
                
                Detalle.TextMatrix(i, 9) = IIf(RsOCd!Pendiente, "", m_Saldada) '28/10/1999
                
                'Detalle.TextMatrix(i, 10) = IIf(RsOCd!certificadoRecibido, "", m_Saldada) '27/06/2014
                
                Detalle.TextMatrix(i, 11) = i ' 04/11/1999
                            
                RsOCd.MoveNext
            Loop
        End If
        
        ComboTipo.SetFocus
        
    End If
End If
End Sub
Private Sub RecepcionTotal_Click()
If RecepcionTotal.Value = 0 Then
    For i = 1 To n_filas
        Detalle.TextMatrix(i, 8) = ""
    Next
Else
    ' puso ticket
    For i = 1 To n_filas
        Detalle.TextMatrix(i, 8) = Detalle.TextMatrix(i, 2)
    Next
End If
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim cambia_titulo As Boolean, grabado As Boolean
cambia_titulo = True
'Accion = ""
Select Case Button.Index
Case 1 ' agregar
    Accion = "Agregando"
    Botones_Enabled 0, 0, 0, 0, 0, 1
    Campos_Enabled False
    
'    Numero.Text = Documento_Numero_Nuevo(RsRMc, "Número")
'    Numero.Enabled = True
'    Numero.SetFocus
    Campos_Enabled True
    Numero.Enabled = False
    Fecha.SetFocus
    btnGrabar.Enabled = True
'    btnSearch.Visible = True

Case 2 ' modificar
    Accion = "Modificando"
    Botones_Enabled 0, 0, 0, 0, 0, 1
    Campos_Enabled False
    Numero.Enabled = True
    Numero.SetFocus
Case 3 ' eliminar
    Accion = "Eliminando"
    Botones_Enabled 0, 0, 0, 0, 0, 1
    Campos_Enabled False
    Numero.Enabled = True
    Numero.SetFocus
Case 4 ' imprimir
    Accion = "Imprimiendo"
    If Numero.Text = "" Then
        Botones_Enabled 0, 0, 0, 1, 0, 1
        Campos_Enabled False
        Numero.Enabled = True
        Numero.SetFocus
    Else
        If MsgBox("¿ Imprime ?", vbYesNo) = vbYes Then
            Doc_Imprimir
        End If
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
    End If
Case 6 ' grabar

    If Doc_Validar Then
        
        If MsgBox("¿ GRABAR " & Obj & " ?", vbYesNo) = vbYes Then
        
            grabado = True
            If Accion = "Agregando" Then
                grabado = Doc_Grabar(True)
            Else
                Doc_Grabar False
            End If
            
            If grabado Then
                If MsgBox("¿ IMPRIME " & Obj & " ?", vbYesNo) = vbYes Then
                    Doc_Imprimir
                End If
            End If
        
            Botones_Enabled 0, 0, 0, 0, 0, 1
            Campos_Limpiar
            
            If Accion = "Agregando" Then
                Campos_Enabled True
                Numero.Enabled = False
                Fecha.SetFocus
                btnGrabar.Enabled = True
'                btnSearch.Visible = True
            Else
                Campos_Enabled False
                Numero.Enabled = True
                Numero.SetFocus
            End If

        End If
    End If
Case 7 ' DesHacer
    If Numero.Text = "" Then
        Botones_Enabled 1, 1, 1, 1, 0, 0
        Campos_Limpiar
        Campos_Enabled False
    Else
        If Accion = "Imprimiendo" Then
            Botones_Enabled 1, 1, 1, 1, 0, 0
            Campos_Limpiar
            Campos_Enabled False
        Else
            If MsgBox("¿Seguro que quiere DesHacer?", vbYesNo) = vbYes Then
                Botones_Enabled 1, 1, 1, 1, 0, 0
                Campos_Limpiar
                Campos_Enabled False
            End If
        End If
    End If
    Accion = ""

End Select

If cambia_titulo Then
    If Accion = "" Then
        Me.Caption = "Mantención de " & StrConv(Objs, vbProperCase)
    Else
        Me.Caption = "Mantención de " & StrConv(Objs, vbProperCase) & " [" & Accion & "]"
    End If
End If

End Sub
Private Sub Botones_Enabled(btn_Agregar As Boolean, btn_Modificar As Boolean, _
                            btn_Eliminar As Boolean, btn_Imprimir As Boolean, _
                             btn_Grabar As Boolean, btn_DesHacer As Boolean)
                            
btnAgregar.Enabled = btn_Agregar
btnModificar.Enabled = btn_Modificar
btnEliminar.Enabled = btn_Eliminar
btnImprimir.Enabled = btn_Imprimir
btnGrabar.Enabled = btn_Grabar
btnDesHacer.Enabled = btn_DesHacer

btnAgregar.Value = tbrUnpressed
btnModificar.Value = tbrUnpressed
btnEliminar.Value = tbrUnpressed
btnImprimir.Value = tbrUnpressed
btnGrabar.Value = tbrUnpressed
btnDesHacer.Value = tbrUnpressed

End Sub
Private Sub Campos_Enabled(Si As Boolean)
Numero.Enabled = Si
Fecha.Enabled = Si
btnBuscarProveedor.Enabled = Si
Oc.Enabled = Si
btnBuscarOC.Enabled = Si
ComboTipo.Enabled = Si
NumeroDoc.Enabled = Si
RecepcionTotal.Enabled = Si
Detalle.Enabled = Si
Obs(0).Enabled = Si
Obs(1).Enabled = Si
Obs(2).Enabled = Si
Obs(3).Enabled = Si
End Sub
Private Sub btnBuscarProveedor_Click()

Search.Muestra data_file, "Proveedores", "RUT", "Razon Social", "Proveedor", "Proveedores"
Rut.Text = Search.Codigo

If Rut.Text <> "" Then
    RsPrv.Seek "=", Rut.Text
    If RsPrv.NoMatch Then
        MsgBox "PROVEEDOR NO EXISTE"
        Rut.SetFocus
    Else
        Razon.Caption = Search.Descripcion
        Direccion.Caption = RsPrv!Direccion
        Comuna.Caption = NoNulo(RsPrv!Comuna)
    End If
End If
End Sub
Private Sub Producto_Buscar()
Dim m_Codigo As String, fi As Integer
fi = Detalle.Row
m_Codigo = Detalle.TextMatrix(fi, 1)

RsPrd.Seek "=", m_Codigo
If RsPrd.NoMatch Then
    MsgBox "PRODUCTO NO EXISTE"
    For i = 1 To n_columnas
        Detalle.TextMatrix(fi, i) = ""
    Next
Else
    Detalle.TextMatrix(fi, 3) = RsPrd![unidad de medida]
    Detalle.TextMatrix(fi, 4) = RsPrd!descripción
    Detalle.TextMatrix(fi, 5) = RsPrd![Precio Compra]
End If

End Sub
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'
' RUTINAS PARA EL FLEXGRID
'
'
Private Sub Detalle_Click()
If Accion = "Imprimiendo" Then Exit Sub
After_Detalle_Click
End Sub
Private Sub After_Detalle_Click()
Select Case Detalle.col
    Case 1 ' codigo
    Case Else
End Select
End Sub
Private Sub Detalle_DblClick()
If Accion = "Imprimiendo" Then Exit Sub
MSFlexGridEdit Detalle, txtEditGD, 32
End Sub
Private Sub Detalle_GotFocus()
If txtEditGD.visible Then
    Detalle = txtEditGD
    txtEditGD.visible = False
End If
End Sub
Private Sub Detalle_LeaveCell()
If txtEditGD.visible Then
    Detalle = txtEditGD
    txtEditGD.visible = False
End If
End Sub
Private Sub Detalle_KeyPress(KeyAscii As Integer)
If Accion = "Imprimiendo" Then Exit Sub
MSFlexGridEdit Detalle, txtEditGD, KeyAscii
End Sub
Private Sub txtEditGD_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCodeP Detalle, txtEditGD, KeyCode, Shift
End Sub
Sub EditKeyCodeP(MSFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
' rutina que se ejecuta con los keydown de Edt
Dim m_col As Integer
m_col = MSFlexGrid.col
Select Case KeyCode
Case vbKeyEscape ' Esc
    Edt.visible = False
    MSFlexGrid.SetFocus
Case vbKeyReturn ' Enter
    Select Case m_col
    Case 1 ' Código producto
        MSFlexGrid.SetFocus
        DoEvents
        Producto_Buscar
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
    End Select
    Cursor_Mueve MSFlexGrid
Case vbKeyUp ' Flecha Arriba
    Select Case m_col
    Case 1 ' Codigo
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
        If MSFlexGrid.Row > MSFlexGrid.FixedRows Then
            MSFlexGrid.Row = MSFlexGrid.Row - 1
        End If
    End Select
Case vbKeyDown ' Flecha Abajo
    Select Case m_col
    Case 1 ' Codigo
    Case Else
        MSFlexGrid.SetFocus
        DoEvents
'        Linea_Actualiza
        If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
            MSFlexGrid.Row = MSFlexGrid.Row + 1
        End If
    End Select
End Select
End Sub
Private Sub txtEditGD_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)
Select Case MSFlexGrid.col
Case 8 'cant a recibir
    Select Case KeyAscii
    Case 0 To 32
        Edt = MSFlexGrid
        Edt.SelStart = 1000
    Case Else
        Edt = Chr(KeyAscii)
        Edt.SelStart = 1
    End Select
    Edt.Move MSFlexGrid.CellLeft + MSFlexGrid.Left, MSFlexGrid.CellTop + MSFlexGrid.Top, MSFlexGrid.CellWidth, MSFlexGrid.CellHeight
    Edt.visible = True
    Edt.SetFocus
    'opGrabar True
Case 9 ' cerrar
    If KeyAscii = vbKeySpace Then
        Detalle.TextMatrix(MSFlexGrid.Row, 9) = IIf(Detalle.TextMatrix(MSFlexGrid.Row, 9) = "", m_Saldada, "")
    End If
Case 10 ' certificado recibido
    If KeyAscii = vbKeySpace Then
        Detalle.TextMatrix(MSFlexGrid.Row, 10) = IIf(Detalle.TextMatrix(MSFlexGrid.Row, 10) = "", m_Saldada, "")
    End If
End Select
End Sub
Private Sub Detalle_KeyDown(KeyCode As Integer, Shift As Integer)
If Accion = "Imprimiendo" Then Exit Sub
Dim m_cod As String, m_Des As String, m_uni As String
Select Case KeyCode
Case vbKeyF1
    If Detalle.col = 1 Then
    
        MousePointer = vbHourglass
        Product_Search.Condicion = ""
        Load Product_Search
        MousePointer = vbDefault
        Product_Search.Show 1
        
        m_cod = Product_Search.CodigoP
        RsPrd.Seek "=", m_cod
        If Not RsPrd.NoMatch Then
            m_uni = RsPrd![unidad de medida]
            m_Des = RsPrd!Descripcion
            
            Detalle.TextMatrix(Detalle.Row, 1) = m_cod
            Detalle.TextMatrix(Detalle.Row, 3) = m_uni
            Detalle.TextMatrix(Detalle.Row, 4) = m_Des
            Detalle.col = 2 'foco en cantidad

        End If

    End If
Case vbKeyF2
    MSFlexGridEdit Detalle, txtEditGD, 32
End Select
End Sub
'Private Sub Detalle_RowColChange()
'MIA
'Posicion = "Lín " & Detalle.Row & ", Col " & Detalle.col
'End Sub
Private Sub Cursor_Mueve(MSFlexGrid As Control)
'MIA
Select Case MSFlexGrid.col
'Case 1
'    MSFlexGrid.col = 2
'Case 2
'    MSFlexGrid.col = 5
'Case 7
'    MSFlexGrid.col = 8
Case 9 'cerrar
    If MSFlexGrid.Row + 1 < MSFlexGrid.Rows Then
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
Private Sub Doc_Imprimir()
MousePointer = vbHourglass
' imprime OT
Dim tab0 As Integer, tab1 As Integer, tab2 As Integer, tab3 As Integer
Dim tab4 As Integer, tab5 As Integer, tab6 As Integer, tab40 As Integer
tab0 = 7 'margen izquierdo
tab1 = tab0 + 0  ' codigo
tab2 = tab1 + 10 ' cantidad total
tab3 = tab2 + 10 ' unidad de medida
tab4 = tab3 + 5  ' descripcion
tab5 = tab4 + 25 ' largo
tab6 = tab5 + 8  ' cant recibida en esta recepcion
tab40 = 39
Dim can_valor As String, linea As String
linea = String(72, "-")

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt
'prt.Orientation = vbPRORLandscape

prt.Font.Size = 15
prt.Print Tab(tab0); "RECEPCION MATERIALES Nº" & Numero.Text;
prt.Print Tab(tab0 + 30); "FECHA : " & Fecha.Text
prt.Font.Size = 12
prt.Print ""
prt.Print ""

' cabecera
prt.Font.Bold = True
prt.Print Tab(tab0); Empresa.Razon;
prt.Font.Bold = False
prt.Print Tab(tab0 + tab40); "SEÑOR(ES) : " & Razon

prt.Print Tab(tab0); "R.U.T. : "; Empresa.Rut;
prt.Print Tab(tab0 + tab40); "RUT       : " & Rut

prt.Print Tab(tab0); "GIRO: " & Empresa.Giro;
prt.Print Tab(tab0 + tab40); "DIRECCIÓN : " & Direccion

prt.Print Tab(tab0); Empresa.Direccion;
prt.Print Tab(tab0 + tab40); "COMUNA    : " & Comuna.Caption

prt.Print Tab(tab0); Empresa.Comuna

prt.Print Tab(tab0); "Teléfono: " & Empresa.Telefono1;


prt.Print Tab(tab0 + tab40); "OBRA      : ";
prt.Font.Bold = True
'prt.Print Format(Mid(ComboNV.Text, 8), ">")
prt.Font.Bold = False

' datos grales orden

prt.Print ""
' detalle
prt.Font.Bold = True
prt.Print Tab(tab1); "CÓDIGO";
prt.Print Tab(tab2); "CAN. TOT";
prt.Print Tab(tab3); "UNI.";
prt.Print Tab(tab4); "DESCRIPCIÓN";
prt.Print Tab(tab5); "LARGO";
prt.Print Tab(tab6); " CAN.REC."
prt.Font.Bold = False

prt.Print Tab(tab1); linea
j = -1
For i = 1 To n_filas

    can_valor = m_CDbl(Detalle.TextMatrix(i, 2))

    If can_valor = 0 Then
    
        j = j + 1
        prt.Print Tab(tab1 + j * 4); "    \"
        
    Else
    
        ' CÓDIGO
        prt.Print Tab(tab1); Detalle.TextMatrix(i, 1);
        
        ' CANTIDAD total
        If can_valor = Int(can_valor) Then
            prt.Print Tab(tab2); m_Format(can_valor, "##,##0.000");
        Else
            prt.Print Tab(tab2); m_Format(can_valor, "##,###.###");
        End If
        
        ' U.MEDIDA
        prt.Print Tab(tab3); Detalle.TextMatrix(i, 3);
        
        ' DESCRIPCION
        prt.Print Tab(tab4); Left(Detalle.TextMatrix(i, 4), 22); ' 50
        
        ' LARGO
        prt.Print Tab(tab5); m_Format(Detalle.TextMatrix(i, 5), "#####");
        
        ' cant recibida en esta recepcion
        can_valor = m_CDbl(Detalle.TextMatrix(i, 8))
        If can_valor = Int(can_valor) Then
            prt.Print Tab(tab6); m_Format(can_valor, "##,##0.000")
        Else
            prt.Print Tab(tab6); m_Format(can_valor, "##,###.###")
        End If
        
    End If
    
Next

prt.Print Tab(tab1); linea

prt.Print Tab(tab0); "OBSERVACIONES :";

prt.Print Tab(tab0); Obs(0).Text
prt.Print Tab(tab0); Obs(1).Text
prt.Print Tab(tab0); Obs(2).Text
prt.Print Tab(tab0); Obs(3).Text

For i = 1 To 2
    prt.Print ""
Next

prt.Print Tab(tab0); Tab(12), "__________________", Tab(56), "__________________"
prt.Print Tab(tab0); Tab(12), "       VºBº       ", Tab(56), "       VºBº       "
prt.Print Tab(tab0); ""

prt.EndDoc

Impresora_Predeterminada "default"

MousePointer = vbDefault

End Sub
Private Sub Oc_Marca_Pendiente(Oc As Double)
' busca oc y la marca como pendiente o saldada
RsOcc.Seek "=", Oc
If Not RsOcc.NoMatch Then

    RsOcc.Edit
    RsOcc!Pendiente = False
    RsOcc.Update

    RsOCd.Seek "=", Oc, 1
    If Not RsOCd.NoMatch Then
    
        Do While Not RsOCd.EOF
            If RsOCd!Numero <> Oc Then Exit Do
            If RsOCd!Cantidad > RsOCd![Cantidad Recibida] And RsOCd!Pendiente Then
                RsOcc.Edit
                RsOcc!Pendiente = True
                RsOcc.Update
                Exit Do
            End If
            RsOCd.MoveNext
        Loop
        
    End If
    
End If

End Sub
