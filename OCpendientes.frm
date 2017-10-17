VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OCpendientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORDENES DE COMPRA PENDIENTES"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Caption         =   "PROVEEDOR"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
      Begin VB.Label lblProvee 
         Caption         =   "TODOS"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3615
      End
   End
   Begin MSComctlLib.ListView ListaOC 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4260
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327680
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MouseIcon       =   "OCpendientes.frx":0000
      NumItems        =   0
   End
End
Attribute VB_Name = "OCpendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Db As Database, Rs As Recordset
Private ColNum As ListItem
Private m_ProveeRut As String, m_ProveeRazon As String
Private m_Oc As Double
Public Property Let Proveedor_RUT(ByVal New_Value As String)
m_ProveeRut = New_Value
End Property
Public Property Let Proveedor_Razon(ByVal New_Value As String)
m_ProveeRazon = New_Value
End Property
Public Property Get Oc() As Double
Oc = m_Oc
End Property
Private Sub Form_Load()
Dim qry As String

lblProvee.Caption = m_ProveeRazon

Set Db = OpenDatabase(Madq_file)
qry = "SELECT * FROM [OC Cabecera] WHERE Pendiente AND NOT Nula"
qry = qry & IIf(m_ProveeRut = "", "", " AND [RUT Proveedor]='" & m_ProveeRut & "'")
qry = qry & " ORDER BY Numero"
Set Rs = Db.OpenRecordset(qry)

Lista_Config

End Sub
Private Sub Lista_Config()
ListaOC.ColumnHeaders.Add , , "N� OC", 700, 0
ListaOC.ColumnHeaders.Add , , "Tipo", 300, 2
ListaOC.ColumnHeaders.Add , , "Fecha", 800, 1
ListaOC.View = lvwReport
Archivo_Leer
End Sub
Private Sub Archivo_Leer()
'lee archivo de usuarios y llena ListView
ListaOC.ListItems.Clear
With Rs
If .RecordCount > 0 Then
    .MoveFirst
    Do While Not .EOF
        Set ColNum = ListaOC.ListItems.Add()
        ColNum.Text = !Numero
        ColNum.SubItems(1) = IIf(!Tipo = "E", "E", "")
        ColNum.SubItems(2) = Format(!Fecha, Fecha_Format)
        .MoveNext
    Loop
End If
End With
End Sub
Private Sub ListaOC_DblClick()
' eligio oc
m_Oc = ListaOC.SelectedItem
Rs.Close
Db.Close
Unload Me
End Sub
