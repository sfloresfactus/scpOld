VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ProvLoc_Lista 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Locales de Proveedor"
   ClientHeight    =   3405
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnNuevo 
      Caption         =   "&Nuevo Usuario"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.ListView LocalesList 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327680
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MouseIcon       =   "ProvLoc_Lista.frx":0000
      NumItems        =   0
   End
   Begin VB.Label lblProveedor 
      Caption         =   "Proveedor"
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label lbl 
      Caption         =   "Pinche con botón secundario del Ratón"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Menu MenuPop 
      Caption         =   ""
      Begin VB.Menu menu 
         Caption         =   "&Agregar Local"
         Index           =   0
      End
      Begin VB.Menu menu 
         Caption         =   "&Modificar Local"
         Index           =   1
      End
      Begin VB.Menu menu 
         Caption         =   "&Eliminar Local"
         Index           =   2
      End
   End
End
Attribute VB_Name = "ProvLoc_Lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private i As Integer, k As Integer
Private Dbc As Database, RsProv As Recordset, RsLoc As Recordset
Private ColCod As ListItem
Private nombre As String
Private m_Rut As String, m_Razon As String
Public Property Let Rut(ByVal New_Value As String)
m_Rut = New_Value
End Property
Public Property Let Razon(ByVal New_Value As String)
m_Razon = New_Value
End Property
Private Sub Form_Load()

lblProveedor.Caption = m_Rut & " " & m_Razon

Set Dbc = OpenDatabase(data_file)

Set RsProv = Dbc.OpenRecordset("Proveedores")
RsProv.Index = "RUT"

Set RsLoc = Dbc.OpenRecordset("Proveedores-Direcciones")
RsLoc.Index = "RUT-Codigo"

LocalesList_Config

End Sub
Private Sub LocalesList_Config()
LocalesList.ColumnHeaders.Add , , "Código", 500
LocalesList.ColumnHeaders.Add , , "Dirección", 3000
LocalesList.ColumnHeaders.Add , , "Comuna", 1000
LocalesList.ColumnHeaders.Add , , "Ciudad", 1000
LocalesList.ColumnHeaders.Add , , "Teléfono", 1000
LocalesList.ColumnHeaders.Add , , "Fax", 1000
LocalesList.ColumnHeaders.Add , , "Contacto", 1000
LocalesList.View = lvwReport
Archivo_Leer
End Sub
Private Sub Archivo_Leer()
LocalesList.ListItems.Clear
With RsLoc
.Seek ">=", m_Rut, 0
If Not .NoMatch Then
    Do While Not .EOF
        If m_Rut <> !Rut Then Exit Do
        Set ColCod = LocalesList.ListItems.Add()
        ColCod.Text = !codigo
        ColCod.SubItems(1) = NoNulo(!Direccion)
        ColCod.SubItems(2) = NoNulo(!Comuna)
        ColCod.SubItems(3) = NoNulo(!Ciudad)
        ColCod.SubItems(4) = NoNulo(![Telefono 1])
        ColCod.SubItems(5) = NoNulo(!Fax)
        ColCod.SubItems(6) = NoNulo(!Contacto)
        .MoveNext
    Loop
End If
End With
End Sub
Private Sub UserList_DblClick()
Propiedades
End Sub
Private Sub UserList_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim OldNombre As String
OldNombre = LocalesList.SelectedItem
RsLoc.Seek "=", NewString
If Not RsLoc.NoMatch Then
    MsgBox "Nombre ya Existe"
    Cancel = True
    Exit Sub
Else
    RsLoc.Seek "=", OldNombre
    If Not RsLoc.NoMatch Then
        RsLoc.Edit
        RsLoc!nombre = NewString
        RsLoc.Update
    End If
End If
End Sub
Private Sub LocalesList_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub
Private Sub LocalesList_DblClick()
Propiedades
End Sub
Private Sub LocalesList_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'nombre = LocalesList.SelectedItem
If Button = 2 Then
    If LocalesList.ListItems.Count = 0 Then
        menu(1).Enabled = False ' modificar
        menu(2).Enabled = False ' eliminar
    Else
        menu(1).Enabled = True ' modificar
        menu(2).Enabled = True ' eliminar
    End If
    PopupMenu MenuPop
End If
End Sub
Private Sub menu_Click(Index As Integer)
Select Case Index
Case 0 ' agragar
    Agregar
Case 1 ' modificar
    Propiedades
Case 2 ' eliminar
    If MsgBox("¿ Elimina " & UCase(nombre) & "?", vbYesNoCancel) = vbYes Then
        RsLoc.Seek "=", nombre
        If Not RsLoc.NoMatch Then
            RsLoc.Delete
            Archivo_Leer
        End If
    End If
End Select
End Sub
Private Sub Agregar()
MousePointer = vbHourglass
ProvLoc_Propiedades.Loc_Rut = m_Rut
ProvLoc_Propiedades.Loc_Id = 0
ProvLoc_Propiedades.Loc_Direccion = ""
ProvLoc_Propiedades.Loc_Comuna = ""
ProvLoc_Propiedades.Loc_Ciudad = ""
ProvLoc_Propiedades.Loc_Telefono = ""
ProvLoc_Propiedades.Loc_Fax = ""
ProvLoc_Propiedades.Loc_Contacto = ""

ProvLoc_Propiedades.Loc_Recordset = RsLoc
ProvLoc_Propiedades.Show 1
Archivo_Leer
MousePointer = vbDefault
End Sub
Private Sub Propiedades()
MousePointer = vbHourglass
nombre = LocalesList.SelectedItem

With RsLoc
.Seek "=", m_Rut, nombre
If .NoMatch Then
    MousePointer = vbDefault
    Exit Sub
Else

    ProvLoc_Propiedades.Loc_Rut = m_Rut
    ProvLoc_Propiedades.Loc_Id = !codigo
    ProvLoc_Propiedades.Loc_Direccion = !Direccion
    ProvLoc_Propiedades.Loc_Comuna = !Comuna
    ProvLoc_Propiedades.Loc_Ciudad = !Ciudad
    ProvLoc_Propiedades.Loc_Telefono = ![Telefono 1]
    ProvLoc_Propiedades.Loc_Fax = !Fax
    ProvLoc_Propiedades.Loc_Contacto = !Contacto
    
    ProvLoc_Propiedades.Loc_Recordset = RsLoc
    ProvLoc_Propiedades.Show 1
    Archivo_Leer

End If

End With
MousePointer = vbDefault
End Sub
