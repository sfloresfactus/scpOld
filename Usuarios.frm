VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Usuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantención de Usuarios"
   ClientHeight    =   3120
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnNuevo 
      Caption         =   "&Nuevo Usuario"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.ListView UserList 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu MenuPop 
      Caption         =   ""
      Begin VB.Menu menu 
         Caption         =   "&Agregar"
         Index           =   0
      End
      Begin VB.Menu menu 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu menu 
         Caption         =   "&Eliminar"
         Index           =   2
      End
      Begin VB.Menu menu 
         Caption         =   "Ca&mbiar Nombre"
         Index           =   3
      End
      Begin VB.Menu menu 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu menu 
         Caption         =   "&Propiedades"
         Index           =   5
      End
   End
End
Attribute VB_Name = "Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' privilegios de usuario
' es decir, opciones del menu que estan disponibles
Option Explicit
Private i As Integer, k As Integer
'Private Dbc As Database
'Private RsUs As Recordset
' microsoft activex data object 2.8 library
Private RsUs As New ADODB.Recordset
Private ColNom As ListItem
Private nombre As String
Private sql As String
Private Sub Form_Load()
Dim m_top As Integer

'Set Dbc = OpenDatabase(Syst_file, False, False, ";pwd=eml")
'Set RsUs = Dbc.OpenRecordset("Usuarios")
'RsUs.Index = "Nombre"

UserList_Config

End Sub
Private Sub UserList_Config()
UserList.ColumnHeaders.Add , , "Nombre", 1000
UserList.ColumnHeaders.Add , , "Descripción", 3000
UserList.View = lvwReport
Archivo_Leer
End Sub
Private Sub Archivo_Leer()
'lee archivo de usuarios y llena ListView
'Set RsUs = CnxSqlServer.OpenRecordset("SELECT * FROM usuarios ORDER BY nombre")
RsUs.Open "SELECT * FROM usuarios ORDER BY nombre", CnxSqlServer_scp0
UserList.ListItems.Clear
With RsUs
'.MoveFirst
Do While Not .EOF
    Set ColNom = UserList.ListItems.Add()
    ColNom.Text = !nombre
    ColNom.SubItems(1) = NoNulo(!Descripcion)
'    Debug.Print !nombre & "," & NoNulo(!descripcion)
    .MoveNext
Loop
.Close
End With
End Sub
Private Sub UserList_DblClick()
Propiedades
End Sub
Private Sub Propiedades()
MousePointer = vbHourglass
nombre = UserList.SelectedItem

'RsUs.Seek "=", nombre

'Set RsUs = CnxSqlServer.OpenRecordset("SELECT * FROM usuarios WHERE nombre='" & nombre & "'")
RsUs.Open "SELECT * FROM usuarios WHERE nombre='" & nombre & "'", CnxSqlServer_scp0

'If RsUs.NoMatch Then
If RsUs.EOF Then
    MousePointer = vbDefault
    Exit Sub
Else
    Usr_Propiedades.Usr_Nombre = nombre
'    Usr_Propiedades.Usr_Descripcion = UserList.SelectedItem.SubItems(1)
    Usr_Propiedades.Usr_Descripcion = NoNulo(RsUs!Descripcion)
    Usr_Propiedades.Usr_Clave = NoNulo(RsUs!clave)
'    Usr_Propiedades.Usr_Privi = NoNulo(RsUs!Privilegios)
End If

'Usr_Propiedades.Usr_Recordset = RsUs
Usr_Propiedades.Show 1

RsUs.Close

Archivo_Leer

MousePointer = vbDefault

End Sub
Private Sub UserList_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim OldNombre As String
OldNombre = UserList.SelectedItem
'RsUs.Seek "=", NewString
sql = "SELECT * FROM usuarios WHERE nombre='" & OldNombre & "'"
If Not RsUs.EOF Then
    MsgBox "Nombre ya Existe"
    Cancel = True
    Exit Sub
Else
    RsUs.Seek "=", OldNombre
    If Not RsUs.EOF Then
    
'        RsUs.Edit
'        RsUs!nombre = NewString
'        RsUs.Update
        sql = ""
        
        CnxSqlServer_scp0.Execute sql
        
    End If
End If
End Sub
Private Sub btnNuevo_Click()
Agregar
End Sub
Private Sub Agregar()
MousePointer = vbHourglass
Usr_Propiedades.Usr_Nombre = ""
Usr_Propiedades.Usr_Descripcion = ""
Usr_Propiedades.Usr_Clave = ""
'Usr_Propiedades.Usr_Recordset = RsUs
Usr_Propiedades.Show 1
Archivo_Leer
MousePointer = vbDefault
End Sub
Private Sub UserList_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
nombre = UserList.SelectedItem
If Button = 2 Then
    PopupMenu MenuPop
End If
End Sub
Private Sub menu_Click(Index As Integer)
Select Case Index
Case 0 ' agragar
    Agregar
Case 1 ' raya
Case 2 ' eliminar
    If MsgBox("¿ Elimina " & UCase(nombre) & "?", vbYesNoCancel) = vbYes Then
    
        sql = "DELETE usuarios WHERE nombre='" & nombre & "'"
    
        CnxSqlServer_scp0.Execute sql
    
        Archivo_Leer

    End If
Case 3 ' cambiar nombre
    UserList.StartLabelEdit
Case 4 ' raya
Case 5 ' propiedades
    Propiedades
End Select
End Sub
