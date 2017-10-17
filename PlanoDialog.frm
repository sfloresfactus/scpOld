VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form PlanoDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plano Nuevo"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ComboNv 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox NvNumero 
      Height          =   300
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5040
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.CommandButton btnImportar 
      Caption         =   "&Importar"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Observacion 
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   4575
   End
   Begin VB.TextBox Revision 
      Height          =   300
      Left            =   120
      MaxLength       =   10
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Plano 
      Height          =   300
      Left            =   120
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1200
      Width           =   3135
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancela"
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   3000
      Width           =   855
   End
   Begin VB.ListBox ListaPlanos 
      Height          =   1620
      Left            =   3360
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lbl 
      Caption         =   "&OBSERVACIÓN"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "P&LANOS"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "&REV"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "&PLANO"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "NV"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "PlanoDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private RsNVc As Recordset
Private RsPla As Recordset
Private i As Integer
Private m_Nv As Double
Private m_obra As String
Private m_Plano As String
Private m_Rev As String
Private m_Obs As String
Private m_RevSN As Boolean
Private Accion As String

' para importacion de excel
Private m_Path As String, m_Archivo As String, m_PathArchivo As String
Private a_Nv(2999, 1) As String, m_NvArea As Integer
'
'////////////////////////////////////////////////////////////////////
Public Property Get Nv() As Double
Nv = m_Nv
End Property
Public Property Get obra() As String
obra = m_obra
End Property
Public Property Get NombrePlano() As String
NombrePlano = m_Plano
End Property
Public Property Get Rev() As String
Rev = m_Rev
End Property
Public Property Get RevSN() As Boolean
RevSN = m_RevSN
End Property
Public Property Get Obs() As String
Obs = m_Obs
End Property
Public Property Get Path() As String
' incluye \
Path = m_Path
End Property
Public Property Get Archivo() As String
Archivo = m_Archivo
End Property
'////////////////////////////////////////////////////////////////////
'Public Sub nuevo(RsNVenta As Recordset, RsPlanos As Recordset)
Public Sub nuevo(RsPlanos As Recordset)
Accion = "Agregar"
Load Me

Variables_Limpiar

'Set RsNVc = RsNVenta
Set RsPla = RsPlanos

For i = 1 To nvTotal
    a_Nv(i, 0) = aNv(i).Numero
    a_Nv(i, 1) = aNv(i).obra
    ComboNv.AddItem Format(aNv(i).Numero, "0000") & " - " & aNv(i).obra
Next

On Error GoTo 0

Me.Show 1

End Sub
'Public Sub Abrir(RsNVenta As Recordset, RsPlanos As Recordset, p_Accion As String)
Public Sub Abrir(RsPlanos As Recordset, p_Accion As String)

Load Me

Variables_Limpiar
'Set RsNVc = RsNVenta
Set RsPla = RsPlanos
Accion = p_Accion
Me.Caption = Accion & " Plano"

For i = 1 To nvTotal
    a_Nv(i, 0) = aNv(i).Numero
    a_Nv(i, 1) = aNv(i).obra
    ComboNv.AddItem Format(aNv(i).Numero, "0000") & " - " & aNv(i).obra
Next

Me.Show 1

End Sub

Private Sub Form_Load()
m_Path = ""
m_Archivo = ""

'btnImportar.Enabled = False

Plano.MaxLength = 25
Revision.MaxLength = 10
Observacion.MaxLength = 50

nvListar False

End Sub
Private Sub Variables_Limpiar()
m_Nv = 0
m_obra = ""
m_Plano = ""
m_Rev = ""
m_RevSN = False
m_Obs = ""
End Sub
Private Sub btnImportar_Click()
Dim p As Integer, m_NombreHoja As String

'cd.DialogTitle = "Buscar Planillas Excel de Plano"
cd.DialogTitle = "Buscar Archivo de Texto de Plano"
'cd.Filter = "Microsoft Excel (*.xls)|*.xls|Texto (*.txt)|*.txt|Todos los Archivos (*.*)|*.*"
cd.Filter = "Texto (*.txt)|*.txt|Microsoft Excel (*.xls)|*.xls|Todos los Archivos (*.*)|*.*"

' busca ultima ruta
m_Path = GetSetting("scp", "planos", "ruta")
If m_Path = "" Then m_Path = "C:"

' si m_path no existe, muestra directorio actual
'cd.InitDir = m_Path
m_Path = Directorio(m_Path)

cd.InitDir = m_Path
cd.ShowOpen

m_PathArchivo = cd.filename

If m_PathArchivo = "" Then
'    MsgBox "Debe escoger Archivo"
    Exit Sub
End If

'MsgBox Cd.filename ' viene con ruta completa
m_PathArchivo = cd.filename

' separa path y archivo
p = InStrLast(m_PathArchivo, "\")
If p > 0 Then

    m_Path = Left(m_PathArchivo, p)
    m_Archivo = Mid(m_PathArchivo, p + 1)
    
    ' guarda ultima ruta usada
    SaveSetting "scp", "planos", "ruta", m_Path
    
    p = InStrLast(m_Archivo, "-")
    
    If p > 0 Then
        Plano.Text = Left(m_Archivo, p - 1)
        Revision.Text = Mid(m_Archivo, p + 1, 1) ' ojo , asumo que revision es siempre de una letra
'        p = InStr(1, m_Archivo, ".")
'        Revision.Text = left(Revision.Text

        ' aqui debe leer observacion desde descripciondemarcas de planilla excel
        m_NombreHoja = Left(m_Archivo, Len(m_Archivo) - 4)
        
        Observacion.Text = Excel_Leer_Descripcion(m_Path, m_Archivo, m_NombreHoja)
'        Observacion.Text = Txt_Leer_Descripcion(m_Path, m_Archivo)

    End If
    
End If

'MsgBox m_Path & vbLf & m_Archivo

End Sub
Private Function Excel_Leer_Descripcion(Path As String, Archivo As String, Hoja As String) As String
Dim Hay_Datos As Boolean
' importa descripciones de marcas desde planilla excel
' lo pone en "observaciones"
Dim Planilla As Object, fi As Integer, co As Integer
Dim m_Can As Integer, m_Des As String, m_Descr As String, ul As String

Dim aDes(99) As String, aCan(99) As Integer, j As Integer, k As Integer

If Not Archivo_Existe(Path, Archivo) Then
    MsgBox "No existe archivo " & vbLf & Path & Archivo
    Exit Function
End If

On Error GoTo NoExcel
Set Planilla = GetObject(Path & Archivo, "Excel.Sheet.8") 'assign sheet object as an OLE excel
On Error GoTo 0

With Planilla.Worksheets(Hoja)

m_Des = ""
Hay_Datos = True
fi = 0

k = 0 ' contador de itemes distintos
Do While Hay_Datos

    fi = fi + 1
    
    m_Can = Val(Trim(.cells(fi, 3).Value))
    
    If m_Can = 0 Then
        ' fila está vacia
        ' fin de lectura de planilla
'        Hay_Datos = False
        Exit Do
    End If
    
'    m_Des = Trim(.cells(fi, 5).Value)
'    aDes(fi) = m_Des

    m_Des = Trim(.cells(fi, 5).Value)
    If k > 0 Then
        For j = 1 To k
            If m_Des = aDes(j) Then
                aCan(j) = aCan(j) + 1
                j = 999
            End If
        Next
        If j <> 1000 Then
            ' agrega
            k = k + 1
            aDes(k) = m_Des
            aCan(k) = 1
        End If
    Else
        ' agrega
        k = k + 1
        aDes(k) = m_Des
        aCan(k) = 1
    End If
    
Loop
End With

' pone plurales a marcas repetidas
m_Descr = ""
If k > 0 Then
    For i = 1 To k
        m_Des = aDes(i)
        If aCan(i) > 1 Then
            ' pone plural
            ul = Right(m_Des, 1)
            If InStr(1, "AEIOUaeiou", ul) = 0 Then
                ' no termina en vocal
                If ul = "." Then
                    ' no pone nada
                Else
                    m_Des = m_Des & "ES"
                End If
            Else
                ' termina en vocal
                m_Des = m_Des & "S"
            End If
        End If
        If m_Descr = "" Then
            m_Descr = m_Des
        Else
            m_Descr = m_Descr & ", " & m_Des
        End If
    Next
End If

Excel_Leer_Descripcion = Left(m_Descr, 50)

Set Planilla = Nothing

Exit Function

NoExcel:
'MsgBox "No Tiene Instalado Microsoft Excel"
Excel_Leer_Descripcion = ""
MsgBox "Nombre de Hoja NO Válido"

End Function
Private Function Txt_Leer_Descripcion(Path As String, Archivo As String) As String
' lee descripciones de archivo de texto
Dim Reg As String, largoreg As Integer, p1 As Integer, m_col As Integer, largocont As Integer
Dim m_Des As String, m_Descr As String, ul As String
Dim aDes(99) As String, aCan(99) As Integer, j As Integer, k As Integer

k = 0 ' contador de itemes distintos

Open Path & Archivo For Input As #1

Do While Not EOF(1)

    Line Input #1, Reg
    
    largoreg = Len(Reg)
    
    p1 = 1
    m_col = 0
    
    For i = 1 To largoreg
    
        If Asc(Mid(Reg, i, 1)) = 9 Then
        
            m_col = m_col + 1
            
            If m_col = 5 Then ' descripcion
            
                largocont = i - p1
                m_Des = Mid(Reg, p1, largocont)
                
                If k > 0 Then
                    For j = 1 To k
                        If m_Des = aDes(j) Then
                            aCan(j) = aCan(j) + 1
                            j = 999
                        End If
                    Next
                    If j <> 1000 Then
                        ' agrega
                        k = k + 1
                        aDes(k) = m_Des
                        aCan(k) = 1
                    End If
                Else
                    ' agrega
                    k = k + 1
                    aDes(k) = m_Des
                    aCan(k) = 1
                    
                    Exit For
                    
                End If
                
            End If
            
            p1 = i + 1
            
        End If
        
    Next
    
Loop
Close #1

' pone plurales a marcas repetidas
m_Descr = ""
If k > 0 Then
    For i = 1 To k
        m_Des = aDes(i)
        If aCan(i) > 1 Then
            ' pone plural
            ul = Right(m_Des, 1)
            If InStr(1, "AEIOUaeiou", ul) = 0 Then
                ' no termina en vocal
                If ul = "." Then
                    ' no pone nada
                Else
                    m_Des = m_Des & "ES"
                End If
            Else
                ' termina en vocal
                m_Des = m_Des & "S"
            End If
        End If
        If m_Descr = "" Then
            m_Descr = m_Des
        Else
            m_Descr = m_Descr & ", " & m_Des
        End If
    Next
End If

Txt_Leer_Descripcion = Left(m_Descr, 50)

End Function
Private Sub btnCancelar_Click()
Variables_Limpiar
Unload Me
End Sub
Private Sub btnOk_Click()

If Plano_Validar Then

    If Accion = "Agregar" Then
        If Plano_Existe Then
        Else
            m_Nv = NvNumero.Text
            m_obra = Mid(ComboNv.Text, 8)
            m_Plano = UCase(Trim(UCase(Plano.Text)))
            m_Rev = UCase(Trim(UCase(Revision.Text)))
            m_RevSN = False
            m_Obs = UCase(Trim(Observacion.Text))
        End If
    Else
        m_Nv = NvNumero.Text
        m_obra = Mid(ComboNv.Text, 8)
        m_Plano = UCase(Trim(UCase(Plano.Text)))
        m_Rev = UCase(Trim(UCase(Revision.Text)))
'        m_RevSN = False o True
        m_Obs = UCase(Trim(Observacion.Text))
    End If
    
    Unload Me

End If
End Sub
Private Function Plano_Validar()
Plano_Validar = False

If ComboNv.ListIndex = -1 Then
    Beep
    MsgBox "Debe Elegir Nota de Venta"
    ComboNv.SetFocus
    Exit Function
End If

If Plano.Text = "" Then
    Beep
    MsgBox "Debe Digitar Número de Plano"
    Plano.SetFocus
    Exit Function
End If

If Accion = "Agregar" Then
    If Plano_Existe Then Exit Function
Else
    If Not Plano_Correcto Then Exit Function
End If

If Revision.Text = "" Then
    Beep
    MsgBox "Debe Digitar Revisión"
    Revision.SetFocus
    Exit Function
End If

If Accion = "Agregar" Then
    If Plano_Existe Then Exit Function
Else
    If Not Plano_Correcto Then Exit Function
End If

Plano_Validar = True
End Function
Private Function Plano_Existe() As Boolean
Plano_Existe = True

RsPla.Index = "NV-Plano"
RsPla.Seek "=", Val(NvNumero), m_NvArea, Plano.Text
If RsPla.NoMatch Then
    'ok numero plano nuevo
Else
    MsgBox "PLANO " & RsPla![Plano] & " YA EXISTE" & vbCr & "USE OPCIÓN ""MODIFICAR PLANO"""
    Exit Function
End If
RsPla.Index = "NV-Plano"

Plano_Existe = False
End Function
Private Function Plano_Correcto() As Boolean
Plano_Correcto = False

RsPla.Index = "NV-Plano"
RsPla.Seek "=", Val(NvNumero), m_NvArea, Plano.Text
If RsPla.NoMatch Then
    MsgBox "PLANO " & Plano.Text & " NO EXISTE"
    Exit Function
Else
    If RsPla("Editable") Then
        'OK
        m_RevSN = False
    Else
        ' DIGITAR REVISIÓN DE PLANO
        m_RevSN = True
    End If
End If
RsPla.Index = "NV-Plano"

Plano_Correcto = True
End Function
Private Sub ComboNV_Click()
Nota_Seleccionar ComboNv.ListIndex
End Sub
'Private Sub ListaNotas_Click()
'ListaPlanos.Clear
'NvNumero.Text = ""
'ComboNv.ListIndex = -1
'Plano.Text = ""
'Revision.Text = "" ' vacio para primera revisión
'Observacion.Text = "" ' vacio para primera revisión
'End Sub
Private Sub Nota_Seleccionar(s As Integer)
'NvNumero.Text = a_Nv(s, 0)
NvNumero.Text = a_Nv(s + 1, 0)
ListaPlanos.Clear

RsPla.Seek ">", NvNumero.Text, ""  ' 0
If Not RsPla.NoMatch Then
    Do While Not RsPla.EOF
        If RsPla!Nv = NvNumero.Text Then
            ListaPlanos.AddItem RsPla![Plano] & ", " & RsPla![Rev] & ", " & RsPla!Observacion
        End If
        RsPla.MoveNext
    Loop
End If

'NvNombre.Caption = ListaNotas.Text

End Sub

Private Sub ListaPlanos_Click()
Dim p1 As Integer, p2 As Integer, mi_Plano As String
mi_Plano = ListaPlanos.Text
p1 = InStr(1, mi_Plano, ",")
p2 = InStr(p1 + 1, mi_Plano, ",")

Plano.Text = Left(mi_Plano, p1 - 1)
Observacion.Text = Mid(mi_Plano, p2 + 2)

If Accion = "Modificar" Then
    Revision.Text = ""
Else
    Revision.Text = Mid(mi_Plano, p1 + 2, p2 - p1 - 2)
End If

End Sub
Private Sub NvNumero_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub NvNumero_LostFocus()
Dim m_Nv As Integer
m_Nv = Val(NvNumero.Text)
If m_Nv = 0 Then Exit Sub
' busca nv en combo
i = 0 '1

For i = 1 To nvTotal
'Do Until a_Nv(i, 0) = ""
    If Val(a_Nv(i, 0)) = m_Nv Then
    
        ComboNv.ListIndex = i - 1
        Nota_Seleccionar ComboNv.ListIndex
   
        Exit Sub
        
    End If
'    i = i + 1
'Loop
Next

MsgBox "NV no existe"
NvNumero.SetFocus

End Sub
Private Sub Plano_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Revision_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Observacion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then btnOk.SetFocus
End Sub
