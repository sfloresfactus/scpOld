VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form planosEnExcelImportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importación de Planos desde Excel"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ComboHoja 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox Nv 
      Height          =   300
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin Crystal.CrystalReport cr 
      Left            =   4200
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton btnImportar 
      Caption         =   "&Importar"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton btnAbrir 
      Caption         =   "&Examinar"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox CbNv 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4560
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Label lblHoja 
      Caption         =   "Hoja Excel"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   " NV | PLANO | REV | MARCA | CANTIDAD | DESCRIPCION | PESOUNI | SUPUNI | OBS"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   "   A         B         C          D                E                      F                   G                H           I"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   6495
   End
   Begin VB.Label lblExplicacion 
      Alignment       =   2  'Center
      Caption         =   "Opcion para importar planos desde una planilla Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6495
   End
   Begin VB.Label lblArchivo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Label lblCarpeta 
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label lbl 
      Caption         =   "OBRA"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   1080
      Width           =   495
   End
End
Attribute VB_Name = "planosEnExcelImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' importacion de archivos desde una planilla excel
' requerida por kheniquez, 07/01/2014
' solo para piezas nuevas ??
' //////////////////////////////////////////////////////////////////////////////////
' -----------------------------------------------------------------------------------
'  A     B      C      D        E           F            G        H       I
' NV | PLANO | REV | MARCA | CANTIDAD | DESCRIPCION | PESOUNI | SUPUNI | OBS
'////////////////////////////////////////////////////////////////////////////////////

Option Explicit
Private Dbm As Database ', RsNVc As Recordset
Private RsPc As Recordset, RsPd As Recordset, RsOTfd As Recordset, RsITOfd As Recordset

Private m_Path As String
Private m_Nv As Integer, m_Plano As String, m_Rev As String, m_Marca As String
' para detalle
Private m_Can As Integer, m_Mar As String, m_Des As String, m_KgU As Double, m_KgT As Double, m_m2U As Double, m_m2T As Double, m_Den As Integer

Private m_PesoTotal As Double, m_SuperficieTotal As Double

Private m_CantidadPlanosNuevos As Integer
Private i As Integer, j As Integer, existe As Boolean, m_Obs As String
' 0: numero,  1: nombre obra
Private a_Nv(2999, 1) As String, m_NvArea As Integer
Private m_PathArchivo As String

Private sql As String
Private SqlRsPaso As New ADODB.Recordset
Private a_Marcas(1999, 2) As String, m_MarcaIndice As Integer
Private Planilla As Object, hojaNombre As String
Private nombreArchivoXls As String, nHojas As Integer

Private Sub CbNv_Click()
Nv.Text = Left(CbNv.Text, 4)
End Sub
Private Sub Form_Load()

Set Dbm = OpenDatabase(mpro_file)

'Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
'RsNVc.Index = Nv_Index ' "Número"

nvListar Usuario.Nv_Activas

' Combo obra
CbNv.AddItem " "
For i = 1 To nvTotal
    a_Nv(i, 0) = aNv(i).Numero
    a_Nv(i, 1) = aNv(i).obra
    CbNv.AddItem Format(aNv(i).Numero, "0000") & " - " & aNv(i).obra
Next

Set RsPc = Dbm.OpenRecordset("Planos Cabecera")
RsPc.Index = "NV-Plano"

Set RsPd = Dbm.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Marca"

Set RsOTfd = Dbm.OpenRecordset("OT fab Detalle")
RsOTfd.Index = "NV-Plano-Marca"

Set RsITOfd = Dbm.OpenRecordset("ITO fab Detalle")
RsITOfd.Index = "NV-Plano-Marca"

'Set DbR = OpenDatabase(repo_file)
'Set RsR = DbR.OpenRecordset("importa_planos")

m_NvArea = 0

End Sub
Private Sub btnAbrir_Click()
' abre carpeta

Dim p As Integer

cd.DialogTitle = "Buscar Carpeta"
cd.Filter = "Archivos Excel (*.xls)|*.xls|Todos los Archivos (*.*)|*.*"

' busca ultima ruta
m_Path = "" 'GetSetting("scp", "planos", "ruta")
If m_Path = "" Then m_Path = "C:"

' si m_path no existe, muestra directorio actual
'cd.InitDir = m_Path

'm_Path = "C:" ' por mientras

m_Path = Directorio(m_Path)

cd.InitDir = m_Path
cd.ShowOpen

m_PathArchivo = cd.filename

If m_PathArchivo = "" Then
'    MsgBox "Debe escoger Archivo"
    Exit Sub
End If

'MsgBox Cd.filename ' viene con ruta clompleta
m_PathArchivo = cd.filename

' separa path y archivo
p = InStrLast(m_PathArchivo, "\")
If p > 0 Then

    m_Path = Left(m_PathArchivo, p)
'    m_Archivo = Mid(m_PathArchivo, p + 1)
    
    ' separa path y archivo
    p = InStrLast(m_PathArchivo, "\")
    If p > 0 Then
        m_Path = Left(m_PathArchivo, p)
        lblCarpeta.Caption = m_Path
    End If
    
    nombreArchivoXls = Mid(m_PathArchivo, p + 1)
    
    lblCarpeta.Caption = m_Path
    
    Set Planilla = GetObject(m_Path & nombreArchivoXls, "Excel.Sheet.8") 'assign sheet object as an OLE excel

    ' busca hojas
    'Debug.Print Planilla.Worksheets.Count
    nHojas = Planilla.Worksheets.Count
    For i = 1 To nHojas
       hojaNombre = Planilla.Worksheets(i).Name
    '      Debug.Print m_Hoja
       ComboHoja.AddItem hojaNombre
    Next

    
    ' guarda ultima ruta usada
    'SaveSetting "scp", "planos", "ruta", m_Path
    
End If

End Sub
Private Sub btnImportar_Click()

Dim p As Integer
Dim p_pg As Integer ' posicion primer guion
Dim p_ug As Integer ' posicion ultimo guion
Dim p_raya As Integer ' posicion de raya vertical |

If CbNv.ListIndex = -1 Then
    MsgBox "Debe Escoger Obra"
    CbNv.SetFocus
    Exit Sub
End If

If lblCarpeta.Caption = "" Then
    MsgBox "Debe Escoger CARPETA de ORIGEN"
    btnAbrir.SetFocus
    Exit Sub
End If

m_Nv = Left(CbNv.Text, 4)
m_CantidadPlanosNuevos = 0
'm_MarcaIndice = 0

'm_PathArchivo ='ruta y archivo

' aquui iba separa archivo

'Txt_Leer m_Path, NombreArchTXT
excelLeer m_Path, nombreArchivoXls

MsgBox "Piezas Importadas con Exito"

End Sub
Private Sub Nv_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub Nv_LostFocus()
Dim m_Nv As Integer
m_Nv = Val(Nv.Text)
If m_Nv = 0 Then Exit Sub
' busca nv en combo
i = 1
Do Until a_Nv(i, 0) = ""
    If Val(a_Nv(i, 0)) = m_Nv Then
        CbNv.ListIndex = i
        Exit Sub
    End If
    i = i + 1
Loop

MsgBox "NV no existe"
Nv.SetFocus

End Sub
Private Sub excelLeer(Path As String, Archivo As String)

Dim m_Designacion As String ' nuevo campo 29/10/07
Dim s_Paso As String, d_Paso As Double, Reg As String
Dim fi As Integer, m_Contenido As String, i As Integer
'Dim m_Contenido As String, i As Integer
Dim PlanoNuevo As Boolean, PlanoconMovimientos As Boolean
Dim editaplano As Boolean ' para 2335
Dim c As Integer, a_ini(19, 2) As Integer, ca As String, cs As String, ultima_pos As Integer, inicio As Integer, cuenta_caracteres As Integer
Dim RsPdConsulta As Recordset

editaplano = False

m_Obs = ""

m_PesoTotal = 0
m_SuperficieTotal = 0

Set Planilla = GetObject(Path & Archivo, "Excel.Sheet.8") 'assign sheet object as an OLE excel

' busca hojas
'Debug.Print Planilla.Worksheets.Count
nHojas = Planilla.Worksheets.Count
For i = 1 To nHojas
   hojaNombre = Planilla.Worksheets(i).Name
'      Debug.Print m_Hoja
   ComboHoja.AddItem hojaNombre
Next

If nHojas = 1 Then
   ComboHoja.ListIndex = 1
Else
   ComboHoja.ListIndex = 0
End If
   
hojaNombre = ComboHoja.Text


' lee excel

With Planilla.Worksheets(hojaNombre)

fi = 1
Do While True

    fi = fi + 1
    
    ' toma la NV elegida por el usuario
    'm_Nv = Val(Trim(.cells(fi, 1).Value))    ' A
    
    If .cells(fi, 1).Value = "" Then
        Exit Do
    End If
    
    m_Plano = Trim(.cells(fi, 2).Value)      ' B
    m_Rev = Trim(.cells(fi, 3).Value)        ' C
    m_Mar = Trim(.cells(fi, 4).Value)        ' D
    m_Can = Val(Trim(.cells(fi, 5).Value))   ' E
    m_Des = Trim(.cells(fi, 6).Value)        ' F
    m_KgU = m_CDbl(Trim(.cells(fi, 7).Value)) ' G
    m_m2U = m_CDbl(Trim(.cells(fi, 8).Value)) ' H
    m_Obs = Trim(.cells(fi, 9).Value)        ' I

    'Debug.Print "|" & fi & "|" & m_Nv & "|" & m_Plano & "|" & m_Rev & "|" & m_Mar & "|" & m_Can & "|" & m_Des & "|" & m_KgU & "|" & m_m2U & "|" & m_Obs & "|"
    
    m_PesoTotal = m_PesoTotal + m_Can * m_KgU
    m_SuperficieTotal = m_SuperficieTotal + m_Can * m_m2U
        
    '////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////
    ' busca si existe plano
    PlanoconMovimientos = False
    RsPc.Seek "=", m_Nv, m_NvArea, m_Plano
    If RsPc.NoMatch Then
    
        PlanoNuevo = True
        
        m_CantidadPlanosNuevos = m_CantidadPlanosNuevos + 1
        
    Else
    
        PlanoNuevo = False
        
        ' busca si plano se puede eliminar
        If Plano_Borrable(RsPd, m_Nv, m_NvArea, m_Plano) Then
            
            Plano_Eliminar Dbm, m_Nv, m_Plano, m_Mar
            
        Else
            
            ' existen ot fab para este plano
            PlanoconMovimientos = True
        
        End If
        
    End If
    '////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////
    
    If False Then
        ' primero cuenta numero de marcas ya grabadas del mismo plano
        'fi = 0 ' comentado para rodrigo nuñez, pernos, 08/05/12
        Set RsPdConsulta = Dbm.OpenRecordset("SELECT * FROM [planos detalle] WHERE nv=" & m_Nv & " AND plano='" & m_Plano & "' ORDER BY item")
        Do While Not RsPdConsulta.EOF
            If fi < RsPdConsulta!item Then
                fi = RsPdConsulta!item
            End If
            RsPdConsulta.MoveNext
        Loop
        RsPdConsulta.Close
    End If
    
    'Debug.Print Archivo & "|" & fi & "|" & Reg & "|" & m_Plano, m_Mar
    
    '/////////////////////////////////////////////////////////
    If PlanoconMovimientos Then
        
        ' busca marca antigua
        RsPd.Seek "=", m_Nv, m_NvArea, m_Plano, m_Mar
        If RsPd.NoMatch Then
            RsPd.AddNew
        Else
            
            ' si hay mov en esta marca
            If RsPd![OT fab] > 0 Then
                
                ' revisa variacion de peso
                If m_KgU = RsPd![Peso] Then
                    ' no hay drama con OTs ni ITOs (con los kilos)
                Else
                    ' marca con peso modificado
                    Marcas_Agregar a_Marcas, m_MarcaIndice, m_Plano, m_Mar
    '                    Exit Sub
                End If
                
                ' revisa disminucion de cantidad
                If m_Can - RsPd![Cantidad Total] >= 0 Then
                    ' no hay drama con OTs
                Else
                    ' marca disminuyó cantidad
                    Marcas_Agregar a_Marcas, m_MarcaIndice, m_Plano, m_Mar
                End If
            
            End If
            
            RsPd.Edit
            editaplano = True
            
        End If
        
    Else
        ' solo graba
        RsPd.AddNew
        
    End If
    
    ' graba linea de detalle de plano
    
    With RsPd
    
    '    .AddNew
    !Nv = m_Nv
    !NvArea = m_NvArea
    
    !Plano = UCase(m_Plano)
    m_Rev = Trim(UCase(m_Rev))
    !Rev = m_Rev
    
    'fi = fi + 1
    If Not editaplano Then
        !item = fi
    End If
     
    !Marca = m_Mar
    !Descripcion = m_Des
    
    ![Cantidad Total] = m_Can
    
    ![Peso] = m_KgU
    ![Superficie] = m_m2U
    ![Observaciones] = m_Obs ' Left(m_Designacion, 30)
    '    ![OT fab] = 0
    '    ![ITO fab] = 0
    '    ![ITO pyg] = 0
    '    ![GD] = 0
    !Chequeada = True ' False ??
    !densidad = m_Den
    .Update
    End With

'        Debug.Print m_Plano


    ' graba cabecera
    With RsPc
    
    If PlanoconMovimientos Then
        .Edit
    Else
        .AddNew
        !Nv = m_Nv
        !NvArea = m_NvArea
        !Editable = True
        !Plano = UCase(m_Plano)
    End If
    !Rev = UCase(m_Rev)
    ![Peso Total] = m_PesoTotal
    ![Superficie Total] = m_SuperficieTotal
    ![Fecha Modificacion] = Format(Now, Fecha_Format)
    !Observacion = m_Obs
    .Update
    End With
    
    m_PesoTotal = 0
    m_SuperficieTotal = 0

Loop

End With

End Sub
