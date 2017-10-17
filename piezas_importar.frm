VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form piezas_importar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importación Desglose de Piezas"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Nv 
      Height          =   300
      Left            =   720
      MaxLength       =   4
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin Crystal.CrystalReport cr 
      Left            =   3480
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton btnImportar 
      Caption         =   "&Importar"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton btnAbrir 
      Caption         =   "&Examinar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox CbNv 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3720
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Label lblExplicacion 
      Alignment       =   2  'Center
      Caption         =   "Opcion para importar archivos de METTAL en formato XSR"
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
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4095
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
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   4335
   End
   Begin VB.Label lblCarpeta 
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label lbl 
      Caption         =   "OBRA"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "piezas_importar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' importacion de archivos de planos desde formato MCR?
' archivo de texto plano, extension ...
' este es el formato de ancho fijo:
' en un solo archivo vienen varios planos-marcas
' version de julio 2012
' //////////////////////////////////////////////////////////////////////////////////
' LISTA DE CONJUNTOS
' M.C.R. ASESORIA EN INGENIERIA TECNICA
' CLIENTE:  METTAL
' MANDANTE :CAROZZI                                             FECHA:10.04.2012
' PROYECTO :AMPLIACION GALLETAS                                 HORA: 17:34:27
' -----------------------------------------------------------------------------------
'                                                 LARGO Area  Area    PESO     PESO
' Marca   N° CONJUNTO       Perfil                      Unit. Total   Unit.    TOTAL
' ....+....1....+....2....+....3....+....4....+....5....+....6....+....7....+....8....+....9
' -----------------------------------------------------------------------------------
' BA1      1 BARANDA        CAÑ1_1/4_SCH40         3187   2.5   2.5    61.8     61.8
' BA2      1 BARANDA        CAÑ1_1/4_SCH40         3457   2.7   2.7    65.0     65.0
' LE1      1 LIMON          C250X50X5              4104  20.8  20.8   301.9    301.9
' RL6      2 RIOSTRA_LATERALCJ50X50X3              1335   0.3   0.7     6.7     13.4
' -----------------------------------------------------------------------------------
'             TOTAL PARA  20   CONJUNTOS:                      92.8           1543.0
' -----------------------------------------------------------------------------------
'                                  END OF REPORT
'////////////////////////////////////////////////////////////////////////////////////

' cada vez que se importen piezas, (que no vienen con plano ni revision)
' debe buscar en planos detalle, para buscar si existe plano y revision, para grabarlo en tabla piezas
' solo agrega piezas nuevas, las antiguas las deja tal cual

Option Explicit
Private Dbm As Database ', RsNVc As Recordset
Private RsPc As Recordset, RsPd As Recordset, RsOTfd As Recordset, RsITOfd As Recordset

Private m_Path As String
Private m_Nv As Integer, m_Plano As String, m_Rev As String, m_Marca As String
' para detalle
Private m_Can As Integer, m_Mar As String, m_Des As String, m_KgU As Double, m_KgT As Double, m_m2U As Double, m_m2T As Double, m_Den As Integer

Private m_PesoTotal As Double, m_SuperficieTotal As Double

'Private a_Planos(999, 2) As String
' 0: plano
' 1: rev
' 2: N: nuevo
'    R: revision (no afecto a ninguna ot, ito, gd)

'Private a_Marcas(1999, 2) As String, m_MarcaIndice As Integer, m_Orden As Integer

' a_otf: 0: numero ot   1: plano   2: marca
'Private a_OTf(1999, 2) As Variant, m_OTfIndice As Integer
'Private a_ITOf(1999, 2) As Variant, m_ITOfIndice As Integer

Private m_CantidadPlanosNuevos As Integer
Private i As Integer, j As Integer, existe As Boolean, m_Obs As String
' 0: numero,  1: nombre obra
Private a_Nv(2999, 1) As String, m_NvArea As Integer
Private m_PathArchivo As String

Private sql As String
Private SqlRsPaso As New ADODB.Recordset
Private a_Marcas(1999, 2) As String, m_MarcaIndice As Integer
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
cd.Filter = "Archivos XSR (*.xsr)|*.xsr|Texto (*.txt)|*.txt|Todos los Archivos (*.*)|*.*"

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
    
    lblCarpeta.Caption = m_Path
    
    ' guarda ultima ruta usada
    SaveSetting "scp", "planos", "ruta", m_Path
    
End If

End Sub

Private Sub btnImportar_Click()

Dim p As Integer
Dim NombreArchTXT As String
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

' separa path y archivo
p = InStrLast(m_PathArchivo, "\")
If p > 0 Then
    m_Path = Left(m_PathArchivo, p)
    lblCarpeta.Caption = m_Path
End If

NombreArchTXT = Mid(m_PathArchivo, p + 1)

Txt_Leer m_Path, NombreArchTXT

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
Private Sub Txt_Leer(Path As String, Archivo As String)

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
   
' abre archivo
Open Path & Archivo For Input As #1

'fi = 0

Do While Not EOF(1)

    Line Input #1, Reg
    
    Reg = Trim(Reg)
    
    If Len(Reg) = 82 Then
       
        If lineaMettal(Reg) Then
            
            ' claudia delgado autorizo a sacarle espacios en blanco a plano y marca, nv 2216
            m_Plano = Replace(m_Plano, " ", "")
            m_Mar = Replace(m_Mar, " ", "")
            
            ' revisa si hay canbio de marcas , en cuyo caso envia email
            '        planos_revisiones_email Nv.Text, m_Plano, m_Rev, m_Mar
                
            ' agrega a tabla scp0.planos_detalle_revisiones
            Planos_Revisiones_Grabar Nv.Text, m_Plano, m_Rev, m_Mar, m_Can, m_Des, m_KgU, m_m2U
                
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
            
            fi = fi + 1
            If Not editaplano Then
                !item = fi
            End If
             
            !Marca = m_Mar
            !Descripcion = m_Des
            
            ![Cantidad Total] = m_Can
            
            ![Peso] = m_KgU
            ![Superficie] = m_m2U
            ![Observaciones] = Left(m_Designacion, 30)
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
        
        End If ' lineaMettal=true
        
    End If ' if len(reg)=82
        
Loop

Close #1

m_Obs = Left(m_Obs, 50)

End Sub
Private Function lineaMettal(Reg As String) As Boolean
Dim prefijo As String
prefijo = "001-"
' linea exclusiva para METTAL
' si linea no es valida, devuelve false
' si linea es valida devuelve true, y puebla arreglo con campos
' ej NV 2993 y NV 3011
' -----------------------------------------------------------------------------------
'                                                 LARGO Area  Area    PESO     PESO
' Marca   N° CONJUNTO       Perfil                      Unit. Total   Unit.    TOTAL
' -----------------------------------------------------------------------------------
' ....+....1....+....2....+....3....+....4....+....5....+....6....+....7....+....8....+....9
' BA1      1 BARANDA        CAÑ1_1/4_SCH40         3187   2.5   2.5    61.8     61.8
' BA2      1 BARANDA        CAÑ1_1/4_SCH40         3457   2.7   2.7    65.0     65.0
' LE1      1 LIMON          C250X50X5              4104  20.8  20.8   301.9    301.9
' RL6      2 RIOSTRA_LATERALCJ50X50X3              1335   0.3   0.7     6.7     13.4

Dim d_Paso As Double

lineaMettal = False
If Len(Reg) <> 82 Then
    Exit Function
End If

'Private m_Pla, m_Can As Integer, m_Mar As String, m_Des As String, m_KgU As Double, m_KgT As Double, m_m2U As Double, m_m2T As Double, m_Den As Integer
On Error GoTo DatumError
m_Mar = Trim(Mid(Reg, 1, 8))
m_Plano = prefijo & m_Mar
m_Can = Mid(Reg, 9, 2)
m_Des = Trim(Mid(Reg, 12, 15))
d_Paso = m_CDbl(Trim(Mid(Reg, 54, 6)))
m_m2U = d_Paso
d_Paso = m_CDbl(Trim(Mid(Reg, 60, 6)))
m_m2T = d_Paso
d_Paso = m_CDbl(Trim(Mid(Reg, 66, 8)))
m_KgU = d_Paso
d_Paso = m_CDbl(Trim(Mid(Reg, 74, 9)))
m_KgT = d_Paso
On Error GoTo 0

lineaMettal = True

Exit Function

DatumError:

End Function
