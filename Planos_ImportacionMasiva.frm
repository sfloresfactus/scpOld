VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Planos_ImportacionMasiva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importación Masiva de Planos"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Nv 
      Height          =   300
      Left            =   720
      MaxLength       =   4
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin Crystal.CrystalReport cr 
      Left            =   3480
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton btnImportar 
      Caption         =   "&Importar"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton btnAbrir 
      Caption         =   "Abrir Carpeta"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox CbNv 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3720
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
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
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Label lblCarpeta 
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label lbl 
      Caption         =   "OBRA"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Planos_ImportacionMasiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'####################################################################
'
' importa planos desde archivos planos a scp
' el usuario debe escoger carpeta (archivo .txt) desde donde importar
' todas los planos (1 por cada archivo plano (ascii) ) de esa carpeta
'
' si el plano es nuevo: importa plano sin problemas
'
' si plano ya existe
' => si plano no afecta a ninguna ot, ito, gd => se informa de este plano (solo es revision)
'

' rutina nueva septiembre 2010
' si se importa plano masivamente, ademas debe buscar si existen piezas(sql.scp0.piezas) asociadas
' para grabarles el plano y revision, si es que no la tienen
' tanto esta rutina como el importar piezas suponen que cada plano cuenta con una sola marca

' rutina nueva septiembre 2010
' incorpora nuevos planos a archivo que registra todos los detalles
' de revisiones de planos
' sqlserver.controlserver\bdms\scp0.planos_detalle_revisiones

'####################################################################
Option Explicit
Private Dbm As Database, RsNVc As Recordset, RsPc As Recordset, RsPd As Recordset
Private DbR As Database, RsR As Recordset
Private RsOTfd As Recordset, RsITOfd As Recordset
Private m_Path As String
Private m_Nv As Integer, m_Plano As String, m_Rev As String, m_Marca As String
Private m_PesoTotal As Double, m_SuperficieTotal As Double

'Private a_Planos(999, 2) As String
' 0: plano
' 1: rev
' 2: N: nuevo
'    R: revision (no afecto a ninguna ot, ito, gd)

Private a_Marcas(1999, 2) As String, m_MarcaIndice As Integer, m_Orden As Integer

' a_otf: 0: numero ot   1: plano   2: marca
Private a_OTf(1999, 2) As Variant, m_OTfIndice As Integer

Private a_ITOf(1999, 2) As Variant, m_ITOfIndice As Integer

Private m_CantidadPlanosNuevos As Integer
Private i As Integer, j As Integer, existe As Boolean, m_Obs As String
' 0: numero,  1: nombre obra
Private a_Nv(2999, 1) As String, m_NvArea As Integer
Private NV2335 As Boolean
Private a_NV2335(3, 200) As String, Total_Marcas_NV2335 As Integer
Private Sub CbNv_Click()
Nv.Text = Left(CbNv.Text, 4)
End Sub
Private Sub Form_Load()

NV2335 = False

If NV2335 Then
    nv2335_Poblar "completo"
End If

Set Dbm = OpenDatabase(mpro_file)

nvListar Usuario.Nv_Activas

' Combo obra
CbNv.AddItem " "
For i = 1 To nvTotal
    a_Nv(i, 0) = aNv(i).Numero
    a_Nv(i, 1) = aNv(i).Obra
    CbNv.AddItem Format(aNv(i).Numero, "0000") & " - " & aNv(i).Obra
Next

' old additem
'With RsNVc
'Do While Not .EOF
'    CbNv.AddItem !Número & " " & !Obra
'    .MoveNext
'Loop
'End With

Set RsPc = Dbm.OpenRecordset("Planos Cabecera")
RsPc.Index = "NV-Plano"

Set RsPd = Dbm.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Marca"

Set RsOTfd = Dbm.OpenRecordset("OT fab Detalle")
RsOTfd.Index = "NV-Plano-Marca"

Set RsITOfd = Dbm.OpenRecordset("ITO fab Detalle")
RsITOfd.Index = "NV-Plano-Marca"

Set DbR = OpenDatabase(repo_file)
Set RsR = DbR.OpenRecordset("importa_planos")

m_NvArea = 0

End Sub
Private Sub btnAbrir_Click()
' abre carpeta

Dim m_PathArchivo As String, p As Integer

cd.DialogTitle = "Buscar Carpeta"
cd.Filter = "Texto (*.txt)|*.txt|Microsoft Excel (*.xls)|*.xls|Todos los Archivos (*.*)|*.*"

' busca ultima ruta
'm_Path = GetSetting("scp", "planos", "ruta")
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

Dim NombreArchTXT As String
Dim p_pg As Integer ' posicion primer guion
Dim p_ug As Integer ' posicion ultimo guion
Dim p_raya As Integer ' posicion de raya vertical |
Dim m_Plano_NV2335 As String, m_Marca_NV2335 As String

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
m_MarcaIndice = 0

DbR.Execute "DELETE * FROM [importa_planos]"

' busca archivos planos que hay en la carpeta
NombreArchTXT = Dir(m_Path & "*.TXT")

i = 0
Do

'    Debug.Print m_NombreArchTXT
    lblArchivo.Caption = NombreArchTXT
    Me.Refresh
    
    p_pg = InStr(NombreArchTXT, "-")
    p_ug = InStrLast(NombreArchTXT, "-")
    
    If p_pg = 0 Or p_pg = p_ug Then
        ' no hay guion o hay un solo guion
        m_Plano = Left(NombreArchTXT, Len(NombreArchTXT) - 4)
'        m_Rev = Mid(NombreArchTXT, p + 1, 1) ' ojo , asumo que revision es siempre de una letra
    End If
    If p_pg < p_ug Then
        ' hay mas de un guion
        m_Plano = Left(NombreArchTXT, p_ug - 1)
    End If

    m_Plano_NV2335 = ""
    m_Marca_NV2335 = ""
    If NV2335 Then
        m_Plano = NV2335_Equivalencia(m_Plano, m_Marca)
        If m_Plano = "" Then
            ' plano nuevo
        Else
            p_raya = InStr(1, m_Plano, "|")
            m_Marca_NV2335 = Mid(m_Plano, p_raya + 1)
            m_Plano_NV2335 = Left(m_Plano, p_raya - 1)
        End If
    End If

'    Txt_Leer_Separador m_Path, NombreArchTXT, m_Plano_NV2335, m_Marca_NV2335
    Txt_Leer_Separador m_Path, NombreArchTXT
    
    i = i + 1

    NombreArchTXT = Dir()
'    NombreArchTXT = "" ' ojo ojo

Loop Until NombreArchTXT = ""

'//////////////////////////////////////////////
' ordena arreglo de marcas
For i = 1 To m_MarcaIndice - 1
For j = i + 1 To m_MarcaIndice

    ' compra planos
    Select Case a_Marcas(i, 0)
    Case Is > a_Marcas(j, 0)
    
        ' cambio
        Marcas_Swap i, j
        
    Case a_Marcas(j, 0)
        
        ' compara marcas
        If a_Marcas(i, 1) > a_Marcas(j, 1) Then
        
            ' cambio
            Marcas_Swap i, j
            
        End If
    
    Case Else
        ' nada
    End Select
    
Next
Next
'//////////////////////////////////////////////
' busca otfab afectadas
m_OTfIndice = 0
With RsOTfd

For i = 1 To m_MarcaIndice

    m_Plano = a_Marcas(i, 0)
    m_Marca = a_Marcas(i, 1)
    .Seek "=", m_Nv, m_NvArea, m_Plano, m_Marca
    
    If Not .NoMatch Then
    
        Do While Not .EOF
        
            If !Nv <> m_Nv Or !Plano <> m_Plano Or !Marca <> m_Marca Then Exit Do
            
            ' busca si existe plano-marca
            existe = False
            For j = 1 To m_OTfIndice
            
                If a_OTf(j, 1) = m_Plano And a_OTf(j, 2) = m_Marca Then
                
                    existe = True
                    j = 999
                    
                End If
                
            Next
            
            If Not existe Then
            
                m_OTfIndice = m_OTfIndice + 1
                a_OTf(m_OTfIndice, 0) = !Numero
                a_OTf(m_OTfIndice, 1) = !Plano
                a_OTf(m_OTfIndice, 2) = !Marca
            
            End If
            
            .MoveNext
            
        Loop
    
    End If
    
Next
End With
'//////////////////////////////////////////////
' ordena arreglo de OTf
For i = 1 To m_OTfIndice - 1
For j = i + 1 To m_OTfIndice

    Select Case a_OTf(i, 0)
    Case Is > a_OTf(j, 0)
    
        ' cambio
        OTf_Swap i, j
        
    Case a_OTf(j, 0)
        
        ' compara planos
        If a_OTf(i, 1) > a_OTf(j, 1) Then
        
            ' cambio
            OTf_Swap i, j
            
        End If
    
    Case Else
        ' nada
    End Select
    
Next
Next
'//////////////////////////////////////////////
' busca ITOfab afectadas
m_ITOfIndice = 0
With RsITOfd

For i = 1 To m_MarcaIndice

    m_Plano = a_Marcas(i, 0)
    m_Marca = a_Marcas(i, 1)
    .Seek "=", m_Nv, m_NvArea, m_Plano, m_Marca
    
    If Not .NoMatch Then
    
        Do While Not .EOF
        
            If !Nv <> m_Nv Or !Plano <> m_Plano Or !Marca <> m_Marca Then Exit Do
            
            ' busca si existe plano-marca
            existe = False
            For j = 1 To m_ITOfIndice
            
                If a_ITOf(j, 1) = m_Plano And a_ITOf(j, 2) = m_Marca Then
                
                    existe = True
                    j = 999
                    
                End If
                
            Next
            
            If Not existe Then
            
                m_ITOfIndice = m_ITOfIndice + 1
                a_ITOf(m_ITOfIndice, 0) = !Numero
                a_ITOf(m_ITOfIndice, 1) = !Plano
                a_ITOf(m_ITOfIndice, 2) = !Marca
            
            End If
            
            .MoveNext
            
        Loop
    
    End If
    
Next
End With
'//////////////////////////////////////////////
' ordena arreglo de ITOf
For i = 1 To m_ITOfIndice - 1
For j = i + 1 To m_ITOfIndice

    Select Case a_ITOf(i, 0)
    Case Is > a_ITOf(j, 0)
    
        ' cambio
        ITOf_Swap i, j
        
    Case a_ITOf(j, 0)
        
        ' compara marcas
        If a_ITOf(i, 1) > a_ITOf(j, 1) Then
        
            ' cambio
            ITOf_Swap i, j
            
        End If
    
    Case Else
        ' nada
    End Select
    
Next
Next
'//////////////////////////////////////////////

' graba reporte

' graba linea de planos nuevos
With RsR
.AddNew
!orden = 1
Select Case m_CantidadPlanosNuevos
Case 0
    !Plano = "NO hay Planos Nuevos"
Case 1
    !Plano = "1 Plano Nuevo"
Case Else
    !Plano = m_CantidadPlanosNuevos & " Planos Nuevos"
End Select
.Update

' linea en blanco
.AddNew
!orden = 1
.Update

For i = 1 To m_OTfIndice

    .AddNew
    !orden = 2
    !td = "OTf"
    !Nd = a_OTf(i, 0)
    !Plano = a_OTf(i, 1)
    !Marca = a_OTf(i, 2)
    .Update

Next

For i = 1 To m_ITOfIndice

    .AddNew
    !orden = 3
    !td = "ITO" 'f"
    !Nd = a_ITOf(i, 0)
    !Plano = a_ITOf(i, 1)
    !Marca = a_ITOf(i, 2)
    .Update

Next

End With

'lblArchivo.Caption = i & " Archivos Importados"
lblArchivo.Caption = "Fin Importación"
Me.Refresh

Piezas_PlanoGrabar Nv.Text

cr.WindowState = crptMaximized
cr.WindowTitle = "Importacion de Planos"
cr.DataFiles(0) = repo_file & ".MDB"
cr.Formulas(0) = "RAZON=""" & Empresa.Razon & """"
cr.Formulas(1) = "TITULO=""" & "Importación de Planos NV " & m_Nv & """"
cr.ReportSource = crptReport
cr.ReportFileName = Drive_Server & Path_Rpt & "planos_importar.rpt"
cr.Action = 1

End Sub

Private Sub Txt_Leer_Tabxxx(Path As String, Archivo As String)
' importa datos desde archivo plano
' los campos del archivo plano estan separados por tabuladores, es decir ascii=9
Dim m_Can As Integer, m_Mar As String, m_Des As String, m_KgU As Double, m_KgT As Double, m_m2U As Double, m_m2T As Double
Dim s_Paso As String, d_Paso As Double, Reg As String
Dim largoreg As Integer, m_col As Integer, largocont As Integer
Dim p1 As Integer, fi As Integer, m_Contenido As String, i As Integer
Dim PlanoNuevo As Boolean, PlanoconMovimientos As Boolean

'If Not Archivo_Existe(Path, Archivo) Then
'    MsgBox "No existe archivo " & vbLf & Path & Archivo
'    Exit Sub
'End If

m_PesoTotal = 0
m_SuperficieTotal = 0

' busca si existe plano
PlanoconMovimientos = False
RsPc.Seek "=", m_Nv, m_Plano
If RsPc.NoMatch Then

    PlanoNuevo = True
    
    m_CantidadPlanosNuevos = m_CantidadPlanosNuevos + 1
    
Else

    PlanoNuevo = False
    
    ' busca si plano se puede eliminar
    If Plano_Borrable(RsPd, m_Nv, m_NvArea, m_Plano) Then
        
        Plano_Eliminar Dbm, m_Nv, m_Plano, m_Marca
        
    Else
        
        ' existen ot fab para este plano
        PlanoconMovimientos = True
    
    End If
    
End If

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
    
    m_PesoTotal = m_PesoTotal + m_Can * m_KgU
    m_SuperficieTotal = m_SuperficieTotal + m_Can * m_m2U
    
    '/////////////////////////////////////////////////////////
    If PlanoconMovimientos Then
        
        ' busca marca antigua
        RsPd.Seek "=", m_Nv, m_Plano, m_Mar
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
            
        End If
        
    Else
        ' solo graba
        RsPd.AddNew
        
    End If
    
    ' graba linea de detalle de plano
    
    With RsPd
    
'    .AddNew
    !Nv = m_Nv
    !Plano = UCase(m_Plano)
    !Rev = UCase(m_Rev)
    fi = fi + 1
    !item = fi
    !Marca = m_Mar
    !descripción = m_Des

    ![Cantidad Total] = m_Can
    
    ![Peso] = m_KgU
    ![Superficie] = m_m2U
    ![Observaciones] = ""
'    ![OT fab] = 0
'    ![ITO fab] = 0
'    ![ITO pyg] = 0
'    ![GD] = 0
    !Chequeada = True ' False ??
    .Update
        
    End With
    
Loop

Close #1

' graba cabecera
With RsPc

If PlanoconMovimientos Then
    .Edit
Else
    .AddNew
    !Nv = m_Nv
    !Editable = True
    !Plano = UCase(m_Plano)
End If
!Rev = UCase(m_Rev)
![Peso Total] = m_PesoTotal
![Superficie Total] = m_SuperficieTotal
![Fecha Modificación] = Format(Now, Fecha_Format)
!Observación = m_Obs
.Update
End With

End Sub
Private Sub oldTxt_Leer_Space(Path As String, Archivo As String)

' importa 1 solo plano
' importa datos desde archivo plano
' los campos del archivo plano tiene largo y posiciones fijas

Dim m_Can As Integer, m_Mar As String, m_Des As String, m_KgU As Double, m_KgT As Double, m_m2U As Double, m_m2T As Double
Dim m_Designacion As String ' nuevo campo 29/10)07
Dim s_Paso As String, d_Paso As Double, Reg As String
Dim fi As Integer, m_Contenido As String, i As Integer
Dim PlanoNuevo As Boolean, PlanoconMovimientos As Boolean

'Dim c As Integer, a_ini(9, 1) As Integer, ca As String, ch As String, ultima_pos As Integer, inicio As Integer
Dim c As Integer, a_ini(19, 1) As Integer, ca As String, ch As String, ultima_pos As Integer, inicio As Integer
' a_ini, 0: posicion inicial, 1: posicion final

m_Obs = ""

m_PesoTotal = 0
m_SuperficieTotal = 0

' abre archivo
Open Path & Archivo For Input As #1

fi = 0

Do While Not EOF(1)

    Line Input #1, Reg
    
    Reg = Trim(Reg)
    
    ' busca en archivo de texto
    ' si cada linea esta separado por tab o espacios
    
    ch = "" ' este va a ser el caracter separador de campos
    ' busca tab ( ascii 9 )
    i = InStr(1, Reg, Chr(9))
    If i > 0 Then
        ch = Chr(9)
    Else
        ' busca espacio ( ascii 32 )
        i = InStr(1, Reg, Chr(32))
        If i > 0 Then
            ch = Chr(32)
        End If
    End If
    
    ' caracter palito empieza a regir desde 13/04/09
    ch = Chr(124) ' | raya vertical
    
'    If Left(Reg, 10) = "021A-CI500" Then
'        MsgBox ""
'    End If
    
    ' el registro(o linea de texto) debe ser de a lo menos 75 caracteres
    If Len(Reg) > 10 Then ' 75
    
    If False Then
    
        m_Can = Mid(Reg, 13, 4)
        m_Mar = Trim(Mid(Reg, 17, 15))
        m_Des = Trim(Mid(Reg, 32, 15))
        
        m_Obs = m_Obs & m_Des
        
        d_Paso = m_CDbl(Trim(Mid(Reg, 47, 10)))
        m_KgU = d_Paso
        
        d_Paso = m_CDbl(Trim(Mid(Reg, 57, 10)))
        m_KgT = d_Paso
        
        d_Paso = m_CDbl(Trim(Mid(Reg, 67, 10)))
        m_m2U = d_Paso
        
        m_PesoTotal = m_PesoTotal + m_Can * m_KgU
        m_SuperficieTotal = m_SuperficieTotal + m_Can * m_m2U
        
    Else
    
        ca = ch ' caracter anterior
        i = 0
'        Printer.Print Reg
'        Printer.Print "....+....1....+....2....+....3....+....4....+....5....+....6....+....7....+....8" '....+....9....+....0"

        ' formato actualizado al 29/10/07
        ' Plano
        ' rev
        ' cantidad
        ' marca
        ' descripcion
        ' kg unitario
        ' kg total
        ' m2 unitario
        ' m2 total
        ' designacion

        '//////////////////////////////////////
        ' parche para plano pegado con revision
        ' nv:1289
        ' plano: 021A-CI500
        ' rev : A
        ' viene con "021A-CI500A"
        '            12345678901

'If m_Plano = "021A-CI500" Then
'If m_Plano = "CS1-004-F104" Then
'MsgBox ""
'End If

        ca = ch ' caracter anterior
        c = InStr(1, Reg, " ") ' busca primer espacio en blanco

        If c <= planoLargo + 1 Then ' 18/11/08 por cambio en formato, se alargo el largo de plano de 20 a 25 caracteres
            ' nombre de plano menor a veintiseis caracteres
            i = 0
            inicio = 1
        Else
'            Debug.Print Reg
            ' nombre de plano es de diez caracteres
            i = 1
'            inicio = 11
'            inicio = 13
            inicio = 21
            a_ini(i, 0) = 1
'            a_ini(i, 1) = 11
'            a_ini(i, 1) = 13
'            a_ini(i, 1) = 21
            a_ini(i, 1) = planoLargo + 1
        End If
        '//////////////////////////////////////
        ' busca espacios en blanco, para ver separacion de campos
        For c = inicio To Len(Reg)
        
            If i > 10 Then
                m_Designacion = Trim(Mid(Reg, a_ini(10, 0)))
                Exit For
            End If
            
            If Mid(Reg, c, 1) <> ch Then
                If ca = ch Then
                    i = i + 1
                    a_ini(i, 0) = c
'                    Printer.Print c; ",";
                End If
                ultima_pos = c
                
                ' verifica que largo maximo de nombre de plano sea diez
'                If i = 1 Then
'                    ' campo 1 es el nombre del plano
'                    If ultima_pos > 11 Then
'                        ultima_pos = 11
'                        GoTo Ultima
'                    End If
'                End If
                
            Else
'                a_ini(i - 1, 1) = ultima_pos
'                Printer.Print "up:"; ultima_pos;
'                Printer.Print ultima_pos - a_ini(i, 0) + 1; "/"; ' largo
'Ultima:
                a_ini(i, 1) = ultima_pos - a_ini(i, 0) + 1 ' largo
            End If
            ca = Mid(Reg, c, 1)
        Next
        
'        Printer.Print "up:"; ultima_pos
'        Printer.Print ultima_pos - a_ini(i, 0) + 1; "/"  ' largo
'        Printer.Print " "
        a_ini(i, 1) = ultima_pos - a_ini(i, 0) + 1 ' largo
    
'        Debug.Print Reg
'        For i = 1 To 9
'            Debug.Print a_ini(i, 0); a_ini(i, 1);
'        Next
'        Debug.Print "-"
        
'        Debug.Print "....+....1....+....2....+....3....+....4....+....5....+....6....+....7....+....8" '....+....9....+....0"
'        Debug.Print Reg

If True Then
'If False Then

        m_Plano = Mid(Reg, a_ini(1, 0), a_ini(1, 1)) ' 01/09/08
        m_Rev = Mid(Reg, a_ini(2, 0), a_ini(2, 1)) ' 01/09/08
        m_Can = Val(Mid(Reg, a_ini(3, 0), a_ini(3, 1)))
        
        If m_Can = 0 Then
        
            ' si cantidad es 0, entonces viene malo el formato, ej: plano viene pegado con revision
            ' NV 2127, 2129, CBI
            m_Rev = Right(m_Plano, 1)
            m_Plano = Left(m_Plano, Len(m_Plano) - 1)

            m_Can = Val(Mid(Reg, a_ini(2, 0), a_ini(2, 1)))
            
            m_Mar = Trim(Mid(Reg, a_ini(3, 0), a_ini(3, 1)))
            m_Des = Trim(Mid(Reg, a_ini(4, 0), a_ini(4, 1)))
            
            m_Obs = m_Obs & m_Des
            
            d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(5, 0), a_ini(5, 1))))
            m_KgU = d_Paso
            
            d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(6, 0), a_ini(6, 1))))
            m_KgT = d_Paso
            
            d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(7, 0), a_ini(7, 1))))
            m_m2U = d_Paso
            
        Else
        
            m_Mar = Trim(Mid(Reg, a_ini(4, 0), a_ini(4, 1)))
            m_Des = Trim(Mid(Reg, a_ini(5, 0), a_ini(5, 1)))
            
            m_Obs = m_Obs & m_Des
            
            d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(6, 0), a_ini(6, 1))))
            m_KgU = d_Paso
            
            d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(7, 0), a_ini(7, 1))))
            m_KgT = d_Paso
            
            d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(8, 0), a_ini(8, 1))))
            m_m2U = d_Paso
                        
        End If
        
Else

' rutina solo para obra 2130, porque algunos planos y marcas vienen con espacios entremedio, ej:
' PLANO-----------R C   MARCA------
' 7100-DETAIL F1-1A 1   DETAIL F1-1    SOPORTE        200,8     200,8     5,42      5,4

    m_Plano = Mid(Reg, a_ini(1, 0), a_ini(1, 1)) & " " & Mid(Reg, a_ini(2, 0), a_ini(2, 1))
    m_Rev = Right(m_Plano, 1)
    m_Plano = Left(m_Plano, Len(m_Plano) - 1)
    m_Plano = Replace(m_Plano, " ", "")
'    If m_Plano = "7399-SS-120-0001(A810)-A" Then
'    MsgBox ""
'    End If
    m_Can = Val(Mid(Reg, a_ini(3, 0), a_ini(3, 1)))
    m_Mar = Mid(Reg, a_ini(4, 0), a_ini(4, 1)) & " " & Mid(Reg, a_ini(5, 0), a_ini(5, 1))
    m_Mar = Replace(m_Mar, " ", "")
    
    ' solo para nv 2216
    ' marca viene pegada con descricion
    Dim pos As Integer, paso As String
    paso = Right(m_Plano, 5)
    pos = InStr(1, m_Mar, paso)
    m_Des = Mid(m_Mar, pos + 5)
    m_Mar = Left(m_Mar, pos + 4)
'    m_Des = Trim(Mid(Reg, a_ini(6, 0), a_ini(6, 1)))
    '//////////////////
'    m_Obs = m_Obs & m_Des
    If False Then
        d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(7, 0), a_ini(7, 1))))
        m_KgU = d_Paso
        d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(8, 0), a_ini(8, 1))))
        m_KgT = d_Paso
        d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(9, 0), a_ini(9, 1))))
        m_m2U = d_Paso
    Else
        d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(7, 0), a_ini(7, 1))))
        m_KgU = d_Paso
        d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(8, 0), a_ini(8, 1))))
        m_KgT = d_Paso
        d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(9, 0), a_ini(9, 1))))
        m_m2U = d_Paso
    End If
End If
'Debug.Print Archivo, m_Plano
        
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
        
        Plano_Eliminar Dbm, m_Nv, m_Plano, m_Marca
        
    Else
        
        ' existen ot fab para este plano
        PlanoconMovimientos = True
    
    End If
    
End If
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
        
    
    End If
        
'    Debug.Print m_Plano
        
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
        
        ' claudia delgado autorizo a sacarle espacios en blanco a plano y marca, nv 2216
        m_Plano = Replace(m_Plano, " ", "")
        
        !Plano = UCase(m_Plano)
        m_Rev = Trim(UCase(m_Rev))
        !Rev = m_Rev
        fi = fi + 1
        !item = fi
        m_Mar = Replace(m_Mar, " ", "")
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
'        .Update
'Debug.Print m_Plano
        End With
    
    End If
    
Loop

Close #1

m_Obs = Left(m_Obs, 50)

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

End Sub
Private Sub Txt_Leer_Separador(Path As String, Archivo As String)
'Private Sub Txt_Leer_Separador(Path As String, archivo As String, m_Plano_NV2335 As String, m_Marca_NV2335 As String)

Dim xxx As Integer

Dim m_Can As Integer, m_Mar As String, m_Des As String, m_KgU As Double, m_KgT As Double, m_m2U As Double, m_m2T As Double, m_Den As Integer
Dim m_Designacion As String ' nuevo campo 29/10)07
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
    
    If Len(Reg) > 20 Then ' registro de al menos 20 caracteres
       
        ' formato
        ' plano | rev | cantidad | marca | descripcion | kg unitario | kg total | m2 unitario | m2 total | designacion
        
        ' caracter separador
        cs = Chr(124) ' | raya vertical
        
        ' cuenta numero de caracteres separadores
        cuenta_caracteres = CharCount(Reg, cs)
        If cuenta_caracteres < 9 Then
        
            ' viene separado por espacios
            i = InStr(1, Reg, Chr(32))
            If i > 0 Then
                cs = Chr(32)
            End If
            
            ca = cs ' caracter anterior
            c = InStr(1, Reg, " ")
            i = 0
            inicio = 1
            
            For c = inicio To Len(Reg)
        
                If i > 10 Then
                    m_Designacion = Trim(Mid(Reg, a_ini(10, 0)))
                    Exit For
                End If
                
                If Mid(Reg, c, 1) <> cs Then
                    If ca = cs Then
                        i = i + 1
                        a_ini(i, 0) = c
                    End If
                    ultima_pos = c
                Else
                    a_ini(i, 1) = ultima_pos - a_ini(i, 0) + 1 ' largo
                End If
                ca = Mid(Reg, c, 1)
            Next
            
            m_Plano = Mid(Reg, a_ini(1, 0), a_ini(1, 1)) ' 01/09/08
            m_Rev = Mid(Reg, a_ini(2, 0), a_ini(2, 1)) ' 01/09/08
            m_Can = Val(Mid(Reg, a_ini(3, 0), a_ini(3, 1)))
            
            If m_Can > 0 Then
            
                m_Mar = Trim(Mid(Reg, a_ini(4, 0), a_ini(4, 1)))
                m_Des = Trim(Mid(Reg, a_ini(5, 0), a_ini(5, 1)))
                
                m_Obs = m_Obs & m_Des
                
                d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(6, 0), a_ini(6, 1))))
                m_KgU = d_Paso
                
                d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(7, 0), a_ini(7, 1))))
                m_KgT = d_Paso
                
                d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(8, 0), a_ini(8, 1))))
                m_m2U = d_Paso
            
                
            End If

        Else ' viene con separadores |
                    
            ca = cs ' caracter anterior
            i = 0
            inicio = 1
            '//////////////////////////////////////
            For c = inicio To Len(Reg)
            
                If Mid(Reg, c, 1) = cs Then
                
                    i = i + 1
                    a_ini(i, 1) = c
                    
                End If
                
            Next
            
            m_Plano = Trim(Mid(Reg, 1, a_ini(1, 1) - 1))
            m_Rev = Trim(Mid(Reg, a_ini(1, 1) + 1, a_ini(2, 1) - a_ini(1, 1) - 1))
            m_Can = Val(Mid(Reg, a_ini(2, 1) + 1, a_ini(3, 1) - a_ini(2, 1) - 1))
            
            m_Mar = Trim(Mid(Reg, a_ini(3, 1) + 1, a_ini(4, 1) - a_ini(3, 1) - 1))
            m_Des = Trim(Mid(Reg, a_ini(4, 1) + 1, a_ini(5, 1) - a_ini(4, 1) - 1))
            
            m_Obs = m_Obs & m_Des
            
            d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(5, 1) + 1, a_ini(6, 1) - a_ini(5, 1) - 1)))
            m_KgU = d_Paso
            
            d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(6, 1) + 1, a_ini(7, 1) - a_ini(6, 1) - 1)))
            m_KgT = d_Paso
            
            d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(7, 1) + 1, a_ini(8, 1) - a_ini(7, 1) - 1)))
            m_m2U = d_Paso
            
            If i = 10 Then ' si existe decima raya vertical, es formato de archivo 16-12-09
                d_Paso = m_CDbl(Trim(Mid(Reg, a_ini(10, 1) + 1)))
                m_Den = d_Paso
            Else
                m_Den = 0
            End If
            
        End If
            
'If m_Plano = "1-F4" Then
'MsgBox ""
'End If

'If m_Marca_NV2335 <> "" Then
'    m_Plano = m_Plano_NV2335
'    m_Mar = m_Marca_NV2335
'End If

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
'        fi = 0 ' comentado para rodrigo nuñez, pernos, 08/05/12
        Set RsPdConsulta = Dbm.OpenRecordset("SELECT * FROM [planos detalle] WHERE nv=" & m_Nv & " AND plano='" & m_Plano & "' ORDER BY item")
        Do While Not RsPdConsulta.EOF
            If fi < RsPdConsulta!item Then
                fi = RsPdConsulta!item
            End If
            RsPdConsulta.MoveNext
        Loop
        RsPdConsulta.Close

'        Debug.Print Archivo & "|" & fi & "|" & Reg & "|" & m_Plano, m_Mar

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
        
If Len(m_Plano) > planoLargo Then
    xxx = xxx + 1
    Debug.Print xxx & "|" & Reg
    MsgBox ""
Else

If m_Plano = "206-3140-1-49046822-23-F532" Then
'    MsgBox ""
End If


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
'        Debug.Print m_Plano
        End With
        
End If

    End If
    
        
Loop

Close #1

m_Obs = Left(m_Obs, 50)

If Len(m_Plano) > planoLargo Then

Else

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

End If

End Sub
'Private Sub Marcas_Agregar(Marca As String, Cantidad As Integer, PesoU As Double)
Public Sub Marcas_Swap(i As Integer, j As Integer)
' cambia marcas en arreglo
Dim m_Paso As String

' plano
m_Paso = a_Marcas(i, 0)
a_Marcas(i, 0) = a_Marcas(j, 0)
a_Marcas(j, 0) = m_Paso

' marca
m_Paso = a_Marcas(i, 1)
a_Marcas(i, 1) = a_Marcas(j, 1)
a_Marcas(j, 1) = m_Paso

End Sub
Private Sub OTf_Swap(i As Integer, j As Integer)
' cambia numeros,planos,marcas en arreglo
Dim m_Paso As String, k As Integer

For k = 0 To 2
    m_Paso = a_OTf(i, k)
    a_OTf(i, k) = a_OTf(j, k)
    a_OTf(j, k) = m_Paso
Next

End Sub
Private Sub ITOf_Swap(i As Integer, j As Integer)
' cambia marcas en arreglo
Dim m_Paso As String, k As Integer

For k = 0 To 2
    m_Paso = a_ITOf(i, k)
    a_ITOf(i, k) = a_ITOf(j, k)
    a_ITOf(j, k) = m_Paso
Next

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
Private Sub nv2335_Poblar(Hoja As String)

Dim Planilla As Object, fi As Integer
Dim m_Plano_IPE As String, m_Marca_IPE As String
Dim m_Plano_Cliente As String, m_Marca_Cliente As String

' puebla arreglo con equivalencias entre plano-marca IPE y Cliente

' 0 plano antiguo IPE
' 1 marca antigua IPE
' 2 plano nuevo
' 3 marca nueva

Set Planilla = GetObject("E:\scp-01\nv2335\nv2335_cambio_090907.xls", "Excel.Sheet.8")

fi = 1

With Planilla.Worksheets(Hoja)
Do While True

    fi = fi + 1

    m_Plano_IPE = Trim(.cells(fi, 1).Value)
    m_Marca_IPE = Trim(.cells(fi, 2).Value)

    m_Plano_Cliente = Trim(.cells(fi, 3).Value)
    m_Marca_Cliente = Trim(.cells(fi, 5).Value)
    
    If Len(m_Plano_IPE & m_Marca_IPE & m_Plano_Cliente & m_Marca_Cliente) = 0 Then
        Exit Do
    End If
    
    a_NV2335(0, fi - 2) = m_Plano_IPE
    a_NV2335(1, fi - 2) = m_Marca_IPE
    a_NV2335(2, fi - 2) = m_Plano_Cliente
    a_NV2335(3, fi - 2) = m_Marca_Cliente
    
Loop

Total_Marcas_NV2335 = fi - 2

Set Planilla = Nothing

End With

End Sub
Private Function NV2335_Equivalencia(plano_IPE As String, marca_IPE As String)
' transforma planos desde IPE a FORMATO CLIENTE
' solo para esta nota de venta
' esta funcion devuelve dos "variables" en una sola, el plano y la marca
' dividida por el caracter |
Dim i As Integer
NV2335_Equivalencia = ""
' busca en plano IPE
For i = 1 To Total_Marcas_NV2335
    If plano_IPE = a_NV2335(0, i) Then
'        If marca_IPE = a_NV2335(1, i) Then
            NV2335_Equivalencia = a_NV2335(2, i) & "|" & a_NV2335(3, i)
            Exit Function
'        End If
    End If
Next
'NV2335_Equivalencia = NV2335_Equivalencia
' busca en plano cliente
For i = 1 To Total_Marcas_NV2335
    If plano_IPE = a_NV2335(2, i) Then
'        If marca_IPE = a_NV2335(3, i) Then
            NV2335_Equivalencia = a_NV2335(2, i) & "|" & a_NV2335(3, i)
            Exit Function
'        End If
    End If
Next
End Function
Private Sub Piezas_PlanoGrabar(ByVal Nv As Double)
' para DESPIECE MARCOS ESCOBAR
' graba planos y revision de planos importados de la nv
' en tabla sql.scp0.piezas, para la nv completa
Dim sql As String, RsPd As Recordset ', RsPz As ADODB.Recordset

sql = "SELECT * FROM [planos detalle]"
sql = sql & " WHERE nv=" & Nv
Set RsPd = Dbm.OpenRecordset(sql)
With RsPd
Do While Not .EOF
    sql = "UPDATE piezas SET plano='" & !Plano & "'" & " WHERE nv=" & Nv & " AND marca='" & !Marca & "'"
    CnxSqlServer_scp0.Execute sql
    .MoveNext
Loop
.Close
End With
'MsgBox "piezas_planograbar"
End Sub
Private Sub xxxPlanos_Revisiones_EmailSend(ByVal Nv As Double, ByVal Plano As String, _
 ByVal Revision As String, ByVal Marca As String, ByVal Cantidad As Integer, ByVal Descripcion As String, _
 ByVal Peso As Double, ByVal Superficie As Double)
' para KARINA HENRIQUEZ/Luis Banda/Erwin Manriquez
' verifica si al importar planos hay revisiones nuevas
' en cuyo caso envia emails
'tabla scp0.planos_detalle_revisiones

Dim m_Tabla As String, sql As String
Dim RsPaso As New ADODB.Recordset

Dim s_PUNI As String, s_SUNI As String

m_Tabla = "planos_detalle_revisiones"

' primero veo si exsite revision
sql = "SELECT * FROM " & m_Tabla
sql = sql & " WHERE nv=" & Nv
sql = sql & " AND plano='" & Plano & "'"
sql = sql & " AND marca='" & Marca & "'"
sql = sql & " AND rev='" & Revision & "'"

RsPaso.Open sql, CnxSqlServer_scp0
If RsPaso.EOF Then
    
    ' no existe revision
    ' por lo tanto es rivision nueva
    s_PUNI = Replace(str(Peso), ",", ".")
    s_SUNI = Replace(str(Superficie), ",", ".")
    
End If

End Sub
