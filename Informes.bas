Attribute VB_Name = "Informes"
Option Explicit
Private qry As String, m_NvArea As Integer
Public Function Repo_Planos_General(Nv As Double, Plano As String) As String
Dim r As String
r = ""
If Nv = 0 Then
    If Plano = "" Then
    Else
        r = "{Planos Cabecera.Plano}='" & Trim(Plano) & "'"
    End If
Else
    If Plano = "" Then
        r = "{Planos Cabecera.NV}=" & Format(Nv)
    Else
        r = "{Planos Cabecera.NV}=" & Format(Nv) & " AND {Planos Cabecera.Plano}='" & Trim(Plano) & "'"
    End If
End If
Repo_Planos_General = r
End Function
Public Sub Repo_Planos_Detalle(Nv As Double, Plano As String)
Dim Dbm As Database, RsNVc As Recordset, RsPd As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String, m_obra As String

NomTabla = "Planos"

qry = MyQuery(Nv, "", "", Plano, "", Fecha_Vacia, Fecha_Vacia, 0)
qry = "SELECT * FROM [Planos Detalle]" & qry

Set Dbm = OpenDatabase(mpro_file)
Set RsPd = Dbm.OpenRecordset(qry)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

Do While Not RsPd.EOF

    m_obra = ""
    RsNVc.Seek "=", RsPd!Nv, RsPd!NvArea
    If Not RsNVc.NoMatch Then
        m_obra = RsNVc![obra]
    End If
    
    RsRepo.AddNew
    RsRepo!Nv = RsPd!Nv
    RsRepo!obra = m_obra
    RsRepo!Plano = RsPd!Plano
    RsRepo!Rev = RsPd!Rev
    RsRepo!item = RsPd!item
    RsRepo![Cantidad Total] = RsPd![Cantidad Total]
    RsRepo!Descripcion = RsPd!Descripcion
    RsRepo!Marca = RsPd!Marca
    RsRepo![Peso Unitario] = RsPd![Peso]
    RsRepo![Peso Total] = RsPd![Cantidad Total] * RsPd![Peso]
    RsRepo![m2 Unitario] = RsPd![Superficie]
    RsRepo![m2 Total] = RsPd![Cantidad Total] * RsPd![Superficie]
    RsRepo!Observaciones = RsPd!Observaciones
    RsRepo!densidad_n = RsPd!densidad
    
    RsRepo.Update
    RsPd.MoveNext
Loop

Dbm.Close
Dbi.Close

End Sub
Public Sub Repo_Planos_Detalle_Revisiones(Nv As Double) ', Plano As String)
' muestra detalle de planos y sus revisiones
Dim Dbm As Database, RsNVc As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String, m_obra As String

Dim RsPdRev As ADODB.Recordset

NomTabla = "Planos"

'qry = MyQuery(Nv, "", "", Plano, "", Fecha_Vacia, Fecha_Vacia, 0)
qry = "SELECT * FROM [planos_detalle_revisiones]" ' & qry
If Nv > 0 Then
    qry = qry & " WHERE nv=" & Nv
End If

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set RsPdRev = New ADODB.Recordset

RsPdRev.Open qry, CnxSqlServer_scp0


Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

With RsPdRev

Do While Not RsPdRev.EOF

    m_obra = ""
    RsNVc.Seek "=", !Nv, !NvArea
    If Not RsNVc.NoMatch Then
        m_obra = RsNVc![obra]
    End If
    
    RsRepo.AddNew
    RsRepo!Nv = !Nv
    RsRepo!obra = m_obra
    RsRepo!Plano = !Plano
    RsRepo!Rev = !Rev
'    RsRepo!Item = !Item
    RsRepo![Cantidad Total] = ![cantidad_total]
    RsRepo!Descripcion = !Descripcion
    RsRepo!Marca = !Marca
    RsRepo![Peso Unitario] = ![Peso]
    RsRepo![Peso Total] = ![cantidad_total] * ![Peso]
    RsRepo![m2 Unitario] = ![Superficie]
    RsRepo![m2 Total] = ![cantidad_total] * ![Superficie]
    RsRepo!Observaciones = !Observacion
    RsRepo!Fecha = !Fecha
'    RsRepo!densidad_n = !densidad
    
    RsRepo.Update
    .MoveNext
    
Loop

End With

Dbm.Close
Dbi.Close

End Sub
Public Sub Repo_Planos_GralxObra() 'NV As Double)
' agosto 2004
' informe general x obra
' KILOS POR OBRA
Dim DbD As Database, RsSc As Recordset
Dim Dbm As Database, RsNVc As Recordset, RsPd As Recordset, RsOTfd As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String, m_Nv As Double ', m_Obra As String
'Dim i As Integer, a(99, 10) As String, m_Minimo As Integer
Dim i As Integer, a_s(2999, 10) As String, a_d(2999, 10) As Double, m_Minimo As Integer
Dim sql As String

Dim m_KgxFab As Double, m_KgenFab As Double, m_KgenPin As Double, m_KgparaDes As Double, m_KgDes As Double

Dim a_Contratistas(99, 2) As String
' n,1 : rut
' n,2 : kg x recib

Dim m_ano As Integer, m_Mes As Integer, meses(12) As String, s_Mes As String
Dim m_double As Double

Dim m_TotalKilos As Double
m_TotalKilos = 0

' tabla nueva creada 14/06/04
NomTabla = "Planos Kilos"

Set DbD = OpenDatabase(data_file)
Set RsSc = DbD.OpenRecordset("Contratistas")
RsSc.Index = "Rut"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = "Numero"
'RsNVc.Index = Nv_Index

'sql = "SELECT Cantidad,[Cantidad Recibida]"
'sql = sql & " FROM [OT Fab Detalle]"
'sql = sql & " WHERE Cantidad > [Cantidad Recibida]"
sql = "OT Fab Detalle"
Set RsOTfd = Dbm.OpenRecordset(sql)
RsOTfd.Index = "NV-Plano-Marca"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)
RsRepo.Index = "Rut"

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

qry = "SELECT NV,"
'qry = "SELECT NV,Plano,Marca,"
'qry = qry & "SUM([Cantidad Total]) AS CantidadPlanos, "
'qry = qry & "SUM(GD) AS CantidadGD, "
qry = qry & "SUM([Cantidad Total]*Peso) AS KgTot, "
qry = qry & "SUM([OT Fab]*Peso) AS KgAsi, "
qry = qry & "SUM([ITO Fab]*Peso) AS KgRec, "
qry = qry & "SUM([ITO pyg]*Peso) AS Kgpyg, "
qry = qry & "SUM([GD]*Peso) AS KgDes "
qry = qry & " FROM [Planos Detalle]"
qry = qry & " GROUP BY NV"

qry = ""

'If NV <> 0 Then
'    qry = qry & " WHERE NV=" & NV
'End If

Set RsPd = Dbm.OpenRecordset("SELECT * FROM [Planos Detalle] ORDER BY nv")

i = 0
With RsPd
Do While Not .EOF

    m_Nv = !Nv
    
'    Debug.Print m_NV, !CantidadPlanos, !CantidadGD, !KgTot, !KgAsi, !KgRec, !KgDes
    
    RsNVc.Seek "=", m_Nv, m_NvArea
    
    If Not RsNVc.NoMatch Then
    If RsNVc!Activa Then
    
        i = Obra_Buscar_Indice(a_d, m_Nv)
        
'        If m_Nv = 1420 Then
'        MsgBox ""
'        End If
        
'        Debug.Print i, m_Nv
        
        m_TotalKilos = 0
        
'        If m_Nv = 2098 Then
        m_Minimo = IIf(!GD > ![Cantidad Total], ![Cantidad Total], !GD)
'        End If
        
        a_d(i, 1) = m_Nv
        a_s(i, 2) = RsNVc!obra
        
        a_s(i, 3) = "Neg"
        If RsNVc!pintura Then
            a_s(i, 3) = "PINT"
        Else
            If RsNVc!galvanizado Then
                a_s(i, 3) = "GALV"
            End If
        End If
'        a(i, 3) = IIf(RsNVc!Pintura = True, "PINT", "GALV")
        
        If RsNVc![Lista Pernos Incluida] Then
            If RsNVc![Lista Pernos Recibida] Then
                a_s(i, 4) = "OK"
            End If
        Else
            a_s(i, 4) = "no lleva"
        End If
        
'        a(i, 5) = m_CDbl(a(i, 5)) + ![Cantidad Total] * !Peso
        a_d(i, 5) = a_d(i, 5) + ![Cantidad Total] * !Peso
        
'If m_Nv = 913 Then
'    MsgBox ""
'    m_TotalKilos = m_TotalKilos + ![Cantidad Total] * !Peso
'   Debug.Print ![Cantidad Total], !Peso, ![Cantidad Total] * !Peso, m_TotalKilos, a_d(i, 5)
'End If

'        m_KgxRec = (![OT fab] - ![ITO fab]) * !Peso   ' Kg x Recibir
'        a(i, 6) = m_CDbl(a(i, 6)) + m_KgxRec ' Kg x Recibir

        m_KgxFab = (![Cantidad Total] - ![OT fab]) * !Peso   ' Kg x fabricar
        
'        a(i, 6) = m_CDbl(a(i, 6)) + m_KgxFab ' Kg x fabricar
        a_d(i, 6) = a_d(i, 6) + m_KgxFab ' Kg x fabricar
        
        If ![OT fab] <> ![ito fab] Then
        
            ' busca marca en ot
            RsOTfd.Seek "=", !Nv, m_NvArea, !Plano, !Marca
            
'            m_KgxRec = (![OT fab] - ![ITO fab]) * !Peso   ' Kg x Recibir
            
            If Not RsOTfd.NoMatch Then
            
                Do While Not RsOTfd.EOF
                
                    If !Marca <> RsOTfd!Marca Then Exit Do
                    
                    ' agregada para corregir
                    m_KgenFab = (RsOTfd!Cantidad - RsOTfd![Cantidad Recibida]) * !Peso
                    
'                    If Trim(RsOTfd![RUT Contratista]) = "5291371-3" Then
                       ' Debug.Print !NV, !Plano, !Marca, ![Cantidad Total], !Peso
'                    End If
                    
                    RsRepo.Seek "=", RsOTfd![Rut contratista]
                    
                    If RsRepo.NoMatch Then
                    
                        RsRepo.AddNew
                        RsRepo!Tipo = "2"
                        RsRepo!rut = RsOTfd![Rut contratista]
                        
                        RsSc.Seek "=", RsOTfd![Rut contratista]
                        If Not RsSc.NoMatch Then
                            RsRepo!Descripcion = StrConv(RsSc![Razon Social], vbProperCase)
                        End If
                        
                        RsRepo!KgenFabxC = m_KgenFab
                        RsRepo.Update
                        
                    Else
                    
                        RsRepo.Edit
                        RsRepo!KgenFabxC = RsRepo!KgenFabxC + m_KgenFab
                        RsRepo.Update
                        
                    End If
                    
                    RsOTfd.MoveNext
                    
                Loop
                
            End If
        
        End If
        
'        a(i, 7) = m_CDbl(a(i, 7)) + (![ITO fab] - m_Minimo) * !Peso ' kg x Despachar

'        a(i, 7) = m_CDbl(a(i, 7)) + (![OT fab] - ![ITO fab]) * !Peso ' kg en fab
        a_d(i, 7) = a_d(i, 7) + (![OT fab] - ![ito fab]) * !Peso ' kg en fab
        
        If ![ITO pyg] < m_Minimo Then
            ' si se ha despachado sin ito pg
'            a(i, 8) = m_CDbl(a(i, 8)) + (![ITO fab] - ![ITO pyg] - m_Minimo) * !Peso ' en pint
            a_d(i, 8) = a_d(i, 8) + (![ito fab] - ![ITO pyg] - m_Minimo) * !Peso ' en pint
'            a(i, 9) = m_CDbl(a(i, 9)) + (![ITO pyg] - m_Minimo) * !Peso  ' para desp
            a_d(i, 9) = a_d(i, 9) + (![ITO pyg] - m_Minimo) * !Peso ' para desp
'            If m_Nv = 798 Then
'                MsgBox ""
'            End If
        Else
'            a(i, 8) = m_CDbl(a(i, 8)) + (![ITO pyg] - m_Minimo) * !Peso  ' en pint

'            a(i, 8) = m_CDbl(a(i, 8)) + (![ITO fab] - ![ITO pyg]) * !Peso  ' en pint
            a_d(i, 8) = a_d(i, 8) + (![ito fab] - ![ITO pyg]) * !Peso  ' en pint
            
'            a(i, 9) = m_CDbl(a(i, 9)) + (![ITO fab] - ![ITO pyg] - m_Minimo) * !Peso ' para desp

'            a(i, 9) = m_CDbl(a(i, 9)) + (![ITO pyg] - m_Minimo) * !Peso  ' para desp
            a_d(i, 9) = a_d(i, 9) + (![ITO pyg] - m_Minimo) * !Peso  ' para desp
        End If
        
'        a(i, 10) = m_CDbl(a(i, 10)) + m_Minimo * !Peso ' desp
        a_d(i, 10) = a_d(i, 10) + m_Minimo * !Peso ' desp
        
'        a(i, 5) = m_TotalKilos ' ??
        
    End If
    End If
    
    .MoveNext
    
    
Loop
End With

For i = 1 To 299 '199 '99

    If a_d(i, 1) = 0 Then Exit For
    
    RsRepo.AddNew
    RsRepo!Tipo = "1"
    RsRepo!Nv = a_d(i, 1)
    RsRepo!Descripcion = a_s(i, 2)
    RsRepo!esquema = a_s(i, 3)
    RsRepo!ListadoPernos = a_s(i, 4)
    RsRepo!KgTotales = a_d(i, 5)
    RsRepo!KgxFab = a_d(i, 6)  ' Kg x Fabricar
   
    ' porcentaje de avance de piezas fabricadas
    m_double = 0
'    If a(i, 7) > 0 Then
'        m_double = 100 - Int((a(i, 5) - a(i, 7) - a(i, 8)) * 100 / a(i, 5))
'        m_double = (Val(a(i, 7)) + Val(a(i, 8))) * 100 / a(i, 7)
        ' (total - xfab - enfab) / total
    If a_d(i, 5) <> 0 Then
        m_double = (Val(a_d(i, 5)) - Val(a_d(i, 6)) - Val(a_d(i, 7))) * 100 / a_d(i, 5)
    End If
'    End If
    
    m_double = Int(m_double + 0.5)
    Select Case m_double
    Case Is < 0
        RsRepo!pfab2 = "-0%"
    Case 100
        '  100% azul
        RsRepo!pfab2 = "100%"
    Case Else
        ' menores que 100% rojo
        RsRepo!pfab1 = Trim(str(m_double)) & "%"
    End Select
    
    RsRepo!Kgenfab = a_d(i, 7)  ' kg en fabricacion
    
'    RsRepo!Kgenpg = m_CDbl(a(i, 8)) ' en pin
    RsRepo!Kgenpg = a_d(i, 8) ' en pin
    
'    m_double = m_CDbl(a(i, 9))
    m_double = a_d(i, 9)
    If m_double >= 0 Then
        RsRepo!KgparaDes1 = m_double ' para desp (recibido de pintura)
    Else
        If a_s(i, 3) <> "Neg" Then
            RsRepo!KgparaDes2 = m_double ' para desp (recibido de pintura) valor negativo en rojo
        End If
    End If
    
    RsRepo!KgDes = a_d(i, 10) ' para desp (recibido de pintura)
    
    RsRepo.Update
    
Next

'Debug.Print t

RsOTfd.Close

GoTo Cierra

' busca kilos recibidos ultimos meses

sql = "TRANSFORM sum([ITO Fab Detalle].cantidad*[ITO Fab Detalle].[Peso Unitario]) AS [El Valor]"
sql = sql & "SELECT [ITO Fab Detalle].[RUT Contratista]"
sql = sql & "From [ITO Fab Detalle]"
sql = sql & "GROUP BY [ITO Fab Detalle].[RUT Contratista]"
sql = sql & "PIVOT Year([Fecha])&Month(Fecha);"

Set RsOTfd = Dbm.OpenRecordset(sql)

m_ano = Year(Date)
m_Mes = Month(Date)

'm_mes = 10

For i = 1 To 12
    m_Mes = m_Mes - 1
    If m_Mes = 0 Then
        m_Mes = 12
        m_ano = m_ano - 1
    End If
    meses(i) = m_ano & m_Mes
'    Debug.Print meses(i)
Next

On Error Resume Next
With RsOTfd
Do While Not .EOF

'    Debug.Print ![RUT Contratista],
    
    RsRepo.Seek "=", ![Rut contratista]
    
    If Not RsRepo.NoMatch Then
        
        RsRepo.Edit
        For i = 1 To 12
            s_Mes = meses(i)
'            Debug.Print "|" & s_mes & "|",
'            Debug.Print NoNulo_Double(RsOTfd(s_mes)),
            RsRepo("mes" & i) = NoNulo_Double(RsOTfd(s_Mes))
        Next
        RsRepo.Update
        Debug.Print " "
    
    End If
    
    .MoveNext
    
Loop
End With

Cierra:

Dbm.Close
Dbi.Close

End Sub
Public Sub Repo_NV(ByVal Rut_Cliente As String, ByVal FechaIni, ByVal FechaTer)
' reporte general de nv, version 2.0

Dim Dbm As Database, RsNVc As Recordset
Dim Dbi As Database, RsRepo As Recordset

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = "Numero"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset("generico")
RsRepo.Index = "texto10_0"

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & "generico" & "]"

With RsNVc
Do While Not .EOF

    RsRepo.AddNew
    RsRepo!texto10_0 = PadL(!Numero, 4)
    RsRepo!texto10_1 = !Fecha
    
    ' busca cliente
    If True Then
        RsRepo!texto50_0 = ""
    End If
    
    RsRepo.Update
    
    .MoveNext
    
Loop
End With

End Sub
Public Sub Repo_NVxCliente(Rut_Cliente As String, FechaIni, FechaTer)
' siempre es de un solo cliente

'Usuario.Nv_Activas

Dim DbD As Database, RsCl As Recordset
Dim Dbm As Database, RsNVc As Recordset, RsPd As Recordset
Dim DbScc As Database, RsScc As Recordset ', Qry As QueryDef
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String, m_Nv As Double ', m_Obra As String
'Dim i As Integer, a(99, 10) As String, m_Minimo As Integer
Dim i As Integer, a_s(2999, 10) As String, a_d(2999, 13) As Double, m_Minimo As Integer
Dim sql As String
Dim aKgCortados(4000) As Double ' hasta nv 4000

Dim m_KgxFab As Double, m_KgenFab As Double
Dim m_KgenGR As Double, m_KgenPP As Double, m_KgenPin As Double
Dim m_KgparaDes As Double, m_K_gDes As Double

Dim a_Contratistas(99, 2) As String
' n,1 : rut
' n,2 : kg x recib

Dim m_ano As Integer, m_Mes As Integer, meses(12) As String, s_Mes As String
Dim m_double As Double

Dim m_TotalKilos As Double
m_TotalKilos = 0

Dim m_RazonSocial As String

'Dim m_Ct As Integer, m_Gd As Integer

' tabla nueva creada 14/06/04
NomTabla = "Planos Kilos"

Set DbD = OpenDatabase(data_file)
Set RsCl = DbD.OpenRecordset("Clientes")
RsCl.Index = "Rut"

m_RazonSocial = ""
RsCl.Seek "=", Rut_Cliente
If Not RsCl.NoMatch Then
    m_RazonSocial = RsCl![Razon Social]
End If

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = "Numero"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)
RsRepo.Index = "Rut"

' ///////////////////////////////
' trae kilos cortados desde scc
' lo carga en un arrregloe
Set DbScc = OpenDatabase(scc_file)
sql = "SELECT documentos.doc_nv, sum(documentos.doc_cantidad * piezas.pz_pesounitario) AS totalKilos FROM documentos"
sql = sql & " INNER JOIN piezas ON (documentos.doc_nv = piezas.pz_nv) AND (documentos.doc_tipopieza = piezas.pz_tipo) AND (documentos.doc_pieza = piezas.pz_codigo)"
sql = sql & " WHERE documentos.doc_tipo='EPLA' or documentos.doc_tipo='EANG'"
sql = sql & " GROUP BY documentos.doc_nv"
Set RsScc = DbScc.OpenRecordset(sql)
With RsScc
Do While Not .EOF
    aKgCortados(!doc_Nv) = !TotalKilos
    .MoveNext
Loop
.Close
End With
DbScc.Close
'////////////////////////////////

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

qry = "SELECT NV,"
'qry = "SELECT NV,Plano,Marca,"
'qry = qry & "SUM([Cantidad Total]) AS CantidadPlanos, "
'qry = qry & "SUM(GD) AS CantidadGD, "
qry = qry & "SUM([Cantidad Total]*Peso) AS KgTot, "
qry = qry & "SUM([OT Fab]*Peso) AS KgAsi, "
qry = qry & "SUM([ITO Fab]*Peso) AS KgRec, "
qry = qry & "SUM([ITO pyg]*Peso) AS Kgpyg, "
qry = qry & "SUM([GD]*Peso) AS KgDes "
qry = qry & " FROM [Planos Detalle]"
qry = qry & " GROUP BY NV"

qry = ""

'If NV <> 0 Then
'    qry = qry & " WHERE NV=" & NV
'End If

Set RsPd = Dbm.OpenRecordset("SELECT * FROM [Planos Detalle] ORDER BY nv")

'Set RsPd = Dbm.OpenRecordset("SELECT * FROM [Planos Detalle] WHERE nv=1807 ORDER BY plano,marca")

'Do While Not RsPd.EOF
'    If RsPd![ito pyg] < RsPd!GD Then
'        Debug.Print RsPd!Plano, RsPd!Marca, RsPd![ito pyg], RsPd!GD
'    End If
'    RsPd.MoveNext
'Loop

i = 0
With RsPd
Do While Not .EOF

    m_Nv = !Nv
    
'    Debug.Print m_NV, !CantidadPlanos, !CantidadGD, !KgTot, !KgAsi, !KgRec, !KgDes
    
    'If m_Nv = "3054" Then
    '    MsgBox "ok"
    'End If
    
    RsNVc.Seek "=", m_Nv, m_NvArea
    
    If Not RsNVc.NoMatch Then
    
        If Usuario.Nv_Activas Then
            If RsNVc![Activa] Then    ' filtro solo activas 02/10/06
                ' incluye
            Else
                GoTo NoIncluir
            End If
        Else
            ' incluye
        End If
        
        If RsNVc![RUT CLiente] = Rut_Cliente Then
        
            If FechaIni = "__/__/__" Then
'                GoTo Incluir
            Else
                If CDate(FechaIni) <= RsNVc!Fecha Then
'                    GoTo Sigue
                Else
                    GoTo NoIncluir
                End If
            End If
'Sigue:
            If FechaTer = "__/__/__" Then
'                GoTo Incluir
            Else
                If RsNVc!Fecha <= CDate(FechaTer) Then
'                    GoTo Incluir
                Else
                    GoTo NoIncluir
                End If
            End If

'Incluir:
    
            i = Obra_Buscar_Indice(a_d, m_Nv)
            
    '        Debug.Print i, m_Nv
            
            m_TotalKilos = 0
            
'            If m_Nv = 2130 Then
'                Dim m_Ct As Integer, m_Gd As Integer
'                m_Ct = m_Ct + ![cantidad total]
'                m_Gd = m_Gd + ![GD]
''''''''''''''''                m_Minimo = IIf(!GD > ![Cantidad Total], ![Cantidad Total], !GD)
'                If !Plano = "7100-F48" Then
'                MsgBox ""
'                End If
'                Debug.Print !Plano, !Marca, ![cantidad total], m_Ct, !GD, m_Gd
                
                If ![Cantidad Total] < !GD Then
                    Debug.Print !Nv, !Plano, !Marca, ![Cantidad Total], !GD
                End If
                
'            End If
            
            a_d(i, 1) = m_Nv
            a_s(i, 2) = RsNVc!obra
            
            If RsNVc!pintura Then
                a_s(i, 3) = "PINT"
            Else
                If RsNVc!galvanizado Then
                    a_s(i, 3) = "GALV"
                End If
            End If
    '        a(i, 3) = IIf(RsNVc!Pintura = True, "PINT", "GALV")
            
            If RsNVc![Lista Pernos Incluida] Then
                If RsNVc![Lista Pernos Recibida] Then
                    a_s(i, 4) = "OK"
                End If
            Else
                a_s(i, 4) = "no lleva"
            End If
            
    '        a(i, 5) = m_CDbl(a(i, 5)) + ![Cantidad Total] * !Peso
            a_d(i, 5) = a_d(i, 5) + ![Cantidad Total] * !Peso
            
    'If m_Nv = 913 Then
    '    MsgBox ""
    '    m_TotalKilos = m_TotalKilos + ![Cantidad Total] * !Peso
    '   Debug.Print ![Cantidad Total], !Peso, ![Cantidad Total] * !Peso, m_TotalKilos, a_d(i, 5)
    'End If
    
    '        m_KgxRec = (![OT fab] - ![ITO fab]) * !Peso   ' Kg x Recibir
    '        a(i, 6) = m_CDbl(a(i, 6)) + m_KgxRec ' Kg x Recibir
            
            m_KgxFab = (![Cantidad Total] - ![OT fab]) * !Peso   ' Kg x fabricar

    '        a(i, 6) = m_CDbl(a(i, 6)) + m_KgxFab ' Kg x fabricar
            a_d(i, 6) = a_d(i, 6) + m_KgxFab ' Kg x fabricar
            
    '        a(i, 7) = m_CDbl(a(i, 7)) + (![ITO fab] - m_Minimo) * !Peso ' kg x Despachar
    
    '        a(i, 7) = m_CDbl(a(i, 7)) + (![OT fab] - ![ITO fab]) * !Peso ' kg en fab
            a_d(i, 7) = a_d(i, 7) + (![OT fab] - ![ito fab]) * !Peso ' kg en fab
            'If m_Nv = 3274 Then
            '    Debug.Print (![OT fab] - ![ITO fab]) * !Peso & "|" & a_d(i, 7)
            'End If
'If m_Nv = 2885 Then
'MsgBox ""
'End If

            ' correccion 27/03/09
            'm_Minimo = (![ITO fab] - ![ITO pyg])
            
            ' si nv es negro , correccion 22/11/11
            If a_s(i, 3) = "" Then m_Minimo = 0

'If m_Nv = 2885 Then
'MsgBox ""
'Debug.Print m_Nv, m_Minimo
'End If
            
            'If !GD >= ![Cantidad Total] Then m_Minimo = 0

            a_d(i, 8) = a_d(i, 8) + (![ito fab] - ![ito gr]) * !Peso ' en negro
            a_d(i, 9) = a_d(i, 9) + (![ito gr] - ![ito pp]) * !Peso ' en PP
            a_d(i, 10) = a_d(i, 10) + (![ito pp] - ![ITO pyg]) * !Peso ' en Pin
            a_d(i, 11) = a_d(i, 11) + (![ito pp] - !GD) * !Peso ' para Despacho
            a_d(i, 12) = a_d(i, 12) + !GD * !Peso  ' despachado
        
        End If
        
    End If
NoIncluir:
    
    .MoveNext
    
Loop
.Close
End With

'GoTo SinGdEsp
' busca kilos despachados en guias especiales
'////////////////////////////////////////////
Set RsPd = Dbm.OpenRecordset("SELECT * FROM [GD Cabecera] WHERE tipo='E' ORDER BY nv,numero")

i = 0
With RsPd
Do While Not .EOF

    m_Nv = !Nv
    
'    If m_Nv = 1471 Then
'    MsgBox ""
'    End If
    
    RsNVc.Seek "=", m_Nv, m_NvArea
    
    If Not RsNVc.NoMatch Then
            
        If Usuario.Nv_Activas Then
        
            If RsNVc![Activa] Then ' filtro solo activas 02/10/06
            
                ' incluye
                
            Else
            
                GoTo NoIncluirGd
            
            End If
            
        Else
            
            ' incluye
            
        End If
            
        If RsNVc![RUT CLiente] = Rut_Cliente Then
        
            If FechaIni = "__/__/__" Then
'                GoTo IncluirGd
            Else
                If CDate(FechaIni) <= RsNVc!Fecha Then
'                    GoTo SigueGd
                Else
                    GoTo NoIncluirGd
                End If
            End If
            
'SigueGd:
            If FechaTer = "__/__/__" Then
'                GoTo IncluirGd
            Else
                If RsNVc!Fecha <= CDate(FechaTer) Then
'                    GoTo IncluirGd
                Else
                    GoTo NoIncluirGd
                End If
            End If
            
            
'IncluirGd:
        
'            Debug.Print m_Nv, RsPd!Número
            ' buscar nv, si es que existe en arreglo
            
            i = Obra_Buscar_Indice(a_d, m_Nv)
            a_d(i, 1) = m_Nv
            a_s(i, 2) = RsNVc!obra
            
            If RsNVc!pintura Then
                a_s(i, 3) = "PINT"
            Else
                If RsNVc!galvanizado Then
                    a_s(i, 3) = "GALV"
                End If
            End If
    '        a(i, 3) = IIf(RsNVc!Pintura = True, "PINT", "GALV")
            
            If RsNVc![Lista Pernos Incluida] Then
                If RsNVc![Lista Pernos Recibida] Then
                    a_s(i, 4) = "OK"
                End If
            Else
                a_s(i, 4) = "no lleva"
            End If
            
            a_d(i, 13) = a_d(i, 13) + RsPd![Peso Total] ' kg desp con gd especial
            
        End If
        
    End If
    
NoIncluirGd:
    
    .MoveNext

Loop
.Close
End With
SinGdEsp:
'////////////////////////////////////////////
' busca nv que no tengan nada, solo el nombre
'GoTo SinNombre
With RsNVc
.MoveFirst
Do While Not .EOF

    If RsNVc![RUT CLiente] = Rut_Cliente Then
    
        m_Nv = ![Numero]
    
'If m_Nv = 1479 Then
'MsgBox ""
'End If
    
    
        If FechaIni = "__/__/__" Then
'                GoTo IncluirGd
        Else
            If CDate(FechaIni) <= RsNVc!Fecha Then
'                    GoTo SigueGd
            Else
                GoTo NoIncluirVacia
            End If
        End If
        
        If FechaTer = "__/__/__" Then
'                GoTo IncluirGd
        Else
            If RsNVc!Fecha <= CDate(FechaTer) Then
'                    GoTo IncluirGd
            Else
                GoTo NoIncluirVacia
            End If
        End If
        
'            Debug.Print m_Nv, RsPd!Número
        ' buscar nv, si es que existe en arreglo
        
        i = Obra_Buscar_Indice(a_d, m_Nv)
        a_d(i, 1) = m_Nv
        a_s(i, 2) = RsNVc!obra
        
        If RsNVc!pintura Then
            a_s(i, 3) = "PINT"
        Else
            If RsNVc!galvanizado Then
                a_s(i, 3) = "GALV"
            End If
        End If
'        a(i, 3) = IIf(RsNVc!Pintura = True, "PINT", "GALV")
        
        If RsNVc![Lista Pernos Incluida] Then
            If RsNVc![Lista Pernos Recibida] Then
                a_s(i, 4) = "OK"
            End If
        Else
            a_s(i, 4) = "no lleva"
        End If
        
'        a_d(i, 11) = a_d(i, 11) + RsPd![Peso Total] ' kg desp con gd especial
        
    End If

NoIncluirVacia:

    .MoveNext
    
Loop
End With
'////////////////////////////////////////////
SinNombre:

For i = 1 To 99

    If a_d(i, 1) = 0 Then Exit For
    
    RsRepo.AddNew
    RsRepo!Tipo = "1"
    RsRepo!Nv = a_d(i, 1)
    RsRepo!Descripcion = a_s(i, 2)
    RsRepo!esquema = a_s(i, 3)
    RsRepo!ListadoPernos = a_s(i, 4)
    RsRepo!KgTotales = a_d(i, 5)
    RsRepo!KgxFab = a_d(i, 6)  ' Kg x Fabricar
   
    ' porcentaje de avance de piezas fabricadas
    m_double = 0
'    If a(i, 7) > 0 Then
'        m_double = 100 - Int((a(i, 5) - a(i, 7) - a(i, 8)) * 100 / a(i, 5))
'        m_double = (Val(a(i, 7)) + Val(a(i, 8))) * 100 / a(i, 7)
        ' (total - xfab - enfab) / total
''''        m_double = (Val(a_d(i, 5)) - Val(a_d(i, 6)) - Val(a_d(i, 7))) * 100 / a_d(i, 5)
'    End If
    
'    m_double = Int(m_double + 0.5)
'    Select Case m_double
'   Case Is < 0
'        RsRepo!pfab2 = "-0%"
'    Case 100
'        '  100% azul
'        RsRepo!pfab2 = "100%"
'    Case Else
'        ' menores que 100% rojo
'        RsRepo!pfab1 = Trim(Str(m_double)) & "%"
'    End Select
    
    RsRepo!Kgenfab = a_d(i, 7)  ' kg en fabricacion
    
    RsRepo!KgenGR = a_d(i, 8) ' en GR
    RsRepo!KgenPP = a_d(i, 9) ' en PP
    RsRepo!KgenPin = a_d(i, 10) ' en Pin
    
'    If a_d(i, 1) = 2885 Then
'        MsgBox ""
'    End If
    
    m_double = a_d(i, 11)
    If m_double >= 0 Then
        RsRepo!KgparaDes1 = m_double ' para desp (recibido de pintura)
    Else
        RsRepo!KgparaDes2 = m_double ' para desp (recibido de pintura) valor negativo en rojo
    End If
    
    RsRepo!KgDes = a_d(i, 12) ' para desp (recibido de pintura)
    
    RsRepo!KgenFabxC = a_d(i, 13)
    
    RsRepo.Update
    
Next

'Debug.Print t

' busca kilos cortados
With RsRepo
'.MoveLast
If .RecordCount > 0 Then
    .MoveFirst
    Do While Not .EOF
        .Edit
        !kgEntregados = aKgCortados(!Nv)
        .Update
        .MoveNext
    Loop
End If
End With
'/////////////////////

DbD.Close
Dbm.Close
Dbi.Close

End Sub
Public Sub Repo_FabricacionyDespacho()

' 05/2006
' informe de lo fabricado (itof) y despachado (gd) agrupado por obra
' por rango de fechas
' incluye archivo historico

'Dim DbD As Database, RsSc As Recordset
Dim Dbm As Database, RsNVc As Recordset, RsPd As Recordset, RsITOfc As Recordset, RsGDc As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String, m_Nv As Integer ', m_Obra As String
Dim i As Integer, Archivo_Movs(1) As String
Dim sql As String

Dim m_KgFab As Double, m_SFab As Double, m_KgDes As Double, m_SDes As Double

'Dim m_ano As Integer, m_mes As Integer, meses(12) As String, s_mes As String
Dim m_double As Double

' tabla nueva creada 14/06/04
NomTabla = "Planos Kilos"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)
RsRepo.Index = "Rut"

' borra tabla de paso
'Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

Archivo_Movs(0) = Movs_Path(Empresa.rut, True)
Archivo_Movs(1) = mpro_file

For i = 1 To 1

    Set Dbm = OpenDatabase(Archivo_Movs(i))
    Debug.Print Archivo_Movs(i)
    
    Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
    RsNVc.Index = "Numero"

    qry = "SELECT NV,"
    qry = qry & "SUM([Peso Total]) AS KgTot, "
    qry = qry & "SUM([Precio Total]) AS STot "
    qry = qry & " FROM [ITO Fab Cabecera]"
    qry = qry & " GROUP BY NV"
'    qry = qry & " WHERE year(fecha)=2006"

    qry = "SELECT * FROM [ITO Fab Cabecera]"
    If True Then
        qry = qry & " WHERE fecha<=#01/01/06#"
    End If
    qry = qry & "Order by NV"

    Set RsITOfc = Dbm.OpenRecordset(qry)

    With RsITOfc
    m_Nv = !Nv
    m_KgFab = ![Peso Total]
    m_SFab = ![Precio Total]
    Do While Not .EOF

        If m_Nv = !Nv Then
            m_KgFab = m_KgFab + ![Peso Total]
            m_SFab = m_SFab + ![Precio Total]
        Else
            If m_KgFab <> 0 Then
                m_double = m_SFab / m_KgFab
                Debug.Print m_Nv, Format(m_KgFab, "#,###"), Format(m_SFab, "#,###"), Format(m_double, "#,###.###")
            End If
            
            'graba en repo
            
            m_Nv = !Nv
            m_KgFab = ![Peso Total]
            m_SFab = ![Precio Total]
            
        End If

'        RsNVc.Seek "=", m_Nv
    
        .MoveNext
    
    Loop
    End With

    RsITOfc.Close
    
    ' guias de despacho
    
    RsNVc.Close
    
    Dbm.Close

Next

Dbi.Close

End Sub
Private Function Obra_Buscar_Indice(arreglo, Nv As Double) As Integer
Dim i As Integer
Obra_Buscar_Indice = 1
For i = 1 To 2999
    If arreglo(i, 1) = 0 Or arreglo(i, 1) = Nv Then
        Obra_Buscar_Indice = i
        Exit For
    End If
Next
End Function
Public Sub Repo_OTf(Nv As Double, RUT_SubC As String, Fecha_Ini As String, Fecha_Fin As String)
'Dim DbD As Database, RsSc As Recordset
Dim Dbm As Database, RsNVc As Recordset, RsOTc As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String
Dim m_obra As String, m_sc As String

NomTabla = "OT"

'Set DbD = OpenDatabase(data_file)
'Set RsSc = DbD.OpenRecordset("Contratistas")
'RsSc.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

qry = MyQuery(Nv, RUT_SubC, "", "", "Fecha", Fecha_Ini, Fecha_Fin, 0)
qry = "SELECT * FROM [OT Fab Cabecera]" & qry
Set RsOTc = Dbm.OpenRecordset(qry)

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

Do While Not RsOTc.EOF
    
    m_obra = ""
    RsNVc.Seek "=", RsOTc!Nv, RsOTc!NvArea
    If Not RsNVc.NoMatch Then
        m_obra = RsNVc![obra]
    End If
    
    m_sc = ""
'    RsSc.Seek "=", RsOTc![RUT Contratista]
'    If Not RsSc.NoMatch Then
'        m_sc = RsSc![Razon Social]
'    End If
    m_sc = Contratista_Lee(SqlRsSc, RsOTc![Rut contratista])

    RsRepo.AddNew
    RsRepo!Nv = RsOTc!Nv
    RsRepo!obra = m_obra
    RsRepo!Contratista = m_sc
    RsRepo![Nº OT] = RsOTc!Numero
    RsRepo!Fecha = RsOTc!Fecha
    RsRepo![Kg Total] = RsOTc![Peso Total]
    RsRepo![$ Total] = RsOTc![Precio Total]
    
    If RsOTc![Peso Total] = 0 Then
        RsRepo![$ Promedio] = 0
    Else
        RsRepo![$ Promedio] = RsOTc![Precio Total] / RsOTc![Peso Total]
    End If
    RsRepo.Update
    
    RsOTc.MoveNext
    
Loop

'DbD.Close
Dbm.Close
Dbi.Close

End Sub
Public Function F_Repo_OTfd(Nv As Double, RUT_SubC As String, Fecha_Ini As String, Fecha_Fin As String) As String
Dim r As String, param As Boolean
r = ""

param = False
If Nv <> 0 Then
    param = True
    r = "{OT Fab Cabecera.NV}=" & Format(Nv)
End If

If RUT_SubC <> "" Then
    If param Then r = r & " AND "
    r = r & "{OT Fab Detalle.RUT Contratista}='" & RUT_SubC & "'"
    param = True
End If

If Fecha_Ini <> "__/__/__" Then
    If param Then r = r & " AND "
    r = r & "{OT Fab Detalle.Fecha}>=Date(" & Format(Fecha_Ini, "yyyy,mm,dd") & ")"
'    Date (1999, 11, 18) -> 18/11/1999
    param = True
End If

If Fecha_Fin <> "__/__/__" Then
    If param Then r = r & " AND "
    r = r & "{OT Fab Detalle.Fecha}<=Date(" & Format(Fecha_Fin, "yyyy,mm,dd") & ")"
End If

F_Repo_OTfd = r

End Function
Public Sub Repo_OTe(Nv As Double, RUT_SubC As String, Fecha_Ini As String, Fecha_Fin As String, Optional Tabla As String, Optional Tipo As String)
Dim DbD As Database, RsSc As Recordset
Dim Dbm As Database, RsNVc As Recordset, RsOTc As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String
Dim m_obra As String, m_sc As String

If IsNull(Tabla) Then
    NomTabla = "OT"
Else
    If Tabla = "" Then
        NomTabla = "OT"
    Else
        NomTabla = Tabla
    End If
End If

Set DbD = OpenDatabase(data_file)
Set RsSc = DbD.OpenRecordset("Contratistas")
RsSc.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

qry = MyQuery(Nv, RUT_SubC, "", "", "Fecha", Fecha_Ini, Fecha_Fin, 0)
qry = "SELECT * FROM [OT Esp Cabecera]" & qry
Set RsOTc = Dbm.OpenRecordset(qry)

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

Do While Not RsOTc.EOF
    
    m_obra = ""
    RsNVc.Seek "=", RsOTc!Nv, RsOTc!NvArea
    If Not RsNVc.NoMatch Then
        m_obra = RsNVc![obra]
    End If
    
    m_sc = ""
    RsSc.Seek "=", RsOTc![Rut contratista]
    If Not RsSc.NoMatch Then
        m_sc = RsSc![Razon Social]
    End If
    
    RsRepo.AddNew
    RsRepo!Nv = RsOTc!Nv
    RsRepo!obra = m_obra
    RsRepo!Contratista = m_sc
    RsRepo![Nº OT] = RsOTc!Numero
    RsRepo!Fecha = RsOTc!Fecha
'    RsRepo![Kg Total] =  RsOTc![Precio TotaUnitario]
    RsRepo![$ Total] = RsOTc![Precio Total]
'    RsRepo!Montaje = RsOTc!Montaje

    RsRepo![Tipo] = RsOTc!Tipo
    
'    If RsOTc![Peso Total] = 0 Then
'        RsRepo![$ Promedio] = 0
'    Else
'        RsRepo![$ Promedio] = RsOTc![Precio Total] / RsOTc![Peso Total]
'    End If
    RsRepo.Update
    
    RsOTc.MoveNext
    
Loop

DbD.Close
Dbm.Close
Dbi.Close

End Sub
Public Sub Repo_ITOc(Tipo As String, Nv As Double, RUT_SubC As String, Fecha_Ini As String, Fecha_Fin As String) ', RUT_Operador As String)

' itos cabecera

Dim DbD As Database, RsSc As Recordset
Dim Dbm As Database, RsNVc As Recordset, RsITOc As Recordset, RsITOd As Recordset, RsOTd As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String
Dim m_obra As String, m_sc As String, m_PG As String

NomTabla = "ITO"

Set DbD = OpenDatabase(data_file)
Set RsSc = DbD.OpenRecordset("Contratistas")
RsSc.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

qry = MyQuery(Nv, RUT_SubC, "", "", "Fecha", Fecha_Ini, Fecha_Fin, 0)

'//////////////////////
' // tipo = Pin  o  Gal
m_PG = ""
If Tipo = "Pin" Then
    m_PG = "Tipo='P'"
    Tipo = "PG"
End If
If Tipo = "Gal" Then
    m_PG = "Tipo='G'"
    Tipo = "PG"
End If
If Tipo = "Gra" Then
    m_PG = "Tipo='R'"
    Tipo = "PG"
End If
'//////////////////////

'qry = "SELECT * FROM [ITO " & Tipo & " Cabecera]"

If Len(m_PG) > 0 Then
    If Len(qry) > 0 Then
        qry = qry & " AND " & m_PG
    Else
        qry = " WHERE " & m_PG
    End If
End If

'If Len(RUT_Operador) > 0 Then
'If Len(qry) > 0 Then
'    qry = qry & " AND [RUT Operador]='" & RUT_Operador & "'"
'Else
'    qry = " WHERE [RUT Operador]='" & RUT_Operador & "'"
'End If
'End If

qry = "SELECT * FROM [ITO " & Tipo & " Cabecera]" & qry

Set RsITOc = Dbm.OpenRecordset(qry)

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' si es fabricaciopn abre otrs tablas para buscar percios unitarios
'If Tipo = "Fab" Then

'    Set RsITOd = Dbm.OpenRecordset("ITO Fab Detalle")
'    RsITOd.Index = "Número-Línea"
    
'    Set RsOTd = Dbm.OpenRecordset("OT Fab Detalle")
'    RsOTd.Index = "Número-Línea"
    
'End If

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

Do While Not RsITOc.EOF
    
'    If RsITOc![RUT Contratista] <> "89784800-7" Then ' excluye estructuras metalicas
    If True Then ' excluye estructuras metalicas
    
        m_obra = ""
        RsNVc.Seek "=", RsITOc!Nv, RsITOc!NvArea
        If Not RsNVc.NoMatch Then
            m_obra = RsNVc![obra]
        End If
        
        m_sc = ""
        If Tipo = "Fab" Then
            RsSc.Seek "=", RsITOc![Rut contratista]
        Else
            RsSc.Seek "=", RsITOc![Rut contratista]
        End If
        
        If Not RsSc.NoMatch Then
            m_sc = RsSc![Razon Social]
        End If
        
        RsRepo.AddNew
        RsRepo!Nv = RsITOc!Nv
        RsRepo!obra = m_obra
        RsRepo!Contratista = m_sc
        
        If Tipo = "Fab" Then
            RsRepo![Nº ITO] = RsITOc!Numero
        Else
            RsRepo![Nº ITO] = RsITOc!Numero
        End If
        
        RsRepo!Fecha = RsITOc!Fecha
        
        RsRepo![Kg Total] = RsITOc![Peso Total]
        If Tipo = "Fab" Then
        Else
            RsRepo![m2 Total] = RsITOc![m2 Total]
        End If
        
        RsRepo![$ Total] = RsITOc![Precio Total]
        
        If Tipo = "Fab" Then
            If RsITOc![Peso Total] > 0 Then
                RsRepo![m2 Total] = RsITOc![Precio Total] / RsITOc![Peso Total]
            End If
        End If
        
        RsRepo.Update
    
    End If
    
    RsITOc.MoveNext
    
Loop

DbD.Close
Dbm.Close
Dbi.Close

End Sub
Public Sub Repo_ITOfd(Nv As Double, RUT_SubC As String, Fecha_Ini As String, Fecha_Fin As String, TipoITO As String)

' ito fabricaciion detalle

Dim DbD As Database, RsTr As Recordset
Dim Dbm As Database, RsNVc As Recordset, RsITOd As Recordset, RsPld As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String
Dim m_obra As String ', m_TrabNombre As String, m_pr As String
Dim m_NvArea As Integer, m_Tipo As String

NomTabla = "ITOpr_det" ' detalle

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set RsPld = Dbm.OpenRecordset("planos detalle")
RsPld.Index = "nv-plano-marca"

qry = MyQuery(Nv, RUT_SubC, "", "", "Fecha", Fecha_Ini, Fecha_Fin, 0)

If TipoITO = "FAB" Then
    qry = "SELECT * FROM [ITO fab Detalle]" & qry
End If
If TipoITO = "PYG" Then
    qry = "SELECT * FROM [ITO pg Detalle]" & qry
    qry = qry & " AND tipo='P'"
End If

Debug.Print qry

Set RsITOd = Dbm.OpenRecordset(qry)


Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

With RsITOd

Do While Not .EOF
    
    m_obra = ""
    RsNVc.Seek "=", RsITOd!Nv, RsITOd!NvArea
    If Not RsNVc.NoMatch Then
        m_obra = RsNVc![obra]
    End If
    
    RsRepo.AddNew
    
    RsRepo!Numero = !Numero
    RsRepo!Nv = !Nv
    RsRepo!Fecha = !Fecha
    RsRepo!obra = m_obra
    RsRepo!Plano = !Plano
    RsRepo!Rev = !Rev
    RsRepo!Marca = !Marca
    RsRepo!cantidad_pr = !Cantidad
    ' agregado el 05/04/16
    RsRepo![Precio Unitario] = ![Precio Unitario]

    RsPld.Seek "=", !Nv, m_NvArea, !Plano, !Marca
    If Not RsPld.NoMatch Then
        RsRepo!cantidad_total = RsPld![Cantidad Total]
        RsRepo!Descripcion = RsPld!Descripcion
        RsRepo![m2 Unitario] = RsPld![Superficie]
        'RsRepo![Peso Unitario] = RsPld![Peso]
    End If
    RsRepo![Peso Unitario] = ![Peso Unitario]
    'RsRepo![m2 Unitario] = ![m2 Unitario]

    RsRepo.Update

    .MoveNext

Loop

End With

Dbm.Close
Dbi.Close

End Sub
Public Sub Repo_ITOd(Tipo As String, Nv As Double, RUT_SubC As String, Fecha_Ini As String, Fecha_Fin As String, RUT_Operador As String, TipoGranalla As String, Maquina As String, Turno As Integer)

' itos detalle Pintura y gRanalla

Dim DbD As Database, RsTr As Recordset
Dim Dbm As Database, RsNVc As Recordset, RsITOd As Recordset, RsPld As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String
Dim m_obra As String, m_TrabNombre As String, m_pr As String
Dim m_NvArea As Integer, m_Tipo As String

NomTabla = "ITOpr_det" ' detalle, solo para itos pintura y granalla

Set DbD = OpenDatabase(data_file)
Set RsTr = DbD.OpenRecordset("Trabajadores")
RsTr.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set RsPld = Dbm.OpenRecordset("planos detalle")
RsPld.Index = "nv-plano-marca"

qry = MyQuery(Nv, RUT_SubC, "", "", "Fecha", Fecha_Ini, Fecha_Fin, 0)

'//////////////////////
' // tipo = Pin  o  Gal
m_pr = ""
If Tipo = "Fab" Then
    m_pr = ""
    m_Tipo = "Fab"
End If
If Tipo = "Pin" Then
    m_pr = "Tipo='P'"
    m_Tipo = "PG"
End If
If Tipo = "Gra" Then
    m_pr = "Tipo='R'"
    m_Tipo = "PG"
End If
If Tipo = "GraEsp" Then
    m_pr = "Tipo='S'"
    m_Tipo = "PG"
End If
If Tipo = "pp" Then
    m_pr = "Tipo='T'"
    m_Tipo = "PG"
End If
If Tipo = "ppesp" Then
    m_pr = "Tipo='U'"
    m_Tipo = "PG"
End If
'//////////////////////

'qry = "SELECT * FROM [ITO " & Tipo & " Cabecera]"

If Len(m_pr) > 0 Then
    If Len(qry) > 0 Then
        qry = qry & " AND " & m_pr
    Else
        qry = " WHERE " & m_pr
    End If
End If

If Len(RUT_Operador) > 0 Then
    If Len(qry) > 0 Then
        qry = qry & " AND [RUT Operador]='" & RUT_Operador & "'"
    Else
        qry = " WHERE [RUT Operador]='" & RUT_Operador & "'"
    End If
End If

If Len(TipoGranalla) > 0 Then
    If Len(qry) > 0 Then
        qry = qry & " AND [tipo2]='" & TipoGranalla & "'"
    Else
        qry = " WHERE tipo2='" & TipoGranalla & "'"
    End If
End If

If Len(Maquina) > 0 Then
    If Len(qry) > 0 Then
        qry = qry & " AND maquina='" & Maquina & "'"
    Else
        qry = " WHERE maquina='" & Maquina & "'"
    End If
End If

If Turno > 0 Then
    If Len(qry) > 0 Then
        qry = qry & " AND turno=" & Turno
    Else
        qry = " WHERE turno=" & Turno
    End If
End If

qry = "SELECT * FROM [ITO " & m_Tipo & " Detalle]" & qry

Set RsITOd = Dbm.OpenRecordset(qry)

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

With RsITOd

Do While Not .EOF
    
    m_obra = ""
    RsNVc.Seek "=", RsITOd!Nv, RsITOd!NvArea
    If Not RsNVc.NoMatch Then
        m_obra = RsNVc![obra]
    End If
    
    RsRepo.AddNew
    
    RsRepo!Numero = !Numero
    RsRepo!Nv = !Nv
    RsRepo!Fecha = !Fecha
    RsRepo!obra = m_obra
    RsRepo![Nombre Operador] = Left(m_TrabNombre, 30)
    RsRepo!Plano = !Plano
    RsRepo!Rev = !Rev
    RsRepo!Marca = !Marca
    RsRepo!cantidad_pr = !Cantidad
    
    RsPld.Seek "=", !Nv, m_NvArea, !Plano, !Marca
    If Not RsPld.NoMatch Then
        RsRepo!cantidad_total = RsPld![Cantidad Total]
        RsRepo!Descripcion = RsPld!Descripcion
        RsRepo![m2 Unitario] = RsPld![Superficie]
        RsRepo![Peso Unitario] = RsPld![Peso]
    End If
    
    If Tipo = "ppesp" Or Tipo = "GraEsp" Then
        RsRepo!Descripcion = !Descripcion
        RsRepo![m2 Unitario] = ![m2 Unitario]
    End If

    ' solo pintura
    RsRepo!manos1 = !manos1
    RsRepo!manos2 = !manos2
    
    ' solo granalla
    RsRepo!tipo2 = !tipo2
    RsRepo!Maquina = !Maquina
    
    RsRepo!Turno = !Turno
    
    m_TrabNombre = ""
    RsTr.Seek "=", RsITOd![RUT Operador]
    If Not RsTr.NoMatch Then
        m_TrabNombre = RsTr![appaterno] & " " & RsTr![apmaterno] & " " & RsTr![nombres]
    End If
    RsRepo![Nombre Operador] = Left(m_TrabNombre, 30)
        
    RsRepo.Update
    
    .MoveNext
    
Loop

End With

DbD.Close
Dbm.Close
Dbi.Close

End Sub
Public Sub Repo_BonoProd(Nv As Double, RUT_SubC As String, Fecha_Ini As String, Fecha_Fin As String)
Dim DbD As Database, RsSc As Recordset, RsTabla As Recordset
Dim Dbm As Database, RsNVc As Recordset, RsITOc As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String
Dim m_Nv As Double, m_obra As String, m_Rut As String, m_sc As String, m_Peso As Double
Dim m_Clas As String, total_peso As Double

'Dim A_Clasi() As String

NomTabla = "ITO"

Set DbD = OpenDatabase(data_file)

Set RsTabla = DbD.OpenRecordset("Tabla Bono Produccion")
RsTabla.Index = "Clasificacion-Tramo"

' llena arreglo de clasificaciones
Set RsSc = DbD.OpenRecordset("Clasificacion de Contratistas")
RsSc.Index = "Codigo"
'ReDim A_Clasi(0)
'A_Clasi(0) = ""
Dim n As Integer
If RsSc.RecordCount <> 0 Then
    n = 0
    Do While Not RsSc.EOF
        n = n + 1
'        ReDim Preserve A_Clasi(n)
'        A_Clasi(n) = RsSc!Código
        RsSc.MoveNext
    Loop
End If
RsSc.Close
'/////////

Set RsSc = DbD.OpenRecordset("Contratistas")
RsSc.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

qry = MyQuery(Nv, RUT_SubC, "", "", "Fecha", Fecha_Ini, Fecha_Fin, 0)
qry = "SELECT * FROM [ITO Fab Cabecera]" & qry & " ORDER BY [RUT Contratista],[NV]"
Set RsITOc = Dbm.OpenRecordset(qry)

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

If RsITOc.RecordCount = 0 Then GoTo FinBono
m_Rut = RsITOc![Rut contratista]
m_Nv = RsITOc!Nv
m_obra = Descripcion_BuscarNV(RsNVc, RsITOc!Nv, "Obra", 0)
m_sc = Descripcion_Buscar(RsSc, RsITOc![Rut contratista], "Razon Social")
' aprovechando el puntero
m_Clas = NoNulo(RsSc!Clasificacion)
m_Peso = 0 'RsITOc![Peso Total]
total_peso = 0
Do While Not RsITOc.EOF
    
'    If RsITOc![NV] = 635 Then
'        Debug.Print RsITOc![RUT Contratista]
'    End If
    
    If m_Rut = RsITOc![Rut contratista] Then
    
        If m_Nv = RsITOc!Nv Then
        
            m_Peso = m_Peso + RsITOc![Peso Total]
            
        Else
        
            Bono_Prod_Lin RsRepo, m_Nv, m_obra, m_sc, m_Clas, m_Peso
            total_peso = total_peso + m_Peso
            '
            m_Rut = RsITOc![Rut contratista]
            m_Nv = RsITOc!Nv
            m_obra = Descripcion_BuscarNV(RsNVc, RsITOc!Nv, "Obra", 0)
            m_sc = Descripcion_Buscar(RsSc, RsITOc![Rut contratista], "Razon Social")
'            If m_Rut = "76114058-2" Then
'                MsgBox ""
'            End If
            m_Clas = NoNulo(RsSc!Clasificacion)
            m_Peso = RsITOc![Peso Total]
            
        End If
        
    Else
        
        Bono_Prod_Lin RsRepo, m_Nv, m_obra, m_sc, m_Clas, m_Peso
        total_peso = total_peso + m_Peso
        '
        m_Rut = RsITOc![Rut contratista]
        m_Nv = RsITOc!Nv
        m_obra = Descripcion_BuscarNV(RsNVc, RsITOc!Nv, "Obra", 0)
        m_sc = Descripcion_Buscar(RsSc, RsITOc![Rut contratista], "Razon Social")
        
'        Debug.Print "|" & m_Rut & "|" & RsITOc!Numero & "|"
        
        ' busca clasificacion
'        If m_Rut = "76114058-2" Then
'            MsgBox ""
'        End If
        
        m_Clas = ""
        RsSc.Seek "=", m_Rut
        If Not RsSc.NoMatch Then
'            Debug.Print m_Rut
            m_Clas = NoNulo(RsSc!Clasificacion)
        End If
        
        m_Peso = RsITOc![Peso Total]
        
    End If
    
    RsITOc.MoveNext
    
Loop

Bono_Prod_Lin RsRepo, m_Nv, m_obra, m_sc, m_Clas, m_Peso
total_peso = total_peso + m_Peso

RsRepo.MoveFirst
Do While Not RsRepo.EOF
    RsRepo.Edit
    RsRepo![$ Total] = RsRepo![Kg Total] * SubC_TablaBono(RsTabla, NoNulo(RsRepo!Clasificacion), total_peso)
    RsRepo.Update
    RsRepo.MoveNext
Loop

FinBono:

DbD.Close
Dbm.Close
Dbi.Close

End Sub
Public Sub Repo_Produccion_Mensual_EMLex(FechaIni As String, FechaTer As String)
' informe de produccion mensual
' de todas las obras; terminadas y en proceso
Dim Db_M As Database, RsITOfc As Recordset, qry As String
Dim Db_I As Database, RsRepo As Recordset
Dim Arch(1) As String, i As Integer, m_Fecha As Date

qry = "SELECT * FROM [ITO Fab Cabecera]"

Arch(0) = "ScpHist": Arch(1) = "ScpMovs"

Set Db_I = OpenDatabase(repo_file)

' borra tabla de paso
Db_I.Execute "DELETE * FROM [Produccion Mensual]"

Set RsRepo = Db_I.OpenRecordset("Produccion Mensual")
RsRepo.Index = "Fecha"

For i = 0 To 1

    Set Db_M = OpenDatabase(Drive_Server & Path_Mdb & Arch(i))
    Set RsITOfc = Db_M.OpenRecordset(qry)
    
    With RsITOfc
    
    Do While Not .EOF
        
        If Format(FechaIni, "yyyy/mm") <= Format(!Fecha, "yyyy/mm") And _
           Format(!Fecha, "yyyy/mm") <= Format(FechaTer, "yyyy/mm") Then
        
            m_Fecha = CDate("01/" & Format(!Fecha, "mm/yy"))
            RsRepo.Seek "=", m_Fecha
            If RsRepo.NoMatch Then
                RsRepo.AddNew
                RsRepo!Fecha = m_Fecha
                RsRepo![Mes Año] = UCase(Format(m_Fecha, "mmm yyyy"))
                RsRepo("SubC " & Contratista(RsRepo, ![Rut contratista])) = ![Peso Total]
'                RsRepo("SubC " & !Orden) = ![Peso Total]
            Else
                RsRepo.Edit
                RsRepo("SubC " & Contratista(RsRepo, ![Rut contratista])) = RsRepo("SubC " & Contratista(RsRepo, ![Rut contratista])) + ![Peso Total]
'                RsRepo("SubC " & !Orden) = RsRepo("SubC " & !Orden) + ![Peso Total]
            End If
            RsRepo.Update
        
        End If
        
        .MoveNext
    Loop
    .Close
    End With
    Db_M.Close
Next

With RsRepo
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    .Edit
    !Total = ![SubC 0] + ![SubC 1] + ![SubC 2] + ![SubC 3] + ![SubC 4] + ![SubC 5] + ![SubC 6] + ![SubC 7] + ![SubC 8] + ![SubC 9]
    .Update
    .MoveNext
Loop
End If
End With

RsRepo.Close
Db_I.Close

End Sub
Public Sub Repo_Produccion_Mensual(FechaIni As String, FechaTer As String)
' informe de produccion mensual
' de todas las obras; terminadas y en proceso
Dim Db_D As Database, RsSc As Recordset
Dim Db_M As Database, RsNVc As Recordset, RsITOfc As Recordset, qry As String
Dim Db_I As Database, RsRepo As Recordset
Dim Arch(1) As String, i As Integer, m_Fecha As Date, m_Orden As Integer

Dim a_Contratistas(9) As String

qry = "SELECT * FROM [ITO Fab Cabecera]"
Arch(0) = "ScpHist": Arch(1) = "ScpMovs"

Set Db_D = OpenDatabase(data_file)

Set RsSc = Db_D.OpenRecordset("SELECT * FROM contratistas ORDER BY orden")
With RsSc
Do While Not .EOF
    If !orden > 0 Then
        a_Contratistas(!orden) = Left(![Razon Social], 15)
    End If
    .MoveNext
Loop
.Close
End With

Set RsSc = Db_D.OpenRecordset("contratistas")
RsSc.Index = "rut"

Set Db_I = OpenDatabase(repo_file)

' borra tabla de paso
Db_I.Execute "DELETE * FROM [Produccion Mensual]"

Set RsRepo = Db_I.OpenRecordset("Produccion Mensual")
RsRepo.Index = "Fecha"

For i = 1 To 1

    Set Db_M = OpenDatabase(Drive_Server & Path_Mdb & Arch(i))
    Set RsITOfc = Db_M.OpenRecordset(qry)
        
    Set RsNVc = Db_M.OpenRecordset("nv cabecera")
    RsNVc.Index = "numero"
    
    With RsITOfc
    
    Do While Not .EOF
        
        If Format(FechaIni, "yyyy/mm") <= Format(!Fecha, "yyyy/mm") And _
           Format(!Fecha, "yyyy/mm") <= Format(FechaTer, "yyyy/mm") Then
        
            RsNVc.Seek "=", !Nv, 0
            
            If Not RsNVc.NoMatch Then
            
                If NoNulo(RsNVc!Tipo) <> "SV" Then
                
                    m_Fecha = CDate("01/" & Format(!Fecha, "mm/yy"))
                    
                    RsRepo.Seek "=", m_Fecha
                    
                    If RsRepo.NoMatch Then
                    
                        RsRepo.AddNew
                        RsRepo!Fecha = m_Fecha
                        RsRepo![Mes Año] = UCase(Format(m_Fecha, "mmm yyyy"))
        '                RsRepo("SubC " & Contratista(![RUT Contratista])) = ![Peso Total] ''''''''''''''
        
                        For m_Orden = 1 To 9
                            RsRepo("nombre" & m_Orden) = a_Contratistas(m_Orden)
                        Next
                        
                        RsSc.Seek "=", ![Rut contratista]
                        If Not RsSc.NoMatch Then
                        
                            m_Orden = RsSc!orden
                            RsRepo("SubC " & m_Orden) = ![Peso Total]
                            If m_Orden > 0 Then '''''''''''''''''''
                                RsRepo("nombre" & m_Orden) = Left(RsSc![Razon Social], 15) '''''''''''''''''''
                            End If '''''''''''''''''''
                        
                        End If
                        
                    Else
                    
                        RsRepo.Edit
        '                RsRepo("SubC " & Contratista(![RUT Contratista])) = RsRepo("SubC " & Contratista(![RUT Contratista])) + ![Peso Total] ''''''''''''''''''
        
                        RsSc.Seek "=", ![Rut contratista]
                        If Not RsSc.NoMatch Then
                        
                            m_Orden = RsSc!orden
                            RsRepo("SubC " & m_Orden) = RsRepo("SubC " & m_Orden) + ![Peso Total]
                            If m_Orden > 0 Then ''''''''''''''''''''''
                                RsRepo("nombre" & m_Orden) = Left(RsSc![Razon Social], 15) '''''''''''''''''
                            End If ''''''''''''''''
                        
                        End If
                        
                    End If
                    
                    RsRepo.Update
                    
                End If ' If RsNvC!Tipo <> "SV" Then
            
            End If ' If Not RsNvC.NoMatch Then
        
        End If
        
        .MoveNext
        
    Loop
    .Close
    End With
    
    RsNVc.Close
    
    Db_M.Close
    
Next

With RsRepo
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    .Edit
    !Total = ![SubC 0] + ![SubC 1] + ![SubC 2] + ![SubC 3] + ![SubC 4] + ![SubC 5] + ![SubC 6] + ![SubC 7] + ![SubC 8] + ![SubC 9]
    .Update
    .MoveNext
Loop
End If
End With

RsRepo.Close
Db_I.Close

End Sub
Private Function ContratistaOLD(rut As String) As String
Dim Sc As String
Sc = "0" ' otros
Select Case Trim(rut)
Case "7001669-9"  'calquin
    Sc = "1"
Case "6466293-7"  'elisa
    Sc = "2"
Case "78961100-9" 'ralfu
    Sc = "3"
Case "6142349-4"  'manquiñir
    Sc = "4"
Case "78651710-9" 'pyp
    Sc = "5"
Case "4076128-4"  'morales
    Sc = "6"
Case "8853594-4"  'palacios
    Sc = "7"
Case "89784800-7" 'prod eml
    Sc = "8"
End Select
ContratistaOLD = Sc
End Function
Private Function Contratista_YaNo(rut As String) As String
Dim Sc As String
Sc = "0" ' otros
Select Case Trim(rut)
Case "7001669-9"  'calquin
    Sc = "1"
Case "6466293-7"  'elisa
    Sc = "2"
Case "78961100-9" 'ralfu
    Sc = "3"
Case "4076128-4"  'morales
    Sc = "4"
Case "8853594-4"  'palacios
    Sc = "5"
Case "89784800-7" 'prod eml
    Sc = "6"
Case "14251706-K" 'daniel olivos
    Sc = "7"
Case "11880244-6" 'blanca toledo
    Sc = "8"
Case "76406180-2" 'a y d
    Sc = "9"
End Select
Contratista_YaNo = Sc
End Function
Private Function Contratista(Rs As Recordset, rut As String) As String
Dim Sc As String
Sc = "0" ' otros
Rs.Seek "=", rut
If Not Rs.NoMatch Then
    Sc = Rs!orden
End If
Contratista = Sc
End Function
Private Function Descripcion_Buscar(Rs As Recordset, Codigo As String, CampoDescripcion As String)
Descripcion_Buscar = ""
Rs.Seek "=", Codigo
If Not Rs.NoMatch Then
    Descripcion_Buscar = Rs(CampoDescripcion)
End If
End Function
Private Function Descripcion_BuscarNV(Rs As Recordset, Codigo As String, CampoDescripcion As String, NvArea As Integer)
Descripcion_BuscarNV = ""
Rs.Seek "=", Codigo, NvArea
If Not Rs.NoMatch Then
    Descripcion_BuscarNV = Rs(CampoDescripcion)
End If
End Function
Private Sub Bono_Prod_Lin(Rs As Recordset, Nv As Double, obra As String, Sc As String, Clasif As String, Peso As Double)
' agrega linea en bono de prod
If Trim(Clasif) <> "" Then
Rs.AddNew
Rs!Nv = Nv
Rs!obra = obra
Rs!Contratista = Sc
Rs![Kg Total] = Peso
Rs!Clasificacion = Clasif
Rs.Update
End If
End Sub
Private Function SubC_TablaBono(RsT As Recordset, Clasificacion As String, Peso_Total As Double) As Double
SubC_TablaBono = 0
If Clasificacion = "" Then Exit Function
'desde , hasta,Clasificacion 1
RsT.Seek ">=", Clasificacion, 1
If Not RsT.NoMatch Then
    Do While Not RsT.EOF
        If RsT!Clasificacion <> Clasificacion Then Exit Do
        If RsT!Desde * 1000 < Peso_Total And Peso_Total <= RsT!Hasta * 1000 Then
            SubC_TablaBono = RsT!Valor
            Exit Do
        End If
        RsT.MoveNext
    Loop
End If

End Function
Public Sub Repo_GDs_Correlativos(Fecha_Ini As String, Fecha_Fin As String)
Dim DbD As Database, RsCl As Recordset
Dim Dbm As Database, RsNVc As Recordset, RsGDc As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String, m_obra As String, m_cliente As String
Dim i As Double, m_primera As Double, m_ultima As Double
Dim Arch(1) As String ', m_Arch As String
Dim m_NvArea As Integer

m_NvArea = 0

NomTabla = "GD"

'Suma_Limpiar

Set DbD = OpenDatabase(data_file)
Set RsCl = DbD.OpenRecordset("Clientes")
RsCl.Index = "RUT"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)
RsRepo.Index = "GD"

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

'Arch(0) = "ScpHist": Arch(1) = "ScpMovs": Arch(3) = "ScpMovs"

Arch(0) = Movs_Path(EmpOC.rut, True)  ' ScpHist
Arch(1) = Movs_Path(EmpOC.rut, False) ' ScpMovs

If UCase(EmpOC.Fantasia) = "EML" Then
    m_primera = 0
    m_ultima = 1
Else
    m_primera = 1
    m_ultima = 1
End If

For i = m_primera To m_ultima

'    m_Arch = Drive_Server & Path_Server & Arch(i)
    qry = MyQuery(0, "", "", "", "Fecha", CStr(Fecha_Ini), CStr(Fecha_Fin), 0)
    qry = "SELECT * FROM [GD Cabecera]" & qry
    
'    Set Dbm = OpenDatabase(m_Arch)
    Set Dbm = OpenDatabase(Arch(i))
    
    Set RsGDc = Dbm.OpenRecordset(qry)
    
    Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
    RsNVc.Index = Nv_Index ' "Número"
    
    Do While Not RsGDc.EOF
    
        m_obra = ""
        RsNVc.Seek "=", RsGDc!Nv, RsGDc!NvArea
        If Not RsNVc.NoMatch Then
            m_obra = RsNVc![obra]
        End If
        
        m_cliente = ""
        RsCl.Seek "=", RsGDc![RUT CLiente]
        If Not RsCl.NoMatch Then
            m_cliente = RsCl![Razon Social]
        End If
        
        RsRepo.AddNew
        
        RsRepo!Nv = RsGDc!Nv
        RsRepo!obra = m_obra
        RsRepo!Cliente = Left(m_cliente, 30)
        RsRepo!GD = RsGDc!Numero
        RsRepo!Tipo = IIf(RsGDc!Tipo = "E", "E", "")
        RsRepo!Fecha = RsGDc!Fecha
        RsRepo![Kg Total] = RsGDc![Peso Total]
        RsRepo![$ Total] = RsGDc![Precio Total]
        RsRepo!chofer = Left(RsGDc![Observacion 1], 30)
        RsRepo!patente = Left(RsGDc![Observacion 2], 30)
        RsRepo![Contenido1] = RsGDc![Observacion 3]
        RsRepo![Contenido2] = RsGDc![Observacion 4]
        
        RsRepo.Update
        
        RsGDc.MoveNext
        
    Loop
Next

' crea registros para guias saltadas
If RsRepo.RecordCount > 0 Then

    RsRepo.MoveFirst
    m_primera = RsRepo!GD
    RsRepo.MoveLast
    m_ultima = RsRepo!GD
    
    For i = m_primera To m_ultima
        RsRepo.Seek "=", i
        If RsRepo.NoMatch Then
        
            RsRepo.AddNew
            RsRepo!GD = i
            RsRepo!obra = Space(25) & "*****"
            RsRepo.Update
            
        End If
    Next

End If

End Sub
Public Sub Repo_ITOs_de_OTs_General(Nv As Double, RUT_SubC As String, OT As Double)
Dim DbD As Database, RsSc As Recordset
Dim Dbm As Database, RsNVc As Recordset, RsOTc As Recordset, RsITOd As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String
Dim m_obra As String, m_sc As String, m_kilos_ito As Double

NomTabla = "ITOs de OTs"

Set DbD = OpenDatabase(data_file)
Set RsSc = DbD.OpenRecordset("Contratistas")
RsSc.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

qry = MyQuery(Nv, RUT_SubC, "", "", "", Fecha_Vacia, Fecha_Vacia, OT)
'qry = ""
qry = "SELECT * FROM [OT Fab Cabecera]" & qry

Set RsOTc = Dbm.OpenRecordset(qry)
'RsOTc.Index = "Número"

Set RsITOd = Dbm.OpenRecordset("ITO Fab Detalle")
RsITOd.Index = "OT"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

Do While Not RsOTc.EOF
    
    m_obra = ""
    RsNVc.Seek "=", RsOTc!Nv, RsOTc!NvArea
    If Not RsNVc.NoMatch Then
        m_obra = RsNVc![obra]
    End If
    
    m_sc = ""
    RsSc.Seek "=", RsOTc![Rut contratista]
    If Not RsSc.NoMatch Then
        m_sc = RsSc![Razón Social]
    End If
    
    RsITOd.Seek "=", RsOTc!Número
    m_kilos_ito = 0
    If Not RsITOd.NoMatch Then
        Do While Not RsITOd.EOF
            If RsOTc!Número <> RsITOd![Número OT] Then Exit Do
            m_kilos_ito = m_kilos_ito + RsITOd!Cantidad * RsITOd![Peso Unitario]
            RsITOd.MoveNext
        Loop
    End If
    RsRepo.AddNew
    
    RsRepo!Nv = RsOTc!Nv
    RsRepo!obra = m_obra
    RsRepo!SubContratista = m_sc
    RsRepo![OT Nº] = RsOTc!Número
    RsRepo![OT Fecha] = RsOTc!Fecha
    RsRepo![OT Kg Total] = RsOTc![Peso Total]
    RsRepo![ITO Kg Total] = m_kilos_ito
    
    RsRepo.Update
    
    RsOTc.MoveNext
    
Loop

DbD.Close
Dbm.Close
Dbi.Close

End Sub
Public Sub Repo_ITOs_de_OTs_Detalle(Nv As Double, RUT_SubC As String, OT As Double)
Dim DbD As Database, RsSc As Recordset
Dim Dbm As Database, RsNVc As Recordset, RsOTc As Recordset, RsITOd As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String
Dim m_obra As String, m_sc As String, m_ito As Double, m_kilos_ito As Double
Dim m_Fecha_ito As Date, primera As Boolean

NomTabla = "ITOs de OTs"

Set DbD = OpenDatabase(data_file)
Set RsSc = DbD.OpenRecordset("Contratistas")
RsSc.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

qry = MyQuery(Nv, RUT_SubC, "", "", "", Fecha_Vacia, Fecha_Vacia, OT)
qry = "SELECT * FROM [OT Fab Cabecera]" & qry

Set RsOTc = Dbm.OpenRecordset(qry)

Set RsITOd = Dbm.OpenRecordset("ITO Fab Detalle")
RsITOd.Index = "OT"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

Do While Not RsOTc.EOF
    
    m_obra = ""
    RsNVc.Seek "=", RsOTc!Nv, RsOTc!NvArea
    If Not RsNVc.NoMatch Then
        m_obra = RsNVc![obra]
    End If
    
    m_sc = ""
    RsSc.Seek "=", RsOTc![Rut contratista]
    If Not RsSc.NoMatch Then
        m_sc = RsSc![Razón Social]
    End If
    
    RsRepo.AddNew
    RsRepo!Nv = RsOTc!Nv
    RsRepo!obra = m_obra
    RsRepo!SubContratista = m_sc
    RsRepo![OT Nº] = RsOTc!Número
    RsRepo![OT Fecha] = RsOTc!Fecha
    RsRepo![OT Kg Total] = RsOTc![Peso Total]
'    RsRepo.Update
    
    RsITOd.Seek "=", RsOTc!Número
    m_kilos_ito = 0
    primera = True
    If Not RsITOd.NoMatch Then
        m_ito = RsITOd!Número
        Do While Not RsITOd.EOF
            If RsOTc!Número <> RsITOd![Número OT] Then Exit Do
            
            'encontró una ito
            
            If m_ito = RsITOd!Número Then
                m_Fecha_ito = RsITOd!Fecha
                m_kilos_ito = m_kilos_ito + RsITOd!Cantidad * RsITOd![Peso Unitario]
            Else
                If primera Then
                    'primera ito
                    RsRepo![ITO Nº] = m_ito
                    RsRepo![ITO Fecha] = m_Fecha_ito
                    RsRepo![ITO Kg Total] = m_kilos_ito
                    RsRepo.Update
                    primera = False
                Else
                    RsRepo.AddNew
                    RsRepo!Nv = RsOTc!Nv
                    RsRepo!obra = m_obra
                    RsRepo!SubContratista = m_sc
                    RsRepo![OT Nº] = RsOTc!Número
                    RsRepo![OT Fecha] = RsOTc!Fecha
                    RsRepo![OT Kg Total] = 0
                    RsRepo![ITO Nº] = m_ito
                    RsRepo![ITO Fecha] = m_Fecha_ito
                    RsRepo![ITO Kg Total] = m_kilos_ito
                    RsRepo.Update
                End If
                m_ito = RsITOd!Número
                m_kilos_ito = RsITOd!Cantidad * RsITOd![Peso Unitario]
            End If
            
            RsITOd.MoveNext
        Loop
    End If
    
    If m_kilos_ito = 0 Then
        RsRepo.Update
    Else
        If primera Then
            'primera ito
            RsRepo![ITO Nº] = m_ito
            RsRepo![ITO Fecha] = m_Fecha_ito
            RsRepo![ITO Kg Total] = m_kilos_ito
            RsRepo.Update
            primera = False
        Else
            RsRepo.AddNew
            RsRepo!Nv = RsOTc!Nv
            RsRepo!obra = m_obra
            RsRepo!SubContratista = m_sc
            RsRepo![OT Nº] = RsOTc!Número
            RsRepo![OT Fecha] = RsOTc!Fecha
            RsRepo![OT Kg Total] = 0
            RsRepo![ITO Nº] = m_ito
            RsRepo![ITO Fecha] = m_Fecha_ito
            RsRepo![ITO Kg Total] = m_kilos_ito
            RsRepo.Update
        End If
    End If
    
    RsOTc.MoveNext
    
Loop

DbD.Close
Dbm.Close
Dbi.Close

End Sub
Public Sub PiezasPendientesOLD(NV_Numero As Double, Plano As String, ContratistaRut As String)
'MousePointer = vbHourglass
Dim RsOTd As Recordset, m_Rut As String
Dim Dbm As Database, RsNVc As Recordset, RsPlano As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String
Dim m_fmt As String, m_Condi As Boolean
Dim Nv_enNegro As Boolean, m_Pintura As String, Nv_Galvanizada As Boolean

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set RsOTd = Dbm.OpenRecordset("OT fab Detalle")
RsOTd.Index = "Nv-Plano-Marca"

NomTabla = "Piezas Pendientes"

Dim num As Integer

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

qry = "SELECT * FROM [Planos Detalle]"
m_Condi = False

If NV_Numero > 0 Then
    qry = qry & " WHERE[NV]=" & NV_Numero
    m_Condi = True
End If

If Plano <> "" Then
    If m_Condi Then
        qry = qry & " AND plano='" & Plano & "'"
    Else
        qry = qry & " WHERE plano='" & Plano & "'"
        m_Condi = True
    End If
End If

Set RsPlano = Dbm.OpenRecordset(qry)

Do While Not RsPlano.EOF

    'If RsPlano("Cantidad Total") - RsPlano("OT fab") <> 0 Or _
    RsPlano("Cantidad Total") - RsPlano("ITO fab") <> 0 Or _
    RsPlano("Cantidad Total") - RsPlano("GD") <> 0 Then

'If RsPlano!Marca = "BUZ-1-3113-054-CN36" Then
'If RsPlano!Marca = "221-F220" Then
'MsgBox ""
'End If
    
    m_Rut = ""
    
    RsOTd.Seek "=", RsPlano!Nv, m_NvArea, RsPlano!Plano, RsPlano!Marca
    
    If Not RsOTd.NoMatch Then
    
        m_Rut = RsOTd![Rut contratista]
        
        ' condicion contratista
        If ContratistaRut <> "" Then
            If Trim(ContratistaRut) <> Trim(m_Rut) Then
                GoTo NoIncluir
            End If
        End If
        
    End If

    RsRepo.AddNew

    RsRepo!Contratista = Left(Contratista_Lee(SqlRsSc, m_Rut), 10)
'Left(RsSc![Razon Social], 10)

'    RsSc.Seek "=", m_Rut
'    If Not RsSc.NoMatch Then
'        RsRepo!Contratista = Left(RsSc![Razon Social], 10)
'    End If
    
    RsRepo!Nv = RsPlano!Nv
    
    Nv_enNegro = False
    
    RsNVc.Seek "=", RsPlano!Nv, RsPlano!NvArea
    
    If Not RsNVc.NoMatch Then
    
        Nv_Galvanizada = False
    
        If RsNVc!pintura Then
            m_Pintura = "pintada"
        Else
            If RsNVc!galvanizado Then
                m_Pintura = "galvanizada"
                Nv_Galvanizada = True
            Else
                m_Pintura = "en negro"
            End If
        End If
        m_Pintura = " - (" & m_Pintura & ")"
    
        RsRepo!obra = Left(RsNVc![obra] & m_Pintura, 30)
        
        If RsNVc!pintura Or RsNVc!galvanizado Then
        Else
            Nv_enNegro = True
        End If
        
    End If
    
    RsRepo!Plano = RsPlano!Plano
    RsRepo!Rev = RsPlano!Rev
    RsRepo!Marca = RsPlano!Marca
    RsRepo!densidad = RsPlano!densidad
    
    RsOTd.Seek "=", RsPlano!Nv, m_NvArea, RsPlano!Plano, RsPlano!Marca
    If Not RsOTd.NoMatch Then
        m_Rut = RsOTd![Rut contratista]
'        RsSc.Seek "=", m_Rut
'        If Not RsSc.NoMatch Then
'            RsRepo!Contratista = Left(RsSc![Razon Social], 10)
'        End If
        RsRepo!Contratista = Left(Contratista_Lee(SqlRsSc, m_Rut), 10)

    End If
    
    RsRepo!Descripcion = RsPlano!Descripcion
    RsRepo![Peso Unitario] = RsPlano![Peso]
    RsRepo![m2 Unitario] = RsPlano![Superficie]
    
    RsRepo!Total = RsPlano![Cantidad Total]
    
    num = RsPlano![Cantidad Total] - RsPlano![OT fab]
    RsRepo!xFab = num
    
    num = RsPlano![OT fab] - RsPlano![ito fab]
    RsRepo!enFab = num

    If RsPlano![ITO pyg] = 0 Then
        ' no siquiera esta pintada
        num = 0
        'num = RsPlano![ito fab]
    Else
        If RsPlano![ITO pyg] > RsPlano![ito gr] Then ' pintadas sin granallado (sistema antiguo)
            num = 0
        Else
            num = RsPlano![ito fab] - RsPlano![ito gr]
        End If
    End If
    RsRepo!enGR = num

    If RsPlano![ITO pyg] = 0 Then
        ' ni siquiera esta pintada
        num = 0
    Else
        If RsPlano![ITO pyg] > RsPlano![ito pp] Then ' pintadas sin produccion pintura (sistema antiguo)
            num = 0
        Else
            num = RsPlano![ito gr] - RsPlano![ito pp]
        End If
    End If
    RsRepo!enPP = num

    num = 0
    If RsPlano![ITO pyg] = 0 Then
        num = 0
    Else
        num = RsPlano![ito pp] - RsPlano![ITO pyg]
    End If
'    If num > 0 Then ' puede haber pintadas sin itopp (antiguo sistema)
'    End If
    RsRepo!enPin = num
       
    num = RsPlano![ITO pyg] - RsPlano![GD]
    If RsPlano![ITO pyg] = 0 Then
        num = 0 ' para cuando existe solo otfab y solo itofab y solo gd
    End If
    RsRepo!xDesp = num
    
    RsRepo!Desp = RsPlano![GD]
    
    RsRepo.Update
    
'    End If

NoIncluir:

    RsPlano.MoveNext
    
Loop

Dbm.Close
Dbi.Close
                             
End Sub
Public Sub PiezasPendientes(NV_Numero As Double, Plano As String, ContratistaRut As String, PiezasPendientes As String)

' PiezasPendientes:
' T: todas
' P: solo pendientes, todas menos la que estan 100% despachadas

'MousePointer = vbHourglass
Dim RsOTd As Recordset, m_Rut As String
Dim Dbm As Database, RsNVc As Recordset, RsPlano As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String
Dim m_fmt As String, m_Condi As Boolean
Dim Nv_enNegro As Boolean, m_Pintura As String, Nv_Galvanizada As Boolean

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

Set RsOTd = Dbm.OpenRecordset("OT fab Detalle")
RsOTd.Index = "Nv-Plano-Marca"

NomTabla = "Piezas Pendientes"

Dim num As Integer

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

qry = "SELECT * FROM [Planos Detalle]"
m_Condi = False

If NV_Numero > 0 Then
    qry = qry & " WHERE[NV]=" & NV_Numero
    m_Condi = True
End If

If Plano <> "" Then
    If m_Condi Then
        qry = qry & " AND plano='" & Plano & "'"
    Else
        qry = qry & " WHERE plano='" & Plano & "'"
        m_Condi = True
    End If
End If

If PiezasPendientes = "P" Then
    If m_Condi Then
        qry = qry & " AND [cantidad total]>[gd]"
    Else
        qry = qry & " WHERE [cantidad total]>[gd]"
        m_Condi = True
    End If
End If

'Debug.Print qry

Set RsPlano = Dbm.OpenRecordset(qry)

Do While Not RsPlano.EOF
    
    m_Rut = ""
    
    RsOTd.Seek "=", RsPlano!Nv, m_NvArea, RsPlano!Plano, RsPlano!Marca
    
    If Not RsOTd.NoMatch Then
    
        m_Rut = RsOTd![Rut contratista]
        
        ' condicion contratista
        If ContratistaRut <> "" Then
            If Trim(ContratistaRut) <> Trim(m_Rut) Then
                GoTo NoIncluir
            End If
        End If
        
    End If

    RsRepo.AddNew

    RsRepo!Contratista = Left(Contratista_Lee(SqlRsSc, m_Rut), 10)
    
    RsRepo!Nv = RsPlano!Nv
    
    Nv_enNegro = False
    
    RsNVc.Seek "=", RsPlano!Nv, RsPlano!NvArea
    
    If Not RsNVc.NoMatch Then
    
        Nv_Galvanizada = False
    
        If RsNVc!pintura Then
            m_Pintura = "pintada"
        Else
            If RsNVc!galvanizado Then
                m_Pintura = "galvanizada"
                Nv_Galvanizada = True
            Else
                m_Pintura = "en negro"
            End If
        End If
        m_Pintura = " - (" & m_Pintura & ")"
    
        RsRepo!obra = Left(RsNVc![obra] & m_Pintura, 30)
        
        If RsNVc!pintura Or RsNVc!galvanizado Then
        Else
            Nv_enNegro = True
        End If
        
    End If
    
    RsRepo!Plano = RsPlano!Plano
    RsRepo!Rev = RsPlano!Rev
    RsRepo!Marca = RsPlano!Marca
    RsRepo!densidad = RsPlano!densidad
    
    RsOTd.Seek "=", RsPlano!Nv, m_NvArea, RsPlano!Plano, RsPlano!Marca
    If Not RsOTd.NoMatch Then
        m_Rut = RsOTd![Rut contratista]
'        RsSc.Seek "=", m_Rut
'        If Not RsSc.NoMatch Then
'            RsRepo!Contratista = Left(RsSc![Razon Social], 10)
'        End If
        RsRepo!Contratista = Left(Contratista_Lee(SqlRsSc, m_Rut), 10)

    End If
    
    RsRepo!Descripcion = RsPlano!Descripcion
    RsRepo![Peso Unitario] = RsPlano![Peso]
    RsRepo![m2 Unitario] = RsPlano![Superficie]
    
    RsRepo!Total = RsPlano![Cantidad Total]
    
    
    '////////////////////////////////////////////////////////////////
    If RsPlano![Cantidad Total] <= RsPlano![GD] Then
        ' esta todo despachado
        GoTo TodoDespachado
    End If
        
    ' no esta todo despachado
    
If RsPlano![Marca] = "M2" Then
    MsgBox ""
End If
    ' se despachó
    If RsPlano![GD] > 0 Then
        ' y no hay granallado, ni produccion pintura
        'If RsPlano![ito gr] = 0 And RsPlano![ito pp] = 0 And RsPlano![ito pyg] = 0 Then
        If RsPlano![ito gr] = 0 And RsPlano![ito pp] = 0 Then
        
            ' caso antiguo
            If RsPlano![ITO pyg] = 0 Then
                ' no [ito gr], no [ito pp], no [ito pyg]
                ' es decir antiguas nv no pintadas, pero despachadas parcialmente
                RsRepo!xDesp = RsPlano![GD] - RsPlano![ito fab]
            Else
                ' no [ito gr], no [ito pp], pero si [ito pyg]
                ' es decir antiguas nv pintadas, pero despachadas parcialmente
                RsRepo!xDesp = RsPlano![GD] - RsPlano![ITO pyg]
            End If

            GoTo Grabar
            
        End If
    End If
      
Grabar:

    RsRepo!xFab = RsPlano![Cantidad Total] - RsPlano![OT fab]
    If RsPlano![OT fab] >= RsPlano![ito fab] Then
        RsRepo!enFab = RsPlano![OT fab] - RsPlano![ito fab]
    End If
    
    RsRepo!enGR = RsPlano![ito fab] - RsPlano![ito gr]
    
    RsRepo!enPP = RsPlano![ito gr] - RsPlano![ito pp]
    
    'RsRepo!enPin = RsPlano![ITO pp] - RsPlano![ITO pyg]
    
    RsRepo!xDesp = RsPlano![ito pp] - RsPlano![GD]
    
TodoDespachado:

    RsRepo!Desp = RsPlano![GD]
    
    RsRepo.Update
    
'    End If

NoIncluir:

    RsPlano.MoveNext
    
Loop

Dbm.Close
Dbi.Close
                             
End Sub
Public Sub PiezasxAsignar(Nv As Double, Plano As String)

Dim Dbm As Database, RsNVc As Recordset, RsPd As Recordset
Dim Dbi As Database, RsRepo As Recordset

Dim NomTabla As String
Dim m_NVnom As String
Dim num As Integer

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

qry = MyQuery(Nv, "", "", Plano, "", Fecha_Vacia, Fecha_Vacia, 0)
qry = "SELECT * FROM [Planos Detalle]" & qry
Set RsPd = Dbm.OpenRecordset(qry)

NomTabla = "Piezas x Asignar"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

Do While Not RsPd.EOF
    
    num = Val(RsPd![Cantidad Total]) - Val(RsPd![OT fab])
    If num <> 0 Then
        
        RsNVc.Seek "=", RsPd!Nv, RsPd!NvArea
        If Not RsNVc.NoMatch Then
            m_NVnom = RsNVc![obra]
        End If
        
        RsRepo.AddNew
        RsRepo!Nv = RsPd!Nv
        RsRepo!obra = m_NVnom
        RsRepo!Plano = RsPd!Plano
        RsRepo!Rev = RsPd!Rev
        RsRepo!Marca = RsPd!Marca
        RsRepo!Descripcion = Left(RsPd!Descripcion, 20)
        RsRepo!Total = RsPd![Cantidad Total]
        RsRepo![x Asignar] = num
        RsRepo![Peso Unitario] = RsPd![Peso]
        RsRepo![Peso Total] = num * RsPd![Peso]
        RsRepo.Update
            
    End If
    
    RsPd.MoveNext
    
Loop

Dbm.Close
Dbi.Close

End Sub
Public Sub PiezasxRecibir(Nv As Double, RUT_SubC As String, Plano As String, Fecha_Ini As String, Fecha_Fin As String)

Dim DbD As Database, RsSc As Recordset
Dim Dbm As Database, RsPd As Recordset, RsNVc As Recordset, RsOTc As Recordset, RsOTd As Recordset, RsITOfd As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String
Dim m_NVnom As String, m_ScNom As String, m_desc As String, m_Rev As String
Dim num As Integer, recib As Integer

Dim Muestra As Boolean

Set DbD = OpenDatabase(data_file)
Set RsSc = DbD.OpenRecordset("Contratistas")
RsSc.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)
Set RsPd = Dbm.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Marca"

Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Numero"

Set RsOTc = Dbm.OpenRecordset("OT Fab Cabecera")
RsOTc.Index = "Numero"

Set RsITOfd = Dbm.OpenRecordset("ITO Fab Detalle")
RsITOfd.Index = "NV-Plano-Marca"

qry = MyQuery(Nv, RUT_SubC, "", Plano, "Fecha Entrega", Fecha_Ini, Fecha_Fin, 0)
qry = "SELECT * FROM [OT Fab Detalle]" & qry
'qry = " ORDER BY nv" ' agregado 21/10/10
Set RsOTd = Dbm.OpenRecordset(qry)

NomTabla = "Piezas x Recibir"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

Do While Not RsOTd.EOF

    Muestra = False
    If Usuario.Nv_Activas Then
        RsNVc.Seek "=", RsOTd!Nv, 0
        If Not RsNVc.NoMatch Then
            Muestra = RsNVc!Activa
        End If
    Else
        Muestra = True
    End If

    If Muestra Then

    '    If RsOTd!Marca = "PE1" Then
    '    MsgBox ""
    '    End If
        
        num = Val(RsOTd!Cantidad) - Val(RsOTd![Cantidad Recibida])
    '    num = Val(RsOTd!Cantidad)
    '    If num < 0 Then
    '        MsgBox num
           ' revisar ITO RECEPCION
    '        num = 0 'Val(RsOTd!Cantidad)
    '    End If
    'If RsOTd!Numero = 13082 Then
    '   MsgBox ""
    'End If
        recib = 0
        If False Then
        'If True Then
          ' busca recibidas en itof
          RsITOfd.Seek "=", RsOTd!Nv, RsOTd!Plano, RsOTd!Marca
          If Not RsITOfd.NoMatch Then
             Do While Not RsITOfd.EOF
                If RsOTd!Nv <> RsITOfd!Nv Or RsOTd!Plano <> RsITOfd!Plano Or RsOTd!Marca <> RsITOfd!Marca Then Exit Do
                   If RsOTd!Numero = RsITOfd![Numero OT] Then
                      recib = recib + RsITOfd!Cantidad
                   End If
                RsITOfd.MoveNext
             Loop
          End If
        End If

        If num <> 0 Then
    '    If num > recib Then
            
            RsNVc.Seek "=", RsOTd!Nv, RsOTd!NvArea
            If Not RsNVc.NoMatch Then
                m_NVnom = RsNVc![obra]
            End If
            
            m_ScNom = Contratista_Lee(SqlRsSc, RsOTd![Rut contratista])

    '        If RsOTd!Marca = "F2033V2" Then
    '        MsgBox ""
    '      endif

            RsPd.Seek "=", RsOTd!Nv, m_NvArea, RsOTd!Plano, RsOTd!Marca
            m_desc = ""
            If Not RsPd.NoMatch Then
                m_Rev = RsPd!Rev
                m_desc = RsPd!Descripcion
            End If

            RsRepo.AddNew
            RsRepo!Nv = RsOTd!Nv
            RsRepo!Plano = RsOTd!Plano
    '        RsRepo!Rev = RsOTd!Rev ' revision en OT
            RsRepo!Rev = m_Rev ' revision den palons detalle
            RsRepo!Marca = RsOTd!Marca
            RsRepo!Descripcion = m_desc
            RsRepo!obra = m_NVnom
            RsRepo!Contratista = m_ScNom
            RsRepo![Nº OT] = RsOTd!Numero
            RsRepo!Fecha = RsOTd!Fecha
    '        RsRepo![x Recibir] = num - recib ' ojo
            RsRepo![x Recibir] = num  ' ojo
            RsRepo![Fecha Entrega] = RsOTd![Fecha Entrega]
            RsRepo![Peso Unitario] = RsOTd![Peso Unitario]
    '        RsRepo![Peso Total] = (num - recib) * RsOTd![Peso Unitario]
            RsRepo![Peso Total] = (num) * RsOTd![Peso Unitario]
            RsRepo.Update

        End If

    End If ' if muestra (x nv activas)

    RsOTd.MoveNext

Loop

DbD.Close
Dbm.Close
Dbi.Close

End Sub
Public Sub PiezasxPintar(Nv As Double, Plano As String)

Dim Dbm As Database, RsNVc As Recordset, RsPd As Recordset
Dim Dbi As Database, RsRepo As Recordset

Dim NomTabla As String
Dim m_NVnom As String
Dim num As Integer

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

qry = MyQuery(Nv, "", "", Plano, "", Fecha_Vacia, Fecha_Vacia, 0)
qry = "SELECT * FROM [Planos Detalle]" & qry
Set RsPd = Dbm.OpenRecordset(qry)

NomTabla = "Piezas x Asignar"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

Do While Not RsPd.EOF
    
    num = Val(RsPd![ito fab]) - Val(RsPd![ITO pyg])
    If num <> 0 Then
        
        RsNVc.Seek "=", RsPd!Nv, RsPd!NvArea
        If Not RsNVc.NoMatch Then
            m_NVnom = RsNVc![obra]
        End If
        
        RsRepo.AddNew
        RsRepo!Nv = RsPd!Nv
        RsRepo!obra = m_NVnom
        RsRepo!Plano = RsPd!Plano
        RsRepo!Rev = RsPd!Rev
        RsRepo!Marca = RsPd!Marca
        RsRepo!descripción = Left(RsPd!descripción, 20)
        RsRepo!Total = RsPd![Cantidad Total]
        RsRepo![x Asignar] = num
        RsRepo![Peso Unitario] = RsPd![Peso]
        RsRepo![Peso Total] = num * RsPd![Peso]
        RsRepo.Update
            
    End If
    
    RsPd.MoveNext
    
Loop

Dbm.Close
Dbi.Close

End Sub
Public Sub PiezasxDespachar(Nv As Double, Plano As String, xDespachar As Boolean, TipoRep As String)

Dim Dbm As Database, RsNVc As Recordset, RsPd As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String, m_cantidad As Integer, m_NVnom As String, m_NVenNegro As Boolean

Set Dbm = OpenDatabase(mpro_file)

Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
'RsNVc.Index = Nv_Index ' "Número"
RsNVc.Index = "Numero"

qry = MyQuery(Nv, "", "", Plano, "", Fecha_Vacia, Fecha_Vacia, 0)
qry = "SELECT * FROM [Planos Detalle]" & qry
Set RsPd = Dbm.OpenRecordset(qry)
'RsPd.Index = "NV-Plano-Marca"

NomTabla = "Piezas x Despachar"
Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

Do While Not RsPd.EOF

'    If RsPd![ITO pyg] < RsPd![GD] Then
'        Debug.Print RsPd!Nv, RsPd!Plano, RsPd!Marca, RsPd![GD], RsPd![ITO pyg]
'    End If

    m_NVenNegro = False
    RsNVc.Seek "=", RsPd!Nv, RsPd!NvArea
    If Not RsNVc.NoMatch Then
        m_NVnom = RsNVc![obra]
        If RsNVc!pintura Or RsNVc!galvanizado Then
            m_NVenNegro = False
        Else
            m_NVenNegro = True
        End If
    End If

'    If RsPd!Nv = 769 And RsPd!Plano = "P101-F1" And RsPd!Marca = "F1" Then
'        MsgBox ""
'    End If

    Select Case TipoRep
    Case "ennegro"
'    If xDespachar Then ' repo en negro y repo pintadas
'        m_cantidad = RsPd![ITO fab] - RsPd![GD]

        If m_NVenNegro Then
            m_cantidad = RsPd![ito fab] - RsPd![GD]
        Else
'            m_cantidad = RsPd![ITO pyg] - RsPd![GD]
            m_cantidad = RsPd![ito fab] - RsPd![ITO pyg]
            
            ' correccion 27/03/09
            If RsPd!GD >= RsPd![Cantidad Total] Then m_cantidad = 0
            
        End If
        
        ' independiente de lo que diga NV
        ' pedido por sruz 24/06/14
        'If False Then
            'm_cantidad = 0
            'If RsPd![ito gr] > 0 Then
                'm_cantidad = RsPd![Cantidad Total] - RsPd![ito pp]
                'm_cantidad = RsPd![ito gr] - RsPd![ito pp]
                m_cantidad = RsPd![ito fab] - RsPd![ito gr]
                'm_cantidad = RsPd![ito fab] - RsPd![ITO pp]
                'Debug.Print RsPd!Marca, RsPd![ito gr], RsPd![ito pp], m_cantidad
                
            'End If
        'End If
        
    Case "pintadas"

        If m_NVenNegro Then
            m_cantidad = 0
        Else
            m_cantidad = RsPd![ITO pyg] - RsPd![GD]
            If m_cantidad < 0 Then m_cantidad = 0
        End If
        
    Case "engalvanizado"

        If m_NVenNegro Then
            m_cantidad = 0
        Else
            m_cantidad = RsPd![GD gal] - RsPd![ITO pyg]
            If m_cantidad < 0 Then m_cantidad = 0
        End If
        
    Case Else ' despachadas
        m_cantidad = RsPd![GD] ' solo para listado de piezas despachadas
    End Select
        
    If m_cantidad <> 0 Then
    
        RsRepo.AddNew
        RsRepo!Nv = RsPd!Nv
        RsRepo!Plano = RsPd!Plano
        RsRepo!Rev = RsPd!Rev
        RsRepo!Marca = RsPd!Marca
        RsRepo!Descripcion = Left(RsPd!Descripcion, 20)
        RsRepo!obra = m_NVnom
        RsRepo!Total = RsPd![Cantidad Total]
        RsRepo![x Despachar] = m_cantidad
        RsRepo![Peso Unitario] = RsPd![Peso]
        If m_cantidad > 0 Then
            RsRepo![Peso Total] = m_cantidad * RsPd![Peso]
        End If
        RsRepo.Update
    End If
    
    RsPd.MoveNext
    
Loop

Dbm.Close
Dbi.Close

Exit Sub

'Dim Dbm As Database,RsNVc As Recordset
Dim RsITOd As Recordset, RsGDd As Recordset
'Dim Dbi As Database, RsRepo As Recordset
'Dim NomTabla As String
'Dim m_NVnom As String
Dim m_Nv As Double, m_Plano As String, m_Marca As String ', m_Cantidad As Integer

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

'qry = MyQuery(NV, "", "", "", fecha_vacia, fecha_vacia)
'qry = "SELECT * FROM [ITO Fab Detalle]" & qry

Set RsITOd = Dbm.OpenRecordset("ITO Fab Detalle")
RsITOd.Index = "NV-Plano-Marca"

Set RsGDd = Dbm.OpenRecordset("GD Detalle")
RsGDd.Index = "NV-Plano-Marca"

NomTabla = "Piezas x Despachar"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

m_Nv = RsITOd!Nv
m_Plano = RsITOd!Plano
m_Marca = RsITOd!Marca

Do While Not RsITOd.EOF

    If m_Nv = RsITOd!Nv And m_Plano = RsITOd!Plano And m_Marca = RsITOd!Marca Then
        m_cantidad = m_cantidad + RsITOd!Cantidad
    Else
    
        m_cantidad = m_cantidad - Despachada_Buscar(m_Nv, m_Plano, m_Marca)
        
        If m_cantidad <> 0 Then
        
            RsNVc.Seek "=", RsITOd!Nv, RsITOd!NvArea
            If Not RsNVc.NoMatch Then
                m_NVnom = RsNVc!obra
            End If
        
            RsRepo.AddNew
            RsRepo!Nv = m_Nv
            RsRepo!Plano = m_Plano
'            RsRepo!Rev = RsITOd!Rev
            RsRepo!Marca = m_Marca
            RsRepo!obra = m_NVnom
'            RsRepo("Nº ITO") = RsITOd!Número
'            RsRepo!Fecha = RsITOd!Fecha
            RsRepo![x Despachar] = m_cantidad
            RsRepo!Peso = m_cantidad * RsITOd![Peso Unitario]
            RsRepo.Update
            
            m_Nv = RsITOd!Nv
            m_Plano = RsITOd!Plano
            m_Marca = RsITOd!Marca
            m_cantidad = RsITOd!Cantidad
            
        End If
        
    End If
    
    RsITOd.MoveNext
    
Loop

Dbm.Close
Dbi.Close

End Sub
Private Function Despachada_Buscar(p_NV As Double, p_Plano As String, p_Marca As String)
Despachada_Buscar = 0
End Function
Public Sub xPiezasenGalvanizado(Nv As Double, Plano As String, xDespachar As Boolean, TipoRep As String)

Dim Dbm As Database, RsNVc As Recordset, RsPd As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String, m_cantidad As Integer, m_NVnom As String, m_NVenNegro As Boolean

Set Dbm = OpenDatabase(mpro_file)

Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
'RsNVc.Index = Nv_Index ' "Número"
RsNVc.Index = "Número"

qry = MyQuery(Nv, "", "", "", "", Fecha_Vacia, Fecha_Vacia, 0)
qry = "SELECT * FROM [Planos Detalle]" & qry
Set RsPd = Dbm.OpenRecordset(qry)
'RsPd.Index = "NV-Plano-Marca"

NomTabla = "Piezas x Despachar"
Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

Do While Not RsPd.EOF

'    If RsPd![ITO pyg] < RsPd![GD] Then
'        Debug.Print RsPd!Nv, RsPd!Plano, RsPd!Marca, RsPd![GD], RsPd![ITO pyg]
'    End If

    m_NVenNegro = False
    RsNVc.Seek "=", RsPd!Nv, RsPd!NvArea
    If Not RsNVc.NoMatch Then
        m_NVnom = RsNVc![obra]
        If RsNVc!pintura Or RsNVc!galvanizado Then
            m_NVenNegro = False
        Else
            m_NVenNegro = True
        End If
    End If

'    If RsPd!Nv = 769 And RsPd!Plano = "P101-F1" And RsPd!Marca = "F1" Then
'        MsgBox ""
'    End If

    Select Case TipoRep
    Case "ennegro"
'    If xDespachar Then ' repo en negro y repo pintadas
'        m_cantidad = RsPd![ITO fab] - RsPd![GD]

        If m_NVenNegro Then
            m_cantidad = RsPd![ito fab] - RsPd![GD]
        Else
'            m_cantidad = RsPd![ITO pyg] - RsPd![GD]
            m_cantidad = RsPd![ito fab] - RsPd![ITO pyg]
        End If
        
    Case "pintadas"

        If m_NVenNegro Then
            m_cantidad = 0
        Else
            m_cantidad = RsPd![ITO pyg] - RsPd![GD]
            If m_cantidad < 0 Then m_cantidad = 0
        End If
        
    Case Else ' despachadas
        m_cantidad = RsPd![GD] ' solo para listado de piezas despachadas
    End Select
        
    If m_cantidad <> 0 Then
    
        RsRepo.AddNew
        RsRepo!Nv = RsPd!Nv
        RsRepo!Plano = RsPd!Plano
        RsRepo!Rev = RsPd!Rev
        RsRepo!Marca = RsPd!Marca
        RsRepo!descripción = Left(RsPd!descripción, 20)
        RsRepo!obra = m_NVnom
        RsRepo!Total = RsPd![Cantidad Total]
        RsRepo![x Despachar] = m_cantidad
        RsRepo![Peso Unitario] = RsPd![Peso]
        If m_cantidad > 0 Then
            RsRepo![Peso Total] = m_cantidad * RsPd![Peso]
        End If
        RsRepo.Update
    End If
    
    RsPd.MoveNext
    
Loop

Dbm.Close
Dbi.Close

Exit Sub

'Dim Dbm As Database,RsNVc As Recordset
Dim RsITOd As Recordset, RsGDd As Recordset
'Dim Dbi As Database, RsRepo As Recordset
'Dim NomTabla As String
'Dim m_NVnom As String
Dim m_Nv As Double, m_Plano As String, m_Marca As String ', m_Cantidad As Integer

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

'qry = MyQuery(NV, "", "", "", fecha_vacia, fecha_vacia)
'qry = "SELECT * FROM [ITO Fab Detalle]" & qry

Set RsITOd = Dbm.OpenRecordset("ITO Fab Detalle")
RsITOd.Index = "NV-Plano-Marca"

Set RsGDd = Dbm.OpenRecordset("GD Detalle")
RsGDd.Index = "NV-Plano-Marca"

NomTabla = "Piezas x Despachar"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

m_Nv = RsITOd!Nv
m_Plano = RsITOd!Plano
m_Marca = RsITOd!Marca

Do While Not RsITOd.EOF

    If m_Nv = RsITOd!Nv And m_Plano = RsITOd!Plano And m_Marca = RsITOd!Marca Then
        m_cantidad = m_cantidad + RsITOd!Cantidad
    Else
    
        m_cantidad = m_cantidad - Despachada_Buscar(m_Nv, m_Plano, m_Marca)
        
        If m_cantidad <> 0 Then
        
            RsNVc.Seek "=", RsITOd!Nv, RsITOd!NvArea
            If Not RsNVc.NoMatch Then
                m_NVnom = RsNVc!obra
            End If
        
            RsRepo.AddNew
            RsRepo!Nv = m_Nv
            RsRepo!Plano = m_Plano
'            RsRepo!Rev = RsITOd!Rev
            RsRepo!Marca = m_Marca
            RsRepo!obra = m_NVnom
'            RsRepo("Nº ITO") = RsITOd!Número
'            RsRepo!Fecha = RsITOd!Fecha
            RsRepo![x Despachar] = m_cantidad
            RsRepo!Peso = m_cantidad * RsITOd![Peso Unitario]
            RsRepo.Update
            
            m_Nv = RsITOd!Nv
            m_Plano = RsITOd!Plano
            m_Marca = RsITOd!Marca
            m_cantidad = RsITOd!Cantidad
            
        End If
        
    End If
    
    RsITOd.MoveNext
    
Loop

Dbm.Close
Dbi.Close

End Sub
Public Sub Repo_OTfxPlanoG(Nv As Double, Plano As String, RUT_SubC As String)
Dim DbD As Database, RsSc As Recordset
Dim Dbm As Database, RsNVc As Recordset, RsOTfd As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String
Dim m_Nv As Double, m_obra As String, m_Plano As String, m_Rev As String
Dim m_ot As Double, m_OTFecha As Date, m_OTEntrega As Date, m_sc As String, m_RutSc As String
Dim m_KTotal As Double, m_KRecib As Double, m_Nueva As String

NomTabla = "OTf x Plano"

Set DbD = OpenDatabase(data_file)
Set RsSc = DbD.OpenRecordset("Contratistas")
RsSc.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

qry = MyQuery(Nv, RUT_SubC, "", "", "", Fecha_Vacia, Fecha_Vacia, 0)
If Len(qry) = 0 Then
qry = "SELECT * FROM [OT Fab Detalle]" & qry & " WHERE Cantidad<>[Cantidad Recibida] ORDER BY NV,Plano,Número"
Else
qry = "SELECT * FROM [OT Fab Detalle]" & qry & " AND Cantidad<>[Cantidad Recibida] ORDER BY NV,Plano,Número"
End If

Set RsOTfd = Dbm.OpenRecordset(qry)

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

If RsOTfd.RecordCount > 0 Then
    m_Nv = RsOTfd!Nv
    m_Plano = RsOTfd!Plano
    m_Rev = RsOTfd!Rev
    m_ot = RsOTfd!Número
    m_RutSc = RsOTfd![Rut contratista]
    m_OTFecha = RsOTfd!Fecha
    m_OTEntrega = RsOTfd![Fecha Entrega]
    m_KTotal = 0
    m_KRecib = 0
    
    m_Nueva = m_Nv & m_Plano & m_ot
    
    Do While Not RsOTfd.EOF
        
        If m_Nueva = RsOTfd!Nv & RsOTfd!Plano & RsOTfd!Número Then
            If m_OTEntrega > RsOTfd![Fecha Entrega] Then m_OTEntrega = RsOTfd![Fecha Entrega]
            m_KTotal = m_KTotal + RsOTfd!Cantidad * RsOTfd![Peso Unitario]
            m_KRecib = m_KRecib + RsOTfd![Cantidad Recibida] * RsOTfd![Peso Unitario]
        Else
        
            ' graba OT
            m_obra = ""
            RsNVc.Seek "=", m_Nv, m_NvArea
            If Not RsNVc.NoMatch Then
                m_obra = RsNVc!obra
            End If
            
            m_sc = ""
            RsSc.Seek "=", m_RutSc
            If Not RsSc.NoMatch Then
                m_sc = RsSc![Razón Social]
            End If
        
            RsRepo.AddNew
            
            RsRepo!Nv = m_Nv
            RsRepo!obra = m_obra
            RsRepo!Plano = m_Plano
            RsRepo!Rev = m_Rev
            RsRepo![Nº OT] = m_ot
            RsRepo![Fecha OT] = m_OTFecha
            RsRepo![Fecha Entrega] = m_OTEntrega
            RsRepo!SubContratista = m_sc
            RsRepo![Kg Total] = m_KTotal
            RsRepo![Kg Recibidos] = m_KRecib
            If m_KTotal <> 0 Then RsRepo![% Avance] = m_KRecib * 100 / m_KTotal
            
            RsRepo.Update
            '////////////
            
            m_Nv = RsOTfd!Nv
            m_Plano = RsOTfd!Plano
            m_Rev = RsOTfd!Rev
            m_ot = RsOTfd!Número
            m_RutSc = RsOTfd![Rut contratista]
            m_OTFecha = RsOTfd!Fecha
            m_OTEntrega = RsOTfd![Fecha Entrega]
            m_KTotal = RsOTfd!Cantidad * RsOTfd![Peso Unitario]
            m_KRecib = RsOTfd![Cantidad Recibida] * RsOTfd![Peso Unitario]
            m_Nueva = RsOTfd!Nv & RsOTfd!Plano & RsOTfd!Número
            
        End If
        
        RsOTfd.MoveNext
        
    Loop

    ' graba OT
    '////////////
    m_obra = ""
    RsNVc.Seek "=", m_Nv, m_NvArea
    If Not RsNVc.NoMatch Then
        m_obra = RsNVc!obra
    End If
    
    m_sc = ""
    RsSc.Seek "=", m_RutSc
    If Not RsSc.NoMatch Then
        m_sc = RsSc![Razón Social]
    End If

    RsRepo.AddNew
    
    RsRepo!Nv = m_Nv
    RsRepo!obra = m_obra
    RsRepo!Plano = m_Plano
    RsRepo!Rev = m_Rev
    RsRepo![Nº OT] = m_ot
    RsRepo![Fecha OT] = m_OTFecha
    RsRepo![Fecha Entrega] = m_OTEntrega
    RsRepo!SubContratista = m_sc
    RsRepo![Kg Total] = m_KTotal
    RsRepo![Kg Recibidos] = m_KRecib
    If m_KTotal <> 0 Then RsRepo![% Avance] = m_KRecib * 100 / m_KTotal
    RsRepo.Update
    '////////////

End If

DbD.Close
Dbm.Close
Dbi.Close

End Sub
Public Sub Repo_OTfxPlanoD(Nv As Double, Plano As String, RUT_SubC As String)
Dim DbD As Database, RsSc As Recordset
Dim Dbm As Database, RsNVc As Recordset, RsOTfd As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim NomTabla As String
Dim m_obra As String, m_sc As String
Dim m_KTotal As Double, m_KRecib As Double

NomTabla = "OTf x Plano"

Set DbD = OpenDatabase(data_file)
Set RsSc = DbD.OpenRecordset("Contratistas")
RsSc.Index = "RUT"

Set Dbm = OpenDatabase(mpro_file)
Set RsNVc = Dbm.OpenRecordset("NV Cabecera")
RsNVc.Index = Nv_Index ' "Número"

qry = MyQuery(Nv, RUT_SubC, "", "", "", Fecha_Vacia, Fecha_Vacia, 0)
If Len(qry) = 0 Then
qry = "SELECT * FROM [OT Fab Detalle]" & qry & " WHERE Cantidad<>[Cantidad Recibida] ORDER BY NV,Plano,Número"
Else
qry = "SELECT * FROM [OT Fab Detalle]" & qry & " AND Cantidad<>[Cantidad Recibida] ORDER BY NV,Plano,Número"
End If

Set RsOTfd = Dbm.OpenRecordset(qry)

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

If RsOTfd.RecordCount > 0 Then

    Do While Not RsOTfd.EOF
        
        ' graba OT
        m_obra = ""
        RsNVc.Seek "=", RsOTfd!Nv, RsOTfd!NvArea
        If Not RsNVc.NoMatch Then
            m_obra = RsNVc!obra
        End If
        
        m_sc = ""
        RsSc.Seek "=", RsOTfd![Rut contratista]
        If Not RsSc.NoMatch Then
            m_sc = RsSc![Razon Social]
        End If
    
        RsRepo.AddNew
        
        RsRepo!Nv = RsOTfd!Nv
        RsRepo!obra = m_obra
        RsRepo!Plano = RsOTfd!Plano
        RsRepo!Rev = RsOTfd!Rev
        RsRepo!Marca = RsOTfd!Marca
        RsRepo![Nº OT] = RsOTfd!Número
        RsRepo![Fecha OT] = RsOTfd!Fecha
        RsRepo![Fecha Entrega] = RsOTfd![Fecha Entrega]
        RsRepo!SubContratista = m_sc
        m_KTotal = RsOTfd!Cantidad * RsOTfd![Peso Unitario]
        RsRepo![Kg Total] = m_KTotal
        m_KRecib = RsOTfd![Cantidad Recibida] * RsOTfd![Peso Unitario]
        RsRepo![Kg Recibidos] = m_KRecib
        If m_KTotal <> 0 Then
            RsRepo![% Avance] = m_KRecib * 100 / m_KTotal
        End If
        
        RsRepo.Update
        
        RsOTfd.MoveNext
        
    Loop

End If

DbD.Close
Dbm.Close
Dbi.Close

End Sub
Public Sub Repo_Piezas(Nv As Double, RUT_SubC As String)
' 06/09/10
Dim Dbm As Database, RsNVc As Recordset, RsOTfd As Recordset
Dim sql As String, condi As String
'Dim RsNVc As Recordset

'sql = "SELECT * FROM [nv cabecera]"

'Set RsNVc = ""

Set Dbm = OpenDatabase(mpro_file)
sql = "SELECT * FROM [ot fab detalle]"

condi = ""
If Nv > 0 Then
    condi = " WHERE nv=" & Nv
End If
If RUT_SubC <> "" Then
    If condi = "" Then
        condi = " WHERE rut='" & RUT_SubC & "'"
    Else
        condi = " AND rut='" & RUT_SubC & "'"
    End If
End If
sql = sql & condi

Set RsOTfd = Dbm.OpenRecordset(sql)

With RsOTfd
Do While Not .EOF
    Debug.Print !Numero
    .MoveNext
Loop
End With


End Sub
Private Function MyQuery(Nv As Double, RUT_SubC As String, RUT_Client As String, Plano As String, Campo_fecha As String, F_Inicial As String, F_Final As String, OT As Double) As String
Dim p As Integer, qry As String, QryN As String
p = 0: qry = ""

' NOTA DE VENTA
If Nv <> 0 Then
    qry = " NV=CDbl(" & Nv & ")"
    p = p + 1
End If

' CONTRATISTA
If Usuario.Tipo = "C" Then
    RUT_SubC = Usuario.rut
End If
If RUT_SubC <> "" Then
    QryN = " [RUT Contratista]='" & RUT_SubC & "'"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If

' CLIENTE
If RUT_Client <> "" Then
    QryN = " [RUT Cliente]='" & RUT_Client & "'"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If

' PLANO
If Plano <> "" Then
    QryN = " Plano='" & Plano & "'"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If

' FECHA INICIAL
'If F_Inicial <> "__/__/__" Then
If F_Inicial <> "__/__/__" And F_Inicial <> "" Then
    QryN = " [" & Campo_fecha & "]>=CDate('" & F_Inicial & "')"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If

' FECHA FINAL
'If F_Final <> "__/__/__" Then
If F_Final <> "__/__/__" And F_Final <> "" Then
    QryN = " [" & Campo_fecha & "]<=CDate('" & F_Final & "')"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If

' OT
If OT <> 0 Then
    QryN = " Número=CDbl(" & OT & ")"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If

If p > 0 Then qry = " WHERE" & qry
MyQuery = qry

End Function
Private Function SQL_MyQuery(Campo_fecha As String, F_Inicial As String, F_Final As String, campoRut As String, rut As String) As String
' query con sintaxis sql server
Dim p As Integer, qry As String, QryN As String
p = 0: qry = ""

If F_Inicial <> "__/__/__" And F_Inicial <> "" Then
    QryN = " [" & Campo_fecha & "]>='" & Format(F_Inicial, sql_Fecha_Formato) & "'"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If

' FECHA FINAL
If F_Final <> "__/__/__" And F_Final <> "" Then
    QryN = " [" & Campo_fecha & "]<='" & Format(F_Final, sql_Fecha_Formato) & "'"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If

If rut <> "" Then
    QryN = " [" & campoRut & "]='" & rut & "'"
    qry = IIf(p = 0, QryN, qry & " AND " & QryN)
    p = p + 1
End If

If p > 0 Then qry = " WHERE" & qry
SQL_MyQuery = qry

End Function
Public Sub Repo_OcxTipo(Nv As Double)
' puebla archivo repo
' con compras por tipo de producto o mejor dicho tipo proveedor
' acero
' pintura
' otros
Dim DbD As Database, RsPrv As Recordset, RsClaP As Recordset
Dim DbAdq As Database, RsOCdN As Recordset
Dim Dbi As Database, RsRepo As Recordset
Dim m_Rut As String, m_Clas As String, sql As String

Set DbD = OpenDatabase(data_file)
Set RsPrv = DbD.OpenRecordset("Proveedores")
RsPrv.Index = "RUT"
Set RsClaP = DbD.OpenRecordset("Clasificacion de Proveedores")
RsClaP.Index = "Codigo"

Set DbAdq = OpenDatabase(Madq_file)
sql = "SELECT * FROM [OC Detalle] WHERE Nv=" & Nv
Set RsOCdN = DbAdq.OpenRecordset(sql)

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset("Compras")
RsRepo.Index = "Tipo"

' borra tabla de paso
Dbi.Execute "DELETE * FROM compras"

With RsOCdN
Do While Not .EOF
    m_Rut = ![RUT Proveedor]
    RsPrv.Seek "=", m_Rut
    m_Clas = ""
    If Not RsPrv.NoMatch Then
        m_Clas = NoNulo(RsPrv!Clasificacion)
        If m_Clas = "" Then m_Clas = "OTRO"
        RsRepo.Seek "=", m_Clas
        If RsRepo.NoMatch Then
            ' no existe en repo
            RsClaP.Seek "=", m_Clas
            If Not RsClaP.NoMatch Then
                RsRepo.AddNew
                RsRepo!Nv = Nv
                RsRepo!orden = IIf(m_Clas = "OTRO", 99, 0)
                RsRepo!Tipo = m_Clas
                RsRepo!Descripcion = RsClaP!Descripcion
                RsRepo!Neto = RsOCdN![Cantidad Recibida] * RsOCdN![Precio Unitario]
                RsRepo.Update
                
            End If
        Else
            RsRepo.Edit
            RsRepo!Neto = RsRepo!Neto + RsOCdN![Cantidad Recibida] * RsOCdN![Precio Unitario]
            RsRepo.Update
        End If
    End If
    .MoveNext
Loop
End With

End Sub
Public Sub Repo_GeneralNv(Nv As Double, Fecha_Ini As String, Fecha_Fin As String)

Dim i As Integer, sql As String, m_KgTot As Double, m_STot As Double
Dim Dbi As Database, RsRepo As Recordset, RsOt As Recordset
Dim Dbm As Database, RsPd As Recordset
Dim RsCompras As Recordset

Set Dbm = OpenDatabase(mpro_file)
'Set RsPd = Dbi.OpenRecordset("Planos Detalle")

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset("General_Nv")

Set RsCompras = Dbi.OpenRecordset("SELECT * FROM Compras ORDER BY Orden")
' borra tabla de paso
Dbi.Execute "DELETE * FROM General_Nv"
i = 0

'///////////////////////////////////////////////////
' resumen de mano de obra (contratistas)
Repo_OTf Nv, "", "", ""
        
sql = "SELECT contratista, SUM([Kg Total]) as KgT,"
sql = sql & " SUM([$ Total]) as PT"
'sql = sql & " PT/KgT as PP"
sql = sql & " FROM OT group BY contratista"
Set RsOt = Dbi.OpenRecordset(sql)

i = i + 1
RsRepo.AddNew
RsRepo!orden = i
RsRepo!Texto0 = "Mano de Obra"
RsRepo.Update

i = i + 1
RsRepo.AddNew
RsRepo!orden = i
RsRepo!Texto1 = "OT Normal"
RsRepo.Update

i = i + 1
m_KgTot = 0
m_STot = 0

Do While Not RsOt.EOF

    i = i + 1
    RsRepo.AddNew
    RsRepo!orden = i
    RsRepo!Texto2 = RsOt!Contratista
    RsRepo!v1_1 = RsOt!KgT ' [kg total]
    RsRepo!v1_0 = RsOt!PT ' [$ total]
    If RsOt!KgT <> 0 Then
        RsRepo!v1_2 = RsOt!PT / RsOt!KgT
    End If
    RsRepo.Update
    
    m_KgTot = m_KgTot + RsOt!KgT
    m_STot = m_STot + RsOt!PT
    
    RsOt.MoveNext
    
Loop

' totales
i = i + 1
RsRepo.AddNew
RsRepo!orden = i
RsRepo!Texto2 = "Total"
RsRepo!v1_t1 = m_KgTot
RsRepo!v1_t0 = m_STot
If m_KgTot > 0 Then
    RsRepo!v1_t2 = m_STot / m_KgTot
End If
RsRepo.Update

RsOt.Close

'GoTo ResumenOtEsp
'////////////////////////////////////////////////////////
' RESUMEN OT ESP
'Repo_OTe NV, "", "", "" ', "Montaje"
'Repo_OTe Nv, "", "", "", "OTe", "P" ' ojo agregar "tipo" de ot especial
Repo_OTe Nv, "", Fecha_Ini, Fecha_Fin, "OTe", "P" ' ojo agregar "tipo" de ot especial
' y una rutina como esta que sigue, para evitar mucho codigo
OTe_Volcar Dbi, RsOt, RsRepo, i

GoTo ResumenOtEsp

sql = "SELECT contratista, SUM([Kg Total]) as KgT,"
sql = sql & " SUM([$ Total]) as PT,"
sql = sql & " PT/KgT as PP"
sql = sql & " FROM OTe group BY contratista"
Set RsOt = Dbi.OpenRecordset(sql)

i = i + 1
RsRepo.AddNew
RsRepo!orden = i
'RsRepo!Texto1 = "Resumen de OT Especial"
RsRepo!Texto1 = "OT Especial"
RsRepo.Update
i = i + 1
m_KgTot = 0
m_STot = 0
Do While Not RsOt.EOF

    i = i + 1
    RsRepo.AddNew
    RsRepo!orden = i
    RsRepo!Texto2 = RsOt!Contratista
    RsRepo!v1_1 = RsOt!KgT
    RsRepo!v1_0 = RsOt!PT
'    RsRepo!v1_2 = RsOt!PP
    RsRepo.Update
    
    m_KgTot = m_KgTot + RsOt!KgT
    m_STot = m_STot + RsOt!PT
    
    RsOt.MoveNext
    
Loop
' totales
i = i + 1
RsRepo.AddNew
RsRepo!orden = i
RsRepo!Texto2 = "Total"
RsRepo!v1_t1 = m_KgTot
RsRepo!v1_t0 = m_STot
If m_KgTot > 0 Then
    RsRepo!v1_t2 = m_STot / m_KgTot
End If
RsRepo.Update

RsOt.Close

ResumenOtEsp:
GoTo Sigue
'////////////////////////////////////////////////////////
' RESUMEN OT ESP
'sql = "SELECT NV,[RUT SubContratista], SUM(Cantidad*[Precio Unitario]) as PrecioTotal"
'sql = sql & " FROM [OT Esp Detalle] group BY [RUT SubContratista] "
'sql = sql & " WHERE NV=" & NV
sql = "SELECT * FROM [OT Esp Detalle] WHERE NV=" & Nv
sql = sql & " ORDER BY [RUT Contratista]"
Set RsOt = Dbm.OpenRecordset(sql)

Do While Not RsOt.EOF

    RsOt.MoveNext
    
Loop

For i = 1 To i

    i = i + 1
    RsRepo.AddNew
    RsRepo!orden = i
    RsRepo!Texto1 = RsOt![Rut contratista]
    RsRepo!v1_0 = RsOt!Cantidad * RsOt![Precio Unitario]
'    RsRepo!v1_0 = RsOt!PrecioTotal
'    RsRepo!v1_0 = RsOt!PT
'    RsRepo!v1_2 = RsOt!PP
    RsRepo.Update
    
'    m_KgTot = m_KgTot + RsOt!KgT
'    m_STot = m_STot + RsOt!PT
    
    RsOt.MoveNext
    
Next

Sigue:
'////////////////////////////////////////////////////////
' resumen de Kg
qry = "SELECT NV,"
qry = qry & "SUM([Cantidad Total]*Peso) AS KgTot, "
qry = qry & "SUM([OT Fab]*Peso) AS KgAsi, "
qry = qry & "SUM([ITO Fab]*Peso) AS KgRec, "
qry = qry & "SUM([ITO PyG]*Peso) AS KgPg, "
qry = qry & "SUM([GD]*Peso) AS KgDes "
qry = qry & " FROM [Planos Detalle]"
qry = qry & " GROUP BY NV"
Set RsPd = Dbm.OpenRecordset(qry)
i = LineaenBlanco(RsRepo, i)
RsRepo.AddNew
RsRepo!orden = i
RsRepo!Texto0 = "Resumen de Kilos"
RsRepo.Update
i = i + 1

Do While Not RsPd.EOF

    If RsPd!Nv = Nv Then
    
'        i = LineaenBlanco(RsRepo, i)
        
'        i = i + 1
'        RsRepo.AddNew
'        RsRepo!Orden = i
'        RsRepo!Texto1 = "Kg Planos"
'        RsRepo!v1_t1 = RsPd!KgTot
'        RsRepo.Update
        
'        i = LineaenBlanco(RsRepo, i)
        
        i = i + 1
        RsRepo.AddNew
        RsRepo!orden = i
'        RsRepo!Texto1 = "Kg por Asignar"
        RsRepo!Texto1 = "Kg por Fabricar"
        RsRepo!v1_1 = RsPd!KgTot - RsPd!KgAsi
        RsRepo.Update
        
        i = i + 1
        RsRepo.AddNew
        RsRepo!orden = i
'        RsRepo!Texto1 = "Kg por Recibir"
        RsRepo!Texto1 = "Kg en Fabricación"
        RsRepo!v1_1 = RsPd!KgAsi - RsPd!KgRec
        If RsRepo!v1_1 = 0 Then
            RsRepo!Texto3 = "0"
        End If
        RsRepo.Update
        
        i = i + 1
        RsRepo.AddNew
        RsRepo!orden = i
'        RsRepo!Texto1 = "Kg por Despachar"
        RsRepo!Texto1 = "Kg en Negro"
        RsRepo!v1_1 = RsPd!KgRec - RsPd!KgPg
        If RsRepo!v1_1 = 0 Then
            RsRepo!Texto3 = "0"
        End If
        RsRepo.Update
        
        i = i + 1
        RsRepo.AddNew
        RsRepo!orden = i
'        RsRepo!Texto1 = "Kg Pintados o Galvanizados"
        RsRepo!Texto1 = "Kg Pintados o Reprocesados"
        If RsPd!KgPg - RsPd!KgDes >= 0 Then
            RsRepo!v1_1 = RsPd!KgPg - RsPd!KgDes
        End If
        If RsRepo!v1_1 = 0 Then
            RsRepo!Texto3 = "0"
        End If
        RsRepo.Update
        
        i = i + 1
        RsRepo.AddNew
        RsRepo!orden = i
        RsRepo!Texto1 = "Kg Despachados"
        RsRepo!v1_1 = RsPd!KgDes
        If RsRepo!v1_1 = 0 Then
            RsRepo!Texto3 = "0"
        End If
        RsRepo.Update
        
        i = i + 1
        RsRepo.AddNew
        RsRepo!orden = i
        RsRepo!Texto1 = "Kg Total"
        RsRepo!v1_t1 = RsPd!KgTot

        RsRepo.Update
        
        Exit Do
        
    End If
    
    RsPd.MoveNext
    
Loop
RsPd.Close

'/////////////////////////////////////////////
' resumen de m2

' resumen de Compras
i = LineaenBlanco(RsRepo, i)
RsRepo.AddNew
RsRepo!orden = i
RsRepo!Texto0 = "Resumen de Compras"
'RsRepo!Valor0 = RsCompras!Neto
RsRepo.Update
i = i + 1

m_STot = 0
Do While Not RsCompras.EOF

    i = i + 1
    RsRepo.AddNew
    RsRepo!orden = i
    RsRepo!Texto1 = RsCompras!Descripcion
    RsRepo!v1_0 = RsCompras!Neto
    RsRepo.Update
    
    m_STot = m_STot + RsCompras!Neto
    
    RsCompras.MoveNext
    
Loop

RsCompras.Close
' total compras
i = i + 1
RsRepo.AddNew
RsRepo!orden = i
RsRepo!Texto1 = "Total Compras"
RsRepo!v1_t0 = m_STot
RsRepo.Update

RsRepo.Close
Dbi.Close

End Sub
Private Function LineaenBlanco(RsRepo As Recordset, n As Integer)
n = n + 1
RsRepo.AddNew
RsRepo!orden = n
RsRepo.Update
n = n + 1
LineaenBlanco = n
End Function
Private Sub OTe_Volcar(Dbi As Database, RsOt As Recordset, RsRepo As Recordset, i As Integer)
' traspasa de OTe a
Dim sql As String, m_KgTot As Double, m_STot As Double
'i As Integer, m_KgTot As Double, m_STot As Double

sql = "SELECT subcontratista, SUM([Kg Total]) as KgT,"
sql = sql & " SUM([$ Total]) as PT,"
sql = sql & " PT/KgT as PP"
sql = sql & " FROM OTe group BY subcontratista"

sql = "SELECT contratista, tipo, SUM([Kg Total]) as KgT,"
sql = sql & " SUM([$ Total]) as PT,"
sql = sql & " PT/KgT as PP"
sql = sql & " FROM OTe group BY tipo,contratista"

Set RsOt = Dbi.OpenRecordset(sql)

i = i + 1
RsRepo.AddNew
RsRepo!orden = i
'RsRepo!Texto1 = "Resumen de OT Especial"
RsRepo!Texto1 = "OT Especial"
RsRepo.Update
i = i + 1
m_KgTot = 0
m_STot = 0

Do While Not RsOt.EOF

    i = i + 1
    RsRepo.AddNew
    RsRepo!orden = i
    RsRepo!Texto2 = RsOt!Contratista
'    RsRepo!Texto3 = RsOt!Tipo
    RsRepo!v1_1 = RsOt!KgT
    RsRepo!v1_0 = RsOt!PT
'    RsRepo!v1_2 = RsOt!PP
    RsRepo.Update
    
    m_KgTot = m_KgTot + RsOt!KgT
    m_STot = m_STot + RsOt!PT
    
    RsOt.MoveNext
    
Loop
' totales
i = i + 1
RsRepo.AddNew
RsRepo!orden = i
RsRepo!Texto2 = "Total"
RsRepo!v1_t1 = m_KgTot
RsRepo!v1_t0 = m_STot
If m_KgTot > 0 Then
    RsRepo!v1_t2 = m_STot / m_KgTot
End If
RsRepo.Update

RsOt.Close

End Sub
Public Sub Repo_Kardex(CodigoProducto As String, F_Inicial As String, F_Final As String)

Dim Db As Database, Rs As Recordset, qry As String
Dim Dbi As Database, RsI As Recordset
Dim m_SaldoCant As Double, m_SaldoVal As Double
Dim m_PPP As Double, m_ValEntra As Double, m_ValSale As Double

Set Db = OpenDatabase(Madq_file)
qry = "SELECT * FROM Documentos"
qry = qry & " WHERE [Codigo Producto]='" & CodigoProducto & "'"
qry = qry & " ORDER BY fecha,tipo"
Set Rs = Db.OpenRecordset(qry)

Set Dbi = OpenDatabase(repo_file)
Dbi.Execute "DELETE * FROM Kardex"
Set RsI = Dbi.OpenRecordset("Kardex")

m_SaldoCant = 0
m_SaldoVal = 0

m_PPP = -1

With Rs
Do While Not .EOF
    
    RsI.AddNew
    RsI!Fecha = !Fecha
    RsI!Tipo = !Tipo
    RsI!Numero = !Numero
    RsI!precio = ![Precio Unitario]
    
    m_SaldoCant = m_SaldoCant + Rs![Cant_Entra] - Rs![Cant_Sale]
    RsI!Cant_Entra = ![Cant_Entra]
    RsI!Cant_Sale = ![Cant_Sale]
    RsI!Cant_Saldo = m_SaldoCant
    
    ' primer ppp, si no es compra
    If m_PPP = -1 Then m_PPP = 1
    
    ' para compras
    m_ValEntra = ![Cant_Entra] * ![Precio Unitario]
    m_ValSale = ![Cant_Sale] * m_PPP
    m_SaldoVal = m_SaldoVal + m_ValEntra - m_ValSale
    
    RsI!Val_Entra = m_ValEntra
    RsI!Val_Sale = m_ValSale
    RsI!Val_Saldo = m_SaldoVal
    
    If m_SaldoCant <> 0 Then
        If m_SaldoVal <> 0 Then
            m_PPP = Format(m_SaldoVal / m_SaldoCant, "#.####")
        End If
    End If
    
    RsI!PPP = m_PPP
    
    RsI.Update
    
    .MoveNext
    
Loop
End With

End Sub
Public Sub Repo_ProdxDescYANO(Tipo As String)

' informe de produccion por descripcion
' en base a itof y gd
' ej:
'     VIGAS  854.000,2

Dim Db As Database, DbR As Database
Dim RsDoc As Recordset, RsRepo As Recordset

Dim sql As String, m_PesoTotal As Double, li As Integer, m_Descrip As String
Dim Tabla As String

'db_repo = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=c:\inetpub\wwwroot\scp\scp_repo.mdb"

'Set Db = OpenDatabase("\\acr3006-dualpro\wwwroot\scp\scp.mdb")
Set Db = OpenDatabase("\\acr3006-dualpro\scp\scp.mdb")
Set DbR = OpenDatabase(repo_file)

DbR.Execute "delete * from prodxdesc"

Set RsRepo = DbR.OpenRecordset("prodxdesc")
RsRepo.Index = "descripcion"

'response.Write ("plano=" & m_plano )

'If Tipo = "ITOF" Or Tipo = "PINTURA" Then
If Tipo = "ITOF" Then
    Tabla = "[itof_detalle]"
    sql = "SELECT "
    sql = sql & Tabla & ".itofd_nv      ,"
    sql = sql & Tabla & ".itofd_plano   ,"
    sql = sql & Tabla & ".itofd_marca   ,"
    sql = sql & Tabla & ".itofd_cantidad as cant,"
    sql = sql & Tabla & ".[itofd_pesounitario] as kgs,"
    sql = sql & " [planos_detalle].pld_descripcion AS marca_descrip"
    sql = sql & " FROM " & Tabla
    sql = sql & " INNER JOIN [planos_detalle]"
    sql = sql & "     ON     " & Tabla & ".itofd_nv    = [planos_detalle].pld_nv"
    sql = sql & "     AND    " & Tabla & ".itofd_plano = [planos_detalle].pld_plano"
    sql = sql & "     AND    " & Tabla & ".itofd_marca = [planos_detalle].pld_marca"
    sql = sql & " WHERE YEAR(" & Tabla & ".itofd_fecha)=2004"
    'sql = sql & " AND  MONTH(" & Tabla & ".itofd_fecha)=2"
End If

If Tipo = "GD" Then
    Tabla = "[gd_detalle]"
    sql = "SELECT "
    sql = sql & Tabla & ".gdd_nv      ,"
    sql = sql & Tabla & ".gdd_plano   ,"
    sql = sql & Tabla & ".gdd_marca   ,"
    sql = sql & Tabla & ".gdd_cantidad as cant,"
    sql = sql & Tabla & ".gdd_preciounitario,"
    sql = sql & Tabla & ".[gdd_pesounitario] as kgs,"
    sql = sql & " [planos_detalle].pld_descripcion AS marca_descrip"
    sql = sql & " FROM " & Tabla
    sql = sql & " INNER JOIN [planos_detalle]"
    sql = sql & "     ON     " & Tabla & ".gdd_nv    = [planos_detalle].pld_nv"
    sql = sql & "     AND    " & Tabla & ".gdd_plano = [planos_detalle].pld_plano"
    sql = sql & "     AND    " & Tabla & ".gdd_marca = [planos_detalle].pld_marca"
    sql = sql & " WHERE YEAR(" & Tabla & ".gdd_fecha)=2004"
    'sql = sql & " AND  MONTH(" & Tabla & ".itofd_fecha)=2"
End If

Set RsDoc = Db.OpenRecordset(sql)
        
m_PesoTotal = 0
li = 0

Do While Not RsDoc.EOF

'   pl_pesototal = pl_pesototal + m_peso
   m_Descrip = RsDoc("marca_descrip")
   m_Descrip = Left(m_Descrip, 4)

   ' busca en repo
   RsRepo.Seek "=", m_Descrip
   If RsRepo.NoMatch Then
      ' agrega registro
      RsRepo.AddNew
      RsRepo!Descripcion = m_Descrip
      RsRepo!kgs = RsDoc!kgs * RsDoc!Cant
      RsRepo!Numero = RsDoc!gdd_preciounitario * RsDoc!Cant
      RsRepo.Update
   Else
      RsRepo.Edit
      RsRepo!kgs = RsRepo!kgs + RsDoc!kgs * RsDoc!Cant
      RsRepo!Numero = RsRepo!Numero + RsDoc!gdd_preciounitario * RsDoc!Cant
      RsRepo.Update
   End If
                        
   RsDoc.MoveNext
                                
Loop
        
RsDoc.Close

End Sub
Public Sub Repo_OCResumen()

' informe de compras agrupadas por tipo de proveedor
' ACER, FERR, PERN, PINT, SOLD y otros

Dim Db As Database, DbR As Database
Dim RsDoc As Recordset, RsRepo As Recordset

Dim sql As String, m_PesoTotal As Double, li As Integer, m_Descrip As String
Dim Tabla As String

'db_repo = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=c:\inetpub\wwwroot\scp\scp_repo.mdb"

'Set Db = OpenDatabase("\\acr3006-dualpro\wwwroot\scp\scp.mdb")
Set Db = OpenDatabase("\\acr3006-dualpro\scp\scp.mdb")
Set DbR = OpenDatabase(repo_file)

DbR.Execute "delete * from prodxdesc"

Set RsRepo = DbR.OpenRecordset("prodxdesc")
RsRepo.Index = "descripcion"

'response.Write ("plano=" & m_plano )

Tabla = "[oc_cabecera]"
sql = "SELECT "
sql = sql & Tabla & ".occ_neto as nneto ,"
sql = sql & " [proveedores].prv_clasificacion AS clasif"
sql = sql & " FROM " & Tabla
sql = sql & " INNER JOIN [proveedores]"
sql = sql & "     ON     " & Tabla & ".occ_rutproveedor = [proveedores].prv_rut"
'sql = sql & "     AND    " & Tabla & ".itofd_plano = [planos_detalle].pld_plano"
'sql = sql & "     AND    " & Tabla & ".itofd_marca = [planos_detalle].pld_marca"
'sql = sql & " GROUP BY oc_cabecera.occ_rutproveedor"
sql = sql & " WHERE YEAR(" & Tabla & ".occ_fecha)=2004"
'sql = sql & " AND  MONTH(" & Tabla & ".occ_fecha)=12"
'sql = sql & " AND  DAY(" & Tabla & ".occ_fecha)=27"

Set RsDoc = Db.OpenRecordset(sql)
        
m_PesoTotal = 0
li = 0

Do While Not RsDoc.EOF

'   pl_pesototal = pl_pesototal + m_peso
   m_Descrip = NoNulo(RsDoc("clasif"))
   If Trim(m_Descrip) = "" Then m_Descrip = "OTRO"

   ' busca en repo
   RsRepo.Seek "=", m_Descrip
   If RsRepo.NoMatch Then
      ' agrega registro
      RsRepo.AddNew
      RsRepo!Descripcion = m_Descrip
      RsRepo!kgs = RsDoc!nneto
      RsRepo.Update
   Else
      RsRepo.Edit
      RsRepo!kgs = RsRepo!kgs + RsDoc!nneto
      RsRepo.Update
   End If
                        
   RsDoc.MoveNext
                                
Loop
        
RsDoc.Close

End Sub
Public Sub Repo_ItoVal(Nv As Double, Contratista_Rut As String, FechaIni As String, FechaFin As String)

' problema que tiene: no refleja precios de OT modificados

' ito valorizada, nivel agregado
' contratista, peso, precio

Dim sql As String, p As Integer, m_TotalPeso As Double, m_TotalPrecio As Double, NomTabla As String
Dim Dbm As Database, RsITOfc As Recordset
Dim Dbi As Database, RsRepo As Recordset

NomTabla = "ITO"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

sql = "SELECT [rut contratista] AS rut, sum([peso total]) AS peso, sum([precio total]) AS precio"
sql = sql & " FROM [ito fab cabecera]"

p = 0
If FechaIni <> Fecha_Vacia Then
'    MsgBox "ini"
    p = p + 1
    FechaIni = Mid(FechaIni, 4, 3) & Mid(FechaIni, 1, 2) & Mid(FechaIni, 6, 3)
    sql = sql & " WHERE Fecha >= #" & FechaIni & "#"
End If

If FechaFin <> Fecha_Vacia Then
'    MsgBox "fin"
    FechaFin = Mid(FechaFin, 4, 3) & Mid(FechaFin, 1, 2) & Mid(FechaFin, 6, 3)
    If p = 0 Then
        sql = sql & " WHERE Fecha <= #" & FechaFin & "#"
    Else
        sql = sql & " AND Fecha <= #" & FechaFin & "#"
    End If
End If

sql = sql & " GROUP BY [rut contratista]"

Set Dbm = OpenDatabase(mpro_file)
Set RsITOfc = Dbm.OpenRecordset(sql)

m_TotalPeso = 0
m_TotalPrecio = 0

With RsITOfc
Do While Not .EOF

'    Debug.Print !Rut, !Peso, !Precio
    
'    m_TotalPeso = m_TotalPeso + !Peso
'    m_TotalPrecio = m_TotalPrecio + !Precio
    
    RsRepo.AddNew
    RsRepo!Clasificacion = !rut
    RsRepo![Kg Total] = !Peso
    RsRepo![$ Total] = !precio
    RsRepo.Update
    
    .MoveNext
    
Loop
End With

'Debug.Print Format(m_TotalPeso, "#,###"), Format(m_TotalPrecio, "#,###")

Dbi.Close
Dbm.Close

End Sub
Public Sub Repo_ItoVal_new(Nv As Double, Contratista_Rut As String, FechaIni As String, FechaFin As String)

' ito valorizada, nivel agregado
' contratista, peso, precio

Dim sql As String, p As Integer, m_TotalPeso As Double, m_TotalPrecio As Double, NomTabla As String
Dim m_sc As String
Dim Dbm As Database, RsITOfd As Recordset, RsOTfd As Recordset
Dim Dbi As Database, RsRepo As Recordset

NomTabla = "ITO"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

sql = "SELECT * FROM [ito fab detalle]"

p = 0
If FechaIni <> Fecha_Vacia Then
'    MsgBox "ini"
    p = p + 1
    FechaIni = Mid(FechaIni, 4, 3) & Mid(FechaIni, 1, 2) & Mid(FechaIni, 6, 3)
    sql = sql & " WHERE Fecha >= #" & FechaIni & "#"
End If

If FechaFin <> Fecha_Vacia Then
'    MsgBox "fin"
    FechaFin = Mid(FechaFin, 4, 3) & Mid(FechaFin, 1, 2) & Mid(FechaFin, 6, 3)
    If p = 0 Then
        sql = sql & " WHERE Fecha <= #" & FechaFin & "#"
    Else
        sql = sql & " AND Fecha <= #" & FechaFin & "#"
    End If
End If

Set Dbm = OpenDatabase(mpro_file)
Set RsITOfd = Dbm.OpenRecordset(sql)

Set RsOTfd = Dbm.OpenRecordset("OT fab detalle")
RsOTfd.Index = "Número-Línea"

m_TotalPeso = 0
m_TotalPrecio = 0

With RsITOfd
m_sc = ![Rut contratista]
Do While Not .EOF

'    Debug.Print !Rut, !Peso, !Precio
    
'    m_TotalPeso = m_TotalPeso + !Peso
'    m_TotalPrecio = m_TotalPrecio + !Precio
    
    If m_sc = ![Rut contratista] Then
    Else
    End If
    
    RsRepo.AddNew
    RsRepo!Clasificacion = !rut
    RsRepo![Kg Total] = !Peso
    RsRepo![$ Total] = !precio
    RsRepo.Update
    
    .MoveNext
    
Loop
End With

'Debug.Print Format(m_TotalPeso, "#,###"), Format(m_TotalPrecio, "#,###")

Dbi.Close
Dbm.Close

End Sub
Public Sub Repo_ChkLst(CodigoArea As String, RUT_ResponsableArea As String, FechaIni, FechaTer)
' area observada | responsable area | resp evaluacion | cargo evaluador | fecha | punt 1 | punt 2 | %item1 | %item2 | %final |
Dim NomTabla As String, m_Area_Responsable As String, m_Evaluacion_Responsable As String
Dim DbD As Database, RsAreas As Recordset, RsTra As Recordset
Dim Dbm2 As Database, RsChkLst As Recordset
Dim Dbi As Database, RsRepo As Recordset

NomTabla = "chklst"

Set DbD = OpenDatabase(data_file)
Set RsAreas = DbD.OpenRecordset("chklst_areas")
RsAreas.Index = "codigo"

Set RsTra = DbD.OpenRecordset("trabajadores")
RsTra.Index = "rut"

Set Dbm2 = OpenDatabase(mpro2_file)
Set RsChkLst = Dbm2.OpenRecordset("chklst_cabecera")

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

With RsChkLst
Do While Not .EOF

    If FechaIni = "__/__/__" Then
    Else
        If CDate(FechaIni) <= !semana Then
        Else
            GoTo NoIncluir
        End If
    End If
    
    If FechaTer = "__/__/__" Then
    Else
        If !semana <= CDate(FechaTer) Then
        Else
            GoTo NoIncluir
        End If
    End If

    If CodigoArea = "" Then
        If RUT_ResponsableArea = "" Then
            GoTo RegistroAgregar
        Else
            If RUT_ResponsableArea = !responsable_area Then
                GoTo RegistroAgregar
            End If
        End If
    Else
        If CodigoArea = !area Then
            If RUT_ResponsableArea = "" Then
                GoTo RegistroAgregar
            Else
                If RUT_ResponsableArea = !responsable_area Then
                    GoTo RegistroAgregar
                End If
            End If
        End If
    End If
    
    If False Then
RegistroAgregar:
    
        RsRepo.AddNew
        
        m_Area_Responsable = ""
        m_Evaluacion_Responsable = ""
        
        RsAreas.Seek "=", !area
        If Not RsAreas.NoMatch Then
            RsRepo!area = RsAreas!Descripcion
            m_Area_Responsable = RsAreas!responsable_area
            m_Evaluacion_Responsable = RsAreas!responsable_evaluacion
        End If
        
        RsTra.Seek "=", m_Area_Responsable
        If Not RsTra.NoMatch Then
            RsRepo!area_responsable = Left(RsTra!appaterno & " " & RsTra!apmaterno & " " & RsTra!nombres, 30)
        End If
    
        RsTra.Seek "=", m_Evaluacion_Responsable
        If Not RsTra.NoMatch Then
            RsRepo!evaluacion_responsable = Left(RsTra!appaterno & " " & RsTra!apmaterno & " " & RsTra!nombres, 30)
            RsRepo!Cargo = RsTra!Cargo
        End If
    
        RsRepo!Fecha = !Fecha ' emision
        RsRepo!semana_fecha = !semana
        RsRepo!semana_texto = fecha2semana(!semana)
        RsRepo!valor1 = !eval1
        RsRepo!valor2 = !eval2
        RsRepo!valor3 = !porcentaje1
        RsRepo!valor4 = !porcentaje2
        RsRepo!valor5 = !porcentaje1 * 0.2 + !porcentaje2 * 0.8
        RsRepo.Update
    
    End If
    
NoIncluir:

    .MoveNext
    
Loop
End With

DbD.Close
Dbm2.Close
Dbi.Close

End Sub
Public Sub Repo_ChkLstObs(CodigoArea As String, RUT_ResponsableArea As String, FechaIni, FechaTer)
' area observada | sem1 | sem2 | sem3 | sem4 | sem5 |
Dim NomTabla As String, m_Area_Responsable As String, m_Evaluacion_Responsable As String
Dim DbD As Database, RsAreas As Recordset, RsTra As Recordset
Dim Dbm2 As Database, RsChkLst As Recordset
Dim Dbi As Database, RsRepo As Recordset

Dim m_Semana As String, s_Mes As String, d_Mes As Date, i As Integer, j As Integer

NomTabla = "chklst_obs"

Set DbD = OpenDatabase(data_file)
Set RsAreas = DbD.OpenRecordset("chklst_areas")
RsAreas.Index = "codigo"

Set RsTra = DbD.OpenRecordset("trabajadores")
RsTra.Index = "rut"

Set Dbm2 = OpenDatabase(mpro2_file)
Set RsChkLst = Dbm2.OpenRecordset("chklst_detalle")
RsChkLst.Index = "area-semana-item"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)
RsRepo.Index = "area-mes-item"

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

With RsChkLst
Do While Not .EOF

    If FechaIni = "__/__/__" Then
    Else
        If CDate(FechaIni) <= !semana Then
        Else
            GoTo NoIncluir
        End If
    End If

    If FechaTer = "__/__/__" Then
    Else
        If !semana <= CDate(FechaTer) Then
        Else
            GoTo NoIncluir
        End If
    End If

    If CodigoArea = "" Then
        If RUT_ResponsableArea = "" Then
            GoTo RegistroAgregar
        Else
            If RUT_ResponsableArea = !responsable_area Then
                GoTo RegistroAgregar
            End If
        End If
    Else
        If CodigoArea = !area Then
            If RUT_ResponsableArea = "" Then
                GoTo RegistroAgregar
            Else
                If RUT_ResponsableArea = !responsable_area Then
                    GoTo RegistroAgregar
                End If
            End If
        End If
    End If
    
    If False Then

RegistroAgregar:

        If Trim(!item_obs) <> "" Then

        m_Semana = fecha2semana(!semana)
        s_Mes = Right(m_Semana, 2)
        j = InStr(1, m_Semana, "-")
        m_Semana = Trim(Left(m_Semana, j - 1))

        s_Mes = "01/" & s_Mes & "/" & Year(!semana)
        d_Mes = CDate(s_Mes)
        
        For i = 1 To 50
        
            RsRepo.Seek "=", !area, d_Mes, i
    
            If RsRepo.NoMatch Then
            
                ' crea registro
                RsRepo.AddNew
                
                RsRepo!CodigoArea = !area
                RsRepo!mes_Fecha = d_Mes
                RsRepo!mes_Nombre = Format(d_Mes, "mmm/yy")
                RsRepo!item = i
                            
                RsAreas.Seek "=", !area
                If Not RsAreas.NoMatch Then
                    RsRepo!area = RsAreas!Descripcion
                End If
                
                ' puebla
                RsRepo("obs" & m_Semana) = Trim(!item_obs)
                
                RsRepo.Update
                
                Exit For
                
            Else
            
                ' encontro registro del mes
                ' ahora busco si obsN esta ocupada
                If Trim(NoNulo(RsRepo("obs" & m_Semana))) = "" Then
                
                    RsRepo.Edit
                    RsRepo("obs" & m_Semana) = Trim(!item_obs)
                    RsRepo.Update
                    
                    Exit For

                Else

                    ' busca siguiente linea

                End If

            End If
            
        Next

        
        End If

    End If
    
NoIncluir:

    .MoveNext
    
Loop
End With

DbD.Close
Dbm2.Close
Dbi.Close

End Sub
Public Sub Repo_ChkLstFinal(CodigoArea As String, RUT_ResponsableArea As String, FechaIni, FechaTer)
' area observada | responsable area | resp evaluacion | cargo evaluador | %sem1 | %sem2 | %sem3 | %sem4 | %sem5 | %final |
Dim NomTabla As String, m_Area_Responsable As String, m_Evaluacion_Responsable As String
Dim DbD As Database, RsAreas As Recordset, RsTra As Recordset
Dim Dbm2 As Database, RsChkLst As Recordset
Dim Dbi As Database, RsRepo As Recordset

Dim m_Semana As String, m_Mes As String, i As Integer, j As Integer, pFinal As Double

NomTabla = "chklst"

Set DbD = OpenDatabase(data_file)
Set RsAreas = DbD.OpenRecordset("chklst_areas")
RsAreas.Index = "codigo"

Set RsTra = DbD.OpenRecordset("trabajadores")
RsTra.Index = "rut"

Set Dbm2 = OpenDatabase(mpro2_file)
Set RsChkLst = Dbm2.OpenRecordset("chklst_cabecera")

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(NomTabla)
RsRepo.Index = "area-fecha"

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & NomTabla & "]"

With RsChkLst
Do While Not .EOF

    If FechaIni = "__/__/__" Then
    Else
        If CDate(FechaIni) <= !semana Then
        Else
            GoTo NoIncluir
        End If
    End If
    
    If FechaTer = "__/__/__" Then
    Else
        If !semana <= CDate(FechaTer) Then
        Else
            GoTo NoIncluir
        End If
    End If

    If CodigoArea = "" Then
        If RUT_ResponsableArea = "" Then
            GoTo RegistroAgregar
        Else
            If RUT_ResponsableArea = !responsable_area Then
                GoTo RegistroAgregar
            End If
        End If
    Else
        If CodigoArea = !area Then
            If RUT_ResponsableArea = "" Then
                GoTo RegistroAgregar
            Else
                If RUT_ResponsableArea = !responsable_area Then
                    GoTo RegistroAgregar
                End If
            End If
        End If
    End If
    
    If False Then
    
RegistroAgregar:

        m_Semana = fecha2semana(!semana)
        m_Mes = Right(m_Semana, 2)
        j = InStr(1, m_Semana, "-")
        m_Semana = Trim(Left(m_Semana, j - 1))
        
        m_Mes = "01/" & m_Mes & "/" & Year(!semana)
        RsRepo.Seek "=", !area, m_Mes
        
        If RsRepo.NoMatch Then
        
            RsRepo.AddNew
            
            RsRepo!CodigoArea = !area
            RsRepo!Fecha = m_Mes
            
            m_Area_Responsable = ""
            m_Evaluacion_Responsable = ""
            
            RsAreas.Seek "=", !area
            If Not RsAreas.NoMatch Then
                RsRepo!area = RsAreas!Descripcion
                m_Area_Responsable = RsAreas!responsable_area
                m_Evaluacion_Responsable = RsAreas!responsable_evaluacion
            End If
            
            RsTra.Seek "=", m_Area_Responsable
            If Not RsTra.NoMatch Then
                RsRepo!area_responsable = Left(RsTra!appaterno & " " & RsTra!apmaterno & " " & RsTra!nombres, 30)
            End If
        
            RsTra.Seek "=", m_Evaluacion_Responsable
            If Not RsTra.NoMatch Then
                RsRepo!evaluacion_responsable = Left(RsTra!appaterno & " " & RsTra!apmaterno & " " & RsTra!nombres, 30)
                RsRepo!Cargo = RsTra!Cargo
            End If
            
            RsRepo!semana_texto = Format(!semana, "mmm-yy")
        
        Else
    
            RsRepo.Edit
    
        End If
        
'        RsRepo!semana_fecha = !semana
        RsRepo("valor" & m_Semana) = !porcentajefinal
        
        j = 0
        pFinal = 0
        For i = 1 To 5
            If RsRepo("valor" & i) Then
                j = j + 1
                pFinal = pFinal + RsRepo("valor" & i)
            End If
        Next
        pFinal = pFinal / j
        
        RsRepo!valor6 = pFinal
        
'        RsRepo!semana_texto = fecha2semana(!semana)
        
'        RsRepo!valor1 = !eval1
        
        RsRepo.Update
    
    End If
    
NoIncluir:

    .MoveNext
    
Loop
End With

DbD.Close
Dbm2.Close
Dbi.Close

End Sub
Public Sub Repo_as_PiezasxTurno(Nv As Double, Fecha_Ini As String, Fecha_Fin As String, OpRut As String, OpNom As String)
' arco sumergido, resumen piezas por turno

Dim Dbm As Database, RsAs As Recordset
Dim DbR As Database, RsRas As Recordset

Dim sql As String, np As Integer, m_Metros As Double
Dim m_DiaC As Integer, m_NocheC As Integer, m_DiaM As Double, m_NocheM As Double, m_DiaK As Double, m_NocheK As Double

Set Dbm = OpenDatabase(mpro_file)

sql = "SELECT * FROM [arco sumergido]"

np = 0
If Nv > 0 Then
    sql = sql & " WHERE nv=" & Nv
    np = np + 1
End If
If Fecha_Ini <> "__/__/__" Then
    If np = 0 Then
        sql = sql & " WHERE fecha>=CDate('" & Fecha_Ini & "')"
    Else
        sql = sql & " AND fecha>=CDate('" & Fecha_Ini & "')"
    End If
    np = np + 1
End If
If Fecha_Fin <> "__/__/__" Then
    If np = 0 Then
        sql = sql & " WHERE fecha<=CDate('" & Fecha_Fin & "')"
    Else
        sql = sql & " AND fecha<=CDate('" & Fecha_Fin & "')"
    End If
    np = np + 1
End If
If OpRut <> "" Then
    If np = 0 Then
        sql = sql & " WHERE [rut operador1]='" & OpRut & "'"
    Else
        sql = sql & " AND [rut operador1]='" & OpRut & "'"
    End If
    np = np + 1
End If

'sql = sql & " ORDER BY tipopieza"

Set RsAs = Dbm.OpenRecordset(sql)

Set DbR = OpenDatabase(repo_file)
Set RsRas = DbR.OpenRecordset("as_piezasxturno")
RsRas.Index = "tipo"
DbR.Execute "DELETE * FROM as_piezasxturno"

RsRas.AddNew
RsRas!orden = 1
RsRas!Tipo = "V"
RsRas!Descripcion = "VIGA"
RsRas!Factor = 4
RsRas.Update
RsRas.AddNew
RsRas!orden = 2
RsRas!Tipo = "T"
RsRas!Descripcion = "TUBULAR"
RsRas!Factor = 2
RsRas.Update
RsRas.AddNew
RsRas!orden = 3
RsRas!Tipo = "S"
RsRas!Descripcion = "TUBEST"
RsRas!Factor = 4
RsRas.Update
RsRas.AddNew
RsRas!orden = 4
RsRas!Tipo = "P"
RsRas!Descripcion = "PLANCHA"
RsRas!Factor = 1
RsRas.Update

'm_tp = RsAs!TipoPieza

Do While Not RsAs.EOF
    
    m_Metros = RsAs!dim7
    
    m_DiaC = 0
    m_NocheC = 0
    
    m_DiaM = 0
    m_NocheM = 0
        
    m_DiaK = 0
    m_NocheK = 0
    
    RsRas.Seek "=", RsAs!TipoPieza
    If Not RsRas.NoMatch Then
        
        If RsAs!Turno = "D" Then
            m_DiaC = RsAs!Cantidad
            m_DiaM = Int(m_DiaC * m_Metros * RsRas!Factor / 1000 + 0.5) ' expresado en metros
            m_DiaK = RsAs![PesoTotal]
        End If
        If RsAs!Turno = "N" Then
            m_NocheC = RsAs!Cantidad
            m_NocheM = Int(m_NocheC * m_Metros * RsRas!Factor / 1000 + 0.5)
            m_NocheK = RsAs![PesoTotal]
        End If
        
        RsRas.Edit
        
        RsRas!dia_cantidad = RsRas!dia_cantidad + m_DiaC
        RsRas!noche_cantidad = RsRas!noche_cantidad + m_NocheC
        RsRas!Total_cantidad = RsRas!Total_cantidad + m_DiaC + m_NocheC
        
        RsRas!dia_metros = RsRas!dia_metros + m_DiaM
        RsRas!noche_metros = RsRas!noche_metros + m_NocheM
        RsRas!Total_metros = RsRas!Total_metros + m_DiaM + m_NocheM
        
        RsRas!dia_kilos = RsRas!dia_kilos + m_DiaK
        RsRas!noche_kilos = RsRas!noche_kilos + m_NocheK
        RsRas!Total_kilos = RsRas!Total_kilos + m_DiaK + m_NocheK
        
        RsRas.Update
        
    End If
    
    RsAs.MoveNext

Loop

End Sub
Public Sub Repo_as_Bono(Nv As Double, Fecha_Ini As String, Fecha_Fin As String, OpRut As String, OpNom As String)
' arco sumergido, piezas por turno

Dim DbD As Database, RsTrab As Recordset
Dim Dbm As Database, RsAs As Recordset
Dim DbR As Database, RsRas As Recordset

'        Select Case m_BaseFamiliar
'        Case RsTaf![Desde 1] To RsTaf![Hasta 1]

' arreglo de tramos para pago de bono
Dim a_Tramos(30, 3) As Double, m_TipoEstructura As Integer

Dim sql As String, np As Integer, m_Factor As Double

Set DbD = OpenDatabase(data_file)

Set RsTrab = DbD.OpenRecordset("Tabla Bono Arco Sumergido")
RsTrab.Index = "nv-estructura-item"
Do While Not RsTrab.EOF

    m_TipoEstructura = TipoEstructura(RsTrab!estructura)

    m_TipoEstructura = m_TipoEstructura + RsTrab!item
    a_Tramos(m_TipoEstructura, 0) = m_TipoEstructura
    a_Tramos(m_TipoEstructura, 1) = RsTrab!Desde
    a_Tramos(m_TipoEstructura, 2) = RsTrab!Hasta
    a_Tramos(m_TipoEstructura, 3) = RsTrab!Valor
    
    RsTrab.MoveNext
    
Loop
RsTrab.Close


Set RsTrab = DbD.OpenRecordset("trabajadores")
RsTrab.Index = "rut"

Set Dbm = OpenDatabase(mpro_file)

sql = "SELECT * FROM [arco sumergido]"

np = 0
If Nv > 0 Then
    sql = sql & " WHERE nv=" & Nv
    np = np + 1
End If
If Fecha_Ini <> "__/__/__" Then
    If np = 0 Then
        sql = sql & " WHERE fecha>=CDate('" & Fecha_Ini & "')"
    Else
        sql = sql & " AND fecha>=CDate('" & Fecha_Ini & "')"
    End If
    np = np + 1
End If
If Fecha_Fin <> "__/__/__" Then
    If np = 0 Then
        sql = sql & " WHERE fecha<=CDate('" & Fecha_Fin & "')"
    Else
        sql = sql & " AND fecha<=CDate('" & Fecha_Fin & "')"
    End If
    np = np + 1
End If
If OpRut <> "" Then
    If np = 0 Then
        sql = sql & " WHERE [rut operador1]='" & OpRut & "'"
    Else
        sql = sql & " AND [rut operador1]='" & OpRut & "'"
    End If
    np = np + 1
End If

'sql = sql & " ORDER BY tipopieza"

' [pesounitario]
' [dim7] es el largo en mm

Set RsAs = Dbm.OpenRecordset(sql)

Set DbR = OpenDatabase(repo_file)
Set RsRas = DbR.OpenRecordset("as_bono")
RsRas.Index = "te-rut"
DbR.Execute "DELETE * FROM as_bono"

Do While Not RsAs.EOF
    
'If RsAs![RUT Operador1] = "16930532-3" Then
'    MsgBox ""
'End If
    
    m_Factor = -1
    If RsAs![dim7] Then
        m_Factor = 1000 * RsAs![PesoUnitario] / RsAs!dim7
    End If
       
    m_TipoEstructura = TipoEstructura(RsAs!TipoPieza)
    RsRas.Seek "=", m_TipoEstructura, RsAs![RUT Operador1]
    
    If RsRas.NoMatch Then
        ' nuevo rut
        RsRas.AddNew
        RsRas!TipoEstructura = m_TipoEstructura
        RsRas![rut] = RsAs![RUT Operador1]
        
        RsTrab.Seek "=", RsAs![RUT Operador1]
        If Not RsTrab.NoMatch Then
            RsRas!nombre = RsTrab![appaterno] & " " & RsTrab![apmaterno] & " " & RsTrab![nombres]
        End If
        AsBono_Linea RsRas, a_Tramos, RsAs!TipoPieza, RsAs!Cantidad, RsAs!PesoUnitario, m_Factor
        RsRas.Update
    Else
        ' rut ya existe
        RsRas.Edit
        AsBono_Linea RsRas, a_Tramos, RsAs!TipoPieza, RsAs!Cantidad, RsAs!PesoUnitario, m_Factor
        RsRas.Update
    End If
    
    ' para cuadro total de V, T, y S
    m_TipoEstructura = 99
    RsRas.Seek "=", m_TipoEstructura, RsAs![RUT Operador1]
    If RsRas.NoMatch Then
        ' nuevo rut
        RsRas.AddNew
        RsRas!TipoEstructura = m_TipoEstructura
        RsRas![rut] = RsAs![RUT Operador1]
        
        RsTrab.Seek "=", RsAs![RUT Operador1]
        If Not RsTrab.NoMatch Then
            RsRas!nombre = RsTrab![appaterno] & " " & RsTrab![apmaterno] & " " & RsTrab![nombres]
        End If
        AsBono_Linea RsRas, a_Tramos, RsAs!TipoPieza, RsAs!Cantidad, RsAs!PesoUnitario, m_Factor
        RsRas.Update
    Else
        ' rut ya existe
        RsRas.Edit
        AsBono_Linea RsRas, a_Tramos, RsAs!TipoPieza, RsAs!Cantidad, RsAs!PesoUnitario, m_Factor
        RsRas.Update
    End If
    
    RsAs.MoveNext

Loop

End Sub
Private Sub AsBono_Linea(RsRas As Recordset, a_Tramos, TipoPieza As String, Cantidad As Integer, PesoUnitario As Double, Factor As Double)
' graba linea de bono de arco sumergido
Dim m_Bono As Double, m_TipoEstructura As Integer

'Dim f As Integer, c As Integer
'For f = 0 To 30
'For c = 0 To 3
'    Debug.Print a_Tramos(f, c),
'Next
'Debug.Print
'Next

'    a_Tramos(m_TipoEstructura, 0) = m_TipoEstructura
'    a_Tramos(m_TipoEstructura, 1) = RsTrab!desde
'    a_Tramos(m_TipoEstructura, 2) = RsTrab!hasta
'    a_Tramos(m_TipoEstructura, 3) = RsTrab!valor
    
m_Bono = 0

With RsRas
Select Case TipoPieza
Case "V"

    m_TipoEstructura = 0

    !n_v = !n_v + Cantidad

    Select Case Factor
'    Case Is <= 30
    Case a_Tramos(m_TipoEstructura + 1, 1) To a_Tramos(m_TipoEstructura + 1, 2)
        !cef_1 = !cef_1 + Cantidad
        !pte_1 = !pte_1 + Cantidad * PesoUnitario
'        m_Bono = 5
        m_Bono = a_Tramos(m_TipoEstructura + 1, 3)
'    Case Is <= 60
    Case a_Tramos(m_TipoEstructura + 2, 1) To a_Tramos(m_TipoEstructura + 2, 2)
        !cef_2 = !cef_2 + Cantidad
        !pte_2 = !pte_2 + Cantidad * PesoUnitario
'        m_Bono = 4
        m_Bono = a_Tramos(m_TipoEstructura + 2, 3)
'    Case Is <= 100
    Case a_Tramos(m_TipoEstructura + 3, 1) To a_Tramos(m_TipoEstructura + 3, 2)
        !cef_3 = !cef_3 + Cantidad
        !pte_3 = !pte_3 + Cantidad * PesoUnitario
'        m_Bono = 3
        m_Bono = a_Tramos(m_TipoEstructura + 3, 3)
    Case Else
        !cef_4 = !cef_4 + Cantidad
        !pte_4 = !pte_4 + Cantidad * PesoUnitario
'        m_Bono = 2
        m_Bono = a_Tramos(m_TipoEstructura + 4, 3)
    End Select
    
Case "T"

    m_TipoEstructura = 10

    !n_t = !n_t + Cantidad
    Select Case Factor
'    Case Is <= 30
    Case a_Tramos(m_TipoEstructura + 1, 1) To a_Tramos(m_TipoEstructura + 1, 2)
        !cef_1 = !cef_1 + Cantidad
        !pte_1 = !pte_1 + Cantidad * PesoUnitario
'        m_Bono = 7
        m_Bono = a_Tramos(m_TipoEstructura + 1, 3)
'    Case Is <= 60
    Case a_Tramos(m_TipoEstructura + 2, 1) To a_Tramos(m_TipoEstructura + 2, 2)
        !cef_2 = !cef_2 + Cantidad
        !pte_2 = !pte_2 + Cantidad * PesoUnitario
'        m_Bono = 6
        m_Bono = a_Tramos(m_TipoEstructura + 2, 3)
'    Case Is <= 100
    Case a_Tramos(m_TipoEstructura + 3, 1) To a_Tramos(m_TipoEstructura + 3, 2)
        !cef_3 = !cef_3 + Cantidad
        !pte_3 = !pte_3 + Cantidad * PesoUnitario
'        m_Bono = 5
        m_Bono = a_Tramos(m_TipoEstructura + 3, 3)
    Case Else
        !cef_4 = !cef_4 + Cantidad
        !pte_4 = !pte_4 + Cantidad * PesoUnitario
'        m_Bono = 4
        m_Bono = a_Tramos(m_TipoEstructura + 4, 3)
    End Select
Case "S"

    m_TipoEstructura = 20
    
    !n_s = !n_s + Cantidad
    
    Select Case Factor
'    Case Is <= 30
    Case a_Tramos(m_TipoEstructura + 1, 1) To a_Tramos(m_TipoEstructura + 1, 2)
        !cef_1 = !cef_1 + Cantidad
        !pte_1 = !pte_1 + Cantidad * PesoUnitario
'        m_Bono = 4
        m_Bono = a_Tramos(m_TipoEstructura + 1, 3)
'    Case Is <= 60
    Case a_Tramos(m_TipoEstructura + 2, 1) To a_Tramos(m_TipoEstructura + 2, 2)
        !cef_2 = !cef_2 + Cantidad
        !pte_2 = !pte_2 + Cantidad * PesoUnitario
'        m_Bono = 4
        m_Bono = a_Tramos(m_TipoEstructura + 2, 3)
'    Case Is <= 100
    Case a_Tramos(m_TipoEstructura + 3, 1) To a_Tramos(m_TipoEstructura + 3, 2)
        !cef_3 = !cef_3 + Cantidad
        !pte_3 = !pte_3 + Cantidad * PesoUnitario
'        m_Bono = 4
        m_Bono = a_Tramos(m_TipoEstructura + 3, 3)
    Case Else
        !cef_4 = !cef_4 + Cantidad
        !pte_4 = !pte_4 + Cantidad * PesoUnitario
'        m_Bono = 4
        m_Bono = a_Tramos(m_TipoEstructura + 4, 3)
    End Select
    
End Select

!totalbono = !totalbono + m_Bono * Cantidad * PesoUnitario

End With
End Sub
Private Function TipoEstructura(te As String) As Integer
' solo para informe de bono arco sumergido
Select Case te
Case "V" ' viga
    TipoEstructura = 0
Case "T" ' tubular
    TipoEstructura = 10
Case "S" ' tubest
    TipoEstructura = 20
End Select
End Function
Public Sub Stock_Recalcular(Fecha As Date)
Dim DbD As Database, RsPrd As Recordset
Dim Dbm As Database, RsMB As Recordset
Dim sql As String, m_Cp As String, m_Cant As Double
' recalcula stock de productos de acuerdo a moviminetosde bodega
' y actualiza tabla "Productos" -> stock1
Set DbD = OpenDatabase(data_file)
Set RsPrd = DbD.OpenRecordset("Productos")
RsPrd.Index = "codigo"
Set Dbm = OpenDatabase(Madq_file)
sql = "SELECT * FROM documentos WHERE fecha <= #" & Month(Fecha) & "-" & Day(Fecha) & "-" & Year(Fecha) & "# ORDER BY [codigo producto],fecha"
Set RsMB = Dbm.OpenRecordset(sql)

m_Cp = RsMB![codigo producto]
m_Cant = 0

Do While Not RsMB.EOF

    If m_Cp = RsMB![codigo producto] Then
        
        If RsMB!Tipo = "IN" Then ' inventario
            m_Cant = 0
        End If
        
        m_Cant = m_Cant + RsMB![Cant_Entra] - RsMB![Cant_Sale]
        
    Else
    
        ' graba stock en tabla producto
        
        RsPrd.Seek "=", m_Cp
        If Not RsPrd.NoMatch Then
            RsPrd.Edit
            RsPrd![stock 1] = m_Cant
            RsPrd.Update
        End If
        
        m_Cp = RsMB![codigo producto]
        m_Cant = RsMB![Cant_Entra] - RsMB![Cant_Sale]
        
    End If
    
'    Debug.Print RsMB![codigo producto], RsMB!Fecha
    
    RsMB.MoveNext
    
Loop

RsPrd.Seek "=", m_Cp
If Not RsPrd.NoMatch Then
    RsPrd.Edit
    RsPrd![stock 1] = m_Cant
    RsPrd.Update
End If

End Sub
Public Sub Repo_inc_xgd(AreaCodigo As String, AreaDescripcion As String, Fecha_Ini As String, Fecha_Fin As String, RutEmisor As String)

' informe de no conformidad x area

Dim DbD As Database, RsT As Recordset
Dim RsNC As ADODB.Recordset
Dim RsMae As ADODB.Recordset, sql As String
Dim DbR As Database, RsINC As Recordset
Dim i As Integer
Dim a_Gerencias(19, 1) As String, a_Areas(99, 1), m_Descripcion As String

Set DbD = OpenDatabase(data_file)
Set RsT = DbD.OpenRecordset("trabajadores")
RsT.Index = "rut"

Set RsMae = New ADODB.Recordset
Set RsNC = New ADODB.Recordset

sql = "SELECT * FROM maestros WHERE tipo='GER'"
RsMae.Open sql, CnxSqlServer_scp0
i = 0
Do While Not RsMae.EOF
    i = i + 1
    a_Gerencias(i, 0) = RsMae!Codigo
    a_Gerencias(i, 1) = RsMae!Descripcion
    RsMae.MoveNext
Loop
RsMae.Close

sql = "SELECT * FROM maestros WHERE tipo='GAR'"
RsMae.Open sql, CnxSqlServer_scp0
i = 0
Do While Not RsMae.EOF
    i = i + 1
    a_Areas(i, 0) = RsMae!Codigo
    a_Areas(i, 1) = RsMae!Descripcion
    RsMae.MoveNext
Loop
RsMae.Close

Set DbR = OpenDatabase(repo_file)
DbR.Execute "DELETE * FROM inc_xa"
Set RsINC = DbR.OpenRecordset("inc_xa")
RsINC.Index = "gerencia-area-numero"

qry = SQL_MyQuery("e_fecha", Fecha_Ini, Fecha_Fin, "e_rut", RutEmisor)
sql = "SELECT * FROM noconformidad" & qry
RsNC.Open sql, CnxSqlServer_scp0

Do While Not RsNC.EOF
    
    RsINC.Seek "=", RsNC!e_gerencia, RsNC!e_area, RsNC!e_numero
    
    If RsINC.NoMatch Then
    
        RsINC.AddNew
        RsINC!Gerencia_Codigo = RsNC!e_gerencia
        RsINC!Area_Codigo = RsNC!e_area
    
    Else
    
        RsINC.Edit
    
    End If
    
    m_Descripcion = Arreglo_DescripcionBuscar(a_Gerencias, RsNC!e_gerencia)
    RsINC!gerencia_descripcion = m_Descripcion
    
    m_Descripcion = Arreglo_DescripcionBuscar(a_Areas, RsNC!e_area)
    RsINC!area_descripcion = m_Descripcion
        
    RsINC!Numero = RsNC!e_numero
    
    RsINC!Fecha_emision = RsNC!e_fecha
    
    If Not IsNull(RsNC!r_fechaPrimeraRespuesta) Then
        RsINC!fechaPrimeraRespuesta = RsNC!r_fechaPrimeraRespuesta
        i = RsNC!r_fechaPrimeraRespuesta - RsNC!e_fecha + 1
        RsINC!diasPrimeraRespuesta = i
    End If
    
'    RsT.Seek "=", Trim(RsNC!e_rut)
    RsT.Seek "=", RsNC!e_rut
    If Not RsT.NoMatch Then
        m_Descripcion = RsT!nombres & " " & RsT![appaterno]
        RsINC!Descripcion = StrConv(m_Descripcion, vbProperCase)
    End If
        
    If RsNC!e_Cerrado = "S" Then
        RsINC!Condicion = "Cerrada"
        RsINC!fecha_cierre = RsNC!e_FechaCierre
        i = RsNC!e_FechaCierre - RsNC!e_fecha + 1 ' + 1 ?
    Else
        RsINC!Condicion = "Abierta"
        i = Date - RsNC!e_fecha
    End If
    
    RsINC!costoestimado = RsNC!costoestimado
    RsINC!dias = i
    
    RsINC.Update
    
    RsNC.MoveNext
    
Loop

End Sub
Public Sub Repo_inc_resumen(Fecha_Ini As String, Fecha_Fin As String)
' resumen, incluye todas las gerencias, entre rango de fechas

Dim RsNC As New ADODB.Recordset, sql As String
Dim RsMaestros As New ADODB.Recordset
Dim a_Gerencias(99, 1) As String
Dim a_Areas(99, 1) As String
Dim i As Integer

' puebla gerencias
sql = "SELECT * FROM maestros WHERE tipo='GER'"
RsMaestros.Open sql, CnxSqlServer_scp0
i = 0
With RsMaestros
Do While Not .EOF
    i = i + 1
    a_Gerencias(i, 0) = !Codigo
    a_Gerencias(i, 1) = !Descripcion
    .MoveNext
Loop
.Close
End With

' puebla areas
sql = "SELECT * FROM maestros WHERE tipo='GAR'"
RsMaestros.Open sql, CnxSqlServer_scp0
i = 0
With RsMaestros
Do While Not .EOF
    i = i + 1
    a_Areas(i, 0) = !Codigo
    a_Areas(i, 1) = !Descripcion
    .MoveNext
Loop
.Close
End With


Dim DbR As Database, RsINC As Recordset
Dim Gerencia_Codigo As String, Area_Codigo As String, Numero As Integer

Set DbR = OpenDatabase(repo_file)
DbR.Execute "DELETE * FROM inc_xa"
Set RsINC = DbR.OpenRecordset("inc_xa")
RsINC.Index = "gerencia-area-numero"


qry = SQL_MyQuery("e_fecha", Fecha_Ini, Fecha_Fin, "", "")
sql = "SELECT * FROM noconformidad" & qry
RsNC.Open sql, CnxSqlServer_scp0
With RsNC
Do While Not .EOF

    Gerencia_Codigo = !e_gerencia
    Area_Codigo = !e_area
    
'    RsINC.Seek "=", Gerencia_Codigo, Area_Codigo, 0
    RsINC.Seek "=", Gerencia_Codigo, 0, 0
    If RsINC.NoMatch Then
        ' busca nombre de gerencia y area
    
        RsINC.AddNew
        RsINC!Gerencia_Codigo = Gerencia_Codigo
        RsINC!gerencia_descripcion = Arreglo_DescripcionBuscar(a_Gerencias, Gerencia_Codigo)
        RsINC!Area_Codigo = 0 ' Area_Codigo
        RsINC!area_descripcion = "" ' Arreglo_DescripcionBuscar(a_Areas, Area_Codigo)
        RsINC!Numero = 0
        
    Else
        RsINC.Edit
    End If
    RsINC!Total = RsINC!Total + 1
    If !e_Cerrado = "S" Then
        RsINC!cerradas = RsINC!cerradas + 1
    Else
        RsINC!abiertas = RsINC!abiertas + 1
    End If
    
    RsINC.Update
    
    .MoveNext
    
Loop
.Close
End With

End Sub
Public Sub Repo_gd_detalle(Nv As Double)

Dim Dbm As Database, RsGDd As Recordset, RsPlanoDetalle As Recordset, RsBulto As Recordset
Dim Dbi As Database, RsRepo As Recordset

Dim pu As Double, ct As Integer, des As String, bulto As Double

Dim tablaRepo As String
tablaRepo = "gd_det"

Set Dbi = OpenDatabase(repo_file)
Set RsRepo = Dbi.OpenRecordset(tablaRepo)
'RsRepo.Index = "gd_det"

' borra tabla de paso
Dbi.Execute "DELETE * FROM [" & tablaRepo & "]"

Set Dbm = OpenDatabase(mpro_file)

Set RsPlanoDetalle = Dbm.OpenRecordset("planos detalle")
RsPlanoDetalle.Index = "nv-plano-marca"

'Set RsBulto = Dbm.OpenRecordset("bultos")
'RsBulto.Index = "numero"

Set RsGDd = Dbm.OpenRecordset("SELECT * FROM [gd detalle] WHERE nv=" & Nv)
Set RsBulto = Dbm.OpenRecordset("bultos")

With RsGDd
Do While Not .EOF

    pu = 0
    ct = 0
    des = ""
    RsPlanoDetalle.Seek "=", !Nv, 0, !Plano, !Marca
    If Not RsPlanoDetalle.NoMatch Then
        des = RsPlanoDetalle!Descripcion
        ct = RsPlanoDetalle![Cantidad Total]
        pu = RsPlanoDetalle![Peso]
    End If
    
    ' busca bulto
    bulto = 0
    Set RsBulto = Dbm.OpenRecordset("SELECT * FROM bultos WHERE nv=" & Nv & " AND plano='" & !Plano & "' AND marca='" & !Marca & "'")
    If RsBulto.RecordCount > 0 Then
        Do While Not RsBulto.EOF
            If RsBulto!Cantidad = !Cantidad Then
                bulto = RsBulto!Numero
            End If
            RsBulto.MoveNext
        Loop
    End If
    '////////////

    RsRepo.AddNew
    RsRepo!Plano = !Plano
    RsRepo!Rev = !Rev
    RsRepo!Marca = !Marca
    RsRepo!Descripcion = des
    RsRepo!cantidadTotal = ct
    RsRepo!PesoUnitario = pu
    RsRepo!guia = !Numero
    RsRepo!Fecha = !Fecha
    RsRepo!cantidadDespachada = !Cantidad
    RsRepo!bulto = bulto
    RsRepo.Update

    .MoveNext

Loop

End With

End Sub
Public Sub RepoEnProceso()
' solicitada por
' cristian verdejo, contador general
' 24/03/2017

Dim Dbm As Database, RsPd As Recordset
Dim sql As String

' piezas en fabricacion
Set Dbm = OpenDatabase(mpro_file)
sql = "SELECT"
'sql = sql & " nv,"
sql = sql & " SUM(([ot fab]-[ito fab])*peso) AS ef,"
sql = sql & " SUM(([ito fab]-[ito gr])*peso) AS eg,"
sql = sql & " SUM(([ito gr]-[ito pp])*peso) AS epp,"
sql = sql & " SUM(([ito pp]-[ito pyg])*peso) AS ep,"
sql = sql & " SUM(([ito pyg]-[gd])*peso) AS pd"
sql = sql & " FROM [planos detalle]"
'sql = sql & " GROUP BY nv"
Set RsPd = Dbm.OpenRecordset(sql)

With RsPd
Do While Not .EOF

    'If !ef > 0 Then
    'If !eg < 0 Then

        Debug.Print " "
        'Debug.Print "|nv|" & !Nv & "|"
        Debug.Print "|ef|" & !ef & "|"
        Debug.Print "|eg|" & !eg & "|"
        Debug.Print "|epp|" & !epp & "|"
        Debug.Print "|ep|" & !ep & "|"
        Debug.Print "|pd|" & !pd & "|"
    
    'End If

    .MoveNext

Loop
End With

RsPd.Close
Dbm.Close

End Sub
