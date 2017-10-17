Attribute VB_Name = "Docs_Imprimir"
' imprime documentos
Option Explicit
Private prt As Printer
Private AjusteY As Double, AjusteX As Double
Public Sub GD_PrintLegal(Numero As Double, obra As String)
' imprime GUIA DE DESPACHO
Dim DbD As Database, RsCli As Recordset, RsPrd As Recordset
Dim Dbm As Database, RsGDc As Recordset, RsGDd As Recordset, RsPd As Recordset
Dim Dba As Database, RsDoc As Recordset
Dim fi As Double, m_desc As String, Tipo As String, linea As String
Dim m_Densidad As Integer, a_Den_d(1, 7) As Double, imprime_densidad As Boolean, Cantidad_Itemes As Long
Dim salto As Double, i As Integer

Dim mPlano As String

'a_Den_s(0, 7) = "Super Heavy"
'a_Den_s(0, 6) = "Heavy"
'a_Den_s(0, 5) = "Medium"
'a_Den_s(0, 4) = "Light"
'a_Den_s(0, 3) = "Grating ARS 6"
'a_Den_s(0, 2) = "Handrails"
'a_Den_s(0, 1) = "Stair Treads ARS 6"

linea = String(74, "-")

AjusteY = -2.6
AjusteX = -0.1

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt

prt.ScaleMode = vbCentimeters

Set Dbm = OpenDatabase(mpro_file)
Set RsGDc = Dbm.OpenRecordset("GD Cabecera")
RsGDc.Index = "Numero"

RsGDc.Seek "=", Numero
If RsGDc.NoMatch Then Exit Sub

Set Dba = OpenDatabase(Madq_file)
Set RsDoc = Dba.OpenRecordset("documentos")
RsDoc.Index = "tipo-numero-linea"

Set RsPd = Dbm.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Marca"

Set DbD = OpenDatabase(data_file)
Set RsCli = DbD.OpenRecordset("Clientes")
RsCli.Index = "RUT"

Set RsPrd = DbD.OpenRecordset("productos")
RsPrd.Index = "codigo"

prt.Font.Size = 12

SetC 3.3, 15.8
prt.Print Format(RsGDc!Numero, "000000")

SetC 6.6, 3.8
prt.Print UCase(Empresa.Ciudad & ", " & Format(RsGDc!Fecha, "d mmmm yyyy"))
SetC 6.6, 13.2
prt.Print RsGDc![RUT CLiente]

With RsCli
.Seek "=", RsGDc![RUT CLiente]
If Not .NoMatch Then

    SetC 7.2, 3.8
    prt.Print ![Razon Social]
    SetC 7.8, 3.8
    prt.Print !Direccion
    SetC 8.5, 3.8
    prt.Print !Comuna
    SetC 8.5, 13.2
    prt.Print NoNulo(![Telefono 1])
    SetC 9.1, 3.8
    prt.Print !Giro

End If
End With

SetC 9.5, 10
prt.Print UCase(obra)
'SetC 9.2, 12.2
'prt.Print NoNulo(RsGDc![Observacion 4])

Cantidad_Itemes = 0

' detalle
Select Case RsGDc!Tipo
Case "N", "G"

    linea = String(95, "-")
    'detalle normal
    'cabecera de detalle
'    SetC 10.1, 1.1
'    prt.Print "PLANO     R MARCA     DESCRIP CANT  KG UNI   KG TOTAL     $ UNI   $ TOTAL" ' oficial
    prt.Font.Size = 10
    
    SetC 10.4, 0.1 '10.1
    
    If RsGDc!Nv = 3109 Or RsGDc!Nv = 3110 Or RsGDc!Nv = 3111 Then
        prt.Print "PLANO      R MARCA                       DESCRIP  CANT   KG UNI   KG TOTAL     $ UNI    $ TOTAL"
        '          12345678901234567890 M 12345678901234567 12345678 234   12345,7   123456,8    12.345  1.234.567
    Else
        prt.Print "PLANO                R MARCA             DESCRIP  CANT   KG UNI   KG TOTAL     $ UNI    $ TOTAL"
        '          12345678901234567890 M 12345678901234567 12345678 234   12345,7   123456,8    12.345  1.234.567
    End If
    
    
    SetC 10.5, 0.1
'    prt.Print linea
    fi = 10.55
    Set RsGDd = Dbm.OpenRecordset("GD Detalle")
    RsGDd.Index = "Numero-Linea"
    RsGDd.Seek "=", Numero, 1
    
'    If Not RsGDd.EOF Then
    If Not RsGDd.NoMatch Then
    
        Do While Not RsGDd.EOF
        
            If Numero <> RsGDd!Numero Then Exit Do
            
                'linea
            fi = fi + 0.5
            
            If False Then
                
                SetC fi, 0 '0.2 ' 1.1
                prt.Print RsGDd!Plano
                SetC fi, 3.8 ' 3.65
                prt.Print RsGDd!Rev
                SetC fi, 4.2 ' 4.1
                prt.Print RsGDd!Marca
                
                m_desc = ""
                RsPd.Seek "=", RsGDd!Nv, RsGDd!NvArea, RsGDd!Plano, RsGDd!Marca
                If Not RsPd.NoMatch Then
                    m_desc = RsPd!Descripcion
                End If
                
                SetC fi, 6.6
                prt.Print Left(m_desc, 8)
                
                SetC fi, 8.2
                prt.Print m_Format(RsGDd!Cantidad, "#,###")
                            
                SetC fi, 9.5
                prt.Print m_Format(RsGDd![Peso Unitario], "###,##0.0")
                SetC fi, 12.3
                prt.Print m_Format(RsGDd!Cantidad * RsGDd![Peso Unitario], "###,##0.0")
                SetC fi, 14.8
                prt.Print m_Format(RsGDd![Precio Unitario], "#,###,##0")
                SetC fi, 17.1
                prt.Print m_Format(RsGDd!Cantidad * RsGDd![Precio Unitario] * RsGDd![Peso Unitario], "##,###,###")
            
            Else
            
                If RsGDc!Nv = 3109 Or RsGDc!Nv = 3110 Or RsGDc!Nv = 3111 Then
                            
                    SetC fi, 0.1
                    
                    mPlano = RsGDd!Plano
                    mPlano = Mid(mPlano, InStrLast(mPlano, "-") + 1)
                    prt.Print mPlano
                    
                    SetC fi, 2.5
                    prt.Print RsGDd!Rev
                    SetC fi, 3
                    prt.Print RsGDd!Marca
                
                Else
                
                    SetC fi, 0.1
                    prt.Print RsGDd!Plano
                    SetC fi, 4.6
                    prt.Print RsGDd!Rev
                    SetC fi, 4.9
                    prt.Print RsGDd!Marca
                
                End If
                
                m_desc = ""
                m_Densidad = 0
                RsPd.Seek "=", RsGDd!Nv, RsGDd!NvArea, RsGDd!Plano, RsGDd!Marca
                If Not RsPd.NoMatch Then
                    m_desc = RsPd!Descripcion
                    m_Densidad = RsPd!densidad
                End If
                SetC fi, 8.8
                prt.Print Left(m_desc, 8)
                
                ' clasifica segun densidad de marca
                'Select Case m_Densidad
                'Case 7
                    ' cantidad de marcas
                    a_Den_d(0, m_Densidad) = a_Den_d(0, m_Densidad) + RsGDd!Cantidad
                    ' peso de marcas
                    a_Den_d(1, m_Densidad) = a_Den_d(1, m_Densidad) + RsGDd!Cantidad * RsGDd![Peso Unitario]
                'End Select
                
                SetC fi, 10.4
                prt.Print m_Format(RsGDd!Cantidad, "#,###")
                
                Cantidad_Itemes = Cantidad_Itemes + RsGDd!Cantidad
                
                SetC fi, 11.5
                prt.Print m_Format(RsGDd![Peso Unitario], "###,##0.0")
                SetC fi, 13.9
                prt.Print m_Format(RsGDd!Cantidad * RsGDd![Peso Unitario], "###,##0.0")
                SetC fi, 16
                prt.Print m_Format(RsGDd![Precio Unitario], "#,###,##0")
                SetC fi, 18.05
                prt.Print m_Format(RsGDd!Cantidad * RsGDd![Precio Unitario] * RsGDd![Peso Unitario], "##,###,##0")
            
            End If
            
            RsGDd.MoveNext
            
        Loop
        
    End If
    
    prt.Font.Size = 12
    
Case "E"

    linea = String(90, "-")

    'detalle especial
    'cabecera de detalle
'    SetC 10.1, 1.4

    prt.Font.Size = 10
    
    SetC 10.4, 0.4
    
    If RsGDc!Tipo = "E" Then
        prt.Print " CANT UNI DETALLE                          KG UNI   KG TOTAL  $ UNI  $ TOTAL"
    Else
        prt.Print " CANT UNI  DETALLE                m2 UNI     m2 TOTAL    $ UNI   $ TOTAL"
    End If
    
    SetC 10.5, 0.3
'    prt.Print linea
    fi = 10.55
    Set RsGDd = Dbm.OpenRecordset("GD Especial Detalle")
    RsGDd.Index = "Numero-Linea"
    RsGDd.Seek "=", Numero, 1
'    If Not RsGDd.EOF Then
    If Not RsGDd.NoMatch Then
        Do While Not RsGDd.EOF
        
            If Numero <> RsGDd!Numero Then Exit Do
            
            'linea
            fi = fi + 0.5
            
            SetC fi, 0.3
'            prt.Print m_Format(RsGDd!Cantidad, "#,###")
            prt.Print m_Format(RsGDd!Cantidad, "#####")
            
            Cantidad_Itemes = Cantidad_Itemes + RsGDd!Cantidad
            
'            SetC fi, 1.65
            SetC fi, 2
            prt.Print RsGDd!unidad
            SetC fi, 2.9 '2.7
            prt.Print RsGDd!Detalle
'            SetC fi, 8.5
'            prt.Print m_Format(RsGDd![Peso Unitario], "##,###,##0.0")
            SetC fi, 10.7
            prt.Print m_Format(RsGDd![Peso Unitario], "##,##0.0")
            SetC fi, 12.3
            prt.Print m_Format(RsGDd!Cantidad * RsGDd![Peso Unitario], "###,###,##0.0")
            SetC fi, 15.2
            prt.Print m_Format(RsGDd![Precio Unitario], "#,###,##0")
            SetC fi, 17.2
            prt.Print m_Format(RsGDd!Cantidad * RsGDd![Peso Unitario] * RsGDd![Precio Unitario], "##,###,##0")
            
            RsGDd.MoveNext
            
        Loop
    End If
    
    prt.Font.Size = 12

Case "P"

    'detalle pernos
    'cabecera de detalle
    SetC 10.4, 1.4
    
    prt.Print "CODIGO         DETALLE                           CANT     $ UNI    $ TOTAL"
    
    SetC 10.5, 1.3
'    prt.Print linea
    fi = 10.55
    With RsDoc
    .Seek "=", "GP", Numero, 1
    If Not .EOF Then
        Do While Not .EOF
        
            If !Tipo <> "GP" Or !Numero <> Numero Then Exit Do
            
            'linea
            fi = fi + 0.5
            
            SetC fi, 1.3
            prt.Print m_Format(![codigo producto], "#,###")
            
            RsPrd.Seek "=", ![codigo producto]
            If Not RsPrd.NoMatch Then
                SetC fi, 5
                prt.Print RsPrd![Descripcion]
            End If
            
            SetC fi, 13.5
            prt.Print m_Format(!Cant_Sale, "#,###")
            
            Cantidad_Itemes = Cantidad_Itemes + !Cant_Sale
            
            SetC fi, 15.1
            prt.Print m_Format(![Precio Unitario], "#,###,##0")
'            SetC fi, 18
            SetC fi, 17.5
            prt.Print m_Format(!Cant_Sale * ![Precio Unitario], "##,###,##0")
            
            .MoveNext
            
        Loop
    End If
    End With
    
End Select

If Cantidad_Itemes > 0 Then
    SetC 21.3, 0.1 ' 20.3
    prt.Print "CANT TOT : "; m_Format(Cantidad_Itemes, "##0")
End If
SetC 21.3, 5.8 ' 20.3
prt.Print "PESO TOTAL : "; m_Format(RsGDc![Peso Total], "##,###,##0.0")
SetC 21.3, 13 ' 20.3
prt.Print "PRECIO TOTAL : "; m_Format(RsGDc![Precio Total], "$###,###,##0")

SetC 22, 1.8
prt.Print "CHOFER    : "; NoNulo(RsGDc![Observacion 1])
SetC 22.5, 1.8
prt.Print "PATENTE   : "; NoNulo(RsGDc![Observacion 2])

SetC 23, 1.8
If RsGDc!Tipo = "P" Then
    prt.Print "ESQUEMA : "; NoNulo(RsGDc![Observacion 3])
Else
    prt.Print "CONTENIDO : "; NoNulo(RsGDc![Observacion 3])
End If

' imprime resumen separado por densidad, pedido por francisco cruces para cliente sedgman
If RsGDc![RUT CLiente] = "76300170-9" Then ' verifica si es cliente sedgman

    imprime_densidad = False
    For i = 1 To 7
        If a_Den_d(0, i) > 0 Then
            imprime_densidad = True
        End If
    Next
    
    salto = 23.5 ' 23
    If imprime_densidad Then
        For i = 7 To 1 Step -1
            If a_Den_d(0, i) > 0 Then
            
                salto = salto + 0.5
                
                SetC salto, 9
                prt.Print a_Den_s(0, i)
                
                SetC salto, 14
                prt.Print m_Format(a_Den_d(0, i), "#,##0")
                
                SetC salto, 15
                prt.Print m_Format(a_Den_d(1, i), "##,###,##0.0") & " Kgs"
                
            End If
        Next
    End If

End If
'SetC 24, 0
'prt.Print "            "; NoNulo(RsGDc![Observación 4])

prt.EndDoc

Impresora_Predeterminada "default"

End Sub
Public Sub GD_PrintLegal_AyD(Numero As Double, obra As String)
' imprime GUIA DE DESPACHO
Dim DbD As Database, RsCli As Recordset, RsPrd As Recordset
Dim Dbm As Database, RsGDc As Recordset, RsGDd As Recordset, RsPd As Recordset
Dim Dba As Database, RsDoc As Recordset
Dim fi As Double, m_desc As String, Tipo As String, linea As String
linea = String(74, "-")

AjusteY = -2.5 ' -2.2
AjusteX = 0

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt

prt.ScaleMode = vbCentimeters

Set Dbm = OpenDatabase(mpro_file)
Set RsGDc = Dbm.OpenRecordset("GD Cabecera")
RsGDc.Index = "Numero"

RsGDc.Seek "=", Numero
If RsGDc.NoMatch Then Exit Sub

Set Dba = OpenDatabase(Madq_file)
Set RsDoc = Dba.OpenRecordset("documentos")
RsDoc.Index = "tipo-numero-linea"

Set RsPd = Dbm.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Marca"

Set DbD = OpenDatabase(data_file)
Set RsCli = DbD.OpenRecordset("Clientes")
RsCli.Index = "RUT"

Set RsPrd = DbD.OpenRecordset("productos")
RsPrd.Index = "codigo"

prt.Font.Size = 12

SetC 4.5, 16.8
prt.Print Format(RsGDc!Numero, "000000")

SetC 6.3, 2.5
prt.Print Format(RsGDc!Fecha, "d")
SetC 6.3, 4.8
prt.Print Format(RsGDc!Fecha, "mmmm")
SetC 6.3, 8.7
prt.Print Format(RsGDc!Fecha, "yy")

With RsCli
.Seek "=", RsGDc![RUT CLiente]
If Not .NoMatch Then

    SetC 6.75, 3.5
    prt.Print ![Razon Social]
    
    SetC 6.5, 15
    prt.Print RsGDc![RUT CLiente]

    SetC 7.1, 3.5
    prt.Print !Direccion
    
    SetC 7.4, 15
    prt.Print !Comuna
    
    SetC 8.3, 3.5
    prt.Print NoNulo(![Telefono 1])
    
'    SetC 8.5, 4.3
'    prt.Print !Giro

End If
End With

'SetC 8.7, 13.2
'prt.Print UCase(Obra)

'SetC 9.2, 13.2
'prt.Print NoNulo(RsGDc![Observación 4])

' detalle
Select Case RsGDc!Tipo
Case "N", "G"

    'detalle normal
    'cabecera de detalle
'    SetC 11.5, 1.1
    SetC 9.6, 1.1
    prt.Print "PLANO     R MARCA     DESCRIP CANT  KG UNI   KG TOTAL     $ UNI   $ TOTAL"
'    SetC 12, 1.1
'    prt.Print linea
    fi = 10.6 '12
    Set RsGDd = Dbm.OpenRecordset("GD Detalle")
    RsGDd.Index = "Numero-Linea"
    RsGDd.Seek "=", Numero, 1
    If Not RsGDd.EOF Then
        Do While Not RsGDd.EOF
            If Numero <> RsGDd!Numero Then Exit Do
            
            'linea
            fi = fi + 0.5
            
            SetC fi, 1.1
            prt.Print RsGDd!Plano
            SetC fi, 3.65
            prt.Print RsGDd!Rev
            SetC fi, 4.1
            prt.Print RsGDd!Marca
            
            m_desc = ""
            RsPd.Seek "=", RsGDd!Nv, RsGDd!NvArea, RsGDd!Plano, RsGDd!Marca
            If Not RsPd.NoMatch Then m_desc = RsPd!Descripcion
            SetC fi, 6.6
            prt.Print Left(m_desc, 8)
            
            SetC fi, 8.2
            prt.Print m_Format(RsGDd!Cantidad, "#,###")
            SetC fi, 9.5
            prt.Print m_Format(RsGDd![Peso Unitario], "###,##0.0")
            SetC fi, 12.3
            prt.Print m_Format(RsGDd!Cantidad * RsGDd![Peso Unitario], "###,##0.0")
            SetC fi, 14.8
            prt.Print m_Format(RsGDd![Precio Unitario], "#,###,##0")
            SetC fi, 17.1
            prt.Print m_Format(RsGDd!Cantidad * RsGDd![Precio Unitario] * RsGDd![Peso Unitario], "##,###,###")
            
            RsGDd.MoveNext
        Loop
    End If
    
Case "E"

    'detalle especial
    'cabecera de detalle
'    SetC 11.5, 1.4
    SetC 9.6, 1.4
    
    If RsGDc!Tipo = "E" Then
        prt.Print " CANT UNI  DETALLE                KG UNI     KG TOTAL    $ UNI   $ TOTAL"
    Else
        prt.Print " CANT UNI  DETALLE                m2 UNI     m2 TOTAL    $ UNI   $ TOTAL"
    End If
    
'    SetC 12, 1.3
'    prt.Print linea
    fi = 10.6 ' 12
    Set RsGDd = Dbm.OpenRecordset("GD Especial Detalle")
    RsGDd.Index = "Numero-Linea"
    RsGDd.Seek "=", Numero, 1
    If Not RsGDd.EOF Then
        Do While Not RsGDd.EOF
        
            If Numero <> RsGDd!Numero Then Exit Do
            
            'linea
            fi = fi + 0.5
            
            SetC fi, 1.3
            prt.Print m_Format(RsGDd!Cantidad, "#,###")
            SetC fi, 2.9
            prt.Print RsGDd!unidad
            SetC fi, 4.1
            prt.Print RsGDd!Detalle
'            SetC fi, 8.5
'            prt.Print m_Format(RsGDd![Peso Unitario], "##,###,##0.0")
            SetC fi, 9.5
            prt.Print m_Format(RsGDd![Peso Unitario], "##,##0.0")
            SetC fi, 11.5
            prt.Print m_Format(RsGDd!Cantidad * RsGDd![Peso Unitario], "###,###,##0.0")
            SetC fi, 14.8
            prt.Print m_Format(RsGDd![Precio Unitario], "#,###,##0")
            SetC fi, 17.2
            prt.Print m_Format(RsGDd!Cantidad * RsGDd![Peso Unitario] * RsGDd![Precio Unitario], "##,###,##0")
            
            RsGDd.MoveNext
            
        Loop
    End If
    
Case "P"

    'detalle pernos
    'cabecera de detalle
    SetC 11.5, 1.4
    
    prt.Print "CODIGO         DETALLE                            CANT    $ UNI    $ TOTAL"
    
    SetC 12, 1.3
    prt.Print linea
    fi = 12
    With RsDoc
    .Seek "=", "GP", Numero, 1
    If Not .EOF Then
        Do While Not .EOF
        
            If !Tipo <> "GP" Or !Numero <> Numero Then Exit Do
            
            'linea
            fi = fi + 0.5
            
            SetC fi, 1.3
            prt.Print m_Format(![codigo producto], "#,###")
            
            RsPrd.Seek "=", ![codigo producto]
            If Not RsPrd.NoMatch Then
                SetC fi, 5
                prt.Print RsPrd![descripción]
            End If
            
            SetC fi, 14
            prt.Print m_Format(!Cant_Sale, "#,###")
            SetC fi, 15
            prt.Print m_Format(![Precio Unitario], "#,###,##0")
            SetC fi, 18
            prt.Print m_Format(!Cant_Sale * ![Precio Unitario], "##,###,##0")
            
            .MoveNext
            
        Loop
    End If
    End With
    
End Select

SetC 22.1, 13
prt.Print "PESO TOTAL   : "; m_Format(RsGDc![Peso Total], "##,###,##0.0")
SetC 22.6, 13
prt.Print "PRECIO TOTAL : "; m_Format(RsGDc![Precio Total], "$###,###,##0")

SetC 22.1, 1
prt.Print "CHOFER    : "; NoNulo(RsGDc![Observacion 1])
SetC 22.6, 1
prt.Print "PATENTE   : "; NoNulo(RsGDc![Observacion 2])

SetC 23.1, 1
If RsGDc!Tipo = "P" Then
    prt.Print "ESQUEMA : "; NoNulo(RsGDc![Observacion 3])
Else
    prt.Print "CONTENIDO : "; NoNulo(RsGDc![Observacion 3])
End If

'SetC 24, 0
'prt.Print "            "; NoNulo(RsGDc![Observación 4])

prt.EndDoc

Impresora_Predeterminada "default"

End Sub
Public Sub GD_PrintLegal_Delsa(Numero As Double, obra As String)
' imprime GUIA DE DESPACHO
Dim DbD As Database, RsCli As Recordset, RsPrd As Recordset
Dim Dbm As Database, RsGDc As Recordset, RsGDd As Recordset, RsPd As Recordset
Dim Dba As Database, RsDoc As Recordset
Dim fi As Double, m_desc As String, Tipo As String, linea As String
linea = String(74, "-")

AjusteY = -2.5 ' -2.2
AjusteX = 0

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt

prt.ScaleMode = vbCentimeters

Set Dbm = OpenDatabase(mpro_file)
Set RsGDc = Dbm.OpenRecordset("GD Cabecera")
RsGDc.Index = "Numero"

RsGDc.Seek "=", Numero
If RsGDc.NoMatch Then Exit Sub

Set Dba = OpenDatabase(Madq_file)
Set RsDoc = Dba.OpenRecordset("documentos")
RsDoc.Index = "tipo-numero-linea"

Set RsPd = Dbm.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Marca"

Set DbD = OpenDatabase(data_file)
Set RsCli = DbD.OpenRecordset("Clientes")
RsCli.Index = "RUT"

Set RsPrd = DbD.OpenRecordset("productos")
RsPrd.Index = "codigo"

prt.Font.Size = 12

SetC 4.9, 15.3
prt.Print Format(RsGDc!Numero, "000000")

SetC 7, 2.6
prt.Print Format(RsGDc!Fecha, "d")
SetC 7, 4.8
prt.Print Format(RsGDc!Fecha, "mmmm")
SetC 7, 9.7
prt.Print Format(RsGDc!Fecha, "yyyy")

With RsCli
.Seek "=", RsGDc![RUT CLiente]
If Not .NoMatch Then

    SetC 7.8, 3.5
    prt.Print ![Razon Social]
    
    SetC 8.5, 3.5
    prt.Print !Direccion
    
    SetC 8.5, 16
    prt.Print RsGDc![RUT CLiente]
    
    SetC 9.2, 3.5
    prt.Print !Giro
    
    SetC 9.8, 3.5
    prt.Print NoNulo(![Telefono 1])
    
    SetC 9.8, 14
    prt.Print !Comuna

End If
End With

'SetC 8.7, 13.2
'prt.Print UCase(Obra)

'SetC 9.2, 13.2
'prt.Print NoNulo(RsGDc![Observación 4])

' detalle
Select Case RsGDc!Tipo
Case "N", "G"

    'detalle normal
    'cabecera de detalle
    SetC 12.3, 1.1
'              12345678901 1234567 1234567890123456 1234 1234,5 12.345,6 1.234 1.234.567
    prt.Print "PLANO       MARCA   DESCRIP          CANT KG.UNI KG TOTAL $ UNI   $ TOTAL"
    fi = 12.35
    Set RsGDd = Dbm.OpenRecordset("GD Detalle")
    RsGDd.Index = "Numero-Linea"
    RsGDd.Seek "=", Numero, 1
    If Not RsGDd.EOF Then
        Do While Not RsGDd.EOF
            If Numero <> RsGDd!Numero Then Exit Do
            
            'linea
            fi = fi + 0.55
            
            SetC fi, 1.1
            prt.Print Left(RsGDd!Plano, 11)
'            SetC fi, 3.65
'            prt.Print RsGDd!Rev
            SetC fi, 4.1
            prt.Print Left(RsGDd!Marca, 7)
            
            m_desc = ""
            RsPd.Seek "=", RsGDd!Nv, RsGDd!NvArea, RsGDd!Plano, RsGDd!Marca
            If Not RsPd.NoMatch Then m_desc = RsPd!Descripcion
            SetC fi, 6.2
            prt.Print Left(m_desc, 16)
            
            SetC fi, 10.5
            prt.Print m_Format(RsGDd!Cantidad, "####")
            SetC fi, 11.7
            prt.Print m_Format(RsGDd![Peso Unitario], "###0.0")
            SetC fi, 13.6
            prt.Print m_Format(RsGDd!Cantidad * RsGDd![Peso Unitario], "##,##0.0")
            SetC fi, 15.8
            prt.Print m_Format(RsGDd![Precio Unitario], "#,##0")
            SetC fi, 17.4
            prt.Print m_Format(RsGDd!Cantidad * RsGDd![Precio Unitario] * RsGDd![Peso Unitario], "#,###,###")
            
            RsGDd.MoveNext
        Loop
    End If
    
Case "E"

    'detalle especial
    'cabecera de detalle
    SetC 12.3, 1.3
    
    If RsGDc!Tipo = "E" Then
        prt.Print " CANT UNI  DETALLE                KG UNI     KG TOTAL    $ UNI   $ TOTAL"
    Else
        prt.Print " CANT UNI  DETALLE                m2 UNI     m2 TOTAL    $ UNI   $ TOTAL"
    End If
    
    fi = 12.35
    Set RsGDd = Dbm.OpenRecordset("GD Especial Detalle")
    RsGDd.Index = "Numero-Linea"
    RsGDd.Seek "=", Numero, 1
    If Not RsGDd.EOF Then
        Do While Not RsGDd.EOF
        
            If Numero <> RsGDd!Numero Then Exit Do
            
            'linea
            fi = fi + 0.55
            
            SetC fi, 1.3
            prt.Print m_Format(RsGDd!Cantidad, "#,###")
            SetC fi, 2.9
            prt.Print RsGDd!unidad
            SetC fi, 4.1
            prt.Print RsGDd!Detalle
'            SetC fi, 8.5
'            prt.Print m_Format(RsGDd![Peso Unitario], "##,###,##0.0")
            SetC fi, 9.5
            prt.Print m_Format(RsGDd![Peso Unitario], "##,##0.0")
            SetC fi, 11.5
            prt.Print m_Format(RsGDd!Cantidad * RsGDd![Peso Unitario], "###,###,##0.0")
            SetC fi, 14.8
            prt.Print m_Format(RsGDd![Precio Unitario], "#,###,##0")
            SetC fi, 17.2
            prt.Print m_Format(RsGDd!Cantidad * RsGDd![Peso Unitario] * RsGDd![Precio Unitario], "##,###,##0")
            
            RsGDd.MoveNext
            
        Loop
    End If
    
Case "P"

    'detalle pernos
    'cabecera de detalle
    SetC 11.5, 1.4
    
    prt.Print "CODIGO         DETALLE                            CANT    $ UNI    $ TOTAL"
    
    SetC 12, 1.3
    prt.Print linea
    fi = 12
    With RsDoc
    .Seek "=", "GP", Numero, 1
    If Not .EOF Then
        Do While Not .EOF
        
            If !Tipo <> "GP" Or !Numero <> Numero Then Exit Do
            
            'linea
            fi = fi + 0.5
            
            SetC fi, 1.3
            prt.Print m_Format(![codigo producto], "#,###")
            
            RsPrd.Seek "=", ![codigo producto]
            If Not RsPrd.NoMatch Then
                SetC fi, 5
                prt.Print RsPrd![descripción]
            End If
            
            SetC fi, 14
            prt.Print m_Format(!Cant_Sale, "#,###")
            SetC fi, 15
            prt.Print m_Format(![Precio Unitario], "#,###,##0")
            SetC fi, 18
            prt.Print m_Format(!Cant_Sale * ![Precio Unitario], "##,###,##0")
            
            .MoveNext
            
        Loop
    End If
    End With
    
End Select


SetC 22.05, 13.3
prt.Print "PESO TOTAL   : "; m_Format(RsGDc![Peso Total], "###,##0.0")
SetC 22.6, 13.3
prt.Print "PRECIO TOTAL : "; m_Format(RsGDc![Precio Total], "$#,###,##0")
SetC 23.15, 13.3
prt.Print "VALORES NETOS"


SetC 21, 1.5
prt.Print "CHOFER    : "; NoNulo(RsGDc![Observacion 1])
SetC 21.55, 1.5
prt.Print "PATENTE   : "; NoNulo(RsGDc![Observacion 2])

SetC 22.1, 1.5
If RsGDc!Tipo = "P" Then
    prt.Print "ESQUEMA : "; NoNulo(RsGDc![Observacion 3])
Else
    prt.Print "CONTENIDO : "; Left(NoNulo(RsGDc![Observacion 3]), 30)
End If
SetC 22.65, 1.5
prt.Print NoNulo(RsGDc![Observacion 4])

'SetC 24, 0
'prt.Print "            "; NoNulo(RsGDc![Observación 4])

prt.EndDoc

Impresora_Predeterminada "default"

End Sub
Public Sub GD_PrintLegal_Eiffel(Numero As Double, obra As String)
' imprime GUIA DE DESPACHO
Dim DbD As Database, RsCli As Recordset, RsPrd As Recordset
Dim Dbm As Database, RsGDc As Recordset, RsGDd As Recordset, RsPd As Recordset
Dim Dba As Database, RsDoc As Recordset
Dim fi As Double, m_desc As String, Tipo As String, linea As String
Dim m_Densidad As Integer, a_Den_d(1, 7) As Double, imprime_densidad As Boolean, Cantidad_Itemes As Long
Dim salto As Double, i As Integer

'a_Den_s(0, 7) = "Super Heavy"
'a_Den_s(0, 6) = "Heavy"
'a_Den_s(0, 5) = "Medium"
'a_Den_s(0, 4) = "Light"
'a_Den_s(0, 3) = "Grating ARS 6"
'a_Den_s(0, 2) = "Handrails"
'a_Den_s(0, 1) = "Stair Treads ARS 6"

linea = String(74, "-")

AjusteY = -2.4
AjusteX = -0.5

'Printer_Set "Documentos"
Set prt = Printer
Font_Setear prt

prt.ScaleMode = vbCentimeters

Set Dbm = OpenDatabase(mpro_file)
Set RsGDc = Dbm.OpenRecordset("GD Cabecera")
RsGDc.Index = "Numero"

RsGDc.Seek "=", Numero
If RsGDc.NoMatch Then Exit Sub

Set Dba = OpenDatabase(Madq_file)
Set RsDoc = Dba.OpenRecordset("documentos")
RsDoc.Index = "tipo-numero-linea"

Set RsPd = Dbm.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Marca"

Set DbD = OpenDatabase(data_file)
Set RsCli = DbD.OpenRecordset("Clientes")
RsCli.Index = "RUT"

Set RsPrd = DbD.OpenRecordset("productos")
RsPrd.Index = "codigo"

prt.Font.Size = 12

SetC 4.2, 16.3
prt.Print Format(RsGDc!Numero, "000000")

SetC 6.6, 3.8
prt.Print UCase(Empresa.Ciudad & ", " & Format(RsGDc!Fecha, "d mmmm yyyy"))

SetC 7.2, 16.2
prt.Print RsGDc![RUT CLiente]

With RsCli
.Seek "=", RsGDc![RUT CLiente]
If Not .NoMatch Then

    SetC 7.2, 3.8
    prt.Print ![Razon Social]
    
    SetC 7.8, 3.8
    prt.Print !Direccion
    SetC 7.8, 17
    prt.Print !Comuna
    
    SetC 8.5, 3.8
    prt.Print !Giro
    SetC 8.5, 17
    prt.Print NoNulo(![Telefono 1])
    
End If
End With


SetC 9.2, 17.5
prt.Print Left(obra, 5) ' numero de nv

SetC 9.8, 3
prt.Print UCase(Mid(obra, 8)) ' nombre obra

SetC 9.8, 18
prt.Print NoNulo(RsGDc![Observacion 2]) ' patente

SetC 10.5, 4
prt.Print NoNulo(RsGDc![Observacion 4]) ' direccion obra

Cantidad_Itemes = 0

' detalle
Select Case RsGDc!Tipo
Case "N", "G"

    linea = String(95, "-")
    'detalle normal
    'cabecera de detalle
'    SetC 10.1, 1.1
'    prt.Print "PLANO     R MARCA     DESCRIP CANT  KG UNI   KG TOTAL     $ UNI   $ TOTAL" ' oficial
    prt.Font.Size = 10
    
'    SetC 10.4, 0.1 '10.1
'    prt.Print "PLANO                R MARCA             DESCRIP  CANT   KG UNI   KG TOTAL     $ UNI    $ TOTAL"
    '          12345678901234567890 M 12345678901234567 12345678 234   12345,7   123456,8    12.345  1.234.567
    SetC 10.5, 0.1
'    prt.Print linea
    fi = 12
    Set RsGDd = Dbm.OpenRecordset("GD Detalle")
    RsGDd.Index = "Numero-Linea"
    RsGDd.Seek "=", Numero, 1
    
'    If Not RsGDd.EOF Then
    If Not RsGDd.NoMatch Then
    
        Do While Not RsGDd.EOF
        
            If Numero <> RsGDd!Numero Then Exit Do
            
                'linea
            fi = fi + 0.5
            
            If False Then
                
                SetC fi, 0 '0.2 ' 1.1
                prt.Print RsGDd!Plano
                SetC fi, 3.8 ' 3.65
                prt.Print RsGDd!Rev
                SetC fi, 4.2 ' 4.1
                prt.Print RsGDd!Marca
                
                m_desc = ""
                RsPd.Seek "=", RsGDd!Nv, RsGDd!NvArea, RsGDd!Plano, RsGDd!Marca
                If Not RsPd.NoMatch Then
                    m_desc = RsPd!Descripcion
                End If
                
                SetC fi, 6.6
                prt.Print Left(m_desc, 8)
                
                SetC fi, 8.2
                prt.Print m_Format(RsGDd!Cantidad, "#,###")
                            
                SetC fi, 9.5
                prt.Print m_Format(RsGDd![Peso Unitario], "###,##0.0")
                SetC fi, 12.3
                prt.Print m_Format(RsGDd!Cantidad * RsGDd![Peso Unitario], "###,##0.0")
                SetC fi, 14.8
                prt.Print m_Format(RsGDd![Precio Unitario], "#,###,##0")
                SetC fi, 17.1
                prt.Print m_Format(RsGDd!Cantidad * RsGDd![Precio Unitario] * RsGDd![Peso Unitario], "##,###,###")
            
            Else
            
                m_desc = ""
                m_Densidad = 0
                RsPd.Seek "=", RsGDd!Nv, RsGDd!NvArea, RsGDd!Plano, RsGDd!Marca
                If Not RsPd.NoMatch Then
                    m_desc = RsPd!Descripcion
                    m_Densidad = RsPd!densidad
                End If
                
                SetC fi, 2.1
                prt.Print m_Format(RsGDd!Cantidad, "#,###")
                
                SetC fi, 4
                prt.Print Left(m_desc, 8)
                
                SetC fi, 11
                prt.Print RsGDd!Plano
                SetC fi, 15
                prt.Print RsGDd!Rev
                SetC fi, 15.5
                prt.Print RsGDd!Marca
                                
                Cantidad_Itemes = Cantidad_Itemes + RsGDd!Cantidad
                            
            End If
            
            RsGDd.MoveNext
            
        Loop
        
    End If
    
    prt.Font.Size = 12
    
Case "E"

    linea = String(90, "-")

    'detalle especial
    'cabecera de detalle
'    SetC 10.1, 1.4

    prt.Font.Size = 10
    
    SetC 10.4, 0.4
    
'    If RsGDc!Tipo = "E" Then
'        prt.Print " CANT UNI DETALLE                          KG UNI   KG TOTAL  $ UNI  $ TOTAL"
'    Else
'        prt.Print " CANT UNI  DETALLE                m2 UNI     m2 TOTAL    $ UNI   $ TOTAL"
'    End If
    
    SetC 10.5, 0.3
'    prt.Print linea
    fi = 12
    Set RsGDd = Dbm.OpenRecordset("GD Especial Detalle")
    RsGDd.Index = "Numero-Linea"
    RsGDd.Seek "=", Numero, 1
'    If Not RsGDd.EOF Then
    If Not RsGDd.NoMatch Then
        Do While Not RsGDd.EOF
        
            If Numero <> RsGDd!Numero Then Exit Do
            
            'linea
            fi = fi + 0.5
            
            SetC fi, 2
'            prt.Print m_Format(RsGDd!Cantidad, "#,###")
            prt.Print m_Format(RsGDd!Cantidad, "#####")
            
            Cantidad_Itemes = Cantidad_Itemes + RsGDd!Cantidad
            
'            SetC fi, 1.65
            SetC fi, 3.5
            prt.Print RsGDd!unidad
            SetC fi, 4.5
            prt.Print RsGDd!Detalle
'            SetC fi, 8.5
'            prt.Print m_Format(RsGDd![Peso Unitario], "##,###,##0.0")
            SetC fi, 18.3
            prt.Print m_Format(RsGDd![Peso Unitario], "##,##0.0")
'            SetC fi, 12.3
'            prt.Print m_Format(RsGDd!Cantidad * RsGDd![Peso Unitario], "###,###,##0.0")
'            SetC fi, 15.2
'            prt.Print m_Format(RsGDd![Precio Unitario], "#,###,##0")
'            SetC fi, 17.2
'            prt.Print m_Format(RsGDd!Cantidad * RsGDd![Peso Unitario] * RsGDd![Precio Unitario], "##,###,##0")
            
            RsGDd.MoveNext
            
        Loop
    End If
    
    prt.Font.Size = 12

Case "P"

    'detalle pernos
    'cabecera de detalle
    SetC 10.4, 1.4
    
    prt.Print "CODIGO         DETALLE                           CANT     $ UNI    $ TOTAL"
    
    SetC 10.5, 1.3
'    prt.Print linea
    fi = 10.55
    With RsDoc
    .Seek "=", "GP", Numero, 1
    If Not .EOF Then
        Do While Not .EOF
        
            If !Tipo <> "GP" Or !Numero <> Numero Then Exit Do
            
            'linea
            fi = fi + 0.5
            
            SetC fi, 1.3
            prt.Print m_Format(![codigo producto], "#,###")
            
            RsPrd.Seek "=", ![codigo producto]
            If Not RsPrd.NoMatch Then
                SetC fi, 5
                prt.Print RsPrd![Descripcion]
            End If
            
            SetC fi, 13.5
            prt.Print m_Format(!Cant_Sale, "#,###")
            
            Cantidad_Itemes = Cantidad_Itemes + !Cant_Sale
            
            SetC fi, 15.1
            prt.Print m_Format(![Precio Unitario], "#,###,##0")
'            SetC fi, 18
            SetC fi, 17.5
            prt.Print m_Format(!Cant_Sale * ![Precio Unitario], "##,###,##0")
            
            .MoveNext
            
        Loop
    End If
    End With
    
End Select

If Cantidad_Itemes > 0 Then
    SetC 21.3, 1.8
    prt.Print "CANT TOT : "; m_Format(Cantidad_Itemes, "##0")
End If

SetC 22, 1.8
prt.Print "CHOFER    : "; NoNulo(RsGDc![Observacion 1])

SetC 22.5, 1.8
If RsGDc!Tipo = "P" Then
    prt.Print "ESQUEMA : "; NoNulo(RsGDc![Observacion 3])
Else
    prt.Print "CONTENIDO : "; NoNulo(RsGDc![Observacion 3])
End If

prt.EndDoc

Impresora_Predeterminada "default"

End Sub
Public Sub OC_Prepara(Numero As String, Nv As Double, obra As String, ImpresoraNombre As String)
' prepara archivo de reporte OC para imprimir

Dim DbD As Database, RsProv As Recordset, RsProDir As Recordset, RsProd As Recordset
Dim Dbm As Database, RsOcc As Recordset, RsOCd As Recordset
Dim fi As Double, m_desc As String, Tipo As String, pos As Integer
Dim p_dir As String, p_com As String, p_tel As String, p_fax As String
Dim pDesc_a_Dinero As Double

'////////////////////////////////////////
Dim Dbi As Database, RsOCr As Recordset
Set Dbi = OpenDatabase(repo_file)
Set RsOCr = Dbi.OpenRecordset("OC Legal")
'////////////////////////////////////////

Set Dbm = OpenDatabase(Madq_file)
Set RsOcc = Dbm.OpenRecordset("OC Cabecera")
RsOcc.Index = "Numero"
Set RsOCd = Dbm.OpenRecordset("OC Detalle")
RsOCd.Index = "Numero-Linea"

RsOcc.Seek "=", Numero
If RsOcc.NoMatch Then Exit Sub

ImpresoraNombre = UCase(ImpresoraNombre)
If InStr(1, ImpresoraNombre, "OKI") > 0 Then
    ' graba como impresa si es que es oki
    RsOcc.Edit
    RsOcc!impresa = True
    RsOcc.Update
'    MsgBox "imprimio con oki"
End If


Set DbD = OpenDatabase(data_file)
Set RsProv = DbD.OpenRecordset("Proveedores")
RsProv.Index = "RUT"
Set RsProDir = DbD.OpenRecordset("Proveedores-Direcciones")
RsProDir.Index = "RUT-Codigo"

Set DbD = OpenDatabase(data_file)
Set RsProd = DbD.OpenRecordset("Productos")
RsProd.Index = "Codigo"

Dim can_valor As String, linea As String

With RsOcc

Dbi.Execute "delete * from [OC Legal]"
RsOCr.AddNew
RsOCr!Numero = Numero
RsOCr!Nv = Nv
RsOCr!obra = obra
RsOCr!Emision = !Fecha

If !Tipo = "N" Then
    RsOCr!Cotizacion = !Cotizacion
Else
    RsOCr![Guia de Despacho] = !Cotizacion
End If

' cabecera
p_dir = ""
p_com = ""
p_tel = ""
p_fax = ""
With RsProv
.Seek "=", RsOcc![RUT Proveedor]
If Not .NoMatch Then

    RsOCr!Senores = ![Razon Social]
    RsOCr!rut = !rut
    
    RsProDir.Seek "=", RsOcc![RUT Proveedor], RsOcc![Codigo Direccion]
    If Not RsProDir.NoMatch Then
        p_dir = RsProDir!Direccion
        p_com = RsProDir!Comuna
        p_tel = RsProDir![Telefono 1]
        p_fax = RsProDir!Fax
    Else
        p_dir = !Direccion
        p_com = !Comuna
        p_tel = ![Telefono 1]
        p_fax = !Fax
    End If
    
End If
End With

RsOCr!Direccion = p_dir
RsOCr!Comuna = p_com
RsOCr!telefono = p_tel
RsOCr!Fax = p_fax
'RsOCr!prt.Print Tab(tab0); IIf(!Tipo = "E", "GUÍA DESPACHO Nº: ", "COTIZACIÓN Nº: "); Format(!Cotización, "#")
RsOCr![At Sr] = !atencion
RsOCr!condiciones = ![Condiciones de Pago]
RsOCr![Entregar en] = ![Entregar en]
RsOCr![Fecha Entrega] = ![Fecha a Recibir]
If !fechaModificacion <> "" Then
    RsOCr![fechaModificacion] = "MODIFICADA " & Format(![fechaModificacion], "dd/mm/yyyy")
End If

' detalle
Dim i As Integer, j As Integer
i = 0

Select Case !Tipo
Case "E"

    'detalle especial
    With RsOCd
    .Seek ">=", Numero, 0
    i = 0
    If Not .NoMatch Then
        Do While Not .EOF
            
            If Numero <> !Numero Then Exit Do
            
            'linea
            i = i + 1
            
            If i < !linea Then
'                prt.Print ""
            Else
                RsOCr("Cantidad" & str(i)) = !Cantidad
                RsOCr("Unidad" & str(i)) = !unidad
                RsOCr("Descripcion" & str(i)) = !Descripcion
                RsOCr("Largo" & str(i)) = !largo
                RsOCr("Precio Unitario" & str(i)) = ![Precio Unitario]
                RsOCr("Total" & str(i)) = !Cantidad * ![Precio Unitario]
                .MoveNext

            End If
        Loop
    End If
    End With

Case Else

    ' detalle normal

    With RsOCd
    .Seek ">=", Numero, 0
    i = 0
    If Not .NoMatch Then
        Do While Not .EOF
            If Numero <> !Numero Then Exit Do

            i = i + 1

            ' CÓDIGO
            RsOCr("Codigo Producto" & str(i)) = ![codigo producto]
            RsOCr("Cantidad" & str(i)) = !Cantidad
            RsProd.Seek "=", ![codigo producto]
            If Not RsProd.NoMatch Then
                RsOCr("Unidad" & str(i)) = RsProd![unidad de medida]
                If !largo > 0 Then
'                    prt.Print Tab(tab4); Left(RsProd!Descripción, 16); '15
                Else
'                    prt.Print Tab(tab4); Left(RsProd!Descripción, 27);
                End If
                RsOCr("Descripcion" & str(i)) = RsProd!Descripcion
            End If
            
            If !largo > 0 Then
'                prt.Print Tab(tab5); m_Format(!Largo, "#,###,##0.0");
            End If
                        
            RsOCr("Largo" & str(i)) = !largo
            RsOCr("Precio Unitario" & str(i)) = ![Precio Unitario]
            RsOCr("Total" & str(i)) = !Cantidad * ![Precio Unitario]

            .MoveNext
            
        Loop
        
    End If
    
    End With
    
End Select

RsOCr![Obs 1] = Left(![Observacion 1], 50)
RsOCr![Obs 2] = Left(![Observacion 2], 50)
RsOCr![Obs 3] = Left(![Observacion 3], 50)
RsOCr![Obs 4] = Left(![Observacion 4], 50)
RsOCr!SubTotal = !SubTotal
RsOCr![% Descuento] = ![% Descuento]
RsOCr![Descuento] = Int(!SubTotal * ![% Descuento] / 100 + 0.5)
RsOCr![Otro Descuento] = ![Descuento]
RsOCr!Neto = !Neto
RsOCr!Iva = !Iva
RsOCr!Total = !Total

RsOCr.Update

End With

' cierra recordsets y databases
RsProv.Close
RsProDir.Close
RsProd.Close
DbD.Close

RsOcc.Close
RsOCd.Close
Dbm.Close

End Sub
Public Sub OC_PreparaC2(Numero As String, Nv As Double, obra As String, ImpresoraNombre As String)
' prepara archivo de reporte OC para imprimir

Dim DbD As Database, RsProv As Recordset, RsProDir As Recordset, RsProd As Recordset
Dim Dbm As Database, RsOcc As Recordset, RsOCd As Recordset
Dim fi As Double, m_desc As String, Tipo As String, pos As Integer
Dim p_dir As String, p_com As String, p_tel As String, p_fax As String
Dim pDesc_a_Dinero As Double

'////////////////////////////////////////
Dim Dbi As Database, RsOCr As Recordset, RsOCrD As Recordset
Set Dbi = OpenDatabase(repo_file)
Set RsOCr = Dbi.OpenRecordset("OC Legal")
Set RsOCrD = Dbi.OpenRecordset("OC Legal Detalle")
'////////////////////////////////////////

Set Dbm = OpenDatabase(Madq_file)
Set RsOcc = Dbm.OpenRecordset("OC Cabecera")
RsOcc.Index = "Numero"
Set RsOCd = Dbm.OpenRecordset("OC Detalle")
RsOCd.Index = "Numero-Linea"

RsOcc.Seek "=", Numero
If RsOcc.NoMatch Then Exit Sub

ImpresoraNombre = UCase(ImpresoraNombre)
If InStr(1, ImpresoraNombre, "OKI") > 0 Then
    ' graba como impresa si es que es oki
    RsOcc.Edit
    RsOcc!impresa = True
    RsOcc.Update
'    MsgBox "imprimio con oki"
End If

Set DbD = OpenDatabase(data_file)
Set RsProv = DbD.OpenRecordset("Proveedores")
RsProv.Index = "RUT"
Set RsProDir = DbD.OpenRecordset("Proveedores-Direcciones")
RsProDir.Index = "RUT-Codigo"

Set DbD = OpenDatabase(data_file)
Set RsProd = DbD.OpenRecordset("Productos")
RsProd.Index = "Codigo"

Dim can_valor As String, linea As String

With RsOcc

Dbi.Execute "delete * from [OC Legal]"
Dbi.Execute "delete * from [OC Legal Detalle]"

RsOCr.AddNew
RsOCr!Numero = Numero
RsOCr!Nv = Nv
RsOCr!obra = obra
RsOCr!Emision = !Fecha

If !Tipo = "N" Then
    RsOCr!Cotizacion = !Cotizacion
Else
    RsOCr![Guia de Despacho] = !Cotizacion
End If

' cabecera
p_dir = ""
p_com = ""
p_tel = ""
p_fax = ""
With RsProv
.Seek "=", RsOcc![RUT Proveedor]
If Not .NoMatch Then

    RsOCr!Senores = ![Razon Social]
    RsOCr!rut = !rut
    
    RsProDir.Seek "=", RsOcc![RUT Proveedor], RsOcc![Codigo Direccion]
    If Not RsProDir.NoMatch Then
        p_dir = RsProDir!Direccion
        p_com = RsProDir!Comuna
        p_tel = RsProDir![Telefono 1]
        p_fax = RsProDir!Fax
    Else
        p_dir = !Direccion
        p_com = !Comuna
        p_tel = ![Telefono 1]
        p_fax = !Fax
    End If
    
End If
End With

RsOCr!Direccion = p_dir
RsOCr!Comuna = p_com
RsOCr!telefono = p_tel
RsOCr!Fax = p_fax
'RsOCr!prt.Print Tab(tab0); IIf(!Tipo = "E", "GUÍA DESPACHO Nº: ", "COTIZACIÓN Nº: "); Format(!Cotización, "#")
RsOCr![At Sr] = !atencion
RsOCr!condiciones = ![Condiciones de Pago]
RsOCr![Entregar en] = ![Entregar en]
RsOCr![Fecha Entrega] = ![Fecha a Recibir]

RsOCr![Obs 1] = ![Observacion 1]
RsOCr![Obs 2] = ![Observacion 2]
RsOCr![Obs 3] = ![Observacion 3]
RsOCr![Obs 4] = ![Observacion 4]
RsOCr!SubTotal = !SubTotal
RsOCr![% Descuento] = ![% Descuento]
RsOCr![Descuento] = Int(!SubTotal * ![% Descuento] / 100 + 0.5)
RsOCr![Otro Descuento] = ![Descuento]
RsOCr!Neto = !Neto
RsOCr!Iva = !Iva
RsOCr!Total = !Total
If !fechaModificacion <> "" Then
    RsOCr![fechaModificacion] = "MODIFICADA " & Format(![fechaModificacion], "dd/mm/yyyy")
End If
RsOCr.Update

' detalle
Dim i As Integer, j As Integer
i = 0

Select Case !Tipo
Case "E"

    'detalle especial
    With RsOCd
    .Seek ">=", Numero, 0
    i = 0
    If Not .NoMatch Then
        Do While Not .EOF
            
            If Numero <> !Numero Then Exit Do
            
            'linea
            i = i + 1
            
            If i < !linea Then
'                prt.Print ""
            Else
            
                RsOCrD.AddNew
                RsOCrD("Numero") = !Numero
                RsOCrD("Cantidad") = !Cantidad
                RsOCrD("Unidad") = !unidad
                RsOCrD("Descripcion") = Left(!Descripcion, 50)
                RsOCrD("Largo") = !largo
                RsOCrD("Precio Unitario") = ![Precio Unitario]
                RsOCrD("Total") = !Cantidad * ![Precio Unitario]
                RsOCrD("CuentaContable") = Left(!cuentacontable & " " & cuentaContableBuscarDescripcion(NoNulo(!cuentacontable)), 20)
                RsOCrD("CentroCosto") = Left(!CentroCosto & " " & centroCostoBuscarDescripcion(NoNulo(!CentroCosto)), 20)
                RsOCrD.Update
                
                .MoveNext

            End If
        Loop
    End If
    End With

Case Else

    ' detalle normal

    With RsOCd
    .Seek ">=", Numero, 0
    i = 0
    If Not .NoMatch Then
        Do While Not .EOF
            If Numero <> !Numero Then Exit Do

            i = i + 1

            ' CÓDIGO
            RsOCrD.AddNew
            
            RsOCrD("Numero") = !Numero
            RsOCrD("Codigo Producto") = ![codigo producto]
            RsOCrD("Cantidad") = !Cantidad
            RsProd.Seek "=", ![codigo producto]
            If Not RsProd.NoMatch Then
                RsOCrD("Unidad") = RsProd![unidad de medida]
                If !largo > 0 Then
'                    prt.Print Tab(tab4); Left(RsProd!Descripción, 16); '15
                Else
'                    prt.Print Tab(tab4); Left(RsProd!Descripción, 27);
                End If
                RsOCrD("Descripcion") = Left(RsProd!Descripcion, 50)  '50
            End If
            
            If !largo > 0 Then
'                prt.Print Tab(tab5); m_Format(!Largo, "#,###,##0.0");
            End If

            RsOCrD("Largo") = !largo
            RsOCrD("Precio Unitario") = ![Precio Unitario]
            RsOCrD("Total") = !Cantidad * ![Precio Unitario]
            RsOCrD("CuentaContable") = Left(!cuentacontable & " " & cuentaContableBuscarDescripcion(!cuentacontable), 20)  '20
            RsOCrD("CentroCosto") = Left(![CentroCosto] & " " & centroCostoBuscarDescripcion(![CentroCosto]), 20)  '20
            
            RsOCrD.Update

            .MoveNext
            
            
        Loop
        
    End If
    
    End With
    
End Select

End With

' cierra recordsets y databases
RsProv.Close
RsProDir.Close
RsProd.Close
DbD.Close

RsOcc.Close
RsOCd.Close
Dbm.Close

End Sub
Public Sub GD_Prepara(Numero As String, Nv As Double, obra As String) ', ImpresoraNombre As String)

' prepara archivo de reporte GD para imprimir (packing list)

Dim DbD As Database, RsCli As Recordset, RsProd As Recordset
Dim Dbm As Database, RsGDc As Recordset, RsGDd As Recordset, RsGdE As Recordset, RsPd As Recordset
Dim fi As Double, m_desc As String, Tipo As String, pos As Integer
Dim p_dir As String, p_com As String, p_tel As String, p_fax As String
Dim pDesc_a_Dinero As Double

'////////////////////////////////////////
Dim Dbi As Database, RsOCr As Recordset
Set Dbi = OpenDatabase(repo_file)
Set RsOCr = Dbi.OpenRecordset("GD packinglist")
'////////////////////////////////////////

Dim DbAdq As Database, RsDoc As Recordset
Set DbAdq = OpenDatabase(Madq_file)
Set RsDoc = DbAdq.OpenRecordset("Documentos")
RsDoc.Index = "Tipo-Numero-Linea"

Set Dbm = OpenDatabase(mpro_file)

Set RsPd = Dbm.OpenRecordset("planos detalle")
RsPd.Index = "nv-plano-marca"

Set RsGDc = Dbm.OpenRecordset("GD Cabecera")
RsGDc.Index = "Numero"
Set RsGDd = Dbm.OpenRecordset("GD Detalle")
RsGDd.Index = "Numero-Linea"
Set RsGdE = Dbm.OpenRecordset("GD especial detalle")
RsGdE.Index = "Numero-Linea"

RsGDc.Seek "=", Numero
If RsGDc.NoMatch Then Exit Sub

'ImpresoraNombre = UCase(ImpresoraNombre)
'If InStr(1, ImpresoraNombre, "OKI") > 0 Then
'    ' graba como impresa si es que es oki
'    RsOcc.Edit
'    RsOcc!impresa = True
'    RsOcc.Update
''    MsgBox "imprimio con oki"
'End If

Set DbD = OpenDatabase(data_file)

Set RsCli = DbD.OpenRecordset("Clientes")
RsCli.Index = "RUT"

Set RsProd = DbD.OpenRecordset("Productos")
RsProd.Index = "codigo"

Dim can_valor As String, linea As String

With RsGDc

Dbi.Execute "delete * from [GD packinglist]"
RsOCr.AddNew
RsOCr!Numero = Numero
RsOCr!Nv = Nv
RsOCr!obra = obra
RsOCr!Emision = !Fecha

' cabecera
p_dir = ""
p_com = ""
p_tel = ""
p_fax = ""
With RsCli
.Seek "=", RsGDc![RUT CLiente]
If Not .NoMatch Then

    RsOCr!Senores = ![Razon Social]
    RsOCr!rut = !rut
    
    p_dir = !Direccion
    p_com = !Comuna
    p_tel = ![Telefono 1]
'    p_fax = !Fax
    
End If
End With

RsOCr!Direccion = p_dir
RsOCr!Comuna = p_com
RsOCr!telefono = p_tel
RsOCr!Fax = p_fax
'RsOCr!prt.Print Tab(tab0); IIf(!Tipo = "E", "GUÍA DESPACHO Nº: ", "COTIZACIÓN Nº: "); Format(!Cotización, "#")
'RsOCr![At Sr] = !atencion
'RsOCr!condiciones = ![Condiciones de Pago]
'RsOCr![Entregar en] = ![Entregar en]
'RsOCr![Fecha Entrega] = ![Fecha a Recibir]

' detalle
Dim i As Integer, j As Integer
i = 0

Select Case !Tipo
Case "E"

    'detalle especial
    With RsGdE
    .Seek ">=", Numero, 0
    i = 0
    If Not .NoMatch Then
        Do While Not .EOF
            If Numero <> !Numero Then Exit Do
            
            'linea
            i = i + 1
            If i < !linea Then
'                prt.Print ""
            Else
                RsOCr("Cantidad" & str(i)) = !Cantidad
                RsOCr("Unidad" & str(i)) = !unidad
                RsOCr("Descripcion" & str(i)) = !Detalle
                RsOCr("Peso Unitario" & str(i)) = ![Peso Unitario]
                RsOCr("Peso total" & str(i)) = !Cantidad * ![Peso Unitario]
                RsOCr("precio unitario" & str(i)) = ![Precio Unitario]
                RsOCr("precio Total" & str(i)) = !Cantidad * ![Precio Unitario]
                .MoveNext

            End If
        Loop
    End If
    End With

Case "N", "G"

    ' detalle normal

    With RsGDd
    .Seek ">=", Numero, 0
    i = 0
    If Not .NoMatch Then
    
        Do While Not .EOF
        
            If Numero <> !Numero Then Exit Do
            
            i = i + 1
            
            RsOCr("plano" & str(i)) = ![Plano]
            RsOCr("marca" & str(i)) = ![Marca]
            
            RsOCr("Cantidad" & str(i)) = !Cantidad
            
            RsPd.Seek "=", Nv, 0, ![Plano], ![Marca]
            If Not RsPd.NoMatch Then
                RsOCr("Descripcion" & str(i)) = RsPd!Descripcion
            End If

            RsOCr("Peso Unitario" & str(i)) = ![Peso Unitario]
            RsOCr("Peso Total" & str(i)) = !Cantidad * ![Peso Unitario]
            
            RsOCr("Precio Unitario" & str(i)) = ![Precio Unitario]
            RsOCr("Precio Total" & str(i)) = !Cantidad * ![Precio Unitario]

            .MoveNext
            
        Loop
        
    End If
    
    End With
    
Case "P"

    'detalle pernos
    'cabecera de detalle
    With RsDoc
    .Seek "=", "GP", Numero, 1
    If Not .EOF Then
        Do While Not .EOF
        
            If !Tipo <> "GP" Or !Numero <> Numero Then Exit Do
            
            i = i + 1
            
            If i > 17 Then Exit Do
            
            RsOCr("plano" & str(i)) = ![codigo producto]
'            RsOCr("marca" & Str(i)) = ![Marca]
            
            RsOCr("Cantidad" & str(i)) = !Cant_Sale
            
            RsProd.Seek "=", ![codigo producto]
            If Not RsProd.NoMatch Then
                RsOCr("Descripcion" & str(i)) = Left(RsProd!Descripcion, 50)
            End If

'            RsOCr("Peso Unitario" & Str(i)) = ![Peso Unitario]
'            RsOCr("Peso Total" & Str(i)) = !Cantidad * ![Peso Unitario]
            
            RsOCr("Precio Unitario" & str(i)) = ![Precio Unitario]
            RsOCr("Precio Total" & str(i)) = !Cant_Sale * ![Precio Unitario]
                        
            .MoveNext
            
        Loop
    End If
    End With
   
End Select

RsOCr![Obs 1] = ![Observacion 1]
RsOCr![Obs 2] = ![Observacion 2]
RsOCr![Obs 3] = ![Observacion 3]
RsOCr![Obs 4] = ![Observacion 4]

RsOCr![Peso Total] = ![Peso Total]

'RsOCr!SubTotal = !SubTotal
'RsOCr![% Descuento] = ![% Descuento]
'RsOCr![Descuento] = Int(!SubTotal * ![% Descuento] / 100 + 0.5)
'RsOCr![Otro Descuento] = ![Descuento]
'RsOCr!Neto = !Neto
'RsOCr!Iva = !Iva
'RsOCr!Total = !Total

RsOCr.Update

End With

' cierra recordsets y databases
RsCli.Close
DbD.Close

RsGDc.Close
RsGDd.Close
Dbm.Close

End Sub
Public Sub bulto_Prepara(Numero As String, Nv As Double, obra As String, ImpresoraNombre As String)
' prepara archivo de reporte BULTO para imprimir

Dim DbD As Database, RsCl As Recordset
Dim Dbm As Database, RsB As Recordset, RsPd As Recordset
Dim fi As Double, m_desc As String, Tipo As String, pos As Integer
Dim p_dir As String, p_com As String, p_tel As String, p_fax As String
Dim pDesc_a_Dinero As Double

'////////////////////////////////////////
Dim Dbi As Database, RsOCr As Recordset
Set Dbi = OpenDatabase(repo_file)
Set RsOCr = Dbi.OpenRecordset("bulto")
'////////////////////////////////////////

Set Dbm = OpenDatabase(mpro_file)
Set RsPd = Dbm.OpenRecordset("Planos Detalle")
RsPd.Index = "NV-Plano-Marca"
Set RsB = Dbm.OpenRecordset("bultos")
RsB.Index = "Numero-Linea"

RsB.Seek "=", Numero, 1
If RsB.NoMatch Then Exit Sub

'ImpresoraNombre = UCase(ImpresoraNombre)

Set DbD = OpenDatabase(data_file)
Set RsCl = DbD.OpenRecordset("clientes")
RsCl.Index = "rut"

Dim can_valor As String, linea As String

With RsB

Dbi.Execute "delete * from [bulto]"

RsOCr.AddNew
RsOCr!Numero = Numero
RsOCr!Nv = Nv
RsOCr!obra = obra
RsOCr!Emision = !Fecha

With RsCl
.Seek "=", RsB![RUT CLiente]
If Not .NoMatch Then
    RsOCr!Senores = ![Razon Social]
End If
End With

' detalle
Dim i As Integer, j As Integer
i = 0

With RsB
.Seek ">=", Numero
i = 0
If Not .NoMatch Then
    Do While Not .EOF
        If Numero <> !Numero Then Exit Do

        i = i + 1
        ' CÓDIGO
        RsOCr("plano" & str(i)) = !Plano ' ![codigo producto]
        RsOCr("rev" & str(i)) = ![Rev]
        RsOCr("marca" & str(i)) = ![Marca]
        RsOCr("cantidad" & str(i)) = !Cantidad
        
        RsPd.Seek "=", Nv, 0, !Plano, !Marca
        If Not RsPd.NoMatch Then
        '    RsOCr("Unidad" & str(i)) = RsProd![unidad de medida]
            RsOCr("descripcion" & str(i)) = RsPd!Descripcion
        End If

        'RsOCr("Largo" & str(i)) = !largo
        RsOCr("pesounitario" & str(i)) = ![PesoUnitario]
        RsOCr("pesototal" & str(i)) = !Cantidad * ![PesoUnitario]

        .MoveNext
        
    Loop
    
End If
End With

RsOCr.Update

End With

' cierra recordsets y databases
RsCl.Close
DbD.Close

RsPd.Close
RsB.Close
Dbm.Close

End Sub
Public Sub protocoloPreparar(Numero As Double, nvObra As String) ', obra As String, ImpresoraNombre As String)
' prepara archivo para impresion de protocolo pintura
' 29/05/2013

Dim DbD As Database, RsCl As Recordset
Dim Dbm As Database, RsNv As Recordset, RsPd As Recordset, RsITOd As Recordset, RsPp As Recordset
Dim i As Integer, Nv As Double, obra As String, rutCliente As String
Nv = CDbl(Left(nvObra, 4))
obra = Trim(Mid(nvObra, 5))
'////////////////////////////////////////
Dim Dbi As Database, RsRpp As Recordset
Set Dbi = OpenDatabase(repo_file)
Set RsRpp = Dbi.OpenRecordset("protocoloPintura")
'////////////////////////////////////////

Set Dbm = OpenDatabase(mpro_file)
Set RsNv = Dbm.OpenRecordset("nv cabecera")
RsNv.Index = "numero"

Set RsPd = Dbm.OpenRecordset("planos detalle")
RsPd.Index = "nv-plano-marca"

Set RsPp = Dbm.OpenRecordset("protocoloPintura")
RsPp.Index = "numero"

Set RsITOd = Dbm.OpenRecordset("ITO PG Detalle")
RsITOd.Index = "tipo-numero-linea"

RsPp.Seek "=", Numero
If RsPp.NoMatch Then Exit Sub

Set DbD = OpenDatabase(data_file)
Set RsCl = DbD.OpenRecordset("clientes")
RsCl.Index = "RUT"

Dbi.Execute "delete * from [protocoloPintura]"

With RsRpp

' ya esta parado en el registro pp

' busca cliente
RsNv.Seek "=", Nv, 0
rutCliente = ""
If Not RsNv.NoMatch Then
    rutCliente = RsNv![RUT CLiente]
End If

.AddNew
!Numero = Numero
!numeroProtocolo = RsPp!numeroProtocolo
!pagina = RsPp!paginaNumero & " de " & RsPp!paginaTotal
!Nv = Nv
!Fecha = RsPp!Fecha
!Responsable = trabajadorBuscarNombreCompleto(NoNulo(RsPp!Responsable))
!Proyecto = obra
!Cliente = clienteBuscarRazon(rutCliente)
!esquema = RsPp!esquema
!probetaprevia = IIf(RsPp!probetaprevia, "SI", "NO")
!granallamezclada = IIf(RsPp!granallamezclada, "SI", "NO")
!calibre = RsPp!calibretipo

' detalle
i = 0

RsITOd.Seek ">=", "P", Numero, 0
i = 0
If Not RsITOd.NoMatch Then
    
    Do While Not RsITOd.EOF
        
        If Numero <> RsITOd!Numero Then Exit Do
        
        'i = i + 1
        i = RsITOd!linea

        ' CÓDIGO
        'RsRpp("cantidad" & str(i)) = RsItoD!Cantidad
        RsRpp("cantidad" & str(i)) = RsITOd!Cantidad
        RsRpp("marca" & str(i)) = RsITOd!Marca
        RsPd.Seek "=", Nv, 0, RsITOd!Plano, RsITOd!Marca
        If Not RsPd.NoMatch Then
            RsRpp("descripcion" & str(i)) = RsPd!Descripcion
        End If

        RsITOd.MoveNext
        
    Loop
    
End If
    
.Update

End With

' cierra recordsets y databases
DbD.Close
Dbm.Close

End Sub
Private Sub SetC(Fila As Double, Columna As Double)
prt.CurrentY = AjusteY + Fila
prt.CurrentX = AjusteX + Columna
End Sub
Private Function corte(txt As String, pos As Integer) As Integer
' encuentra el espacio en string el "txt",
' buscando de atras para adelante comenzando por "pos"
Dim c As Integer
corte = pos
If Len(txt) <= pos Then Exit Function

For c = pos + 1 To 1 Step -1
    If Mid(txt, c, 1) = " " Then
        corte = c
        Exit For
    End If
Next

End Function
