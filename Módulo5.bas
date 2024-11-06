Attribute VB_Name = "Módulo5"
Sub extraer_datos_factura()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Sheets("BASE").Activate
    
    Call buscar_oc
    If MsgBox("¿Sigo con las GD?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

    Call buscar_gd
    If MsgBox("¿Sigo con las Fact?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

    Call extraer_datos
    Call buscar_monto_factura
    If MsgBox("¿Sigo con las NC?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

    Call buscar_nc
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub buscar_oc()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim nro As Long
    nro = Application.CountA(Range("G:G")) - 1
    Dim i As Long
    
    For i = 1 To nro
        Application.StatusBar = "Procesando OC: " & Format((i / nro), "0.0%")
        With Range("Y2").Offset(i - 1, 0)
            Dim referencia As String
            referencia = .Offset(0, -18).Value
            Dim posicion As Long
            posicion = InStr(1, referencia, "Tipo:801")
            If posicion > 0 Then
                .Value = Mid(referencia, posicion + 15, 10)
            End If
        End With
    Next i
    
    Call buscar_oc2
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub buscar_oc2()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim largo As Long
    largo = Application.CountA(Range("A:A"))
    Range("H2:H" & largo).formulaR1C1 = "=SI(RC[18]>0,RC[18],RC[17])"
    Range("H:H").Calculate
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub buscar_gd()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim nro As Long
    nro = Application.CountA(Range("G:G")) - 1
    Dim i As Long
    
    For i = 1 To nro
        Application.StatusBar = "Procesando GD: " & Format((i / nro), "0.0%")
        With Range("I2").Offset(i - 1, 0)
            Dim referencia As String
            referencia = .Offset(0, -2).Value
            Dim op_1 As Long, op_2 As Long, gd_1 As String
            op_1 = InStr(1, referencia, "Tipo:52") + 14
            If op_1 > 14 Then
                op_2 = InStr(op_1, referencia, ",")
                gd_1 = Mid(referencia, op_1, op_2 - op_1)
                .Value = gd_1
            Else
                op_1 = InStr(1, referencia, "Tipo:052") + 15
                If op_1 > 15 Then
                    op_2 = InStr(op_1, referencia, ",")
                    gd_1 = Mid(referencia, op_1, op_2 - op_1)
                    .Value = gd_1
                Else
                    .Value = "Sin GD"
                    Call buscar_gd2(i - 1)
                End If
            End If
        End With
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub buscar_gd2(indiceFila As Long)
    Dim rut As String
    With Range("I2").Offset(indiceFila, 0)
        rut = .Offset(0, -7).Value
        If rut = "94668000-1" Or rut = "76066726-9" Or rut = "96803460-K" Then
            If .Value = "Sin GD" Then
                .Value = .Offset(0, -8).Value
            End If
        End If
    End With
End Sub

Sub extraer_datos()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim nro As Long
    nro = Application.CountA(Range("G:G")) - 1
    Dim i As Long
    
    For i = 1 To nro
        Application.StatusBar = "Procesando RUT: " & Format((i / nro), "0.0%")
        With Range("AA2").Offset(i - 1, 0)
            Dim referencia As String
            referencia = .Offset(0, -25).Value
            Dim op_2 As Long
            op_2 = InStr(1, referencia, "-")
            If op_2 > 0 Then
                .Value = Left(referencia, op_2 - 1)
            End If
        End With
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub buscar_monto_factura()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim nro As Long
    nro = Application.CountA(Range("G:G")) - 1
    Dim i As Long
    
    For i = 1 To nro
        Application.StatusBar = "Procesando Monto de Factura: " & Format((i / nro), "0.0%")
        With Range("N2").Offset(i - 1, 0)
            .Value = .Offset(0, -9).Value
        End With
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub buscar_nc()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim nro As Long
    nro = Application.CountA(Range("G:G")) - 1
    Dim i As Long
    
    For i = 1 To nro
        Application.StatusBar = "Procesando NC: " & Format((i / nro), "0.0%")
        With Range("O2").Offset(i - 1, 0)
            Dim rut As String
            rut = .Offset(0, -13).Value
            Dim folio_nc As String
            folio_nc = .Offset(0, -14).Value
            Dim rango_nc_monto As Range, rango_nc_folio As Range, rango_nc_rut As Range
            Set rango_nc_monto = Sheets("NC").Range("L:L")
            Set rango_nc_folio = Sheets("NC").Range("B:B")
            Set rango_nc_rut = Sheets("NC").Range("D:D")
            .Value = Application.SumIfs(rango_nc_monto, rango_nc_folio, folio_nc, rango_nc_rut, rut)
        End With
    Next i
    
    If MsgBox("¿Saco los folios de las NC?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    Call buscar_nc_folios
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub buscar_nc_folios()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim nro As Long
    nro = Application.CountA(Range("N:N")) - 1
    Dim i As Long

    For i = 1 To nro
        Application.StatusBar = "Procesando Folios NC: " & Format((i / nro), "0.0%")
        With Range("O2").Offset(i - 1, 0)
            If IsNumeric(.Value) And .Value > 1 Then
                Dim folio As String
                folio = .Offset(0, -14).Value
                With Sheets("NC")
                    If .FilterMode Then .ShowAllData
                    .Range("$B:$S").AutoFilter Field:=1, Criteria1:=folio
                    Dim rango_copia As Range
                    Set rango_copia = .Range("c2:c" & .Cells(.Rows.Count, "E").End(xlUp).Row)
                    If Not rango_copia Is Nothing Then
                        rango_copia.Copy
                        Sheets("BASE").Activate
                        Range("O2").Offset(i - 1, 17).PasteSpecial Paste:=xlPasteAll, Transpose:=True
                    End If
                    .ShowAllData
                End With
            End If
        End With
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub buscar_fecha_pago()
    Dim hojaBase As Worksheet
    Dim hojaPF0 As Worksheet
    Dim rango_cc As Range
    Dim nro As Long
    Dim nr As Long
    Dim porcentaje As String
    Dim rc As Variant
    Dim folio As Variant
    Dim rut As String
    Dim folio_rut As String
    Dim fecha As Variant
    Dim celda As Range

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set hojaBase = Sheets("BASE")
    Set hojaPF0 = Sheets("PF0")
    nro = hojaBase.Cells(Rows.Count, 15).End(xlUp).Row
    Set rango_cc = hojaPF0.Range("c2:ai2000")

    For nr = 2 To nro
        Set celda = hojaBase.Cells(nr, 16)
        rut = hojaBase.Cells(nr, 1)
        folio = hojaBase.Cells(nr, 2)

        If Len(rut) * Len(folio) > 0 Then
            folio_rut = rut & folio

            On Error Resume Next
            rc = Application.Match(folio_rut, rango_cc.Columns(34), 0)
            On Error GoTo 0

            If IsNumeric(rc) Then
                fecha = rango_cc.Cells(rc, 23)
            Else
                fecha = Application.VLookup(folio, rango_cc, 23, False)
            End If

            If Not IsError(fecha) Then
                celda.Value = fecha
            End If
        End If

        porcentaje = Format((nr - 1) / nro, "0.0%")
        Application.StatusBar = "Cargando fecha de pago: " & porcentaje
    Next nr

    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

Sub buscar_Em()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Ejecutar las macros find_em y find_valorem
    Call find_em
    Call find_valorem

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
