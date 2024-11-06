Attribute VB_Name = "Módulo8"
Sub con_formula()
    Dim largo As Long
    Dim ws As Worksheet

    Set ws = Sheets("BASE")
    largo = Application.CountA(ws.Range("A:A"))

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    With ws
        ' Columna P
        .Range("P2").formulaR1C1 = "=SUMIFS(PF0!C[2],PF0!C[-9],RC[-14],PF0!C[-5],RC[-15],PF0!C[-7],""RC"")"
        .Range("P2:P" & largo).FillDown

        ' Columna Q
        .Range("Q2").formulaR1C1 = "=SUMIFS(PF0!C[1],PF0!C[-10],RC[-15],PF0!C[-6],RC[-16],PF0!C[-8],""KE"")"
        .Range("Q2:Q" & largo).FillDown

        ' Columna R
        .Range("R2").formulaR1C1 = "=SUMIFS(PF0!C, PF0!C[-11], RC[-16], PF0!C[-9], ""CE"", PF0!C[-6], BASE!RC[-17]) + " & _
                                   "SUMIFS(PF0!C18, PF0!C7, RC2, PF0!C9, ""CE"", PF0!C11, RC[14]) + " & _
                                   "SUMIFS(PF0!C18, PF0!C7, RC2, PF0!C9, ""CE"", PF0!C11, RC[15]) + " & _
                                   "SUMIFS(PF0!C18, PF0!C7, RC2, PF0!C9, ""CE"", PF0!C11, RC[16]) + " & _
                                   "SUMIFS(PF0!C18, PF0!C7, RC2, PF0!C9, ""CE"", PF0!C11, RC[17])"
        .Range("R2:R" & largo).FillDown

        ' Columna S
        .Range("S2").formulaR1C1 = "=SUMIFS(PF0!C[-1], PF0!C[-12], RC[-17], PF0!C[-8], RC[-18], PF0!C[-10], ""FQ"")"
        .Range("S2:S" & largo).FillDown

        ' Columna T
        .Range("T2").formulaR1C1 = "=SUMIFS(PF0!C[-2], PF0!C[-9], RC[-19], PF0!C[-13], RC[-18], PF0!C[-11], ""ZK"")"
        .Range("T2:T" & largo).FillDown

        .Range("P2:T" & largo).Calculate
    End With

    Call suma_dif

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub suma_dif()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim nr As Long
    Dim porcen As String
    Dim Fact As Double
    Dim nc As Double
    Dim rc As Double
    Dim ke As Double
    Dim ce1 As Double
    Dim fq As Double
    Dim zk As Double
    Dim dif As Double
    Dim cell As Range

    Set ws = ThisWorkbook.Sheets("BASE")
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For nr = 2 To lastRow
        Set cell = ws.Cells(nr, "U")

        Fact = cell.Offset(0, -7).Value
        nc = cell.Offset(0, -6).Value
        rc = cell.Offset(0, -5).Value
        ke = cell.Offset(0, -4).Value
        ce1 = cell.Offset(0, -3).Value
        fq = cell.Offset(0, -2).Value
        zk = cell.Offset(0, -1).Value

        dif = (Fact - nc) + (rc + (ke + ce1 + zk) + fq)

        If (Fact - nc) < 200 And (Fact - nc) > -200 Then
            cell.Value = "Fact-NC"
        ElseIf dif < 200 And dif > -200 Then
            cell.Value = "Fact Pagada"
        Else
            cell.Value = dif
        End If

        If nr Mod 10 = 0 Then
            porcen = Format((nr / lastRow) * 100, "0.0%")
            Application.StatusBar = "Va en un " & porcen & "% del cálculo de la suma"
        End If
    Next nr

    ws.Columns("U:U").Style = "Currency [0]"

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub




