Attribute VB_Name = "Módulo9"
Sub find_em()
    Dim ran_em As Range
    Dim lcom As String
    Dim em As Variant
    Dim oc As String
    Dim gd As String
    Dim fecha_pago As Variant
    Dim ws As Worksheet
    Dim wsEM As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Definir las hojas de trabajo
    Set ws = ActiveSheet
    Set wsEM = Sheets("EM")

    ' Obtener la última fila con datos en la columna G para determinar el límite del bucle
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

    ' Bucle a través de cada fila comenzando desde la fila 2 hasta la última fila con datos en la columna G
    For i = 2 To lastRow
        em = "Sin Dato" ' Inicializar la variable em a "Sin Dato"
        
        ' Obtener los valores necesarios de las celdas
        fecha_pago = ws.Cells(i, "V").Value ' Columna V (12 columnas a la derecha de J)
        oc = ws.Cells(i, "H").Value ' Columna H (2 columnas a la izquierda de J)
        lcom = ws.Cells(i, "K").Value ' Columna K (1 columna a la derecha de J)
        gd = ws.Cells(i, "I").Value ' Columna I (1 columna a la izquierda de J)
        
        ' Verificar el valor de lcom y establecer el rango y los parámetros de VLookup apropiados
        If lcom = "E599" Then
            Set ran_em = wsEM.Range("B:F")
            em = Application.VLookup(oc, ran_em, 5, False)
        Else
            Set ran_em = wsEM.Range("A:F")
            em = Application.VLookup(oc & gd, ran_em, 6, False)
        End If
        
        ' Manejar errores de VLookup
        If IsError(em) Then
            em = "Sin Dato"
        End If
        
        ' Actualizar la celda activa con el valor encontrado o "Sin Dato"
        ws.Cells(i, "J").Value = em ' Columna J (columna actual)
        
        ' Actualizar el indicador de progreso en la barra de estado
        Application.StatusBar = "Procesando fila " & i - 1 & " de " & lastRow - 1 & " (" & Format((i - 1) / (lastRow - 1), "0%") & " completado)"
    Next i

    ' Restablecer la barra de estado
    Application.StatusBar = False
End Sub




Sub find_valorem()
    Dim wsBase As Worksheet
    Dim wsDinamica As Worksheet
    Dim ran_em As Range
    Dim oc As Variant
    Dim em_1 As Variant
    Dim em_2 As Variant
    Dim em_3 As Variant
    Dim tt As Variant
    Dim cell As Range
    Dim nro As Long
    Dim nr As Long

    ' Define worksheets
    Set wsBase = Sheets("BASE")
    Set wsDinamica = Sheets("Dinamica")

    ' Determine the last row with data in column G
    nro = Application.CountA(wsBase.Range("G:G")) - 1

    ' Define the range for VLOOKUP
    Set ran_em = wsDinamica.Range("A:C")

    ' Disable screen updating and automatic calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Iterate through each row starting from row 2
    For nr = 2 To nro + 1
        Set cell = wsBase.Cells(nr, "M")

        If IsEmpty(cell.Offset(0, -6).Value) Then Exit For

        oc = cell.Offset(0, -5).Value

        em_1 = Application.VLookup(oc, ran_em, 2, False)
        If IsError(em_1) Then em_1 = 0

        em_2 = Application.VLookup(oc, ran_em, 3, False)
        If IsError(em_2) Then em_2 = 0

        tt = em_1 - em_2
        em_3 = tt * 1.19

        cell.Value = em_3

        ' Update progress
        Application.StatusBar = "Processing row " & nr & " of " & nro & " (" & Format(nr / nro * 100, "0.0") & "%)"
    Next nr

    ' Restore screen updating and automatic calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False

    ' Clear variables
    Set ran_em = Nothing
    Set wsBase = Nothing
    Set wsDinamica = Nothing
End Sub

Sub find_fecha_pago()
    Dim wsBase As Worksheet
    Dim wsPF0 As Worksheet
    Dim ran_cc As Range
    Dim nro As Long
    Dim nr As Long
    Dim porcen As String
    Dim rc As Variant
    Dim folio As Variant
    Dim rut As String
    Dim folio_rut As String
    Dim fecha As Variant
    Dim cell As Range

    ' Define worksheets
    Set wsBase = Sheets("BASE")
    Set wsPF0 = Sheets("PF0")

    ' Determine the last row with data in column G
    nro = Application.CountA(wsBase.Range("G:G")) - 1

    ' Define the range for VLOOKUP
    Set ran_cc = wsPF0.Range("A:W")

    ' Disable screen updating and automatic calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Iterate through each row starting from row 2
    For nr = 2 To nro + 1
        Set cell = wsBase.Cells(nr, "W")

        rc = cell.Offset(0, -7).Value

        If rc < 0 Then
            folio = cell.Offset(0, -22).Value
            rut = cell.Offset(0, -21).Value
            folio_rut = folio & rut

            fecha = Application.VLookup(folio_rut, ran_cc, 23, False)
            If IsError(fecha) Or IsEmpty(fecha) Then
                fecha = Application.VLookup(folio_rut, ran_cc, 17, False)
                If IsError(fecha) Or IsEmpty(fecha) Then
                    fecha = Application.VLookup(folio_rut, ran_cc, 16, False)
                End If
            End If

            cell.Value = IIf(IsError(fecha) Or IsEmpty(fecha), "No Hay Pago", fecha)
        Else
            cell.Value = "No Hay Pago"
        End If

        ' Update the status bar every 10 rows
        If nr Mod 10 = 0 Then
            porcen = Format((nr / nro), "0.0%")
            Application.StatusBar = "Va en un " & porcen & "% del cálculo de la fecha pago"
        End If
    Next nr

    ' Restore screen updating and automatic calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False

    ' Clear variables
    Set ran_cc = Nothing
    Set wsBase = Nothing
    Set wsPF0 = Nothing
End Sub

