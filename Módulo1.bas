Attribute VB_Name = "Módulo1"
Sub CopiarFilasConValor33()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim filaDestino As Long
    
    ' Cambia "HojaOrigen" y "HojaDestino" por los nombres reales de tus hojas
    Set wsOrigen = ThisWorkbook.Sheets("PPL")
    Set wsDestino = ThisWorkbook.Sheets("BASE")

    ' Encuentra la última fila con datos en la hoja origen
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, "A").End(xlUp).Row

    filaDestino = 1 ' Inicia en la primera fila de la hoja destino

    ' Copia el encabezado
    wsDestino.Range("A1").Value = wsOrigen.Range("C1").Value
    wsDestino.Range("B1").Value = wsOrigen.Range("D1").Value
    wsDestino.Range("C1").Value = wsOrigen.Range("E1").Value
    wsDestino.Range("D1").Value = wsOrigen.Range("H1").Value
    wsDestino.Range("E1").Value = wsOrigen.Range("L1").Value
    wsDestino.Range("F1").Value = wsOrigen.Range("P1").Value
    wsDestino.Range("G1").Value = wsOrigen.Range("Q1").Value

    ' Recorre las filas de la hoja origen
    For i = 2 To ultimaFila
        If wsOrigen.Cells(i, 1).Value = 33 Then
            filaDestino = filaDestino + 1
            wsDestino.Cells(filaDestino, 1).Value = wsOrigen.Cells(i, 3).Value
            wsDestino.Cells(filaDestino, 2).Value = wsOrigen.Cells(i, 4).Value
            wsDestino.Cells(filaDestino, 3).Value = wsOrigen.Cells(i, 5).Value
            wsDestino.Cells(filaDestino, 4).Value = wsOrigen.Cells(i, 8).Value
            wsDestino.Cells(filaDestino, 5).Value = wsOrigen.Cells(i, 12).Value
            wsDestino.Cells(filaDestino, 6).Value = wsOrigen.Cells(i, 16).Value
            wsDestino.Cells(filaDestino, 7).Value = wsOrigen.Cells(i, 17).Value
        End If
    Next i

    MsgBox "Columnas copiadas exitosamente."
End Sub
Sub CopiarFilasConValor61()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim filaDestino As Long
    Dim ultimaColumna As Long
    Dim filaInicialDestino As Long
    
    ' Cambia "PPL" por el nombre real de tu hoja origen
    Set wsOrigen = ThisWorkbook.Sheets("PPL")
    ' Cambia "NC" por el nombre real de tu hoja destino
    Set wsDestino = ThisWorkbook.Sheets("NC")

    ' Encuentra la última fila con datos en la hoja origen
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, "A").End(xlUp).Row

    filaInicialDestino = 2 ' Comienza desde la segunda fila de la hoja destino
    filaDestino = filaInicialDestino ' Inicializa la fila de destino

    ' Recorre las filas de la hoja origen
    For i = 1 To ultimaFila
        If wsOrigen.Cells(i, 1).Value = 61 Then
            ' Encuentra la última columna con datos en la fila actual de la hoja origen
            ultimaColumna = wsOrigen.Cells(i, wsOrigen.Columns.Count).End(xlToLeft).Column
            ' Copia desde la columna C hasta la última columna con datos en la fila actual
            wsOrigen.Range(wsOrigen.Cells(i, 3), wsOrigen.Cells(i, ultimaColumna)).Copy
            ' Pega en la hoja destino empezando desde la columna C y la fila de destino actual
            wsDestino.Cells(filaDestino, 3).PasteSpecial Paste:=xlPasteAll
            filaDestino = filaDestino + 1 ' Avanza a la siguiente fila de destino
        End If
    Next i

    Application.CutCopyMode = False ' Limpiar el portapapeles
    MsgBox "Filas copiadas exitosamente."
End Sub

Sub CopiarFormula()
    Dim baseWs As Worksheet
    Dim configWs As Worksheet
    Dim lastRow As Long
    Dim formulaR1C1 As String
    Dim filePath As String
    Dim fileDialog As fileDialog
    Dim i As Long

    ' Desactivar actualización de pantalla y alertas para mejorar rendimiento
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Definir la hoja de trabajo "BASE"
    Set baseWs = ThisWorkbook.Sheets("BASE")
    
    ' Verificar si la hoja Configuración existe, si no, crearla
    On Error Resume Next
    Set configWs = ThisWorkbook.Sheets("Configuración")
    On Error GoTo 0
    If configWs Is Nothing Then
        Set configWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        configWs.Name = "Configuración"
    End If

    ' Obtener la última fila de la columna G en la hoja "BASE"
    lastRow = baseWs.Cells(baseWs.Rows.Count, "G").End(xlUp).Row

    ' Definir la fórmula
    formulaR1C1 = "=IF(COUNTIFS('[DTE Reclamados Año 2019-C004.xlsx]Reclamos'!C4, R2C27, '[DTE Reclamados Año 2019-C004.xlsx]Reclamos'!C3, R21) > 0, ""FACT RECLAMADA"", ""Sin Reclamo"")"

    ' Obtener la última ubicación del archivo desde la hoja Configuración
    filePath = configWs.Cells(1, "A").Value ' Asumiendo que la ruta se almacena en la celda A1 de la hoja Configuración
    If Dir(filePath) = "" Then
        ' Si el archivo no se encuentra, solicitar la ruta al usuario
        Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
        With fileDialog
            .Title = "Selecciona el archivo DTE Reclamados Año 2019-C004.xlsx"
            .Filters.Add "Archivos de Excel", "*.xlsx"
            If .Show = -1 Then
                filePath = .SelectedItems(1)
                ' Almacenar la última ubicación del archivo en la hoja Configuración
                configWs.Cells(1, "A").Value = filePath ' Almacenar la ruta en la celda A1
            Else
                MsgBox "No se seleccionó ningún archivo. La macro se detendrá.", vbExclamation
                Exit Sub
            End If
        End With
    End If

    ' Reemplazar la ruta en la fórmula
    formulaR1C1 = Replace(formulaR1C1, "C:\Ruta\Inicial\DTE Reclamados Año 2019-C004.xlsx", filePath)

    ' Copiar la fórmula desde L2 hasta el largo de la columna G en la hoja "BASE"
    For i = 2 To lastRow
        baseWs.Cells(i, "L").formulaR1C1 = formulaR1C1
        ' Actualizar la barra de estado
        Application.StatusBar = "Procesando fila " & i & " de " & lastRow & " (" & Format(i / lastRow, "0%") & ")"
    Next i

    ' Restaurar la barra de estado y reactivar la actualización de pantalla y alertas
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "La fórmula se ha copiado correctamente.", vbInformation
End Sub

