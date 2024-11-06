Attribute VB_Name = "Módulo10"
Sub accion()
    Dim oc As String
    Dim nro As Long, nr As Long
    Dim porcen As String
    Dim em_1 As Variant, em_a As Variant, reclamo_1 As Variant
    Dim suma_fact As Double, rc As Double, ke As Double
    Dim fecha_pago As Variant, fact_nc As Variant
    Dim ws As Worksheet
    Dim currentCell As Range
    
    ' Desactivar actualización de pantalla para mejorar rendimiento
    Application.ScreenUpdating = False
    
    ' Referencia a la hoja BASE
    Set ws = Sheets("BASE")
    nro = Application.CountA(ws.Range("G:G")) - 1
    Set currentCell = ws.Range("V2")
    
    ' Loop a través del rango de celdas
    For nr = 1 To nro
        porcen = Format((nr / nro) * 100, "0.0") & "%"
        Application.StatusBar = "Progreso: " & porcen & " del cálculo completado"
        
        ' Obtener valores de las celdas con offset
        With currentCell
            reclamo_1 = .Offset(0, -10).Value
            oc = .Offset(0, -14).Value
            em_1 = .Offset(0, -12).Value
            em_a = .Offset(0, -9).Value
            suma_fact = Application.WorksheetFunction.SumIfs(ws.Range("U:U"), ws.Range("H:H"), oc)
            rc = .Offset(0, -6).Value
            ke = .Offset(0, -5).Value
            fecha_pago = .Offset(0, 1).Value
            fact_nc = .Offset(0, -1).Value
        End With

        ' Condiciones para actualizar las celdas
        If fact_nc = "Fact-NC" Then
            currentCell.Value = "Fact-NC"
        ElseIf reclamo_1 = "FACT RECLAMADA" Then
            Call recla_factura(currentCell, oc, em_1, em_a, suma_fact, rc, ke, fecha_pago)
        ElseIf rc < 0 Then
            Call apago(currentCell, oc, em_1, em_a, suma_fact, rc, ke, fecha_pago)
        Else
            Call contabiliza(currentCell, oc, em_1, em_a, suma_fact, rc, ke, fecha_pago)
        End If
        
        Set currentCell = currentCell.Offset(1, 0)
        
        ' Actualizar la pantalla periódicamente
        If nr Mod 50 = 0 Then DoEvents
    Next nr
    
    ' Restaurar actualización de pantalla
    Application.ScreenUpdating = True
    Application.StatusBar = False ' Restablecer barra de estado
End Sub

Sub recla_factura(currentCell As Range, oc As String, em_1 As Variant, em_a As Variant, suma_fact As Double, rc As Double, ke As Double, fecha_pago As Variant)
    Dim reclamo_1 As Variant
    Dim em_s As Double
    
    reclamo_1 = currentCell.Offset(0, -10).Value
    
    ' Calcular proporción de EM
    On Error Resume Next
    em_s = em_a / suma_fact
    
    ' Actualizar celda con el resultado apropiado
    If em_s >= 0.95 And reclamo_1 = "FACT RECLAMADA" Then
        currentCell.Value = "FACT RECLAMADA - REFACTURAR"
    ElseIf em_s < 0.95 And reclamo_1 = "FACT RECLAMADA" Then
        currentCell.Value = "FACT RECLAMADA - Enviar GD"
    End If
End Sub

Sub apago(currentCell As Range, oc As String, em_1 As Variant, em_a As Variant, suma_fact As Double, rc As Double, ke As Double, fecha_pago As Variant)
    Dim reclamo_1 As Variant
    Dim em_s As Double
    
    reclamo_1 = currentCell.Offset(0, -10).Value
    
    ' Calcular proporción de EM
    On Error Resume Next
    em_s = em_a / suma_fact
    
    ' Actualizar celda con el resultado apropiado
    If ke > 0 Then
        currentCell.Value = "Factura con dif a pago el dia " & fecha_pago
    ElseIf rc < 0 Then
        currentCell.Value = "Factura a pago el dia " & fecha_pago
    End If
End Sub

Sub contabiliza(currentCell As Range, oc As String, em_1 As Variant, em_a As Variant, suma_fact As Double, rc As Double, ke As Double, fecha_pago As Variant)
    Dim reclamo_1 As Variant
    Dim em_s As Double
    
    reclamo_1 = currentCell.Offset(0, -10).Value
    
    ' Inicializar celda activa a 0
    currentCell.Value = 0
    
    ' Calcular proporción de EM
    On Error Resume Next
    em_s = em_a / suma_fact
    
    ' Actualizar celda con el resultado apropiado
    If em_s >= 0.95 And em_1 <> "Sin Dato" And reclamo_1 <> "FACT RECLAMADA" Then
        currentCell.Value = "Contabilizar"
    ElseIf em_s <= 0.95 And em_1 > 0 And reclamo_1 <> "FACT RECLAMADA" Then
        currentCell.Value = "Sin EM suficiente"
    ElseIf em_s >= 0.95 And em_1 = "Sin Dato" And reclamo_1 <> "FACT RECLAMADA" Then
        currentCell.Value = "Sin EM registrada"
    Else
        currentCell.Value = "Sin EM suficiente"
    End If
End Sub

