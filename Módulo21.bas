Attribute VB_Name = "Módulo21"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

Dim fecha_conta As Date
Dim resp As Integer
Dim filtr As String

resp = MsgBox("Vas a mandar los que dicen Contabilizar?", vbQuestion + vbYesNo + vbDefaultButton2, "Contabilicemos")

If resp = vbYes Then
    filtr = "Contabilizar"
Else
    filtr = Application.InputBox("Cual es el filtro?")
End If
Sheets("BASE").Select
If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData

fecha_conta = Date - 15

    ActiveSheet.Range("$A:$BC").AutoFilter Field:=22, Criteria1:= _
        filtr
    'ActiveSheet.Range("$A:$BC").AutoFilter Field:=4, Criteria1:=
     '   "<=" & Date - 15, Operator:=xlAnd

    Sheets("BASE").Select
    Columns("A:J").Select
    Selection.Copy
    Sheets.Add
    ActiveSheet.Name = "contabilizacion"
     
    
    Range("A1").Select
    ActiveSheet.Paste
    
End Sub

Sub fecha_reg()

Dim nro As String
Dim nr As Integer
Range("k1").Select
ActiveCell = "Fecha Registro"
ActiveCell.Offset(1, 0).Select
nro = (Application.CountA(Range("j:j")) - 1)
For nr = 1 To nro

ActiveCell = Date
ActiveCell.Offset(1, 0).Select
Next
End Sub

Sub rut()

Dim nro As String
Dim nr As Integer
Dim rut As String
Range("n1").Select
ActiveCell = "RUT P02"
ActiveCell.Offset(1, 0).Select
nro = (Application.CountA(Range("j:j")) - 1)
For nr = 1 To nro
rut = ActiveCell.Offset(0, -12)
ActiveCell = rut
ActiveCell.Offset(1, 0).Select
Next

End Sub

Sub razon_social()

Dim nro As String
Dim nr As Integer
Dim razon As String
Range("p1").Select
ActiveCell = "Razon_Social"
ActiveCell.Offset(1, 0).Select
nro = (Application.CountA(Range("j:j")) - 1)
For nr = 1 To nro
razon = ActiveCell.Offset(0, -13)
ActiveCell = razon
ActiveCell.Offset(1, 0).Select
Next
End Sub

Sub nro_fact()

Dim nro As String
Dim nr As Integer
Dim nro_factura As String
Range("q1").Select
ActiveCell = "Nº Factura"
ActiveCell.Offset(1, 0).Select
nro = (Application.CountA(Range("j:j")) - 1)
For nr = 1 To nro
nro_factura = ActiveCell.Offset(0, -16)
ActiveCell = nro_factura
ActiveCell.Offset(1, 0).Select
Next

End Sub

Sub fecha_fact()

Dim nro As String
Dim nr As Integer
Dim fecha_factura As Date
Range("r1").Select
ActiveCell = "Fecha_Factura"
ActiveCell.Offset(1, 0).Select
nro = (Application.CountA(Range("j:j")) - 1)
For nr = 1 To nro
fecha_factura = ActiveCell.Offset(0, -14)
ActiveCell = fecha_factura
ActiveCell.Offset(1, 0).Select
Next

End Sub

Sub monto_neto()

Dim nro As String
Dim nr As Integer
Dim monto_neto As String
Range("s1").Select
ActiveCell = "Monto Neto"
ActiveCell.Offset(1, 0).Select
nro = (Application.CountA(Range("j:j")) - 1)
For nr = 1 To nro
monto_neto = Format((ActiveCell.Offset(0, -14) * 0.81), "$0;(0)")

ActiveCell = monto_neto
ActiveCell.Offset(1, 0).Select
monto_neto = 0
Next

End Sub

Sub monto_iva()

Dim nro As String
Dim nr As Integer
Dim monto_iva As String
Range("t1").Select
ActiveCell = "IVA"
ActiveCell.Offset(1, 0).Select
nro = (Application.CountA(Range("j:j")) - 1)
For nr = 1 To nro
monto_iva = Format((ActiveCell.Offset(0, -15) * 0.19), "$0;(0)")

ActiveCell = monto_iva
ActiveCell.Offset(1, 0).Select
monto_iva = 0
Next

End Sub

Sub monto()

Dim nro As String
Dim nr As Integer
Dim monto As String
Range("u1").Select
ActiveCell = "Monto_Total"
ActiveCell.Offset(1, 0).Select
nro = (Application.CountA(Range("j:j")) - 1)
For nr = 1 To nro
monto = Format((ActiveCell.Offset(0, -16)), "$0;(0)")
ActiveCell = monto
ActiveCell.Offset(1, 0).Select
Next

End Sub

Sub oc()

Dim nro As String
Dim nr As Integer
Dim oc As String
Range("v1").Select
ActiveCell = "Nº OC"
ActiveCell.Offset(1, 0).Select
nro = (Application.CountA(Range("j:j")) - 1)
For nr = 1 To nro
oc = ActiveCell.Offset(0, -14)
ActiveCell = oc
ActiveCell.Offset(1, 0).Select
Next

End Sub

Sub gd()

Dim nro As String
Dim nr As Integer
Dim gd As String
Range("w1").Select
ActiveCell = "GD"
ActiveCell.Offset(1, 0).Select
nro = (Application.CountA(Range("j:j")) - 1)
For nr = 1 To nro
gd = ActiveCell.Offset(0, -14)
ActiveCell = gd
ActiveCell.Offset(1, 0).Select
Next



End Sub

Sub em()

Dim nro As String
Dim nr As Integer
Dim em As String
Range("x1").Select
ActiveCell = "Nº Recepcion"
ActiveCell.Offset(1, 0).Select
nro = (Application.CountA(Range("j:j")) - 1)
For nr = 1 To nro
em = ActiveCell.Offset(0, -14)
ActiveCell = em
ActiveCell.Offset(1, 0).Select
Next

End Sub

Sub ppl()

Dim nro As String
Dim nr As Integer
Dim ppl As String
Range("y1").Select
ActiveCell = "PPL"
ActiveCell.Offset(1, 0).Select
nro = (Application.CountA(Range("j:j")) - 1)
For nr = 1 To nro
ppl = ActiveCell.Offset(0, -19)
ActiveCell = ppl
ActiveCell.Offset(1, 0).Select
Next

End Sub

Sub borra()

Columns("A:A").Select
   
    Columns("A:J").Select
    Selection.Delete Shift:=xlToLeft
End Sub

Sub cta_fico()



Dim nro As String
Dim nr As Integer
Dim cta_fico As String
Dim ran_fico As Range
Dim rut As String
Set ran_fico = Sheets("PF0").Range("b:c")
Range("o1").Select
ActiveCell = "Cta. FICO"
ActiveCell.Offset(1, 0).Select
nro = (Application.CountA(Range("j:j")) - 1)
For nr = 1 To nro
rut = ActiveCell.Offset(0, -1)
cta_fico = Application.VLookup(rut, ran_fico, 2, 0)
ActiveCell = cta_fico
ActiveCell.Offset(1, 0).Select
Next
End Sub

Sub Macro5()
'
' Macro5 Macro
'

'
   Dim wwb As String
   
    
    wwb = ActiveWorkbook.Name
    Columns("k:y").Select
    'Selection.Delete Shift:=xlToLeft
    'Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste

    Sheets("Hoja1").Select
    Application.CutCopyMode = False
     Windows(wwb).Activate
    Sheets("contabilizacion").Select
    Application.DisplayAlerts = False
    Sheets("contabilizacion").Delete
    
End Sub

Sub crea()

Call Macro1
Call fecha_reg
Call fecha_fact
Call em
Call gd
Call monto_neto
Call monto_iva
Call monto
Call nro_fact

Call oc
Call ppl
Call razon_social
Call rut
Call cta_fico
Call Macro5
End Sub
