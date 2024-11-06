Attribute VB_Name = "Módulo2"


Sub limpia_planilla()

If CheckBox1 = True Then
Call limpia_base
End If
If CheckBox2 = True Then
Call limpia_ppl
End If
If CheckBox3 = True Then
Call limpia_nc
End If
If CheckBox4 = True Then
Call limpia_pf0
End If
End Sub

Sub limpia_base()

Sheets("BASE").Select
If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
Range("A3").Select
Range("A3:aj3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
Sheets("Inicio").Select
End Sub

Sub limpia_ppl()

 Sheets("PPL").Select
 If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
Sheets("Inicio").Select
End Sub

  Sub limpia_pf0()
Attribute limpia_pf0.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro10 Macro
'

'
    Sheets("PF0").Select
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    Range("C2:x2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
 Sheets("Inicio").Select
End Sub
Sub limpia_nc()
Attribute limpia_nc.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro11 Macro
'

'
    Sheets("NC").Select
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    Range("C2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Sheets("Inicio").Select
    
End Sub


Sub limpia_em()
Attribute limpia_em.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro12 Macro
'

'
    Sheets("EM").Select
   If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    Sheets("Inicio").Select
End Sub

Sub limpia_base2()
Attribute limpia_base2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro13 Macro
'

'
    Sheets("base 2").Select
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
   
    Sheets("Inicio").Select
End Sub
