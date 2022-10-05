Attribute VB_Name = "Módulo1"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja2").Select
    Sheets("Hoja2").Name = "Hoja Nueva"
    Range("H12").Select
End Sub
