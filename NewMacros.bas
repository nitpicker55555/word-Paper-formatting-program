Attribute VB_Name = "NewMacros"
Sub 브1()
Attribute 브1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.브1"
'
' 브1 브
'
'
End Sub
Sub 브2()
Attribute 브2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.브2"
'
' 브2 브
'
'
Selection.HomeKey Unit:=wdLine
    Selection.Style = ActiveDocument.Styles("깃痙 1")
    Selection.EndKey Unit:=wdLine
End Sub
Sub 브3()
Attribute 브3.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.브3"
'
' 브3 브
'
'
    Selection.HomeKey Unit:=wdLine
    Selection.Style = ActiveDocument.Styles("깃痙 2")
    Selection.EndKey Unit:=wdLine
End Sub
Sub 브4()
Attribute 브4.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.브4"
'
' 브4 브
'
'
Selection.HomeKey Unit:=wdLine
    Selection.Style = ActiveDocument.Styles("깃痙 3")
    Selection.EndKey Unit:=wdLine
End Sub
Sub 브5()
Attribute 브5.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.브5"
'
' 브5 브
'
'
    Options.DefaultHighlightColorIndex = wdYellow
    Selection.Range.HighlightColorIndex = wdYellow
    Selection.Font.Bold = wdToggle
End Sub
Sub 브6()
Attribute 브6.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.브6"
'
' 브6 브
'
'
    Selection.HomeKey Unit:=wdLine
    Selection.EndKey Unit:=wdLine
End Sub
Sub 브7()
Attribute 브7.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.브7"
'
' 브7 브
'
'
    Selection.Font.Superscript = wdToggle
    Selection.TypeText Text:="[1]"
    Selection.Font.Superscript = wdToggle
End Sub
Sub 브8()
Attribute 브8.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.브8"
'
' 브8 브
'
'
    Selection.Font.Superscript = wdToggle
    Selection.TypeText Text:="[]"
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub 브9()
Attribute 브9.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.브9"
'
' 브9 브
'
Selection.HomeKey Unit:=wdLine
    Selection.Style = ActiveDocument.Styles("깃痙 4")
    Selection.EndKey Unit:=wdLine
End Sub
