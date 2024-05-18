Attribute VB_Name = "SurveyData"
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Sub PutSurveyDataOnForm(control As IRibbonControl)
resp = MsgBox("WARNING" & vbNewLine & vbNewLine & "This macro will use your keyboard to input your data from the current selection into the webpage table." & _
vbNewLine & vbNewLine & "Before moving forward, please make sure that:" & vbNewLine & "• You're able to navigate the page's table by using the TAB key on the keyboard" & _
 vbNewLine & "• When you hit tab on the last column of the page's table it goes to the next row" & vbNewLine & _
"• You already have the right selection on your excel sheet (usually without the headers)" & vbNewLine & vbNewLine & _
"If you're ready to move forward, please click the OK button. Please, note that you will have 8 seconds to go to the first cell of the website table where you want your data inserted.", vbOKCancel)

If resp = vbOK Then
    Application.Wait (Now + TimeValue("00:00:08"))
    For Each celula In Selection
        Application.SendKeys celula.text
        Sleep (50)
        Application.SendKeys ("{TAB}")
        Sleep (50)
    Next
End If


End Sub