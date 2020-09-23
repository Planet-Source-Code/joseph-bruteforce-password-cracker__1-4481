Attribute VB_Name = "mdlBruteForce"
Public Sub BFFile()
Open frmMain.CDial.FileName For Input As #1
While Not (EOF(1))
Line Input #1, a
If frmMain.optWaitOff.Value = True Then SendKeys a + frmMain.txtFinal
If frmMain.optWaitOn.Value = True Then SendKeys a + frmMain.txtFinal, wait

DoEvents
Wend
Close #1
frmMain.txtDelay.Text = ticktotal
End Sub
Public Sub BFSequence()
For i = frmLimit.txtLower To frmLimit.txtUpper
If frmMain.optWaitOn.Value = True Then SendKeys Mid$(Str(i), 2) + frmMain.txtFinal, wait
If frmMain.optWaitOff.Value = True Then SendKeys (Mid$(Str(i), 2) + frmMain.txtFinal)



DoEvents
Next i
frmMain.txtDelay.Text = ticktotal

End Sub
Public Sub BFText()
If frmMain.optWaitOff = True Then SendKeys frmMain.txtFinal
If frmMain.optWaitOn = True Then SendKeys frmMain.txtFinal, wait

DoEvents
frmMain.txtDelay.Text = ticktotal

End Sub
