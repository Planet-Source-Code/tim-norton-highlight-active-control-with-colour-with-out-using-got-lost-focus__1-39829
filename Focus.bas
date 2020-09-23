Attribute VB_Name = "Focus"
Dim lastControl As Control
Dim strLastControl As String
Dim oldBackColor As Long

Public Sub CurrentControl(myControl As Control, myForm As Form)
'*************************************************************************
' Author:   Tim Norton
' Date:     5 October 2002
' Purpose:  Highlight the active control in a different colour
' may need to exclude selected controls such as tab strips
' Use the timer event to call this routine
'*************************************************************************
    If Not strLastControl = "" Then 'First time through will be empty
        If myControl.Name = strLastControl Then
            myControl.BackColor = vbGreen 'Set focus colour here
            Exit Sub
        Else
            lastControl.BackColor = oldBackColor 'Restore to original back colour
            oldBackColor = myControl.BackColor 'store the original back colour
            myControl.BackColor = vbGreen 'Set focus colour here
        End If
    Else
        Set lastControl = myControl
        oldBackColor = myControl.BackColor
    End If
    Set lastControl = myControl ' store current control
    strLastControl = lastControl.Name

End Sub
