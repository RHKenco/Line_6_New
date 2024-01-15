Attribute VB_Name = "Status_Messages"

Public Sub statusMsg(newStatus As String, Optional displayStr As String)

'Exit sub if Msg hasn't changed
Static currentMsg
If currentMsg = newStatus Then Exit Sub

' Check if auger or blade
If btn_state = "Active" Then
    If c6kOps.getAugerSet() Then btnState ("Active Auger") Else btnState ("Active Blade")
End If

Dim tempLabel As String
    
Select Case newStatus

    Case "Inactive"
        frmLine6.Var_Label_System_Status.Caption = ""
        frmLine6.Var_Label_System_Status.Visible = False
        frmLine6.Var_Label_System_Status.Refresh
        Exit Sub
        
    Case "Active Blade"
        tempLabel = "Ready to Start Blade"
        
    Case "Active Auger"
        tempLabel = "Ready to Start Auger"
        
    Case "Started"
        tempLabel = "Started - " & displayStr & " Remaining"
        
    Case "Strike"
        tempLabel = "Move to strike location, then press Release"
    Case "Toggle"
        tempLabel = "Please flip switch to proceed"
    Case "Running"
        tempLabel = "Running Pass. Flip Switch to Pause"
    Case "Paused"
        tempLabel = "Pass Paused. Flip Switch to Resume or Press Release to Finish"
    Case "Returned"
        tempLabel = "Pass Completed; Returning to Pass Start." & Chr(13) & "Press Release for Next Pass"
    
    Case "Not-Finish"
        tempLabel = "N/F Pressed - Press Start to Resume"
    
    Case "Finish"
        tempLabel = "Finished " & displayStr & " Parts." & Chr(13) & "Press start for next set of parts," & Chr(13) & "or press clear to enter new WO"
    
    Case "Timeout"
        tempLabel = "System Timed Out - Press Start to Resume"

    Case Else
        MsgBox "Error in Status Message Display"
End Select

frmLine6.Var_Label_System_Status.Caption = tempLabel
frmLine6.Var_Label_System_Status.Visible = True
frmLine6.Var_Label_System_Status.Refresh

currentMsg = newStatus

End Sub
