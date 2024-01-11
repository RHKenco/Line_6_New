Attribute VB_Name = "Status_Messages"

Public Sub statusMsg(newStatus As String, displayStr As String)


' Check if auger or blade
If btn_state = "Active" Then
    If c6kOps.getAugerSet() Then btnState ("Active Auger") Else btnState ("Active Blade")
End If

    
Select Case newStatus

    Case "Inactive"
        frmLine6.Var_Label_System_Status.Caption = ""
        frmLine6.Var_Label_System_Status.Visible = False
        frmLine6.Var_Label_System_Status.Refresh
        Exit Sub
        
    Case "Active Blade"
        frmLine6.Var_Label_System_Status.Caption = "Ready to Start Blade"
        
    Case "Active Auger"
        frmLine6.Var_Label_System_Status.Caption = "Ready to Start Auger"
        
    Case "Started"
        frmLine6.Var_Label_System_Status.Caption = "Started - " & displayStr & " Remaining"
    Case "Running"
    
    Case "Not-Finish"
        frmLine6.Var_Label_System_Status.Caption = "N/F Pressed - Press Start to Resume"
    
    Case "Finish"
        frmLine6.Var_Label_System_Status.Caption = "Finished " & displayStr & " Parts. Press start for next set of parts," & Chr(13) & "or press clear to enter new WO"
    
    Case "Timeout"
        frmLine6.Var_Label_System_Status.Caption = "System Timed Out - Press Start to Resume"

    Case Else
        MsgBox "Error in Status Message Display"
End Select

frmLine6.Var_Label_System_Status.Visible = True
frmLine6.Var_Label_System_Status.Refresh

End Sub
