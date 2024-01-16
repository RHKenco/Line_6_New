Attribute VB_Name = "Button_State"
'-- btnState() - Function to handle visibility and enabling of Enter, Go, Start, N/F, and Finish Buttons -------------------------------------
'   - Inputs: btn_State - string containing desired button state
Public Sub btnState(btn_state As String)
    'Button States:
    '   - "Inactive" - The default state on loading. No work order is active
    '   - "Active" - Calls check to see if blade or auger - NOT USED IN BUTTON STATES
    '   - "Active Blade" - Work order is active in BLADE MODE, but has not been started
    '   - "Active Auger" - Work order is active in AUGER MODE, but has not been started
    '   - "Started" - Work order has been started, waiting for "Go" command
    '   - "Running" - Work order is running
    
    
    ' Check if auger or blade
    If btn_state = "Active" Then
        If c6kOps.getAugerSet() Then btn_state = ("Active Auger") Else btn_state = ("Active Blade")
    End If
    
    Call btnSt_Button_WO_Enter_Clear(btn_state)
    Call btnSt_Button_NF_Fin(btn_state)
    Call btnSt_Button_Start(btn_state)
    Call btnSt_Button_Go(btn_state)
    Call btnSt_Button_Set_Clear_Auger(btn_state)
    
    
    frmLine6.Refresh

End Sub

Private Sub btnSt_Button_WO_Enter_Clear(btn_state As String)

Select Case btn_state       'States: "Inactive", "Active Blade", "Active Auger", "Started", "Running"
    Case "Inactive"
        frmLine6.Button_WO_Enter_Clear.Enabled = True
        frmLine6.Button_WO_Enter_Clear.Caption = "Enter"
        frmLine6.Text_Enter_WO.Locked = False
        
    Case "Active Blade", "Active Auger"
        frmLine6.Button_WO_Enter_Clear.Enabled = True
        frmLine6.Button_WO_Enter_Clear.Caption = "Clear"
        frmLine6.Text_Enter_WO.Locked = True
        
    Case "Started", "Running"
        frmLine6.Button_WO_Enter_Clear.Enabled = False
        
    Case Else
        MsgBox "Unspecified Button State for Enter/Clear Button: " & btn_state
End Select
        
End Sub


Private Sub btnSt_Button_Start(btn_state As String)

Select Case btn_state       'States: "Inactive", "Active Blade", "Active Auger", "Started", "Running"
    Case "Inactive"
        frmLine6.Button_Start.Enabled = False
        frmLine6.Button_Start.Visible = True
    Case "Active Blade", "Active Auger"
        frmLine6.Button_Start.Enabled = True
        frmLine6.Button_Start.Visible = True
    Case "Started", "Running"
        frmLine6.Button_Start.Enabled = False
        frmLine6.Button_Start.Visible = False
    Case Else
        MsgBox "Unspecified Button State for Start Button: " & btn_state
        
End Select

End Sub

Private Sub btnSt_Button_NF_Fin(btn_state As String)

Select Case btn_state       'States: "Inactive", "Active Blade", "Active Auger", "Started", "Running"

        
    Case "Inactive", "Active Blade", "Active Auger"
        frmLine6.Button_NF.Enabled = False
        frmLine6.Button_NF.Visible = False
        
        frmLine6.Button_Fin.Enabled = False
        frmLine6.Button_Fin.Visible = False
    Case "Started"
        frmLine6.Button_NF.Enabled = True
        frmLine6.Button_NF.Visible = True
        
        frmLine6.Button_Fin.Enabled = True
        frmLine6.Button_Fin.Visible = True
    Case "Running"
        frmLine6.Button_NF.Enabled = False
        frmLine6.Button_NF.Visible = True
        
        frmLine6.Button_Fin.Enabled = False
        frmLine6.Button_Fin.Visible = True
    Case Else
        MsgBox "Unspecified Button State for Not Finish & Finish Buttons: " & btn_state
        
End Select

End Sub

Private Sub btnSt_Button_Go(btn_state As String)

Select Case btn_state       'States: "Inactive", "Active Blade", "Active Auger", "Started", "Running"

        
    Case "Inactive", "Active Blade", "Active Auger"
        frmLine6.Button_Go.Enabled = False
        frmLine6.Button_Go.Visible = True
    Case "Started"
        frmLine6.Button_Go.Enabled = True
        frmLine6.Button_Go.Visible = True
    Case "Running"
        frmLine6.Button_Go.Enabled = False
        frmLine6.Button_Go.Visible = True
    Case Else
        MsgBox "Unspecified Button State for GO Button: " & btn_state
        
End Select

End Sub

Private Sub btnSt_Button_Set_Clear_Auger(btn_state As String)

Select Case btn_state       'States: "Inactive", "Active Blade", "Active Auger", "Started", "Running"
    Case "Inactive"
        frmLine6.Button_Set_Auger.Enabled = False
        frmLine6.Button_Set_Auger.Visible = True
        
        frmLine6.Button_Clear_Auger.Enabled = False
        frmLine6.Button_Clear_Auger.Visible = False
    Case "Active Blade"
        frmLine6.Button_Set_Auger.Enabled = True
        frmLine6.Button_Set_Auger.Visible = True
        
        frmLine6.Button_Clear_Auger.Enabled = False
        frmLine6.Button_Clear_Auger.Visible = False
    Case "Active Auger"
        frmLine6.Button_Set_Auger.Enabled = False
        frmLine6.Button_Set_Auger.Visible = False
        
        frmLine6.Button_Clear_Auger.Enabled = True
        frmLine6.Button_Clear_Auger.Visible = True
    Case "Started", "Running"
        frmLine6.Button_Set_Auger.Enabled = False
        
        frmLine6.Button_Clear_Auger.Enabled = False
    Case Else
        MsgBox "Unspecified Button State for Auger Clear Set Buttons: " & btn_state
        
End Select

End Sub





'Private Sub btnSt_(btn_state As String)
'
'Select Case btn_state       'States: "Inactive", "Active Blade", "Active Auger", "Started", "Running"
'
'    Case Else
'        MsgBox "Unspecified Button State for ________ Button"
'
'End Select
'
'End Sub



