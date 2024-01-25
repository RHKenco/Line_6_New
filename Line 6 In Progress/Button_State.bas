Attribute VB_Name = "Button_State"

'Button State Definitions
Public Const btnInactive As Integer = 1
Public Const btnActive As Integer = 2
Public Const btnActiveBlade As Integer = 3
Public Const btnActiveAuger As Integer = 4
Public Const btnStarted As Integer = 5
Public Const btnRunning As Integer = 6
Public Const btnEstop As Integer = 7


'-- btnState() - Function to handle visibility and enabling of Enter, Go, Start, N/F, and Finish Buttons -------------------------------------
'   - Inputs: btn_State - string containing desired button state
Public Sub btnState(btn_state As Integer)
    'Button States:
    '   - "Inactive" - The default state on loading. No work order is active
    '   - "Active" - Calls check to see if blade or auger - NOT USED IN BUTTON STATES
    '   - "Active Blade" - Work order is active in BLADE MODE, but has not been started
    '   - "Active Auger" - Work order is active in AUGER MODE, but has not been started
    '   - "Started" - Work order has been started, waiting for "Go" command
    '   - "Running" - Work order is running
    
    
    ' Check if auger or blade
    If btn_state = btnActive Then
        If c6kOps.getAugerSet() Then btn_state = btnActiveAuger Else btn_state = btnActiveBlade
    End If
    
    Call btnSt_Button_WO_Enter_Clear(btn_state)
    Call btnSt_Button_NF_Fin(btn_state)
    Call btnSt_Button_Start(btn_state)
    Call btnSt_Button_Go(btn_state)
    Call btnSt_Button_Set_Clear_Auger(btn_state)
    
    
    frmLine6.Refresh

End Sub

Private Sub btnSt_Button_WO_Enter_Clear(btn_state As Integer)

Select Case btn_state       'States: "Inactive", "Active Blade", "Active Auger", "Started", "Running"
    Case btnInactive
        frmLine6.Button_WO_Enter_Clear.Enabled = True
        frmLine6.Button_WO_Enter_Clear.Caption = "Enter"
        frmLine6.Text_Enter_WO.Locked = False
        
    Case btnActiveBlade, btnActiveAuger
        frmLine6.Button_WO_Enter_Clear.Enabled = True
        frmLine6.Button_WO_Enter_Clear.Caption = "Clear"
        frmLine6.Text_Enter_WO.Locked = True
        
    Case btnStarted, btnRunning, btnEstop
        frmLine6.Button_WO_Enter_Clear.Enabled = False
        
    Case Else
        MsgBox "Unspecified Button State for Enter/Clear Button: " & btn_state
End Select
        
End Sub


Private Sub btnSt_Button_Start(btn_state As Integer)

Select Case btn_state       'States: "Inactive", "Active Blade", "Active Auger", "Started", "Running"
    Case btnInactive
        frmLine6.Button_Start.Enabled = False
        frmLine6.Button_Start.Visible = True
    Case btnActiveBlade, btnActiveAuger
        frmLine6.Button_Start.Enabled = True
        frmLine6.Button_Start.Visible = True
    Case btnStarted, btnRunning, btnEstop
        frmLine6.Button_Start.Enabled = False
        frmLine6.Button_Start.Visible = False
    Case Else
        MsgBox "Unspecified Button State for Start Button: " & btn_state
        
End Select

End Sub

Private Sub btnSt_Button_NF_Fin(btn_state As Integer)

Select Case btn_state       'States: "Inactive", "Active Blade", "Active Auger", "Started", "Running"

        
    Case btnInactive, btnActiveBlade, btnActiveAuger, btnEstop
        frmLine6.Button_NF.Enabled = False
        frmLine6.Button_NF.Visible = False
        
        frmLine6.Button_Fin.Enabled = False
        frmLine6.Button_Fin.Visible = False
    Case btnStarted
        frmLine6.Button_NF.Enabled = True
        frmLine6.Button_NF.Visible = True
        
        frmLine6.Button_Fin.Enabled = True
        frmLine6.Button_Fin.Visible = True
    Case btnRunning
        frmLine6.Button_NF.Enabled = False
        frmLine6.Button_NF.Visible = True
        
        frmLine6.Button_Fin.Enabled = False
        frmLine6.Button_Fin.Visible = True
    Case Else
        MsgBox "Unspecified Button State for Not Finish & Finish Buttons: " & btn_state
        
End Select

End Sub

Private Sub btnSt_Button_Go(btn_state As Integer)

Select Case btn_state       'States: "Inactive", "Active Blade", "Active Auger", "Started", "Running"

        
    Case btnInactive, btnActiveBlade, btnActiveAuger, btnRunning, btnEstop
        frmLine6.Button_Go.Enabled = False
    Case btnStarted
        frmLine6.Button_Go.Enabled = True
    Case Else
        MsgBox "Unspecified Button State for GO Button: " & btn_state
        
End Select

End Sub

Private Sub btnSt_Button_Set_Clear_Auger(btn_state As Integer)

Select Case btn_state       'States: "Inactive", "Active Blade", "Active Auger", "Started", "Running"
    Case btnInactive, btnEstop
        frmLine6.Button_Set_Auger.Enabled = False
        frmLine6.Button_Set_Auger.Visible = True
        
        frmLine6.Button_Clear_Auger.Enabled = False
        frmLine6.Button_Clear_Auger.Visible = False
    Case btnActiveBlade
        frmLine6.Button_Set_Auger.Enabled = True
        frmLine6.Button_Set_Auger.Visible = True
        
        frmLine6.Button_Clear_Auger.Enabled = False
        frmLine6.Button_Clear_Auger.Visible = False
    Case btnActiveAuger
        frmLine6.Button_Set_Auger.Enabled = False
        frmLine6.Button_Set_Auger.Visible = False
        
        frmLine6.Button_Clear_Auger.Enabled = True
        frmLine6.Button_Clear_Auger.Visible = True
    Case btnStarted, btnRunning
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



