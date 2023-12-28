VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassWO_Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'|====================================================================================================================|
'|
'|              ----- WO Manager Class -----
'|
'|
'|
'|====================================================================================================================|

Option Explicit

Private Const woTimeoutTime = 300   'Work Order Timeout Clock Time in Seconds
Private Const woTimeoutCalib = 10   'Calibration constant - cycle of timer in Hz

Private woActive As Boolean
Private woNum As String
Private woCountdownTime As Integer
Private woStartTime As String
Private woEndTime As String



Public Function isWOactive() As Boolean
    isWOactive = woActive
End Function

Public Function getWOnum() As String
    getWOnum = woNum
End Function

Public Sub loadWO()

'-- Load FORM.ASC Data & Populate Necessary Fields ------------------------------------------
If frmLine6.Text_Enter_WO.Text <> frmLine6.Text_Enter_WO.Tag And frmLine6.Text_Enter_WO.Text <> "" Then

    WORK_ORDER$ = frmLine6.Text_Enter_WO.Text
    ' Load FORM.ASC
    Open "F:\MFG\FORM.ASC" For Input As #1
    '
    Do While WORK_ORDER$ <> WOK$
         Input #1, WOK$, PartNumber, TQTY!, PRICE!, Tc!, COMM$, DAT$, DWG$, UNUSED_1!, UNUSED_2!
            If WORK_ORDER$ = WOK$ Then
                
                Close #1
                
                woNum = WOK$
                
                woActive = True
                
                'Populate Part # Field
                frmLine6.Text_Pop_Part_Num.Text = PartNumber
                'Populate Dwg # Field
                frmLine6.Text_Pop_Dwg_Num.Text = DWG$
                'Check Due Date; if overdue change background to red
                If Date >= CDate(DAT$) Then frmLine6.Text_Pop_Due_Date.BackColor = QBColor(4) Else frmLine6.Text_Pop_Due_Date.BackColor = &HFFFFFF
                'Populate Due Date
                frmLine6.Text_Pop_Due_Date.Text = DAT$
                
                'Update Work Order Status & System Status Labels
                frmLine6.Var_Label_WO_Active.Caption = "Work Order " + WOK$ + " Active"
                frmLine6.Var_Label_WO_Active.Visible = True
                
                frmLine6.Var_Label_System_Status.Caption = "Ready to Run Blade"
                frmLine6.Var_Label_System_Status.Visible = True
                
                'Update Buttons
                btnState ("Active")
                
                frmLine6.Refresh
                            
                'Exit Loop
                Exit Do
            End If
            
            ' Check for End-Of-File
            If EOF(1) Then
            
                'If EOF, Close FORM.ASC, Throw Error Message, and Exit Sub.
                Close #1
                MsgBox "WORK ORDER NOT FOUND!"
                
                'Exit Loop
                Exit Do
                
            End If
            
    Loop
    
    'Close FORM.ASC
    Close #1
End If

End Sub

Public Sub clearWO()
    '-- Clear All Fields and Reset Button to Enter ------------------------------------------------
    frmLine6.Text_Enter_WO.Text = Text_Enter_WO.Tag
    frmLine6.Text_Enter_WO.Locked = False
    
    frmLine6.Text_Pop_Part_Num.Text = ""
    frmLine6.Text_Pop_Dwg_Num.Text = ""
    frmLine6.Text_Pop_Due_Date.Text = ""

    'Update Work Order Status & System Status Labels
    frmLine6.Var_Label_WO_Active.Caption = Var_Label_WO_Active.Tag
    frmLine6.Var_Label_WO_Active.Visible = True
    
    frmLine6.Var_Label_System_Status.Visible = False
    
    woNum = ""
    woActive = False
    
    'Update Buttons
    btnState ("Inactive")
    
    frmLine6.Refresh
End Sub

'-- btnState - Function to handle visibility and enabling of Enter, Go, Start, N/F, and Finish Buttons -------------------------------------
Public Sub btnState(btn_state As String)
    
    If btn_state = "Inactive" Then
        
        frmLine6.Button_WO_Enter_Clear.Enabled = True
        frmLine6.Button_WO_Enter_Clear.Caption = "Enter"
        frmLine6.Text_Enter_WO.Locked = False
        
        frmLine6.Button_Start.Enabled = False
        frmLine6.Button_Start.Visible = True
        frmLine6.Button_NF.Enabled = False
        frmLine6.Button_NF.Visible = False
        frmLine6.Button_Fin.Enabled = False
        frmLine6.Button_Fin.Visible = False
        
        frmLine6.Button_Go.Enabled = False
        frmLine6.Button_Go.Visible = True
        
    ElseIf btn_state = "Active" Then
            
        frmLine6.Button_WO_Enter_Clear.Enabled = True
        frmLine6.Button_WO_Enter_Clear.Caption = "Clear"
        frmLine6.Text_Enter_WO.Locked = True
                    
        frmLine6.Button_Start.Enabled = True
        frmLine6.Button_Start.Visible = True
        frmLine6.Button_NF.Enabled = False
        frmLine6.Button_NF.Visible = False
        frmLine6.Button_Fin.Enabled = False
        frmLine6.Button_Fin.Visible = False
        
        frmLine6.Button_Go.Enabled = False
        
    ElseIf btn_state = "Started" Then
    
        frmLine6.Button_WO_Enter_Clear.Enabled = False
            
        frmLine6.Button_Start.Enabled = False
        frmLine6.Button_Start.Visible = False
        frmLine6.Button_NF.Enabled = True
        frmLine6.Button_NF.Visible = True
        frmLine6.Button_Fin.Enabled = True
        frmLine6.Button_Fin.Visible = True
        
        frmLine6.Button_Go.Enabled = True
        
    ElseIf btn_state = "Running" Then
    
        frmLine6.Button_WO_Enter_Clear.Enabled = False
            
        frmLine6.Button_Start.Enabled = False
        frmLine6.Button_Start.Visible = False
        frmLine6.Button_NF.Enabled = False
        frmLine6.Button_NF.Visible = True
        frmLine6.Button_Fin.Enabled = False
        frmLine6.Button_Fin.Visible = True
        
        frmLine6.Button_Go.Enabled = False
        
    ElseIf btn_state = "Returned" Then
    
        frmLine6.Button_WO_Enter_Clear.Enabled = False
            
        frmLine6.Button_Start.Enabled = False
        frmLine6.Button_Start.Visible = False
        frmLine6.Button_NF.Enabled = True
        frmLine6.Button_NF.Visible = True
        frmLine6.Button_Fin.Enabled = True
        frmLine6.Button_Fin.Visible = True
        
        frmLine6.Button_Go.Enabled = False
        
    Else
        MsgBox ("Error in Button State Function")
        
        Exit Sub
        
    End If
    
    frmLine6.Refresh

End Sub

Public Function woTimer(timerCmd As String) As Integer

Select Case timerCmd
    Case "Start", "Reset"
        woCountdownTime = woTimeoutTime * woTimeoutCalib
    Case "Decrement"
        If woCountdownTime > 0 Then
            woCountdownTime = woCountdownTime - 1
        End If
    Case "Get"
    
    Case Else
        MsgBox ("Inproper Input in woTimer Function")
    
End Select

woTimer = woCountdownTime

End Function
