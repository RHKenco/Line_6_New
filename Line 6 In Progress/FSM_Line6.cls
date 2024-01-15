VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FSM_Line6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'|====================================================================================================================|
'|
'|           ------ Line 6 Main Finite-State Machine ------
'|
'|
'|      States:
'|          - 0 - Setup
'|                  - initializes FSM
'|          - 1 - Idle
'|                  - Waits for input from Form
'|          - 2 - Transition
'|                  - Initializes Run mode FSM
'|          - 3 - Run
'|                  - Operates Run mode FSM
'|
'|
'|====================================================================================================================|



Option Explicit

'Set FSM Input Variables
Public State As Integer

'Set FSM Internal Variables







Public Sub Run()

Select Case State
    Case 0
        Call runState0
    Case 1
        Call runState1
    Case 2
        Call runState2
    Case 3
        Call runState3
End Select

End Sub

Private Sub runState0() 'Setup State - Runs any necessary FSM Setup code and moves immediately to state 1
    State = 1
End Sub

Private Sub runState1() 'Idle State - Joystick motion & UI Operation

'-- Call Fast Status
Call c6kOps.updFastStatus

'-- Check for E-Stop
    If Not c6kOps.chkE_Stop Then
    
    
    '-- Run Joystick if active
        If c6kOps.getJoyActive And Not c6kOps.chkE_Stop() Then
        
            'Run JoyRun function, and if it returns true,
            If c6kOps.runJoy("Free") Then
        
                'Set Joystick Status Message
                frmLine6.Var_Label_Joystick_Status.Caption = "Joystick" & Chr(13) & "Enabled:" & Chr(13) & Chr(13) & c6kOps.getJoyStr() & Chr(13) & "Mode"
                frmLine6.Var_Label_Joystick_Status.Visible = True
        
            Else
                'If the joystick becomes inactive hide label
                frmLine6.Var_Label_Joystick_Status.Visible = False
        
            End If
        
        '--Input State Debug - Uncomment these two lines to enter input debug mode
        'frmLine6.Var_Label_Joystick_Status.Caption = c6kOps.getInputState
        'frmLine6.Var_Label_Joystick_Status.Visible = True
            
        End If
        
        
    '-- Run Timeout Timer if Start has been pressed
        If woMgr.isWOstarted Then
            woMgr.woTimer ("Count")
        End If
        
    End If

End Sub

Private Sub runState2() 'Initialize Pass State - Runs setup for pass operation

'-- Call Fast Status
Call c6kOps.updFastStatus
'-- Check for E-Stop and H2O Flow
    If Not c6kOps.chkE_Stop Then
    
    
    '-- Run subFSM initialization state
    
        fsmRun.State = 0
        Call fsmRun.Run
        
    'Proceed to main run state
    State = 3

    End If

End Sub


Private Sub runState3() 'Run Pass State - Runs sub-FSM to operate the machine while running a pass

'-- Call Fast Status
Call c6kOps.updFastStatus
'-- Check for E-Stop and H2O Flow
    If Not c6kOps.chkE_Stop And Not c6kOps.chkH2Oflow Then
    
    
    '-- Run subFSM for run-state
        Call fsmRun.Run

    End If

End Sub
