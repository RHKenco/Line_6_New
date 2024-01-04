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
'|          - 2 - Run
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
End Select

End Sub

Private Sub runState0()
    State = 1
End Sub

Private Sub runState1()

'-- Call Fast Status
Call c6kOps.updFastStatus

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

'Reset Fast Status Update Flag
Call c6kOps.resetFSupd

End Sub

Private Sub runState2()

End Sub
