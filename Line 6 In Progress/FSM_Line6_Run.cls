VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FSM_Line6_Run"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------
'|
'|           ------ Line 6 Run-Pass Finite-State Machine ------
'|
'|
'|      States:
'|          - 0 - Setup
'|                  - initializes Oscillator
'|          - 1 - Set Torch & Strike Arc
'|                  - Enables Free Joystick to position Torch, Homes Oscillator,
'|                      and Strikes Arc / Starts Osc uppon Transition
'|          - 2 - Pause
'|                  - Stops movement & allows for re-positioning joystick. Finish Pass from this state
'|          - 3 - Run Pass
'|                  - Starts movement & operates mid-pass joystick
'|          - 4 - Return
'|                  - Cuts Arc and returns to beginning of pass
'|
'|
'----------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------



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
    Case 4
        Call runState4
End Select

End Sub

Private Sub runState0() 'Setup State - Runs any necessary FSM Setup code and moves immediately to state 1

'Home Oscillator
c6kOps.homeOsc
    
State = 1
    
End Sub

Private Sub runState1() 'Set Torch & Strike Arc State - Set torch for initial pass, then strike arc

End Sub

Private Sub runState2() 'Pause State - Joystick movement & Oscillator Only

'-- Run Oscillator
Call c6kOps.runOsc


'-- Check for Inputs
'** If SwitchState = Off, Resume Pass & set FSM State to 3 (Run Mode)
'** If Release Is Pressed, set FSM to State 4



End Sub

Private Sub runState3() 'Run State - Joystick in pass adjustment mode, Oscillator running

'-- Run Oscillator
Call c6kOps.runOsc


'-- Check for switch
'** If SwitchState = ON, Pause Pass & set FSM State to 2 (Pause Mode)


'-- Run Pass
Call c6kOps.runPass("Update")


End Sub

Private Sub runState4() 'Return State - moves head back to start of pass, then backs away perpendicularly by set distance

End Sub
