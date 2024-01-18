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
Private stateInner As Integer
Private delayTimer As Integer




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
'Call c6kOps.homeOsc

'Start Water Pump
Call c6kOps.setOutput("Water Pump", True)
    
'Initialize Joystick
Call Joy.runJoy("Enable")
    
'Initialize internal variables
stateInner = 0

'Proceed to next state
State = 1
    
End Sub

Private Sub runState1() 'Set Torch & Strike Arc State - Set torch for initial pass, then strike arc

Select Case stateInner
    Case 0
        'Set status message
        Call statusMsg("Strike", "")
        'Wait until joystick is inactive
        If Joy.runJoy("Free") Then Exit Sub
        stateInner = 1
    Case 1
        'Strike Arc
        Call c6kOps.setMachineHome
        Call c6kOps.strikeArc
        stateInner = 2
    Case 2
        'Verify that the pause switch is thrown, then
        If c6kOps.getJoyToggle Then
            'Proceed to pause state
            State = 2
            stateInner = 0
            
            'Set status message
            Call statusMsg("Paused", "")
        
        Else
            'Set status message
            Call statusMsg("Toggle", "")
        End If
    Case Else
        MsgBox "Error in Pass Initialization"
End Select

End Sub

Private Sub runState2() 'Pause State - Joystick movement & Oscillator Only


'-- Run Joystick
If Not Joy.runJoy("Pause") Then      'If the joystick is released, figure out why

    '-- Check for Inputs
    If Not Joy.getJoyToggle Then 'If switch is not on
        
        'Revert to Run state
        State = 3
        Call c6kOps.RunPass("Start")
        
        'Set status message
        Call statusMsg("Running", "")
        
        Exit Sub
    End If
    
    'If the switch is on, then release was pressed - Switch to Return state
    State = 4
    Call c6kOps.RunPass("Stop")

End If

End Sub

Private Sub runState3() 'Run State - Joystick in pass adjustment mode, Oscillator running

'-- Run Pass
Call c6kOps.RunPass("Update")

'-- Run Joystick
If Not Joy.runJoy("Run") Then        'If the joystick is released, figure out why
    
    '-- Check for switch
    If c6kOps.getJoyToggle Then 'If switch is on
        
        'Revert to Pause State
        State = 2
        Call c6kOps.RunPass("Stop")
        
        'Set status message
        Call statusMsg("Paused", "")
            
        Exit Sub
        
    End If
    
    'If the switch is not on and the joystick was released, an error occurred
    MsgBox "Run Joystick Incorrectly Released"

End If

End Sub

Private Sub runState4() 'Return State - moves head back to start of pass, then backs away perpendicularly by set distance

Call statusMsg("Returned")

'*** Return to Start Position

'-- Check for Release / Toggle
If c6kOps.getJoyToggle And c6kOps.getJoyRelease Then
    'Revert to Pause State
    State = 1
    Exit Sub
    
End If

End Sub
