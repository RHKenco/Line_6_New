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
Call c6kOps.homeOsc

'Start Water Pump
Call c6kOps.setOutput(outH2O, True)
    
'Initialize Joystick
Call Joy.runJoy(joyFree)
    
'Initialize internal variables
stateInner = 0

'Proceed to next state
State = 1
    
End Sub

Private Sub runState1() 'Set Torch & Strike Arc State - Set torch for initial pass, then strike arc

Select Case stateInner
    Case 0  'Verify that the pause switch is thrown
        If Joy.getJoyToggle Then
            stateInner = 1
        Else
            'Set status message
            Call statusMsg(msgToggle)
            Call Joy.getJoyInputs(c6kOps.getInputState)
        End If
    Case 1  'Run Joystick to set strike location
        'Set status message
        Call statusMsg(msgStrike)
        'Wait until joystick is inactive
        If Joy.runJoy(joyFree) Then Exit Sub
        stateInner = 2
    Case 2  'Strike Arc
        Call c6kOps.setMachineHome
        Call c6kOps.strikeArc
        c6k.Write "COMEXC1" & Chr(13)    'Enable continuous Command Processing
        stateInner = 3
    Case 3  'Proceed to pause state
        'Set status message
        Call statusMsg(msgPaused)
        State = 2
        stateInner = 0
    
    Case Else
        MsgBox "Error in Pass Initialization"
End Select

End Sub

Private Sub runState2() 'Pause State - Joystick movement & Oscillator Only

'-- Run Joystick
If Not Joy.runJoy(joyPause) Then      'If the joystick is released, figure out why

    '-- Check for Inputs
    If Not Joy.getJoyToggle Then 'If switch is not on
        
        'Revert to Run state
        State = 3
        Call c6kOps.RunPass("Start")
        
        'Set status message
        Call statusMsg(msgRunning)
        
        'On first run, start carbide
        If stateInner = 0 Then
            Call c6kOps.setOutput(outTcFeed, True)
            stateInner = 1
        End If
        
        GoTo runOsc2
    End If
    
    'If the switch is on, then release was pressed - Switch to Return state
    State = 4
    Call c6kOps.RunPass("Stop")
    Call c6kOps.setOutput(outTcFeed, False)
    stateInner = 0
    Exit Sub
End If

runOsc2:

'Run Oscillator
Call c6kOps.runOsc


End Sub

Private Sub runState3() 'Run State - Joystick in pass adjustment mode, Oscillator running

'-- Run Pass
Call c6kOps.RunPass("Update")

'-- Run Joystick
If Not Joy.runJoy(joyRun) Then        'If the joystick is released, figure out why
    
    '-- Check for switch
    If Joy.getJoyToggle Then 'If switch is on
        
        'Revert to Pause State
        State = 2
        Call c6kOps.RunPass("Stop")
        
        'Set status message
        Call statusMsg(msgPaused)
            
        GoTo runOsc3
        
    End If
    
    'If the switch is not on and the joystick was released, an error occurred
    MsgBox "Run Joystick Incorrectly Released"
    
    Exit Sub

End If

runOsc3:

'Run Oscillator
Call c6kOps.runOsc

End Sub

Private Sub runState4() 'Return State - moves head back to start of pass, then backs away perpendicularly by set distance

Call c6kOps.returnOsc

Call Joy.getJoyInputs(c6kOps.getInputState)

Select Case stateInner
    Case 0
        'Update Status Message
        Call statusMsg(msgCompleted)
        
        'Clear joystick speed adjustments
        Call Joy.resetSpeedMult
        
        c6k.Write "COMEXC0" & Chr(13)    'Disable continuous Command Processing
        
        stateInner = 1
        
    Case 1
        'Wait for input
        If Not Joy.getJoyToggle Then
            Call statusMsg(msgReturning)
            stateInner = 2
        ElseIf Joy.getJoyRelease Then
            'Reset inner State & Set Main FSM state to pass complete
            stateInner = 0
            fsmMain.State = 4
        End If
    Case 2
        '*** Create code to send machine back to 0,0
        stateInner = 3
    Case 3
        Call statusMsg(msgNextPass)
        
        '*** Add timeout function here
        
        If Joy.getJoyToggle Then
            stateInner = 0
            State = 1
        End If
    
End Select

End Sub
