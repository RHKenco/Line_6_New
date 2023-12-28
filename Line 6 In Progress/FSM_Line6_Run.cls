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
'|          - 1 - Set Torch
'|                  - Enables Free Joystick to position Torch
'|          - 2 - Strike Arc
'|                  - Strikes Arc & Begins Oscillation
'|          - 3 - Form Puddle
'|                  - Enables Free Joystick mode to position puddle
'|          - 4 - Run Pass
'|                  - Runs Pass, Joystick in Run mode
'|          - 5 - Terminate
'|                  - Enables Free Joystick mode to end pass
'|          - 6 - Return
'|                  - Cuts Arc and returns to beginning of pass
'|          - 7 - Wait
'|                  - Waits for keypress to return to Set mode for next pass
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
    Case 5
        Call runState5
    Case 6
        Call runState6
    Case 7
        Call runState7
End Select

End Sub

Private Sub runState0()

End Sub

Private Sub runState1()

End Sub

Private Sub runState2()

End Sub

Private Sub runState3()

End Sub

Private Sub runState4()

End Sub

Private Sub runState5()

End Sub

Private Sub runState6()

End Sub

Private Sub runState7()

End Sub
