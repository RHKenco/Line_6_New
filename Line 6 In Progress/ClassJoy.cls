VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassJoy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'|====================================================================================================================|
'|
'|              ----- Joystick Class -----
'|
'|
'|
'|====================================================================================================================|



'-- Class Constants & Variables ---------------------------------------------------------------------------------------

'Drive States
Private Const driveJog As Integer = 1
Private Const driveMcRun As Integer = 2
Private Const driveMcJogLR As Integer = 3
Private Const driveMcJogFB As Integer = 4
Private Const driveMcJogUD As Integer = 5
Private Const driveBoostLR As Integer = 6
Private Const driveBoostFB As Integer = 7
Private Const driveBoostUD As Integer = 8

'Drive Axes
Private Const aX As Integer = 1
Private Const aXstr As String = "X"
Private Const aXb As Integer = 2
Private Const aXbstr As String = "Xb"

Private Const aY As Integer = 3
Private Const aYstr As String = "Y"
Private Const aYb As Integer = 4
Private Const aYbstr As String = "Yb"

Private Const aZ As Integer = 5
Private Const aZstr As String = "Z"
Private Const aZb As Integer = 6
Private Const aZbstr As String = "Zb"

Private Const aZa As Integer = 7
Private Const aZastr As String = "Za"
Private Const aZab As Integer = 8
Private Const aZabstr As String = "Zab"

Private Const aR As Integer = 9
Private Const aRstr As String = "R"
Private Const aRb As Integer = 10
Private Const aRbstr As String = "Rb"

Private Const aYR As Integer = 11
Private Const aYRstr As String = "YR"
Private Const aYRb As Integer = 12
Private Const aYRbstr As String = "YRb"

Private Const aZp As Integer = 13
Private Const aZpstr As String = "Z'"
Private Const aZpb As Integer = 14
Private Const aZpbstr As String = "Z'b"

Private Const aZpR As Integer = 15
Private Const aZpRstr As String = "Z'R"
Private Const aZpRb As Integer = 16
Private Const aZpRbstr As String = "Z'Rb"

Private Const arS As Integer = 18
Private Const arSstr As String = "Run Speed"

'Joystick Jog Velocity
Private Const cJoyJogVel As Single = 5#
Private Const cJoyJogVelStr As String = "5,5,5,5,5,5"

'Internal Variables & Defaults
Private Const inDebCount As Integer = 5     'Time in FSM cycles of input debounce

Private userSpeedMult As Single             'Speed Multiplier to be applied to speed for varying pass speeds on the fly
Private Const speedMultInc As Single = 0.03           'Constant for incrementing the speedmult increase
Private Const speedMultUpperBound As Single = 2.5     ' +/- Speed Mult Allowance
Private Const speedMultLowerBound As Single = 0.4
Private userSpeedBoost(6) As Integer        'Speed Boost Int for each axis - joystick adjustment of multi-axis movements
Private Const userSpeedBoostMod = 0.25      'Constant for speed boost for multi-axis joystick operations


'Read-Only Variables
Private myJoy As joystickState

'Public Variables



'-- Class Commands ----------------------------------------------------------------------------------------------------
Public Function createJoystick() As Boolean

'Set up joy binary array from constants
myJoy.inBin(joyLeft) = bJoyLeft
myJoy.inBin(joyRight) = bJoyRight
myJoy.inBin(joyFront) = bJoyFront
myJoy.inBin(joyBack) = bJoyBack
myJoy.inBin(joyUp) = bJoyUp
myJoy.inBin(joyDown) = bJoyDown
myJoy.inBin(joySelect) = bJoySelect
myJoy.inBin(joyRelease) = bJoyRelease
myJoy.inBin(joyToggle) = bJoyToggle

Call resetSpeedMult

'Initialize joystick jog parameters
Call initJoy

End Function


'-- Joystick Input Commands -------------------------------------------------------------------------------------------


'-- Simple Read-Only Return Functions:
Public Function getJoyActive() As Boolean
    getJoyActive = myJoy.active
End Function

Public Function getJoyStr() As String
    getJoyStr = myJoy.stateStr
End Function

Private Function getStr(axisIn As Integer) As String

Select Case axisIn
    Case aX
        getStr = aXstr
    Case aXb
        getStr = aXbstr

    Case aY
        getStr = aYstr
    Case aYb
        getStr = aYbstr

    Case aZ
        getStr = aZstr
    Case aZb
        getStr = aZbstr

    Case aZa
        getStr = aZastr
    Case aZab
        getStr = aZabstr

    Case aR
        getStr = aRstr
    Case aRb
        getStr = aRbstr

    Case aYR
        getStr = aYRstr
    Case aYRb
        getStr = aYRbstr

    Case aZp
        getStr = aZpstr
    Case aZpb
        getStr = aZpbstr

    Case aZpR
        getStr = aZpRstr
    Case aZpRb
        getStr = aZpRbstr

    Case arS
        getStr = arSstr
    
    Case Else
        MsgBox "Error: " & axisIn & " does not correspond with a valid axis"
End Select

End Function


'-- getJoyInputs - Function to update joystick inputs
'       - Inputs:
'           - c6kInput is the Long containing the Input Status from the c6k
'           - returnIndex is the index corresponding to the return input plus one (1-11 instead of 0-10)
'       - Returns:
'           - Bool corresponding to the joystick input corresponding to returnIndex. Returns False by default
Public Function getJoyInputs(c6kInput As Long, Optional returnIndex As Integer) As Boolean

    'Clear prior update flags
    myJoy.stateJoyUpd = False
    myJoy.stateBtnUpd = False

    'For each input,
    Dim i As Integer
    Dim tempIn As Boolean
    For i = 0 To 8
        
        'Store the current input value
        tempIn = (c6kInput And myJoy.inBin(i))
        
        'Set Normally Open / Normally Closed
        Select Case i
            Case joyLeft, joyRight, joyFront, joyBack, joyUp, joyDown
                'Do nothing to Normally Open inputs
            Case joySelect, joyRelease, joyToggle
                tempIn = Not tempIn     'Normally Closed inputs reversed
        End Select
        
        
        'See if the current input value is different from the previous one
        If tempIn <> myJoy.stateNow(i) Then
            'If it is, Store last input value
            myJoy.stateLast(i) = myJoy.stateNow(i)
            
            Dim isUpdated As Boolean
            isUpdated = False
            
            'If the input is set in the 6k and input debounce is not still counting, then set the state active.
            If tempIn Then
                If Not (myJoy.inDeb(i) > 0) Then
                    myJoy.stateNow(i) = True
                    myJoy.inDeb(i) = inDebCount
                    isUpdated = True
                Else
                    ' If the code reaches this point, the input changed but debounce is preventing it from reading true again
                    myJoy.stateNow(i) = False
                End If
            Else
                myJoy.stateNow(i) = False
                isUpdated = True
            End If
            
            If isUpdated = True And i < 6 Then myJoy.stateJoyUpd = True Else myJoy.stateJoyUpd = False
            If isUpdated = True And i > 5 Then myJoy.stateBtnUpd = True Else myJoy.stateBtnUpd = False
            
        End If
        
        'Decrement inputDebounce when necessary
        If myJoy.inDeb(i) > 0 Then myJoy.inDeb(i) = myJoy.inDeb(i) - 1
        
    Next i
    
    ' If a return index is supplied, return the corresponding value
    If returnIndex = 0 Then
        getJoyInputs = False
    ElseIf returnIndex >= 0 And returnIndex < 11 Then
        getJoyInputs = myJoy.stateNow(returnIndex - 1)
    Else
        MsgBox "Error in getJoyInputs: " & returnIndex & " is not a valid input"
    End If
    
End Function

Public Function getJoyButtons(Index As Integer) As Boolean
    If Index < 9 Then
        getJoyButtons = myJoy.stateNow(Index)
    Else
        MsgBox "Invalid Button Index Request"
    End If
End Function

Public Function getJoySelect() As Boolean
    getJoySelect = myJoy.stateNow(joySelect)
End Function

Public Function getJoyRelease() As Boolean
    getJoyRelease = myJoy.stateNow(joyRelease)
End Function

Public Function getJoyToggle() As Boolean
    getJoyToggle = myJoy.stateNow(joyToggle)
End Function

'-- initJoy - Function to set up basic joystick parameters. Called when initializing joystick
Private Sub initJoy()

'Immediately (!) disable jog mode for all axes
c6k.Write ("!JOG000000" & Chr$(13))

'Set motion parameters for Joystick Motion
c6k.Write ("JOG000000:JOGA4,4,4,4,144,4:JOGAD10,10,10,10,360,10:JOGVH10,10,10,10,360,10:JOGVL5,5,5,5,36,5:V10,10,10,10,360,10" & Chr$(13))

End Sub

'-- stopJoy - Function to shut down joystick
'       - Inputs:
'           - currentJogAxes - String containing the current jog axes that need to be shut down - Format 1XXXXX where 1 indicates enabled axis
'           - currentMcAxes - As above, for continuous motion axes
Private Sub stopJoy(currentJogAxes As String, Optional currentMcAxes As String)

'Disable jog mode on currently active drives
c6k.Write ("!JOG" & currentJogAxes & Chr$(13))

'If MC on any drives, disable
If currentMcAxes <> "" Then c6k.Write ("!MC" & currentMcAxes & Chr$(13))

'Disable inputs
c6k.Write ("1INFNC1-A:1INFNC2-A:1INFNC3-A:1INFNC4-A:1INFNC5-A:1INFNC6-A" & Chr$(13))

'Lower active flag and clear joystick string
myJoy.active = False
myJoy.stateStr = ""

End Sub

Public Sub eStopJoy()
'Stop all axes, Disable all jog, disable all input functions
c6k.Write ("!S:!@JOG0:1INFNC1-A:1INFNC2-A:1INFNC3-A:1INFNC4-A:1INFNC5-A:1INFNC6-A" & Chr$(13))

'Lower active flag and clear joystick string
myJoy.active = False
myJoy.stateStr = ""

End Sub

'-- runJoy() - main function call for joystick mode - calls appropriate functions for called mode & returns boolean status
'       - Inputs:
'           - "Free" - Sets joystick to Free mode
'           - "Pause" - Set joystick to Pause mode
'           - "Run" - Run mode
'           - "Disable" - Disables joystick inputs
Public Function runJoy(mode As Integer) As Boolean
         
'Update the joystick inputs. Calling the getInputState function will update FastStatus
Call getJoyInputs(c6kOps.getInputState)

'Set up end flag & set to default of false
Dim endJoy As Boolean
endJoy = False

'If mode is disable, shut down the joystick. Else,
If mode = joyDisable Then
    Call stopJoy(myJoy.joyJogStr, myJoy.joyMcStr)
    runJoy = False
    Exit Function
Else
    'Set the mode of the joystick. This will pair the correct machine axes with the corresponding joystick axes for the mode requested
    endJoy = Not joyModeSet(mode)
    'When the joystick is released, immediately disable the joystick
    If endJoy Then
        runJoy = False
        Exit Function
    End If
End If

'Run continuous motion mode - does nothing unless MC joy mode is enabled on an axis
Call runJoyInputs

'Return True if joy mode is still active
runJoy = True

End Function



'-- joyModeSet() - Sets joystick mode -------------------------------------------------------------------------------------------
'       - Inputs:
'           - runMode - Run mode of the pass - see runJoy for details (same input)
Private Function joyModeSet(runMode As Integer) As Boolean

If myJoy.stateBtnUpd Or Not myJoy.active Then

    '-- Develop buttons integer - see below
    Dim buttons As Integer
    If myJoy.stateNow(joySelect) Then buttons = buttons Or (2 ^ (1 - 1))
    If myJoy.stateNow(joyRelease) Then buttons = buttons Or (2 ^ (2 - 1))
    If myJoy.stateNow(joyToggle) Then buttons = buttons Or (2 ^ (3 - 1))
    
    'Single integer to describe the state of the buttons on the joystick. The integer is set as below:
        '        Sel  Rel  Tog
        '   0  -  0    0    0
        '   1  -  1    0    0
        '   2  -  0    1    0
        '   3  -  1    1    0
        '   4  -  0    0    1
        '   5  -  1    0    1
        '   6  -  0    1    1
        '   7  -  1    1    1
        
    Dim passType As Integer
    passType = c6kOps.getPassType()
        
    '-- Set joystick mode based on run mode - Free, Pause, or Run
    Select Case runMode
        Case joyFree         'Standard X-Y-Z Mode. Used for Blades & Positioning
        
            Select Case buttons
                Case 0, 4  ' No Keys Pressed or Toggle Only
                    Select Case passType
                        Case passBlade
                            Call setJoyInputs(aX, aY, aZ)
                        Case passAugEdg
                            Call setJoyInputs(aYR, aX, aZ) 'Axis 2 should be aY, changed for test
                        Case passAugFac
                            Call setJoyInputs(aX, aZpR, aZp)
                    End Select
                    
                Case 1, 5 ' Select Key Pressed - Machine Coordinate Mode
                    Select Case passType
                        Case passBlade
                            Call setJoyInputs(aX, aY, aZa)
                        Case passAugEdg
                            Call setJoyInputs(aX, aY, aZa)
                        Case passAugFac
                            Call setJoyInputs(aX, aY, aZa)
                    End Select
                
                Case 2, 3, 6, 7 ' Release pressed in any combination
                    
                    GoTo ReleaseJoystick
                    
                Case Else
                    
                    GoTo joyErrorHand
                
            End Select
            
        Case joyPause
            Select Case buttons     'Toggle must be flipped for this case to occur, therefore only case 4-7
                Case 4      'Toggle Only
                    Select Case passType
                        Case passBlade
                            Call setJoyInputs(aX, aY, aZ)
                        Case passAugEdg
                            Call setJoyInputs(aYR, aY, aZ)
                        Case passAugFac
                            Call setJoyInputs(aZpR, aX, aZp)
                    End Select
                Case 5      'Select Pressed
                    Select Case passType
                        Case passBlade
                            Call setJoyInputs(aX, aY, aZa)
                        Case passAugEdg, passAugFac
                            Call setJoyInputs(aY, aR, aZa)
                    End Select
                Case 6, 7  ' Release Pressed - Advance Program
                    GoTo ReleaseJoystick
                Case 0, 1, 2, 3     'Toggle not flipped - return to run mode
                    GoTo ReleaseJoystick
                Case Else
                    GoTo joyErrorHand
            End Select
        
        Case joyRun
            Select Case buttons
                Case 0    'No buttons Pressed - Horizontal axis will adjust run speed
                    Select Case passType
                        Case passBlade, passAugEdg
                            Call setJoyInputs(arS, aY, aZ)
                        Case passAugFac
                            Call setJoyInputs(arS, aX, aZ)
                    End Select
                Case 1      'Select Pressed - Horizontal axis will nudge run axis
                    Select Case passType
                        Case passBlade
                            Call setJoyInputs(aXb, aY, aZa)
                        Case passAugEdg
                            Call setJoyInputs(aYRb, aX, aZa)
                        Case passAugFac
                            Call setJoyInputs(aZpRb, aY, aZpb)
                    End Select
    
                Case 4, 5, 6, 7 'Toggle flipped
                    GoTo ReleaseJoystick
                Case 2, 3    'Release Pressed, no toggle - Ignore
                    
                Case Else
                    GoTo joyErrorHand
            End Select
        
        
        
        Case Else
            GoTo joyErrorHand
    End Select
        
End If
    
'Call jogOn
joyModeSet = True
myJoy.active = True

Exit Function

joyErrorHand:

MsgBox "Error in joyModeFree Handler:" & Chr(13) & "Mode " & passType & ", Case: " & buttons

ReleaseJoystick:
    'Return False to indicate that user input should end joystick mode
    Call jogOff
    joyModeSet = False
    myJoy.active = False


End Function


Private Sub setJoyInputs(joyLR As Integer, joyFB As Integer, joyUD As Integer)

'Ensure no overlap
If joyLR = joyFB Or joyLR = joyUD Or joyFB = joyUD Then
    MsgBox "Cannot assign multiple joystick directions to the same axis!"
    Exit Sub
End If

Call jogOff

'Set joyString
myJoy.stateStr = getStr(joyLR) & " - " & getStr(joyFB) & " - " & getStr(joyUD)

'Disable inputs to prevent carry-over commands
c6k.Write ("1INFNC1-A:1INFNC2-A:1INFNC3-A:1INFNC4-A:1INFNC5-A:1INFNC6-A")

'Clear outSt array to ensure no carry-over
Dim i As Integer
For i = 1 To 6
    myJoy.outSt(i) = 0
Next i

Dim tempJoyAxis As Integer 'Temp variable containing value of axis
Dim tempJoyInputA As String 'Temp variable containing input value for Jog
Dim tempJoyInputB As String 'As Above
Dim tempMcJogJoy As Integer
Dim tempMcBstJoy As Integer 'MC Boost joyAxis


'For Each Axis
For i = 0 To 2

    '-- Set temps for axis
    Select Case i
        Case 0
            tempJoyAxis = joyLR
            tempJoyInputA = "1"
            tempJoyInputB = "2"
            tempMcJogJoy = driveMcJogLR
            tempMcBstJoy = driveBoostLR
        Case 1
            tempJoyAxis = joyFB
            tempJoyInputA = "3"
            tempJoyInputB = "4"
            tempMcJogJoy = driveMcJogFB
            tempMcBstJoy = driveBoostFB
        Case 2
            tempJoyAxis = joyUD
            tempJoyInputA = "5"
            tempJoyInputB = "6"
            tempMcJogJoy = driveMcJogUD
            tempMcBstJoy = driveBoostUD
    End Select
            
    '-- Determine axis and set accordingly
    Select Case tempJoyAxis
        '-- Run Speed Adjustment ----------------------------------------------------------------------------------rS
        Case arS
            myJoy.inSt(i) = 3
        
        '-- X - Axis -----------------------------------------------------------------------------------------------X
        Case aX
            c6k.Write "1INFNC" & tempJoyInputA & "-" & driveX & "J:1INFNC" & tempJoyInputB & "-" & driveX & "K" & Chr$(13)    'Set jog inputs in 6k
            myJoy.outSt(driveX) = driveJog      'Set axis in Joystick object to drive mode
            myJoy.inSt(i) = driveJog            'Set joystick direction in joystick object to drive mode
        Case aXb
            myJoy.outSt(driveX) = tempMcBstJoy
            myJoy.inSt(i) = tempMcBstJoy
    
        '-- Y - Axis -----------------------------------------------------------------------------------------------Y
        Case aY
            c6k.Write "1INFNC" & tempJoyInputA & "-" & driveY & "K:1INFNC" & tempJoyInputB & "-" & driveY & "J" & Chr$(13)    'Set jog inputs in 6k
            myJoy.outSt(driveY) = driveJog      'Set axis in Joystick object to drive mode
            myJoy.inSt(i) = driveJog            'Set joystick direction in joystick object to drive mode
        Case aYb
            myJoy.outSt(driveY) = tempMcBstJoy
            myJoy.inSt(i) = tempMcBstJoy

        '-- Z - Axis Fuctions --------------------------------------------------------------------------------------Z
        Case aZ
            c6k.Write "1INFNC" & tempJoyInputA & "-" & driveZ & "K:1INFNC" & tempJoyInputB & "-" & driveZ & "J" & Chr$(13)    'Set jog inputs in 6k
            myJoy.outSt(driveZ) = driveJog      'Set axis in Joystick object to drive mode
            myJoy.inSt(i) = driveJog            'Set joystick direction in joystick object to drive mode
        Case aZb
            myJoy.outSt(driveZ) = tempMcBstJoy
            myJoy.inSt(i) = tempMcBstJoy
    
        '-- Za - Axis Fuctions ------------------------------------------------------------------------------------Za
        Case aZa
            c6k.Write "1INFNC" & tempJoyInputA & "-" & driveZa & "K:1INFNC" & tempJoyInputB & "-" & driveZa & "J" & Chr$(13)    'Set jog inputs in 6k
            myJoy.outSt(driveZa) = driveJog      'Set axis in Joystick object to drive mode
            myJoy.inSt(i) = driveJog            'Set joystick direction in joystick object to drive mode
        Case aZab
            myJoy.outSt(driveZa) = tempMcBstJoy
            myJoy.inSt(i) = tempMcBstJoy
    
        '-- R - Axis Fuctions --------------------------------------------------------------------------------------R
        Case aR
            c6k.Write "1INFNC" & tempJoyInputA & "-" & driveR & "K:1INFNC" & tempJoyInputB & "-" & driveR & "J" & Chr$(13)    'Set jog inputs in 6k
            myJoy.outSt(driveR) = driveJog      'Set axis in Joystick object to drive mode
            myJoy.inSt(i) = driveJog            'Set joystick direction in joystick object to drive mode
        Case aRb
            myJoy.outSt(driveR) = tempMcBstJoy
            myJoy.inSt(i) = tempMcBstJoy
    
    
        '-- YR - Axis Fuctions ------------------------------------------------------------------------------------YR
        Case aYR
            myJoy.outSt(driveR) = tempMcJogJoy    'Set axes in Joystick object to continuous jog mode
            myJoy.outSt(driveY) = tempMcJogJoy
            myJoy.inSt(i) = tempMcJogJoy          'Set joystick direction in joystick object to drive mode
        Case aYRb
            myJoy.outSt(driveR) = tempMcBstJoy
            myJoy.outSt(driveY) = tempMcBstJoy
            myJoy.inSt(i) = tempMcBstJoy
    
        '-- Z' - Axis Fuctions ------------------------------------------------------------------------------------Z'
        Case aZp
            myJoy.outSt(driveY) = tempMcJogJoy    'Set axes in Joystick object to continuous jog mode
            myJoy.outSt(driveZ) = tempMcJogJoy
            myJoy.inSt(i) = tempMcJogJoy          'Set joystick direction in joystick object to drive mode
        Case aZpb
            myJoy.outSt(driveY) = tempMcBstJoy
            myJoy.outSt(driveZ) = tempMcBstJoy
            myJoy.inSt(i) = tempMcBstJoy

        '-- Z'R - Axis Fuctions ----------------------------------------------------------------------------------Z'R
        Case aZpR
            myJoy.outSt(driveY) = tempMcJogJoy    'Set axes in Joystick object to continuous jog mode
            myJoy.outSt(driveZ) = tempMcJogJoy
            myJoy.outSt(driveR) = tempMcJogJoy
            myJoy.inSt(i) = tempMcJogJoy          'Set joystick direction in joystick object to drive mode
        Case aZpRb
            myJoy.outSt(driveY) = tempMcBstJoy
            myJoy.outSt(driveZ) = tempMcBstJoy
            myJoy.outSt(driveR) = tempMcBstJoy
            myJoy.inSt(i) = tempMcBstJoy
    End Select
    
Next i

Call jogOn

End Sub

Private Sub jogOn()

'-- Set c6k Jog State if necessary
If myJoy.inSt(0) = driveJog Or myJoy.inSt(1) = driveJog Or myJoy.inSt(2) = driveJog Then
    Dim tempJog As String
    Dim tempErr As Integer
    Dim i As Integer
    'For all values in joyJogState (except index 0, which is unused)
    For i = 1 To 6
        If i <> 4 Then  'For all i except 4 (Osc)
        
            Select Case myJoy.outSt(i)
                Case 0, driveMcJogLR, driveMcJogFB, driveMcJogUD, driveMcRun, driveBoostLR, driveBoostFB, driveBoostUD         ' If jogState is not 1
                    'First loop through, set initial variable value, else append character
                    If i = 1 Then
                        tempJog = "X"
                    Else
                        tempJog = tempJog & "X"
                    End If
                    
                Case driveJog         'If drive is in jog mode, make it active
                    If i = 1 Then
                        tempJog = "1"
                    Else
                        tempJog = tempJog & "1"
                    End If
                    
                Case Else
                    tempErr = 1
                    GoTo JogError
            End Select
        Else    ' If i = 4 (Oscillator Axis)
            tempJog = tempJog & "X"
        End If
    Next i
    
    'Disable Previous Jog mode
    If myJoy.joyJogStr <> "" Then c6k.Write ("JOG" & Replace(myJoy.joyJogStr, "1", "0") & Chr(13))
    'Enable Current Jog Mode
    c6k.Write ("JOG" & tempJog & Chr(13))

    myJoy.joyJogStr = tempJog
End If

'-- Set c6k Mc State if necessary
Dim Check1 As Boolean
For i = 0 To 2
    If myJoy.inSt(i) = driveMcJogLR Or myJoy.inSt(i) = driveMcJogFB Or myJoy.inSt(i) = driveMcJogUD Then Check1 = True
Next i

If Check1 Then
    
    If myJoy.joyMcStr = "" And fsmRun.State <> 2 And fsmRun.State <> 3 Then c6k.Write "COMEXC1" & Chr(13)   'Enable cont. Com ex when enabling MC mode
    
    Dim tempMc As String
    Dim tempMcVel As String
    'For all values in joyJogState (except index 0, which is unused)
    For i = 1 To 6
        If i <> 4 Then  'For all i except 4 (Osc)
        
            Select Case myJoy.outSt(i)
                Case 0, driveJog, driveMcRun, driveBoostLR, driveBoostFB, driveBoostUD         ' If jogState is not 1
                    'First loop through, set initial variable value, else append character
                    If i = 1 Then
                        tempMc = "X"
                        tempMcVel = ""
                    Else
                        tempMc = tempMc & "X"
                        tempMcVel = tempMcVel & ","
                    End If
                    
                Case driveMcJogLR, driveMcJogFB, driveMcJogUD          'If drive is in jog mode, make it active
                    If i = 1 Then
                        tempMc = "1"
                        tempMcVel = "*,"
                    Else
                        tempMc = tempMc & "1"
                        tempMcVel = tempMcVel & ",*"
                    End If
                    
                Case Else
                    tempErr = 2
                    GoTo JogError
            End Select
        Else    ' If i = 4 (Oscillator Axis)
            tempMc = tempMc & "X"
            tempMcVel = tempMcVel & ","
        End If
    Next i
    
    'Disable Previous Jog Jog mode
    If myJoy.joyMcStr <> "" Then c6k.Write ("MC" & Replace(myJoy.joyMcStr, "1", "0") & Chr(13))
    'Enable Current MC Jog Mode
    c6k.Write ("MC" & tempMc & Chr(13))
    'Set velocities of MC axes to 0
    c6k.Write "V" & Replace(tempMcVel, "*", "0") & Chr(13)
    
    'Save strings
    myJoy.joyMcStr = tempMc
    myJoy.joyMcDelStr = tempMcVel
    
'If MC has previously been set, but is no longer set, Disable MC and clear strings
ElseIf myJoy.joyMcStr <> "" Then
    c6k.Write ("MC" & Replace(myJoy.joyMcStr, "1", "0") & Chr(13))
    c6k.Write ("V" & Replace(myJoy.joyMcDelStr, "*", "0") & Chr(13))
    myJoy.joyMcStr = ""
    myJoy.joyMcDelStr = ""
    If fsmRun.State <> 2 And fsmRun.State <> 3 Then c6k.Write "COMEXC0" & Chr(13)   'Disable con com ex first loop after MC disabled
End If

Exit Sub

'********* ERROR HANDLER ****************************************************
JogError:

Dim ErrStr As String

'Build Error Message
ErrStr = "Error in Jog Enable Routine:" & Chr(13)

If tempErr = 1 Then ErrStr = ErrStr & "Jog" Else ErrStr = ErrStr & "Mc Jog"
ErrStr = ErrStr & " failed at Index " & i & " with Case " & myJoy.outSt(i)

'Display Error Msg in msgbox
MsgBox (ErrStr)

'Reset Program
Call c6kOps.progError(ErrStr)

End Sub

Private Sub jogOff()

If myJoy.joyJogStr <> "" Then
    c6k.Write ("JOG" & Replace(myJoy.joyJogStr, "1", "0") & Chr(13))
End If

If myJoy.joyMcStr <> "" Then
    c6k.Write ("S" & myJoy.joyMcStr & Chr(13))
    c6k.Write ("MC" & Replace(myJoy.joyMcStr, "1", "0") & Chr(13))
End If

End Sub


Private Sub runJoyInputs()

Dim i As Integer
Dim runMc As Boolean
Dim runSc As Boolean
Dim runBst As Boolean


'-- If no joystick inputs require continuous motion, exit
If myJoy.inSt(0) = driveJog And myJoy.inSt(1) = driveJog And myJoy.inSt(2) = driveJog Then
        Exit Sub
End If

'-- Determine what other functions need to run
For i = 0 To 2

    If myJoy.inSt(i) = driveMcRun Then runSc = True
    If myJoy.inSt(i) = driveMcJogLR Or myJoy.inSt(i) = driveMcJogFB Or myJoy.inSt(i) = driveMcJogUD Then runMc = True
    If myJoy.inSt(i) = driveBoostLR Or myJoy.inSt(i) = driveBoostFB Or myJoy.inSt(i) = driveBoostUD Then runBst = True

Next i

'-- If any of the axes are set to speed control, call for speed to be adjusted
If runSc Then
    Dim changeSpeed As Integer

    For i = 0 To 2  'For each joystick direction
        If myJoy.inSt(i) = driveMcRun Then
            Select Case i
                Case 0
                    If myJoy.stateNow(joyRight) Then
                        changeSpeed = 1
                    ElseIf myJoy.stateNow(joyLeft) Then
                        changeSpeed = 2
                    End If
                Case 1
                    If myJoy.stateNow(joyFront) Then
                        changeSpeed = 1
                    ElseIf myJoy.stateNow(joyBack) Then
                        changeSpeed = 2
                    End If
                Case 2
                    If myJoy.stateNow(joyUp) Then
                        changeSpeed = 1
                    ElseIf myJoy.stateNow(joyDown) Then
                        changeSpeed = 2
                    End If
            End Select
        End If
    Next i
    
    'Adjust Speed Accordingly, respecting speed bounds
    If changeSpeed = 1 Then
        If userSpeedMult <= speedMultUpperBound Then userSpeedMult = userSpeedMult + speedMultInc
    ElseIf changeSpeed = 2 Then
        If userSpeedMult >= speedMultLowerBound Then userSpeedMult = userSpeedMult - speedMultInc
    End If
    
End If

'-- If any of the axes require Continuous Motion outside of non-run axes
If runMc Then

    'Clear Axis Velocities
    For j = 1 To 6
        myJoy.joyMcAxisVel(j) = 0
    Next j

    'For each joysick axis
    For i = 0 To 2
        'If the axis is in mcJog mode
        If myJoy.inSt(i) = driveMcJogLR Or myJoy.inSt(i) = driveMcJogFB Or myJoy.inSt(i) = driveMcJogUD Then
        
            Dim mcStart As Integer
            
            'Determine what direction the joystick is in (1- Neg, 0- Centered, 2- Pos)
            Select Case i
                Case 0
                    If myJoy.stateNow(joyLeft) Then
                        mcStart = 1
                    ElseIf myJoy.stateNow(joyRight) Then
                        mcStart = 2
                    End If
                Case 1
                    If myJoy.stateNow(joyFront) Then
                        mcStart = 1
                    ElseIf myJoy.stateNow(joyBack) Then
                        mcStart = 2
                    End If
                Case 2
                    If myJoy.stateNow(joyUp) Then
                        mcStart = 1
                    ElseIf myJoy.stateNow(joyDown) Then
                        mcStart = 2
                    End If
            End Select
            
            'Check for Direction Change
            For j = 1 To 6
                'If any axis assigned to the current joystick axis is moving, then
                If (myJoy.outSt(j) = myJoy.inSt(i)) And (fsinfo.AxisStatus(j) And 1) Then
                    'If motion is + and command is - or motion is - and command is + then there is a mid-motion direction change
                    If ((fsinfo.AxisStatus(j) And 2) And mcStart <> 2) Or (Not (fsinfo.AxisStatus(j) And 2) And mcStart <> 1) Then mcStart = 0
                End If
            Next j
            
            'Set Axis Velocities
            For j = 1 To 6
                'With the exception of the Osc. Axis
                If j <> 4 Then
                    'For each drive axis, if the drive output state is equal to the joy input state the drive is assigned to the current axis, Therefore:
                    If myJoy.outSt(j) = myJoy.inSt(i) Then
                        'The velocity of the drive axis should be set based on the joystick position
                        Select Case mcStart
                            Case 1
                                myJoy.joyMcAxisVel(j) = myJoy.joyMcAxisVel(j) + cJoyJogVel
                            Case 2
                                myJoy.joyMcAxisVel(j) = myJoy.joyMcAxisVel(j) - cJoyJogVel
                        End Select
                    End If
                End If
            Next j

        End If
        
    Next i
        
    Dim mcVelocityStr As String
    Dim mcDirectionStr As String
        
    'Build Velocity & Direction Strings
    For i = 1 To 6
        If i = 1 Then
            If myJoy.joyMcAxisVel(i) > 0 Then
                mcDirectionStr = "+"
            ElseIf myJoy.joyMcAxisVel(i) < 0 Then
                mcDirectionStr = "-"
            Else
                mcDirectionStr = ""
            End If
            
            mcVelocityStr = CStr(Abs(myJoy.joyMcAxisVel(i)))
        Else
            If myJoy.joyMcAxisVel(i) > 0 Then
                mcDirectionStr = mcDirectionStr & ",+"
            ElseIf myJoy.joyMcAxisVel(i) < 0 Then
                mcDirectionStr = mcDirectionStr & ",-"
            Else
                mcDirectionStr = mcDirectionStr & ","
            End If
            
            mcVelocityStr = mcVelocityStr & "," & CStr(Abs(myJoy.joyMcAxisVel(i)))
        End If
    Next i
    
    'Update MC Joy Velocities
    c6k.Write "D" & mcDirectionStr & ":V" & mcVelocityStr & ":GO" & myJoy.joyMcStr & Chr(13)

End If

'-- Clear boost array
For i = 1 To 6
    userSpeedBoost(i) = 0
Next i

'-- If any of the axes require Continuous Motion outside of non-run axes
If runBst Then
    
    Dim tempBst As Integer
    
    For i = 0 To 2
            
        Select Case i
            Case 0
                'Determine what direction to boost
                If myJoy.inSt(i) = driveBoostLR Then
                    If myJoy.stateNow(joyRight) Then
                        tempBst = 1
                    ElseIf myJoy.stateNow(joyLeft) Then
                        tempBst = 2
                    End If
                End If
                'Determine which axes to boost
                Dim A As Integer
                For A = 1 To 6
                    If myJoy.outSt(A) = driveBoostLR Then userSpeedBoost(A) = tempBst
                Next A
                
            Case 1
                'Determine what direction to boost
                If myJoy.inSt(i) = driveBoostFB Then
                    If myJoy.stateNow(joyFront) Then
                        tempBst = 1
                    ElseIf myJoy.stateNow(joyBack) Then
                        tempBst = 2
                    End If
                End If
                'Determine which axes to boost
                For A = 1 To 6
                    If myJoy.outSt(A) = driveBoostFB Then userSpeedBoost(A) = tempBst
                Next A
                
            Case 2
                'Determine what direction to boost
                If myJoy.inSt(i) = driveBoostUD Then
                    If myJoy.stateNow(joyUp) Then
                        tempBst = 1
                    ElseIf myJoy.stateNow(joyDown) Then
                        tempBst = 2
                    End If
                End If
                'Determine which axes to boost
                For A = 1 To 6
                    If myJoy.outSt(A) = driveBoostUD Then userSpeedBoost(A) = tempBst
                Next A
                
        End Select
        
    Next i

End If


End Sub

Public Function getUserSpeedMult() As Single
    getUserSpeedMult = userSpeedMult
End Function

Public Sub resetSpeedMult()
    userSpeedMult = 1
End Sub

Public Function getUserSpeedBoost(drive As Integer) As Integer
    getUserSpeedBoost = userSpeedBoost(drive)
End Function

Public Function getUserSpeedBoostMod() As Single
    getUserSpeedBoostMod = userSpeedBoostMod
End Function
