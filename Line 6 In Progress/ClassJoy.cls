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

'Motion state attached to either a joystick axis or a drive axis
Private Enum motionAttach
    driveJog = 1
    driveMcRun
    driveMcJog
End Enum

'Drive Axes
Private Enum motionAxes
    arS = (2 ^ (1 - 1))
    
    axisJog = (2 ^ (2 - 1))
    axisMc = (2 ^ (3 - 1))
    axisBst = (2 ^ (4 - 1))
    
    aX = (2 ^ (5 - 1))
    aXi = aX Or (2 ^ (6 - 1))
    aXj = axisJog Or aX
    aXc = axisMc Or aX
    aXci = axisMc Or aXi
    aXb = axisBst Or aX
    
    aY = (2 ^ (7 - 1))
    aYi = aY Or (2 ^ (8 - 1))
    aYj = axisJog Or aY
    aYc = axisMc Or aY
    aYci = axisMc Or aYi
    aYb = axisBst Or aY

    aZ = (2 ^ (9 - 1))
    aZi = aZ Or (2 ^ (10 - 1))
    aZj = axisJog Or aZ
    aZc = axisMc Or aZ
    aZci = axisMc Or aZi
    aZb = axisBst Or aZ

    aZa = (2 ^ (11 - 1))
    aZai = aZa Or (2 ^ (12 - 1))
    aZaj = axisJog Or aZa
    aZac = axisMc Or aZa
    aZaci = axisMc Or aZai
    aZab = axisBst Or aZa

    aR = (2 ^ (13 - 1))
    aRi = aR Or (2 ^ (14 - 1))
    aRj = axisJog Or aR
    aRc = axisMc Or aR
    aRci = axisMc Or aRi
    aRb = axisBst Or aR

    aYR = aYc Or aRc
    aYRi = aYc Or aRci
    aYRb = aYb Or aRb

    aZp = aYc Or aZc
    aZpb = aYb Or aZb

    aZpR = aYc Or aZc Or aRc
    aZpRi = aYc Or aZc Or aRci
    aZpRb = aYb Or aZb Or aRb
End Enum

Private Const arSstr = "Run Speed"
Private Const aXstr = "X"
Private Const aXbstr = "Xb"
Private Const aYstr = "Y"
Private Const aYbstr = "Yb"
Private Const aZstr = "Z"
Private Const aZbstr = "Zb"
Private Const aZastr = "Za"
Private Const aZabstr = "Zab"
Private Const aRstr = "R"
Private Const aRbstr = "Rb"
Private Const aYRstr = "YR"
Private Const aYRbstr = "YRb"
Private Const aZpstr = "Z'"
Private Const aZpbstr = "Z'b"
Private Const aZpRstr = "Z'R"
Private Const aZpRbstr = "Z'Rb"

'Joystick Jog Velocity
Private Const cJoyJogVel As Single = 1#
Private Const cJoyJogVelStr As String = "1,1,1,1,1,1"

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

Private Function getStr(axisIn As Long) As String

Select Case axisIn
    Case aXj, aXc
        getStr = aXstr
    Case aXb
        getStr = aXbstr

    Case aYj, aYc
        getStr = aYstr
    Case aYb
        getStr = aYbstr

    Case aZj, aZc
        getStr = aZstr
    Case aZb
        getStr = aZbstr

    Case aZaj, aZac
        getStr = aZastr
    Case aZab
        getStr = aZabstr

    Case aRj, aRc
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
                            Call initJoyAxes(aXj, False, aYj, False, aZj, False)
                        Case passAugEdg
                            Call initJoyAxes(aYR, False, aYc, False, aZj, False) 'Axis 2 should be aY, changed for test
                        Case passAugFac
                            Call initJoyAxes(aXj, False, aZpR, False, aZp, False)
                    End Select
                    
                Case 1, 5 ' Select Key Pressed - Machine Coordinate Mode
                    Select Case passType
                        Case passBlade
                            Call initJoyAxes(aXj, False, aYj, False, aZaj, False)
                        Case passAugEdg
                            Call initJoyAxes(aXj, False, aYj, True, aZaj, False)
                        Case passAugFac
                            Call initJoyAxes(aXj, False, aYj, True, aZaj, False)
                    End Select
                
                Case 2, 3, 6, 7 ' Release pressed in any combination
                    
                    GoTo ReleaseJoystick
                    
                Case Else
                    
                    GoTo joyErrorHand
                
            End Select
            
        Case joyPause  '------------------------------------------------------------------------- PAUSE -------------------
            Select Case buttons     'Toggle must be flipped for this case to occur, therefore only case 4-7
                Case 4      'Toggle Only
                    Select Case passType
                        Case passBlade
                            Call initJoyAxes(aXj, False, aYj, False, aZj, False)
                        Case passAugEdg
                            If c6kOps.getAugerDir = "Right" Then
                                Call initJoyAxes(aYR, False, aYj, False, aZj, False)
                            Else
                                Call initJoyAxes(aYRi, False, aYj, False, aZj, False)
                            End If
                        Case passAugFac
                            If c6kOps.getAugerDir = "Right" Then
                                Call initJoyAxes(aZpR, True, aXj, False, aZp, True)
                            Else
                                Call initJoyAxes(aZpRi, True, aXj, False, aZp, True)
                            End If
                    End Select
                Case 5      'Select Pressed
                    Select Case passType
                        Case passBlade
                            Call initJoyAxes(aXj, False, aYj, False, aZaj, False)
                        Case passAugEdg, passAugFac
                            If c6kOps.getAugerDir = "Right" Then
                                Call initJoyAxes(aYj, False, aRj, False, aZaj, False)
                            Else
                                Call initJoyAxes(aYj, False, aRj, True, aZaj, False)
                            End If
                    End Select
                Case 6, 7  ' Release Pressed - Advance Program
                    GoTo ReleaseJoystick
                Case 0, 1, 2, 3     'Toggle not flipped - return to run mode
                    GoTo ReleaseJoystick
                Case Else
                    GoTo joyErrorHand
            End Select
        
        Case joyRun '---------------------------------------------------------------------------- RUN ---------------------
            Select Case buttons
                Case 0    'No buttons Pressed - Horizontal axis will adjust run speed
                    Select Case passType
                        Case passBlade
                            Call initJoyAxes(arS, False, aYj, False, aZj, False)
                        Case passAugEdg
                            Call initJoyAxes(arS, False, aYb, False, aZj, False)
                        Case passAugFac
                            Call initJoyAxes(arS, False, aXj, False, aZj, False)
                    End Select
                Case 1      'Select Pressed - Horizontal axis will nudge run axis
                    Select Case passType
                        Case passBlade
                            Call initJoyAxes(aXb, False, aYj, False, aZaj, False)
                        Case passAugEdg
                                Call initJoyAxes(aYRb, False, aXj, False, aZaj, False)
                        Case passAugFac
                            Call initJoyAxes(aZpRb, False, aYc, False, aZpb, False)
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


Private Sub initJoyAxes(joyLR As Long, LRinv As Boolean, joyFB As Long, FBinv As Boolean, joyUD As Long, UDinv As Boolean)

'Ensure no identical axes
If joyLR = joyFB Or joyLR = joyUD Or joyFB = joyUD Then
    MsgBox "Cannot assign multiple joystick directions to the same axis!"
    Call eStopJoy
    Exit Sub
End If

Dim i As Integer
Dim j As Integer

'Load axes into joystick object
myJoy.inSt(0) = joyLR
myJoy.inSt(1) = joyFB
myJoy.inSt(2) = joyUD
myJoy.inInv(0) = LRinv
myJoy.inInv(1) = FBinv
myJoy.inInv(2) = FBinv

'Resolve conflicts between jog axes and McJog axes
Dim chkFlags As Long
chkFlags = (axisJog Or axisMc Or axisBst)
Dim chkLR As Long
Dim chkFB As Long
Dim chkUD As Long
'Remove flags from conflict check
chkLR = (joyLR Or chkFlags) Xor chkFlags
chkFB = (joyFB Or chkFlags) Xor chkFlags
chkUD = (joyUD Or chkFlags) Xor chkFlags

'Check axes against each other. If any conflicts, then
If (chkLR And chkFB) Or (chkLR And chkUD) Or (chkFB And chkUD) Then
    Dim tempSt As Long
    'For each joystick axis
    For i = 0 To 2
        'Determine if the axis is in jog mode
        tempSt = myJoy.inSt(i)
        If tempSt And axisJog Then
            'If it is, determine if it conflicts with any axes
            For j = 0 To 2
                If tempSt <> myJoy.inSt(j) Then
                    'If it does conflict, then determine which jog axis is conflicting and convert it to a McJog axis
                    If Not CBool((tempSt Xor myJoy.inSt(j)) And tempSt) Then
                        Select Case tempSt
                            Case aXj
                                tempSt = aXc
                            Case aYj
                                tempSt = aYc
                            Case aZj
                                tempSt = aZc
                            Case aZaj
                                tempSt = aZac
                            Case aRj
                                tempSt = aRc
                        End Select
                        myJoy.inSt(i) = tempSt
                    End If
                End If
            Next j
        End If
    Next i

End If

Call jogOff

'Set joyString
myJoy.stateStr = getStr(joyLR) & " - " & getStr(joyFB) & " - " & getStr(joyUD)

'Clear outSt array to ensure no carry-over
For i = 1 To 6
    myJoy.outSt(i) = 0
Next i

'Clear joystick flags
myJoy.joyJogOn = False
myJoy.joyMcJogOn = False
myJoy.joyBstOn = False
myJoy.joySpdOn = False

'Cause a brief delay - I don't know why this needs to be here, but the inputs definitions won't stick without it
Dim d As Long
For d = 1 To 25000000
Next d

Dim tempJoyAxis As Integer 'Temp variable containing value of axis
Dim tempJoyInputA As String 'Temp variable containing input value for Jog
Dim tempJoyInputB As String 'As Above
Dim tempDirA
Dim tempDirB

'For Each Axis
For i = 0 To 2

    tempJoyAxis = myJoy.inSt(i)
    '-- Set temps for axis
    Select Case i
        Case 0
            tempJoyInputA = "1"
            tempJoyInputB = "2"
        Case 1
            tempJoyInputA = "3"
            tempJoyInputB = "4"
        Case 2
            tempJoyInputA = "5"
            tempJoyInputB = "6"
    End Select
    
    'Set axis input inversion
    If myJoy.inInv(i) Then
        tempDirA = "J"
        tempDirB = "K"
    Else
        tempDirA = "K"
        tempDirB = "J"
    End If
    
            
    'Set joystick flags
    If tempJoyAxis And axisJog Then myJoy.joyJogOn = True
    If tempJoyAxis And axisMc Then myJoy.joyMcJogOn = True
    If tempJoyAxis And axisBst Then myJoy.joyBstOn = True
    
            
    '-- Determine axis and set accordingly
    Select Case tempJoyAxis
        '-- Run Speed Adjustment ----------------------------------------------------------------------------------rS
        Case arS
            myJoy.joySpdOn = True
            
        '-- X - Axis -----------------------------------------------------------------------------------------------X
        Case aXj
            
            c6k.Write "1INFNC" & tempJoyInputA & "-" & driveX & tempDirA & ":1INFNC" & tempJoyInputB & "-" & driveX & tempDirB & Chr$(13)    'Set jog inputs in 6k
            myJoy.outSt(driveX) = driveJog       'Set axis in Joystick object to drive mode
        Case aXc
            myJoy.outSt(driveX) = driveMcJog
        Case aXb
    
        '-- Y - Axis -----------------------------------------------------------------------------------------------Y
        Case aYj
            c6k.Write "1INFNC" & tempJoyInputA & "-" & driveY & tempDirA & ":1INFNC" & tempJoyInputB & "-" & driveY & tempDirB & Chr$(13)    'Set jog inputs in 6k
            myJoy.outSt(driveY) = driveJog      'Set axis in Joystick object to drive mode
        Case aYc
            myJoy.outSt(driveY) = driveMcJog
        Case aYb

        '-- Z - Axis Fuctions --------------------------------------------------------------------------------------Z
        Case aZj
            c6k.Write "1INFNC" & tempJoyInputA & "-" & driveZ & "K:1INFNC" & tempJoyInputB & "-" & driveZ & tempDirB & Chr$(13)    'Set jog inputs in 6k
            myJoy.outSt(driveZ) = driveJog      'Set axis in Joystick object to drive mode
        Case aZc
            myJoy.outSt(driveZ) = driveMcJog
        Case aZb
    
        '-- Za - Axis Fuctions ------------------------------------------------------------------------------------Za
        Case aZaj
            c6k.Write "1INFNC" & tempJoyInputA & "-" & driveZa & tempDirA & ":1INFNC" & tempJoyInputB & "-" & driveZa & tempDirB & Chr$(13)    'Set jog inputs in 6k
            myJoy.outSt(driveZa) = driveJog      'Set axis in Joystick object to drive mode
        Case aZac
            myJoy.outSt(driveZa) = driveMcJog
        Case aZab
    
        '-- R - Axis Fuctions --------------------------------------------------------------------------------------R
        Case aRj
            c6k.Write "1INFNC" & tempJoyInputA & "-" & driveR & tempDirA & ":1INFNC" & tempJoyInputB & "-" & driveR & tempDirB & Chr$(13)    'Set jog inputs in 6k
            myJoy.outSt(driveR) = driveJog      'Set axis in Joystick object to drive mode
        Case aRc
            myJoy.outSt(driveR) = driveMcJog
        Case aRb
    
        '-- YR - Axis Fuctions ------------------------------------------------------------------------------------YR
        Case aYR, aYRi
            myJoy.outSt(driveR) = driveMcJog    'Set axes in Joystick object to continuous jog mode
            myJoy.outSt(driveY) = driveMcJog
        Case aYRb
    
        '-- Z' - Axis Fuctions ------------------------------------------------------------------------------------Z'
        Case aZp, aZpi
            myJoy.outSt(driveY) = driveMcJog    'Set axes in Joystick object to continuous jog mode
            myJoy.outSt(driveZ) = driveMcJog
        Case aZpb


        '-- Z'R - Axis Fuctions ----------------------------------------------------------------------------------Z'R
        Case aZpR, aZpRi
            myJoy.outSt(driveY) = driveMcJog    'Set axes in Joystick object to continuous jog mode
            myJoy.outSt(driveZ) = driveMcJog
            myJoy.outSt(driveR) = driveMcJog
        Case aZpRb
        
        Case Else
            c6k.Write "1INFNC" & tempJoyInputA & "-A:1INFNC" & tempJoyInputB & "-A" & Chr$(13)    'Set jog inputs in 6k
    End Select
    
Next i

c6k.Write ("1INFNC" & Chr(13))

Call jogOn

End Sub

Private Sub jogOn()

'-- Set c6k Jog State if necessary
If myJoy.joyJogOn Then
    Dim tempJog As String
    Dim tempErr As Integer
    Dim i As Integer
    'For all values in joyJogState (except index 0, which is unused)
    For i = 1 To 6
        If i <> 4 Then  'For all i except 4 (Osc)
        
            Select Case myJoy.outSt(i)
                Case 0, driveMcJog, driveMcRun         ' If jogState is not 1
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
If myJoy.joyMcJogOn Then
    
    If myJoy.joyMcStr = "" And fsmRun.State <> 2 And fsmRun.State <> 3 Then c6k.Write "COMEXC1" & Chr(13)   'Enable cont. Com ex when enabling MC mode
    
    Dim tempMc As String
    Dim tempMcVel As String
    'For all values in joyJogState (except index 0, which is unused)
    For i = 1 To 6
        If i <> 4 Then  'For all i except 4 (Osc)
        
            Select Case myJoy.outSt(i)
                Case 0, driveJog, driveMcRun         ' If jogState is not 1
                    'First loop through, set initial variable value, else append character
                    If i = 1 Then
                        tempMc = "X"
                        tempMcVel = ""
                    Else
                        tempMc = tempMc & "X"
                        tempMcVel = tempMcVel & ","
                    End If
                    
                Case driveMcJog          'If drive is in jog mode, make it active
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
    
    'Disable Previous MC Jog mode
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

'Disable Jog Axes
If myJoy.joyJogStr <> "" Then
    c6k.Write ("JOG" & Replace(myJoy.joyJogStr, "1", "0") & Chr(13))
End If

'Disable MC Axes
If myJoy.joyMcStr <> "" Then
    c6k.Write ("S" & myJoy.joyMcStr & Chr(13))
    c6k.Write ("MC" & Replace(myJoy.joyMcStr, "1", "0") & Chr(13))
End If

End Sub

Private Function checkJoyAxisDir(i As Integer)

Dim currentJoyDir

If myJoy.inInv(i) Then
    dirA = 2
    dirB = 1
Else
    dirA = 1
    dirB = 2
End If
    
Select Case i
    Case 0
        If myJoy.stateNow(joyRight) Then
            currentJoyDir = dirA
        ElseIf myJoy.stateNow(joyLeft) Then
            currentJoyDir = dirB
        End If
    Case 1
        If myJoy.stateNow(joyFront) Then
            currentJoyDir = dirA
        ElseIf myJoy.stateNow(joyBack) Then
            currentJoyDir = dirB
        End If
    Case 2
        If myJoy.stateNow(joyUp) Then
            currentJoyDir = dirA
        ElseIf myJoy.stateNow(joyDown) Then
            currentJoyDir = dirB
        End If
End Select

checkJoyAxisDir = currentJoyDir
End Function

Private Sub runJoyInputs()

Dim i As Integer

Dim dirA As Integer
Dim dirB As Integer

'-- If no joystick inputs require continuous motion, exit
If myJoy.joyJogOn And Not myJoy.joyMcJogOn And Not myJoy.joyBstOn Then Exit Sub


'-- If any of the axes are set to speed control, call for speed to be adjusted
If myJoy.joySpdOn Then
    Dim changeSpeed As Integer

    For i = 0 To 2  'For each joystick direction
        If myJoy.inSt(i) = arS Then
            changeSpeed = checkJoyAxisDir(i)
        End If
    Next i
    
    'Adjust Speed Accordingly, respecting speed bounds
    If changeSpeed = 1 Then
        If userSpeedMult <= speedMultUpperBound Then userSpeedMult = userSpeedMult + speedMultInc
    ElseIf changeSpeed = 2 Then
        If userSpeedMult >= speedMultLowerBound Then userSpeedMult = userSpeedMult - speedMultInc
    End If
    
End If

'-- If any of the axes require MC Jog motion
If myJoy.joyMcJogOn Then

    'Clear Axis Velocities
    For j = 1 To 6
        myJoy.joyMcAxisVel(j) = 0
    Next j

    'For each joysick axis
    For i = 0 To 2
        'If the axis is in mcJog mode
        If myJoy.inSt(i) And axisMc Then
        
            Dim mcJoyDir As Integer
            Dim axisIsInverted As Boolean
            Dim tempSpeedMult As Single
            tempSpeedMult = 1
            
            'Determine what direction the joystick is in (1- Neg, 0- Centered, 2- Pos)
            mcJoyDir = checkJoyAxisDir(i)
            
            'Set Axis Velocities
            For j = 1 To 6
                If j = 4 Then GoTo VelSetNextAxis
                Select Case j
                    Case driveX
                        If Not CBool(myJoy.inSt(i) And aX) Then GoTo VelSetNextAxis
                        If CBool(myJoy.inSt(i) And aXi) Then axisIsInverted = True
                    Case driveY
                        If Not CBool(myJoy.inSt(i) And aY) Then GoTo VelSetNextAxis
                        If CBool(myJoy.inSt(i) And aYi) Then axisIsInverted = True
                        If passType = passAugEdg Then tempSpeedMult = augerPitch / (PI * augerDiam)
                        If passType = passAugFac Then tempSpeedMult = Sin(augerAngle) * augerPitch / (PI * augerDiam)
                    Case driveZ
                        If Not CBool(myJoy.inSt(i) And aZ) Then GoTo VelSetNextAxis
                        If CBool(myJoy.inSt(i) And aZi) Then axisIsInverted = True
                        If passType = passAugFac Then tempSpeedMult = Cos(augerAngle) * augerPitch / (PI * augerDiam)
                    Case driveZa
                        If Not CBool(myJoy.inSt(i) And aZa) Then GoTo VelSetNextAxis
                        If CBool(myJoy.inSt(i) And aZai) Then axisIsInverted = True
                    Case driveR
                        If Not CBool(myJoy.inSt(i) And aR) Then GoTo VelSetNextAxis
                        If CBool(myJoy.inSt(i) And aRi) Then axisIsInverted = True
                        tempSpeedMult = 60
                        If passType = passAugEdg Or passType = passAugFac Then tempSpeedMult = 360 / (PI * augerDiam)
                End Select
                
                If mcJoyDir <> 0 Then 'If a motion is commanded,
                    'The velocity commanded to the motor is not zero, or is not at constant speed,
                    If (fsinfo.MotorVel(j) <> 0) Or (CBool(fsinfo.AxisStatus(j) And (2 ^ (4 - 1)))) Then
                        'And if the input direction is opposite the motor's current direction
                        If ((CBool(fsinfo.AxisStatus(j) And 2) And mcJoyDir = 1) Or (Not CBool(fsinfo.AxisStatus(j) And 2) And mcJoyDir = 2)) Then
                            'Reset the output such that velocity does not get added for this axis
                            mcJoyDir = 0
                        End If
                    End If
                End If
                
                Dim tempVel As Single
                
                If mcJoyDir Then
                    'The velocity of the drive axis should be set based on the joystick position
                    Select Case mcJoyDir
                        Case 1
                            If Not axisIsInverted Then tempVel = cJoyJogVel Else temVel = -cJoyJogVel
                        Case 2
                            If Not axisIsInverted Then tempVel = -cJoyJogVel Else temVel = cJoyJogVel
                    End Select
                    
                    myJoy.joyMcAxisVel(j) = myJoy.joyMcAxisVel(j) + (tempVel * tempSpeedMult)
                End If
VelSetNextAxis:
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
If myJoy.joyBstOn Then
    
    Dim tempBst As Integer
    
    For i = 0 To 2
        If myJoy.inSt(i) And axisBst Then
            tempBst = checkJoyAxisDir(i)
            
            'Set Axis Boost
            For j = 1 To 6
                If j = 4 Then GoTo BstSetNextAxis
                Select Case j
                    Case driveX
                        If Not CBool(myJoy.inSt(i) And aX) Then GoTo BstSetNextAxis
                    Case driveY
                        If Not CBool(myJoy.inSt(i) And aY) Then GoTo BstSetNextAxis
                    Case driveZ
                        If Not CBool(myJoy.inSt(i) And aZ) Then GoTo BstSetNextAxis
                    Case driveZa
                        If Not CBool(myJoy.inSt(i) And aZa) Then GoTo BstSetNextAxis
                    Case driveR
                        If Not CBool(myJoy.inSt(i) And aR) Then GoTo BstSetNextAxis
                End Select
            
                userSpeedBoost(j) = tempBst
            
BstSetNextAxis:
            Next j
        End If
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
