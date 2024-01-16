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

'Inputs - bINPUT corresponds to the binary from FastStatus, INPUT corresponds to index in input array (joyIn(9))
Private Const bJoyLeft As Long = Input1
Private Const joyLeft As Integer = 0
Private Const bJoyRight As Long = Input2
Private Const joyRignt As Integer = 1
Private Const bJoyFront As Long = Input3
Private Const joyFront As Integer = 2
Private Const bJoyBack As Long = Input4
Private Const joyBack As Integer = 3
Private Const bJoyUp As Long = Input5
Private Const joyUp As Integer = 4
Private Const bJoyDown As Long = Input6
Private Const joyDown As Integer = 5

Private Const bJoyPause As Long = Input21
Private Const joyPause As Integer = 6
Private Const bJoySelect As Long = Input22
Private Const joySelect As Integer = 7
Private Const bJoyRelease As Long = Input23
Private Const joyRelease As Integer = 8
Private Const bJoyToggle As Long = Input24
Private Const joyToggle As Integer = 9

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

'Internal Variables & Defaults
Private Const inDebCount As Integer = 5      'Time in FSM cycles of input debounce

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
myJoy.inBin(joyPause) = bJoyPause
myJoy.inBin(joySelect) = bJoySelect
myJoy.inBin(joyRelease) = bJoyRelease
myJoy.inBin(joyToggle) = bJoyToggle

'Initialize joystick jog parameters
Call initJoy

End Function


'-- Joystick Input Commands -------------------------------------------------------------------------------------------


'-- Simple Read-Only Return Functions:
Public Function getJoyActive() As Boolean
    getJoyActive = myJoy.active
End Function

Public Function getJoyStr() As String
    getJoyStr = myJoy.statusStr
End Function


'-- getJoyInputs - Function to update joystick inputs
'       - Inputs:
'           - c6kInput is the Long containing the Input Status from the c6k
'           - returnIndex is the index corresponding to the return input
'       - Returns:
'           - Bool corresponding to the joystick input corresponding to returnIndex. Returns False by default
Public Function getJoyInputs(c6kInput As Long, Optional returnIndex As Integer) As Bool
    'For each input,
    Dim I As Integer
    Dim tempIn As Boolean
    For I = 0 To 9
        
        'Store the current input value
        tempIn = (c6kInput And myJoy.inBin(I))
        
        'See if the current input value is different from the previous one
        If tempIn <> myJoy.stateNow(I) Then
            'If it is, Store last input value
            myJoy.stateLast(I) = myJoy.stateNow(I)
            
            'If the input is set in the 6k and input debounce is not still counting, then set the state active.
            If tempIn Then
                If Not (myJoy.inDeb(I) > 0) Then
                    myJoy.stateNow(I) = True
                    myJoy.inDeb(I) = inDebCount
                    myJoy.stateUpd = True
                Else
                    ' If the code reaches this point, the input changed but debounce is preventing it from reading true again
                    myJoy.stateNow(I) = False
                End If
            Else
                myJoy.stateNow(I) = False
                myJoy.stateUpd = True
            End If
        End If
        
        'Decrement inputDebounce when necessary
        If myJoy.inDeb(I) > 0 Then myJoy.inDeb(I) = myJoy.inDeb(I) - 1
        
    Next I
    
    ' If a return index is supplied, return the corresponding value
    If IsMissing(returnIndex) Then
        getJoyInputs = False
    ElseIf returnIndex >= 0 And returnIndex < 10 Then
        getJoyInputs = myJoy.stateNow(returnIndex)
    Else
        MsgBox "Error in getJoyInputs: " & returnIndex & " is not a valid input"
    End If
    
End Function

'-- initJoy - Function to set up basic joystick parameters. Called when initializing joystick
Private Sub initJoy()

'Immediately (!) disable jog mode for all axes
c6k.Write ("!JOG000000" & Chr$(13))

'Set motion parameters for Joystick Motion
c6k.Write ("JOG000000:JOGA4,5,5,1,5,5:JOGAD50,99,99,99,99,15:JOGVH8,8,10,2,5,3:JOGVL8,15,10,5,5,5" & Chr$(13))

End Sub


'-- startJoy - Function to set joystick jog axes
'       - Inputs:
'           - newAxes - String containing the new desired jog axis - Format 1XXXXX where 1 indicates enabled axis
'           - oldAxes - String containing the previously enabled joystick axes - Format as above
Private Function startJoy(newAxes As String, Optional oldAxes) As String

If Not IsMissing(oldAxes) Then
    'Immediately disable jog on the previous axes
    c6k.Write ("!JOG" & Replace(oldAxes, "1", "0") & Chr$(13))
End If

'Write new jog axes to 6k
c6k.Write ("JOG" & newAxes & Chr$(13))

myJoy.active = True

startJoy = newAxes

End Function

'-- stopJoy - Function to shut down joystick
'       - Inputs:
'           - currentAxes - String containing the current jog axes that need to be shut down - Format 1XXXXX where 1 indicates enabled axis
Private Sub stopJoy(currentJogAxes As String, Optional currentMcAxes As String)

'Disable jog mode on currently active drives
c6k.Write ("!JOG" & currentJogAxes & Chr$(13))

'If MC on any drives, disable
If Not IsMissing(currentMcAxes) Then c6k.Write ("!MC" & currentMcAxes & Chr$(13))

'Disable inputs
c6k.Write ("1INFC1-A:1INFC2-A:1INFC3-A:1INFC4-A")

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
Public Function runJoy(mode As String) As Boolean
         
'Update the joystick inputs. Calling the getInputState function will update FastStatus
Call getJoyInputs(c6kOps.getInputState)

'Set up end flag & set to default of false
Dim endJoy As Boolean
endJoy = False

'If mode is disable, shut down the joystick. Else,
If mode = "Disable" Then
    Call stopJoy(myJoy.joyJogStr, myJoy.joyMcStr)
    runJoy = False
    Exit Function
Else
    'Set the mode of the joystick. This will pair the correct machine axes with the corresponding joystick axes for the mode requested
    endJoy = Not modeSet(inputs, mode)
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
'           - buttons - integer containing button state of joystick
'           - runMode - Run mode of the pass - see runJoy for details (same input)
Private Function joyModeSet(runMode As String) As Boolean

If (buttons + 1) <> joyLastBtnInput Then
    
    'Set joystick mode based on run mode - Free, Pause, or Run
    Select Case runMode
        Case "Free"         'Standard X-Y-Z Mode. Used for Blades & Positioning
        
            Select Case buttons
                Case 0, 4  ' No Keys Pressed or Toggle Only - X-Y Mode
                    Select Case passType
                        Case "Blade"
                            Call setJoyInputs("X", "Y")
                        Case "Auger Edge"
                            Call setJoyInputs("YR", "Y")
                        Case "Auger Face"
                            Call setJoyInputs("X", "Z'R")
                    End Select
                    
                Case 1 ' Select Key Pressed
                    Select Case passType
                        Case "Blade"
                            Call setJoyInputs("X", "Z")
                        Case "Auger Edge"
                            Call setJoyInputs("YR", "Z")
                        Case "Auger Face"
                            Call setJoyInputs("X", "Z'")
                    End Select
                  
                Case 5 ' Toggle Flipped and Select Pressed
                    Select Case passType
                        Case "Blade"
                            Call setJoyInputs("X", "Za")
                        Case "Auger Edge"
                            Call setJoyInputs("YR", "Za")
                        Case "Auger Face"
                            Call setJoyInputs("Y", "Z'")
                    End Select
                
                Case 2, 3, 6, 7 ' Release pressed in any combination
                    
                    GoTo ReleaseJoystick
                    
                Case Else
                    
                    GoTo joyErrorHand
                
            End Select
            
        Case "Pause"
            Select Case buttons
                Case 4      'Toggle Only
                    Select Case passType
                        Case "Blade"
                            Call setJoyInputs("X", "Y")
                        Case "Auger Edge"
                            Call setJoyInputs("YR", "Y")
                        Case "Auger Face"
                            Call setJoyInputs("Z'R", "X")
                    End Select
                Case 5      'Select Pressed
                    Select Case passType
                        Case "Blade"
                            Call setJoyInputs("X", "Za")
                        Case "Auger Edge"
                            Call setJoyInputs("YR", "Za")
                        Case "Auger Face"
                            Call setJoyInputs("Y", "Za")
                    End Select
                Case 6, 7  ' Release Pressed - Advance Program
                    GoTo ReleaseJoystick
                Case 0, 1, 2, 3     'Toggle not flipped - return to run mode
                    GoTo ReleaseJoystick
                Case Else
                    GoTo joyErrorHand
            End Select
        
        Case "Run"
            Select Case buttons
                Case 0    'No buttons Pressed - Horizontal axis will adjust run speed
                    Select Case passType
                        Case "Blade"
                            Call setJoyInputs("rS", "Y")
                        Case "Auger Edge"
                            Call setJoyInputs("rS", "Y")
                        Case "Auger Face"
                            Call setJoyInputs("rS", "X")
                    End Select
                Case 1      'Select Pressed - Horizontal axis will nudge run axis
                    Select Case passType
                        Case "Blade"
                            Call setJoyInputs("Xb", "Z")
                        Case "Auger Edge"
                            Call setJoyInputs("YRb", "Z")
                        Case "Auger Face"
                            Call setJoyInputs("Z'Rb", "Z'b")
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
    
    
joyModeSet = True
joyLastBtnInput = (buttons + 1)

Exit Function

joyErrorHand:

MsgBox "Error in joyModeFree Handler:" & Chr(13) & "Mode " & passType & ", Case: " & buttons

ReleaseJoystick:
    'Return False to indicate that user input should end joystick mode
    joyModeSet = False
    joyLastBtnInput = 0


End Function


Private Sub setJoyInputs(joyHoriz As String, joyVert As String)    'joyRck As String

If joyHoriz = joyVert Then
    MsgBox "Cannot assign both joystick directions to the same axis!"
    Exit Sub
End If

joyStr = joyHoriz & " - " & joyVert

'Clear joyMode vars and JogState vars
joyModeX = 0
joyModeY = 0

Dim I As Integer
For I = 1 To 6
    joyJogState(I) = 0
Next I

Dim temp As String

' Run Speed Adjustment Axis
temp = "rS"
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as the current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        joyModeX = 3
    ElseIf joyVert = temp Then
        joyModeY = 3
    End If
End If

'-- X - Axis Fuctions --------------------------------------------------------------------------------------X
temp = "X"  'Jog
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as the current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        c6k.Write "1INFNC1-1K:1INFNC2-1J" & Chr$(13)    'Set inputs 1 & 2 for Horizontal joy axis
        joyJogState(1) = 1                              'Set jogState appropriately
        joyModeX = 0
    ElseIf joyVert = temp Then
        c6k.Write "1INFNC3-1K:1INFNC4-1J" & Chr$(13)    'Set inputs 3 & 4 for Vertical joy axis
        joyJogState(1) = 1                              'Set jogState appropriately
        joyModeY = 0
    End If
End If

temp = "Xb"  'Run Speed Boost
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as the current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        joyJogState(1) = 2
        joyModeX = 2
    ElseIf joyVert = temp Then
        joyJogState(1) = 3
        joyModeY = 2
    End If
End If



'-- Y - Axis Fuctions --------------------------------------------------------------------------------------Y
temp = "Y"  'Set Y - Axis Functionality
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as the current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        c6k.Write "1INFNC1-6K:1INFNC2-6J" & Chr$(13)    'Set inputs 1 & 2 for Horizontal joy axis
        joyJogState(6) = 1                              'Set jogState appropriately
        joyModeX = 0
    ElseIf joyVert = temp Then
        c6k.Write "1INFNC3-6K:1INFNC4-6J" & Chr$(13)    'Set inputs 3 & 4 for Vertical joy axis
        joyJogState(6) = 1                              'Set jogState appropriately
        joyModeY = 0
    End If
End If

temp = "Yb"  'Run Speed Boost
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as the current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        joyJogState(6) = 2
        joyModeX = 2
    ElseIf joyVert = temp Then
        joyJogState(6) = 3
        joyModeY = 2
    End If
End If


'-- Z - Axis Fuctions --------------------------------------------------------------------------------------Z
temp = "Z"  'Set Z - Axis Functionality
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as the current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        c6k.Write "1INFNC1-2K:1INFNC2-2J" & Chr$(13)    'Set inputs 1 & 2 for Horizontal joy axis
        joyJogState(2) = 1                              'Set jogState appropriately
        joyModeX = 0
    ElseIf joyVert = temp Then
        c6k.Write "1INFNC3-2K:1INFNC4-2J" & Chr$(13)    'Set inputs 3 & 4 for Vertical joy axis
        joyJogState(2) = 1                              'Set jogState appropriately
        joyModeY = 0
    End If
End If

temp = "Zb"  'Run Speed Boost
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as the current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        joyJogState(2) = 2
        joyModeX = 2
    ElseIf joyVert = temp Then
        joyJogState(2) = 3
        joyModeY = 2
    End If
End If


'-- Za - Axis Fuctions ------------------------------------------------------------------------------------Za
temp = "Za"  'Set Zalt - Axis Functionality
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as teh current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        c6k.Write "1INFNC1-3K:1INFNC2-3J" & Chr$(13)    'Set inputs 1 & 2 for Horizontal joy axis
        joyJogState(3) = 1                              'Set jogState appropriately
        joyModeX = 0
    ElseIf joyVert = temp Then
        c6k.Write "1INFNC3-3K:1INFNC4-3J" & Chr$(13)    'Set inputs 3 & 4 for Vertical joy axis
        joyJogState(3) = 1                              'Set jogState appropriately
        joyModeY = 0
    End If
End If

temp = "Zab"  'Run Speed Boost
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as the current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        joyJogState(3) = 2
        joyModeX = 2
    ElseIf joyVert = temp Then
        joyJogState(3) = 3
        joyModeY = 2
    End If
End If


'-- R - Axis Fuctions --------------------------------------------------------------------------------------R
temp = "R"  'Set Rotation - Axis Functionality
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as teh current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        c6k.Write "1INFNC1-5J:1INFNC2-5K" & Chr$(13)    'Set inputs 1 & 2 for Horizontal joy axis
        joyJogState(5) = 1                              'Set jogState appropriately
        joyModeX = 0
    ElseIf joyVert = temp Then
        c6k.Write "1INFNC3-5J:1INFNC4-5K" & Chr$(13)    'Set inputs 3 & 4 for Vertical joy axis
        joyJogState(5) = 1                              'Set jogState appropriately
        joyModeY = 0
    End If
End If

temp = "Rb"  'Run Speed Boost
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as the current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        joyJogState(5) = 2
        joyModeX = 2
    ElseIf joyVert = temp Then
        joyJogState(5) = 3
        joyModeY = 2
    End If
End If


'-- YR - Axis Fuctions ------------------------------------------------------------------------------------YR
temp = "YR"  'Set Auger Edge - Y&R - Axis Functionality
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as teh current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        joyJogState(5) = 2                              'Set jogState appropriately
        joyJogState(6) = 2
        joyModeX = 1
    ElseIf joyVert = temp Then
        joyJogState(5) = 3                              'Set jogState appropriately
        joyJogState(6) = 3                              'Set jogState appropriately
        joyModeY = 1
    End If
End If

temp = "YRb"  'Run Speed Boost
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as the current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        joyJogState(5) = 2
        joyJogState(6) = 2
        joyModeX = 2
    ElseIf joyVert = temp Then
        joyJogState(5) = 3
        joyJogState(6) = 3
        joyModeY = 2
    End If
End If

'-- Z' - Axis Fuctions ------------------------------------------------------------------------------------Z'
temp = "YR"  'Set Auger Edge - Y&R - Axis Functionality
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as teh current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        joyJogState(2) = 2                              'Set jogState appropriately
        joyJogState(6) = 2
        joyModeX = 1
    ElseIf joyVert = temp Then
        joyJogState(2) = 3
        joyJogState(6) = 3                              'Set jogState appropriately
        joyModeY = 1
    End If
End If

temp = "Z'b"  'Run Speed Boost
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as the current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        joyJogState(2) = 2
        joyJogState(6) = 2
        joyModeX = 2
    ElseIf joyVert = temp Then
        joyJogState(2) = 3
        joyJogState(6) = 3
        joyModeY = 2
    End If
End If

'-- Z'R - Axis Fuctions ----------------------------------------------------------------------------------Z'R
temp = "YR"  'Set Auger Edge - Y&R - Axis Functionality
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as teh current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        joyJogState(2) = 2
        joyJogState(5) = 2                              'Set jogState appropriately
        joyJogState(6) = 2
        joyModeX = 1
    ElseIf joyVert = temp Then
        joyJogState(2) = 3
        joyJogState(5) = 3                              'Set jogState appropriately
        joyJogState(6) = 3                              'Set jogState appropriately
        joyModeY = 1
    End If
End If

temp = "YRb"  'Run Speed Boost
If joyHoriz = temp Or joyVert = temp Then   'If any input axes are indicated as the current axis, then
    'Determine what axis should be set and set jog / MC inputs accordingly
    If joyHoriz = temp Then
        joyJogState(2) = 2
        joyJogState(5) = 2
        joyJogState(6) = 2
        joyModeX = 2
    ElseIf joyVert = temp Then
        joyJogState(2) = 3
        joyJogState(5) = 3
        joyJogState(6) = 3
        joyModeY = 2
    End If
End If



'-- Set c6k Jog State if necessary
If joyModeX = 0 Or joyModeY = 0 Then
    Dim tempJog As String
    
    'For all values in joyJogState (except index 0, which is unused)
    For I = 1 To 6
        If I <> 4 Then  'For all i except 4 (Osc)
        
            Select Case joyJogState(I)
                Case 0, 2, 3, 4         ' If jogState is not 1
                    'First loop through, set initial variable value, else append character
                    If I = 1 Then
                        tempJog = "0"
                    Else
                        tempJog = tempJog & "0"
                    End If
                    
                Case 1          'If jogState is 1 then axis is active
                    If I = 1 Then
                        tempJog = "1"
                    Else
                        tempJog = tempJog & "1"
                    End If
                    
                Case Else
                    MsgBox ("Error joyJogState Conversion:" & Chr(13) & "Index: " & I & ", State: " & joyJogState(I))
            End Select
        Else    ' If i = 4 (Oscillator Axis)
            tempJog = tempJog & "0"
        End If
    Next I
    
    c6k.Write ("JOG" & tempJog & Chr(13))
End If

End Sub

Private Sub runJoyInputs()

Dim I As Integer

'-- If neither joystick input requires continuous motion, set jog mode
If joyModeX = 0 And joyModeY = 0 Then

        'Jog is set in setJoyInputs
End If

'-- If one or both of the axes require Continuous Motion outside of run mode
If joyModeX = 1 Or joyModeY = 1 Then
    
    Call getJoyStkInputs
    
    '-- define temp values - These are static to retain values between calls
    Static TempH As String
    Static dirH As String
    
    Static TempV As String
    Static dirV As String
    
    
    '-- Run based on joystick inputs
    If joyStkInputs <> joyLastStkInputs Then
    
        '-- Stop movement based on prior input
        Select Case joyLastStkInputs
            Case 0
    
            Case 1, 2
                If TempH <> "" Then c6k.Write "!S" & TempH & ":MC" & Replace(TempH, "1", "0") & Chr(13)
            Case 3, 4
                If TempV <> "" Then c6k.Write "!S" & TempV & ":MC" & Replace(TempV, "1", "0") & Chr(13)
            Case Else
                MsgBox "Error in joystick inputs"
        End Select
        
        'Short Circuit for if input is 0 (to avoid unnecessary calculations)
        If joyStkInputs = 0 Then Exit Sub
        
        
        '-- Rebuild Temp Values for Next State
        'For all values in joyJogState (except index 0, which is unused)
        For I = 1 To 6
            If I <> 4 Then  'For all i except 4 (Osc)
            
                Select Case joyJogState(I)
                    Case 0, 1         ' If jogState is 0 (axis inactive) or 1 (standard jog active) MC disabled
                        'First loop through, set initial variable value, else append character
                        If I = 1 Then
                            TempH = "X"
                            TempV = "X"
                            dirH = ","
                            dirV = ","
                        Else
                            TempH = TempH & "X"
                            TempV = TempV & "X"
                            dirH = dirH & ","
                            dirV = dirV & ","
                        End If
                        
                    Case 2          'If jogState is 1 then axis is tied to joy horizontal
                        If I = 1 Then
                            TempH = "1"
                            TempV = "X"
                            dirH = "*,"
                            dirV = ","
                        Else
                            TempH = TempH & "1"
                            TempV = TempV & "X"
                            dirH = dirH & "*,"
                            dirV = dirV & ","
                        End If
                
                    Case 3          'If jogState is 3 then axis is tied to joy vertical
                        If I = 1 Then
                            TempH = "X"
                            TempV = "1"
                            dirH = ","
                            dirV = "*,"
                        Else
                            TempH = TempH & "X"
                            TempV = TempV & "1"
                            dirH = dirH & ","
                            dirV = dirV & "*,"
                        End If
                        
                    Case Else
                        MsgBox ("Error joyJogState Conversion:" & Chr(13) & "Index: " & I & ", State: " & joyJogState(I))
                End Select
            Else    ' If i = 4 (Oscillator Axis)
                TempH = TempH & "X"
                TempV = TempV & "X"
                dirH = dirH & ","
                dirV = dirV & ","
            End If
        Next I
    
    
        '-- Begin motion based on current input
        Dim tempDirection As String
        
        Select Case joyStkInputs
            Case 0
    
            Case 1, 2
                'Set the direction based on the joystick input
                If joyStkInputs = 1 Then tempDirection = Replace(dirH, "*", "+") Else tempDirection = Replace(dirH, "*", "-")
                'Output Continuous motion mode, direction, and go command to 6k
                c6k.Write "MC" & TempH & ":D" & tempDirection & ":GO" & TempH & Chr(13)
                
            Case 3, 4
                'Set the direction based on the joystick input
                If joyStkInputs = 3 Then tempDirection = Replace(dirV, "*", "+") Else tempDirection = Replace(dirV, "*", "-")
                'Output Continuous motion mode, direction, and go command to 6k
                c6k.Write "MC" & TempV & ":D" & tempDirection & ":GO" & TempV & Chr(13)
                
            Case Else
                MsgBox "Error in joystick inputs"
        End Select
    End If
    
End If

'-- If either of the joystick axes require run-mode velocity Boost modifications
If joyModeX = 2 Or joyModeY = 2 Then

    Call getJoyStkInputs

    If joyStkInputs <> joyLastStkInputs Then
    
        '-- Set/Clear boost bool for each state
        'For all values in joyJogState (except index 0, which is unused)
        For I = 1 To 6
            If I <> 4 Then  'For all i except 4 (Osc)
            
                Select Case joyJogState(I)
                    Case 0, 1         ' If jogState is 0 (axis inactive) or 1 (standard jog active) No Boost
                        
                        userSpeedBoost(I) = False
                        
                    Case 2          'If jogState is 1 then axis Boost is tied to joy horizontal
                    
                        'If joy horizontal is in either direction, enable boost, else disable
                        If joyStkInputs = 1 Or 2 Then userSpeedBoost(I) = True Else userSpeedBoost(I) = False
                            
                    Case 3          'If jogState is 3 then axis Boost is tied to joy vertical
                        
                        'If joy vertical is in either direction, enable boost, else disable
                        If joyStkInputs = 3 Or 4 Then userSpeedBoost(I) = True Else userSpeedBoost(I) = False
                        
                    Case Else
                        MsgBox ("Error joyJogState Boost Conversion:" & Chr(13) & "Index: " & I & ", State: " & joyJogState(I))
                End Select
                
            End If
        Next I
    End If
End If

'-- If either of the joystick axes require run-mode pass velocity modifications
If joyModeX = 3 Or joyModeY = 3 Then

    Call getJoyStkInputs

    ' Increase or decrease the speedMult factor based on input. Continuously increase/Decrease while input is held
    Select Case joyStkInputs
        Case 0
            'Stick is centered, no action required
        Case 1
            If joyModeX = 3 Then userSpeedMult = userSpeedMult + speedMultInc
        Case 2
            If joyModeX = 3 Then userSpeedMult = userSpeedMult - speedMultInc
        Case 3
            If joyModeY = 3 Then userSpeedMult = userSpeedMult + speedMultInc
        Case 4
            If joyModeY = 3 Then userSpeedMult = userSpeedMult - speedMultInc
            
        Case Else
            MsgBox ("Error in joySpeed Mult Handling")
    End Select
    
End If

End Sub
