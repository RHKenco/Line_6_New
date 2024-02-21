VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassJoystick"
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


Option Explicit

'-- Class Constants & Variables ---------------------------------------------------------------------------------------

Public Enum iMotionAxes                 'Integer Values for Motion Axes, for entering data into arrays
    axisX = driveX
    axisY = driveY
    axisZ = driveZ
    axisZa = driveZa
    axisR = driveR
    
    axisrS = 9
    
    axisYR
    axisZp
    axisZpR
End Enum

Private Enum bMotionAxes                'Sets bits corresponding to the axes operated by the joystick
    'The first byte of any axis status variable will correspond to the 8 axes on the 6k
    aX = (2 ^ (driveX - 1))
    aY = (2 ^ (driveY - 1))
    aZ = (2 ^ (driveZ - 1))
    aZa = (2 ^ (driveZa - 1))
    aR = (2 ^ (driveR - 1))
    aO = (2 ^ (driveO - 1))
    
    'The Second byte of any axis status variable will correspond to the inversion state of each axis
    aXi = (2 ^ (driveX + 7))
    aYi = (2 ^ (driveY + 7))
    aZi = (2 ^ (driveZ + 7))
    aZai = (2 ^ (driveZa + 7))
    aRi = (2 ^ (driveR + 7))
    aOi = (2 ^ (driveO + 7))
    
    'The first bit of the Third byte indicates if the axis is in speed-adjust mode
    arS = (2 ^ (17 - 1))
    
    'The remainder of the Third byte contains status information
    aUNASSIGNED = (2 ^ (18 - 1))
    aINVERTED = (2 ^ (19 - 1))
    
    
    
    'Multiple drive-axis virtual axes will contain all necessary drive axes
    aYR = aY Or aR
    aYRi = aYi
    aZp = aY Or aZ
    aZpi = aYi Or aZi
    aZpR = aY Or aZ Or aR
    aZpRi = aYi Or aZi Or aRi
End Enum

'Constants for strings associated with each axis
Private Const arSstr = "Run Speed"
Private Const aXstr = "X"
Private Const aYstr = "Y"
Private Const aZstr = "Z"
Private Const aZastr = "Za"
Private Const aRstr = "R"
Private Const aYRstr = "YR"
Private Const aZpstr = "Z'"
Private Const aZpRstr = "Z'R"

Public Enum joyAxisState
    joyOff = 0
    joyPos
    joyNeg
End Enum

Public Enabled As Boolean      'Enabled Boolean - true when the user activates the joystick
Private Active As Boolean       'Active boolean - true when joystick is doing stuff

Private lastInputs As Long      'Long to store the inputs previously seen by the joystick routine

Private state As String         'State String - To display currently active axes

Private AxisFnc(2) As Long      'Axis Function Array - contains data pairing joystick axes with c6k drives

Private AxisState(2) As Integer 'Axis State Array - See enum joyAxisState for options


Public Function getActive() As Boolean
    getActive = Active
End Function

Public Function getState() As String
    getState = state
End Function

Private Function getStr(axisIn As Long) As String

    Select Case axisIn
        Case axisX
            getStr = aXstr
    
        Case axisY
            getStr = aYstr

        Case axisZ
            getStr = aZstr
    
        Case axisR
            getStr = aRstr

        Case axisYR
            getStr = aYRstr

        Case axisZp
            getStr = aZpstr
    
        Case axisZpR
            getStr = aZpRstr
    
        Case axisrS
            getStr = arSstr
        
        Case Else
            MsgBox "Error: " & axisIn & " does not correspond with a valid axis"
    End Select

End Function

Public Sub runJoy()

    If Enabled Or Active Then
        If Enabled And Not Active Then
            lastInputs = 0
            Active = True
        ElseIf Active And Not Enabled Then
            GoTo ReleaseJoystick
        End If
        
        
        'If Enabled And Active Then do the rest of the subroutine'
        
        'Get Current Inputs on the Joystick
        Dim currentInputs As Long
        currentInputs = fsinfo.ProgIn(1) Xor bJoyAll
        
        'Only run the selection algorithm if the button inputs have changed
        If (currentInputs Xor bJoyButtons) <> (lastInputs Xor bJoyButtons) Then
            
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
                
            If Not setJoy(buttons) Then GoTo ReleaseJoystick
  
        End If

        'Only run the velocity set algorithm if the joystick input state has changed
        If (currentInputs Xor bJoyAxes) <> (lastInputs Xor bJoyAxes) Then
            Call Drives.setDriveVelocities
        End If

    End If

    Exit Sub

ReleaseJoystick:

    Active = False
    lastInputs = 0
    


End Sub


Public Function setJoy(buttons As Integer) As Boolean

    Dim passType As Integer
    passType = c6kOps.getPassType

    '-- Set joystick mode based on the current run mode.
    Select Case Drives.RunMode
        Case passNone, passStrike  '------------------------------------------------------------------------- FREE -------------------
        
            Select Case buttons
                Case 0, 4  ' No Keys Pressed or Toggle Only
                    Select Case passType
                        Case passBlade
                            AxisFnc(0) = aX
                            AxisFnc(1) = aY
                            AxisFnc(2) = aZ
                        Case passAugEdg
                            AxisFnc(0) = aYR
                            If Auger.getFlight = augerLeft Then AxisFnc(0) = AxisFnc(0) Or aRi
                            AxisFnc(1) = aY
                            AxisFnc(2) = aZ
                        Case passAugFac
                            AxisFnc(0) = aX
                            AxisFnc(1) = aZpR
                            If Auger.getFlight = augerLeft Then AxisFnc(1) = AxisFnc(1) Or aRi
                            AxisFnc(2) = aZp
                    End Select
                    
                Case 1, 5 ' Select Key Pressed - Machine Coordinate Mode
                    Select Case passType
                        Case passBlade, passAugEdg, passAugFac
                            AxisFnc(0) = aX
                            AxisFnc(1) = aY
                            AxisFnc(2) = aZa
                    End Select
                
                Case 2, 3, 6, 7 ' Release pressed in any combination
                    
                    GoTo joyRelease
                    
                Case Else
                    
                    GoTo joyErrorHand
                
            End Select
            
        Case passPaused  '------------------------------------------------------------------------- PAUSE -------------------
            Select Case buttons     'Toggle must be flipped for this case to occur, therefore only case 4-7
                Case 4      'Toggle Only
                    Select Case passType
                        Case passBlade
                            AxisFnc(0) = aX
                            AxisFnc(1) = aY
                            AxisFnc(2) = aZ
                        Case passAugEdg
                            AxisFnc(0) = aYR
                            If Auger.getFlight = augerLeft Then AxisFnc(0) = AxisFnc(0) Or aRi
                            AxisFnc(1) = aY
                            AxisFnc(2) = aZ
                        Case passAugFac
                            AxisFnc(0) = aX
                            AxisFnc(1) = aZpR
                            If Auger.getFlight = augerLeft Then AxisFnc(1) = AxisFnc(1) Or aRi
                            AxisFnc(2) = aZp
                    End Select
                Case 5      'Select Pressed
                    Select Case passType
                        Case passBlade
                            AxisFnc(0) = aX
                            AxisFnc(1) = aY
                            AxisFnc(2) = aZa
                        Case passAugEdg, passAugFac
                            AxisFnc(0) = aY
                            AxisFnc(1) = aR
                            If Auger.getFlight = augerLeft Then AxisFnc(1) = AxisFnc(1) Or aRi
                            AxisFnc(2) = aZa
                    End Select
                Case 6, 7  ' Release Pressed - Advance Program
                    GoTo joyRelease
                Case 0, 1, 2, 3     'Toggle not flipped - return to run mode
                    GoTo joyRelease
                Case Else
                    GoTo joyErrorHand
            End Select
        
        Case passRunning '---------------------------------------------------------------------------- RUN ---------------------
            Select Case buttons
                Case 0    'No buttons Pressed - Horizontal axis will adjust run speed
                    Select Case passType
                        Case passBlade, passAugEdg
                            AxisFnc(0) = arS
                            AxisFnc(1) = aY
                            AxisFnc(2) = aZ
                        Case passAugFac
                            AxisFnc(0) = arS
                            AxisFnc(1) = aX
                            AxisFnc(2) = aZ
                    End Select
                Case 1      'Select Pressed - Horizontal axis will nudge run axis
                    Select Case passType
                        Case passBlade
                            AxisFnc(0) = aX
                            AxisFnc(1) = aY
                            AxisFnc(2) = aZa
                        Case passAugEdg
                            AxisFnc(0) = aYR
                            If Auger.getFlight = augerLeft Then AxisFnc(0) = AxisFnc(0) Or aRi
                            AxisFnc(1) = aY
                            AxisFnc(2) = aZa
                        Case passAugFac
                            AxisFnc(0) = aZpR
                            If Auger.getFlight = augerLeft Then AxisFnc(0) = AxisFnc(0) Or aRi
                            AxisFnc(1) = aY
                            AxisFnc(2) = aZa
                    End Select
    
                Case 4, 5, 6, 7 'Toggle flipped
                    GoTo joyRelease
                Case 2, 3    'Release Pressed, no toggle - Ignore
                    
                Case Else
                    GoTo joyErrorHand
            End Select
        
        Case Else
            GoTo joyErrorHand
    End Select

    setJoy = True
    
    Exit Function
    
joyErrorHand:
    
    MsgBox "Error In Joystick Mode Set. Button State: " & buttons & ", Pass Type: " & passType

joyRelease:

    setJoy = False

End Function





