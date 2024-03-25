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


'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Setup - Constants & Variables ---===---===---===---===---===---===---===---===---===---===

'--------------------------- Class Enumerations & Constants -----------------------------------------------------------
    
    Private Enum JoyButtonState
    
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
    
    joyNoBtn
    joySelOnly
    joyRelOnly
    joySelRel
    joyTogOnly
    joySelTog
    joyRelTog
    joyPressAll

    End Enum
    
    Private Enum joyAxisState
        joyOff = 0
        joyPos
        joyNeg = -1
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

'--------------------------- Class Variables --------------------------------------------------------------------------

    'Public (Editable) Variables
    Public Enabled As Boolean      'Enabled Boolean - true when the user activates the joystick

    'Read-Only Varibles (Associated with a me.getXXXX() Function
    Private Active As Boolean       'Active boolean - true when joystick is doing stuff
    Private state As String         'State String - To display currently active axes

    'Private Variables
    Private lastInputs As Long      'Long to store the inputs previously seen by the joystick routine

    Private AxisFnc(2) As Long      'Axis Function Array - contains data pairing joystick axes with c6k drives

    Private AxisInv(2) As Boolean   'Axis Inversion Array
    
    Private joyMult As DriveAxes    'Custom Type containing array of drive values
    
'---===---===---===---===--- Setup - Constants & Variables ---===---===---===---===---===---===---===---===---===---===
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=









'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--

'--------------------------- Init & Clear Functions -------------------------------------------------------------------

    Public Sub ClearJoy()

        If Active = True Then MsgBox ("Joystick Still Active!"): Exit Sub
        
        Active = False
        Enabled = False
        
        lastInputs = 0
        
        For i = 0 To 2
            AxisFnc(i) = 0
            AxisInv(i) = False
        Next i

    End Sub


'---------------------------- Read-Only Set & Returns -----------------------------------------------------------------

    Public Function getActive() As Boolean
        getActive = Active
    End Function

    Public Function getState() As String
        getState = state
    End Function


'---------------------------- Basic Internal Functions ----------------------------------------------------------------

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

            Case AxisYR
                getStr = aYRstr

            Case AxisZp
                getStr = aZpstr
        
            Case AxisZpR
                getStr = aZpRstr
        
            Case axisrS
                getStr = arSstr
            
            Case Else
                MsgBox "Error: " & axisIn & " does not correspond with a valid axis"
                Enabled = False
        End Select

    End Function

    Private Sub clearMult()
    
        'Clear Joystick Mult Values
        Dim i As Integer
        For i = 1 To 8
            joyMult.Axis(i) = 0
        Next i
    
    End Sub

'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=








'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Advanced Functions - Joystick Operation ---===---===---===---===---===---===---===---===--

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
                
            If Not setJoy(buttons) Then GoTo ReleaseJoystick
  
        End If

        'Only run the velocity set algorithm if the joystick input state has changed
        If (currentInputs Xor bJoyAxes) <> (lastInputs Xor bJoyAxes) Then
            
            Call clearMult
            
            Dim Direction As Integer
            
            '-- Left-Right --
            'Set each axis according to the axis state
            If currentInputs And bjoyLeft Then
                Direction = joyNeg
            ElseIf currentInputs And bjoyRight Then
                Direction = joyPos
            Else
                Direction = joyOff
            End If
                
            'And set adjust the desired velocity accordingly
            Call runAxis(AxisFnc(1), Direction, AxisInv(1))
            
            
            '-- Front-Back --
            If currentInputs And bjoyBack Then
                Direction = joyNeg
            ElseIf currentInputs And bjoyFront Then
                Direction = joyPos
            Else
                Direction = joyOff
            End If
            
            Call runAxis(AxisFnc(2), Direction, AxisInv(2))
            
            
            '-- Up-Down --
            If currentInputs And bjoyDown Then
                Direction = joyNeg
            ElseIf currentInputs And bjoyUp Then
                Direction = joyPos
            Else
                Direction = joyOff
            End If
            
            Call runAxis(AxisFnc(3), Direction, AxisInv(3))
            
        End If

    End If

    Exit Sub

ReleaseJoystick:

    Call clearMult
    Enabled = False
    Active = False
    lastInputs = 0
    


End Sub


Private Sub runAxis(Axis As Integer, Direction As Integer, Inversion As Boolean)

    Dim temp As Single

    temp = Direction
    
    If Axis < 8 Then

        If Inversion Then temp = -1 * temp

        joyMult.Axis(Axis) = joyMult.Axis(Axis) + temp
    
    Elif Axis = 9
        joyMult.Axis(0) = joyMult.Axis(0) + Direction * cJoySpeedAdjust
        joyMult.Axis(0) = Trim(joyMult.Axis(0), cJoySpeedMin, cJoySpeedMax)
    Else
        
        Dim AugVels As myDriveVels
        
        AugVels = Auger.getVels(Axis)
        
        For i = 1 To 7
            If Inversion Then AugVels.Axis(i) = -1 * AugVels.Axis(i)
        
            joyMult.Axis(i) = joyMult.Axis(i) + AugVels.Axis(i)
        
        Next i
        
    End If

End Sub


Public Function setJoy(buttons As Integer) As Boolean

    Dim passType As Integer
    passType = c6kOps.getPassType

    '-- Set joystick mode based on the current run mode.
    Select Case Drives.RunMode
        Case passNone, passStrike  '------------------------------------------------------------------------- FREE -------------------
        
            Select Case buttons
                Case joyNoBtn, joyTogOnly  ' No Keys Pressed or Toggle Only
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
                    
                Case joySelOnly, joySelTog ' Select Key Pressed - Machine Coordinate Mode
                    Select Case passType
                        Case passBlade, passAugEdg, passAugFac
                            AxisFnc(0) = aX
                            AxisFnc(1) = aY
                            AxisFnc(2) = aZa
                    End Select
                
                Case joyRelOnly, joySelRel, joyRelTog, joyPressAll ' Release pressed in any combination
                    
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



'---===---===---===---===--- Advanced Functions - Joystick Operation ---===---===---===---===---===---===---===---===--
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=


