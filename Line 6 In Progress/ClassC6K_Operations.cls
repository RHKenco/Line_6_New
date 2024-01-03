VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassC6K_Operations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Line 6 Compumotor I/O:
'       - 1  - Input  - JoyLeft
'       - 2  - Input  - JoyRight
'       - 3  - Input  - JoyFront
'       - 4  - Input  - JoyBack
'       - 5  - UNUSED
'       - 6  - UNUSED
'       - 7  - UNUSED
'       - 8  - Input  - UNKNOWN <-----

'       - 9  - Output - Airblade for View
'       - 10 - UNUSED
'       - 11 - UNUSED
'       - 12 - Output - Argon
'       - 13 - Output - Exhaust Fan
'       - 14 - Output - TC Feeder
'       - 15 - Output - Water Pump
'       - 16 - Output - Welder Contact

'       - 17 - UNUSED
'       - 18 - Input  - Osc. Limit Switch
'       - 19 - Input  - Water Flow Switch
'       - 20 - Input  - "E-Stop"
'       - 21 - Input  - Rotate Prox Switch <--------
'       - 22 - Input  - JoySelect
'       - 23 - Input  - JoyRelease
'       - 24 - Input  - JoyToggle

'       - 25 - UNUSED
'       - 26 - UNUSED
'       - 27 - UNUSED
'       - 28 - UNUSED
'       - 29 - UNUSED
'       - 30 - UNUSED
'       - 31 - UNUSED
'       - 32 - UNUSED

'Line 6 c6k Drives:
'       - 1 - X-Axis
'       - 2 - Z-Axis
'       - 3 - Z-Alt Axis
'       - 4 - Oscillator
'       - 5 - Auger Rotation
'       - 6 - Y-Axis - 2 controllers ganged
'       - 7 - Unused?
'       - 8 - Unused

Option Explicit


'-- Constants for readability of Inputs

Private Const inEstop As Long = Input20
Private Const joySelect As Long = Input22
Private Const joyRelease As Long = Input23
Private Const joyToggle As Long = Input24

'Constants for Water Flow Counter
Private Const H2Ocount As Integer = 5   'Time between loss of signal and stopping machine, in seconds
Private Const H2Oconst As Integer = 20  'FSM Period in Hz

'-- Class Private Variables

Private FSinputState As Long      'Variable to store FastStatus input Data
Private FSoutputState As Long     'Variable to store FastStatus output Data
Private fsUpToDate As Boolean       'Boolean flag to limit FastStatus Calls - Reset in FSM

Private joyState As Integer     'Integer containing state of joystick buttons
Private joyActive As Boolean     'Boolean to determine if the joystick is active
Private joyStr As String        'String to describe current joystick state

Private setE_Stop As Boolean    'Local E-stop variable
Private chkH2O As Boolean       'Local Water Sensor Boolean
Private countH2O As Integer     'Water Flow Counter

'============================================================ Class Setup ==================================================================================================
Public Sub Enable()

' Ensure FSEnabled
c6k.FSEnabled = True

'Set up variables
FSinputState = 0
FSoutputState = 0
fsUpToDate = False

joyState = 0
joyActive = False
joyStr = ""

setE_Stop = False

End Sub

'============================================================ c6k Generic =================================================================================================

'-- updFastStatus() - Function to call Fast Status; only runs if up-to-date flag is false
Public Sub updFastStatus()

If Not fsUpToDate Then
    'Pull c6k Fast-Status Data
    Dim temp() As Byte
    temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
    
    FSinputState = fsinfo.ProgIn(1)
    FSoutputState = fsinfo.ProgOut(1)
    
    fsUpToDate = True
    
End If

End Sub

'-- resetFSupd() - Function to clear Fast Status up-to-date flag - used in loop in program
Public Sub resetFSupd()
    fsUpToDate = False
End Sub

'============================================================ Drive Setup & Testing =================================================================================================
'-- bootDrives() - boots up drives in c6k ------------------------------------------------------------------------------
Public Sub bootDrives()

' -- Set Up & Enable Drives

'Set stop behavior - When stop command recieved, will decelerate to stop and clear command buffer
c6k.Write (":COMEXS0" & Chr$(13))

'Reset Motors - Denergizes (Drive0) drives 0-7, Sets Motor Polarity Direction to Standard, Disables Drive Fault Input Check
c6k.Write ("DRIVE0,0,0,0,0,0,0,0:CMDDIR000000:DRFEN00000000" & Chr$(13))

'Set Up Axis Defititions - Sets drives to Stepper mode, Sets Drive Resolution to 250000, Sets Distance Scales, Sets Acceleration Scale for All Axes,
'  Sets Velocity Scale for All Axes, Enables Scale Factors, Disables End-of-Travel Limit Input Checking, Sets Hard Limit Deceleration to 100
c6k.Write "AXSDEF00000000:@DRES250000:SCLD26550,39683,510204,,62500,62500,62500:@SCLA25000:@SCLV25000:SCALE1:LH0,0,0,0,0,0,0,0:LHAD100,100,100,100,100,100,100,100:" + Chr$(13)

'Set Up Encoders - Set Encodor Polarity, Set Encoder Resolution, Set to Capture Encoder Position rather than Commanded Position
c6k.Write ("ENCPOL11011111:ERES4000,4000,4000,4000,4000,4000,4000,4000:ENCCNT11111111:" & Chr$(13))

'Enabling Drives
c6k.Write ("DRIVE1,1,1,1,1,1,1,0" & Chr$(13))

' Enable all inputs on block 1, Set up E-stop pin, Immediate Disable Jog Outputs
c6k.Write "1INEN:1INFNC20-D:!JOG0000000" & Chr$(13)

End Sub

'-- testDrives - Moves all drives by a short distance to verify functionality
Public Sub testDrives()

    c6k.Write "JOG000000:PSET0,0,0,0,0,0" & Chr$(13)
    c6k.Write ("D1,1,1,1,1,1:GO111111:WAIT(MOV=b0000)" & Chr$(13))
    c6k.Write ("T2:D-1,-1,-1,-1,-1,-1:GO111111:WAIT(MOV=b0000)" & Chr$(13))

End Sub

'-- homeMachine() -
Public Sub setMachineHome()
    'Set position to 0 for all axes except the oscillator (drive #4)
    c6k.Write "JOG000000:PSET0,0,0,x,0,0" & Chr$(13)
End Sub

'-- homeOsc() - Homes the oscillator
Public Sub homeOsc()

' Set osc limit switch pin (pin 18) to be the home limit for the oscilator axis (drive #4)
c6k.Write ("1INFNC18-4T:" & Chr$(13))
'Set motion parameters for osc axis and give go command to move the drive away from home
c6k.Write ("A4.10:AD4.10:V4.4:D4.-.3:GO000100:" & Chr$(13))
'set home motion parameters
c6k.Write ("HOMA4.1:HOMAD4.50:@HOMZ0:HOMV4.1:HOMVF4.1:" & Chr$(13))
'Adtl motion parameters, call home function
c6k.Write ("HOMBAC111011:HOMEDG111111:HOMDF000100:HOM,,,0,,:" & Chr$(13))
'Wait until homing motion has stopped, then move back to standby position
c6k.Write ("WAIT(6AS=XXX1X):T.1:D,,,-2.375,,:GO0001000:" & Chr$(13))

End Sub

'============================================================ c6k General Input =========================================================================================
'-- getInputState() - Diagnotsic function to get full binary input state from c6k
Public Function getInputState() As Long

'Refresh input data
If Not fsUpToDate Then
    Call updFastStatus
End If

getInputState = FSinputState
    
End Function

'-- inputStateStr() - Diagnostic function to allow for display of inputs in String format -----------------------------------
'   - Inputs: None
'   - Outputs: String - contains string of input data
Public Function getInputStateStr() As String

'Refresh input data
If Not fsUpToDate Then
    Call updFastStatus
End If

' Generate Header
Dim stateOutput As String
stateOutput = "1 2 3 4 5 6 7 8 " & Chr$(13) & Chr$(13)

' Generate first input block
If (FSinputState And Input1) Then stateOutput = "1 " Else stateOutput = "0 "
If (FSinputState And Input2) Then stateOutput = stateOutput & "1 " Else stateOutput = stateOutput & "0 "
If (FSinputState And Input3) Then stateOutput = stateOutput & "1 " Else stateOutput = stateOutput & "0 "
If (FSinputState And Input4) Then stateOutput = stateOutput & "1 " Else stateOutput = stateOutput & "0 "
If (FSinputState And Input5) Then stateOutput = stateOutput & "1 " Else stateOutput = stateOutput & "0 "
If (FSinputState And Input6) Then stateOutput = stateOutput & "1 " Else stateOutput = stateOutput & "0 "
If (FSinputState And Input7) Then stateOutput = stateOutput & "1 " Else stateOutput = stateOutput & "0 "
If (FSinputState And Input8) Then stateOutput = stateOutput & "1 " Else stateOutput = stateOutput & "0 "

'Line Break
stateOutput = stateOutput & Chr$(13)

'Generate Second Input Block
If (FSinputState And Input17) Then stateOutput = stateOutput & "1" Else stateOutput = stateOutput & "0"
If (FSinputState And Input18) Then stateOutput = stateOutput & "1" Else stateOutput = stateOutput & "0"
If (FSinputState And Input19) Then stateOutput = stateOutput & "1" Else stateOutput = stateOutput & "0"
If (FSinputState And Input20) Then stateOutput = stateOutput & "1" Else stateOutput = stateOutput & "0"
If (FSinputState And Input21) Then stateOutput = stateOutput & "1" Else stateOutput = stateOutput & "0"
If (FSinputState And Input22) Then stateOutput = stateOutput & "1" Else stateOutput = stateOutput & "0"
If (FSinputState And Input23) Then stateOutput = stateOutput & "1" Else stateOutput = stateOutput & "0"
If (FSinputState And Input24) Then stateOutput = stateOutput & "1" Else stateOutput = stateOutput & "0"

'Output State
getInputStateStr = stateOutput

End Function

'============================================================ c6k General Output =========================================================================================
'-- chkOutputs() - Checks c6k FastStatus for Outputs & stores in class ----------------------------------------------------
Public Function getOutputState() As Long

'Refresh input data
If Not fsUpToDate Then
    Call updFastStatus
End If

getOutputState = FSoutputState

End Function

Public Sub stopAllOut()
    c6k.Write "JOG000000:1OUTALL9,16,0:1OUTALL25,32,0:T2" & Chr$(13)
End Sub

Public Sub setOutput(output As String, value As Boolean)

Select Case output
    Case "Airblade"
        If value Then
            c6k.Write ("1OUT.9-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.9-0" & Chr$(13))
        End If
    Case "Argon"
        If value Then
            c6k.Write ("1OUT.12-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.12-0" & Chr$(13))
        End If
    Case "Exhaust"
        If value Then
            c6k.Write ("1OUT.13-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.13-0" & Chr$(13))
        End If
    Case "TC Feeder"
        If value Then
            c6k.Write ("1OUT.14-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.14-0" & Chr$(13))
        End If
    Case "Water Pump"
        If value Then
            c6k.Write ("1OUT.15-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.15-0" & Chr$(13))
        End If
    Case "Welder Contact"
        If value Then
            c6k.Write ("1OUT.16-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.16-0" & Chr$(13))
        End If
    Case Else
        MsgBox ("Error in Output Handler:" & Chr$(13) & "Invalid Input - " & output)
End Select

End Sub

' -- setOutputNum() - Function to set the state of a single output pin
'   - Inputs:
'       - output - Integer corresponding to the output pin to be set. Valid values are 9-16 and 25-32
'       - value - boolean for the desired state of the output (true for on, false for off)
Public Sub setOutputNum(output As Integer, value As Boolean)

Select Case output
    Case 9
        If value Then
            c6k.Write ("1OUT.9-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.9-0" & Chr$(13))
        End If
    Case 10
        If value Then
            c6k.Write ("1OUT.10-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.10-0" & Chr$(13))
        End If
    Case 11
        If value Then
            c6k.Write ("1OUT.11-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.11-0" & Chr$(13))
        End If
    Case 12
        If value Then
            c6k.Write ("1OUT.12-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.12-0" & Chr$(13))
        End If
    Case 13
        If value Then
            c6k.Write ("1OUT.13-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.13-0" & Chr$(13))
        End If
    Case 14
        If value Then
            c6k.Write ("1OUT.14-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.14-0" & Chr$(13))
        End If
    Case 15
        If value Then
            c6k.Write ("1OUT.15-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.15-0" & Chr$(13))
        End If
    Case 16
        If value Then
            c6k.Write ("1OUT.16-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.16-1" & Chr$(13))
        End If
    Case 25
        If value Then
            c6k.Write ("1OUT.25-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.25-0" & Chr$(13))
        End If
    Case 26
        If value Then
            c6k.Write ("1OUT.26-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.26-0" & Chr$(13))
        End If
    Case 27
        If value Then
            c6k.Write ("1OUT.27-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.27-0" & Chr$(13))
        End If
    Case 28
        If value Then
            c6k.Write ("1OUT.28-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.28-0" & Chr$(13))
        End If
    Case 29
        If value Then
            c6k.Write ("1OUT.29-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.29-0" & Chr$(13))
        End If
    Case 30
        If value Then
            c6k.Write ("1OUT.30-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.30-0" & Chr$(13))
        End If
    Case 31
        If value Then
            c6k.Write ("1OUT.31-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.31-0" & Chr$(13))
        End If
    Case 32
        If value Then
            c6k.Write ("1OUT.32-1" & Chr$(13))
        Else
            c6k.Write ("1OUT.32-0" & Chr$(13))
        End If
        
    Case Else
        MsgBox ("Error in Output Handler:" & Chr$(13) & "Invalid Input - " & output)
End Select
End Sub

'============================================================ E-Stop & H2O Sensor Handler ==========================================================================================
'-- chkE_Stop() - Emergency Stop Check function - checks c6k inputs for E-Stop and sets flag accordingly
'   - Inputs: None
'   - Outputs: Boolean containing value of setE_Stop
Public Function chkE_Stop() As Boolean

'Refresh Input Data
If Not fsUpToDate Then
    Call updFastStatus
End If

Dim isEstopPressed As Boolean

isEstopPressed = (FSinputState And inEstop)


If (isEstopPressed And Not setE_Stop) Then          'If Estop is pressed and the was not pressed last time the function ran

    'Raise E-Stop Flag, disable jog mode on all axes, shut down all c6k outputs, and wait 2 seconds before accepting further commands
    setE_Stop = True
    Call stopAllOut
    frmLine6.Label_Estop.Visible = True
    frmLine6.Label_Estop.Refresh
    
ElseIf (Not isEstopPressed And setE_Stop) Then      'If Estop is not pressed but was pressed the last time the function ran

    'Lower E-Stop Flag and re-enable airblade and blower fan
    setE_Stop = False
    Call setOutput("Airblade", True)
    Call setOutput("Exhaust", True)
    frmLine6.Label_Estop.Visible = False
    frmLine6.Label_Estop.Refresh
    
End If
                                                    'If Estop is not pressed and was not pressed last time Or
                                                    'If Estop is pressed and was pressed last time
                                                    ' Then do nothing

chkE_Stop = setE_Stop

End Function

Public Function chkH2O()

    If Not fsUpToDate Then
        Call updFastStatus
    End If
    
    

End Function

'============================================================ Joystick ====================================================================================
Public Function getJoyInputs() As Integer
    
'Refresh Input Data
If Not fsUpToDate Then
    Call updFastStatus
End If

Dim temp As Byte
temp = 0

'Set joyState based on inputs
If (FSinputState And joySelect) Then temp = temp Or (2 ^ (1 - 1))
If (FSinputState And joyRelease) Then temp = temp Or (2 ^ (2 - 1))
If (FSinputState And joyToggle) Then temp = temp Or (2 ^ (3 - 1))

If temp > 7 Then MsgBox ("Input Error:" & Chr$(13) & "inputState Is " & temp)

joyState = 7 And Not temp
getJoyInputs = joyState

'Based inputs, joyInputs will be:
'   - 0 - No keys pressed
'   - 1 - Select only pressed
'   - 2 - Release only pressed
'   - 3 - Select and Release pressed
'   - 4 - Toggle flipped
'   - 5 - Toggle flipped and select pressed
'   - 6 - Toggle flipped and release pressed
'   - 7 - Toggle flipped, Select pressed, and Release pressed

End Function
'-- runJoy() - main function call for joystick mode - calls appropriate functions for called mode & returns boolean status
'       - Inputs:
'           - "Enable" - Sets motors for joystick mode and enables run-mode value
'           - "Free" - Sets joystick to Free mode (Standard X-Y-Z options)
'           - "Auger" - As Free mode, but allows APT Rotation
'           - "Run" - Run mode (PathSpeed-PathTransverse-Z options)
'           - "Disable" - Disables joystick inputs
Public Function runJoy(mode As String) As Boolean
         
'If joystick Enable is called, start joystick
If mode = "Enable" Then
    Call startJoy
End If
         
'If the the joystick is active and the e-stop is not set
If joyActive = True And chkE_Stop = False Then

    '-- Variable Setup
    'Set up temporary input storage variable & get current inputs
    Dim inputs As Integer
    inputs = getJoyInputs
    
    'Set up end flag & set to default of false
    Dim endJoy As Boolean
    endJoy = False
    
    '-- Joystick Jump-Table
    Select Case mode
        Case "Free"
            endJoy = Not joyFree(inputs)
            If endJoy Then GoTo CloseJoystick
        Case "Disable"
            GoTo CloseJoystick
    End Select

End If

'Return True if joy mode is still active
runJoy = True

Exit Function

CloseJoystick:

Call stopJoy

'Return False if joy mode has been disabled
runJoy = False

End Function

Private Function joyFree(mode As Integer) As Boolean

Static mode_Last As Integer

If (mode + 1) <> mode_Last Then

    Select Case mode
        Case 0, 4  ' No Keys Pressed or Toggle Only - X-Y Mode
        
            c6k.Write ("JOG000000" & Chr$(13))
        
            c6k.Write ("1INFNC1-1J:1INFNC2-1K:1INFNC3-6J:1INFNC4-6K" & Chr$(13)) 'Set inputs 1 & 2 for +/- Axis 1 ; Set inputs 3 & 4 for +/- Axis 6

            c6k.Write ("JOG1111111" & Chr$(13)) 'Set Axes 1 & 5 to Jog Mode
            
            joyStr = "X-Y"
            
        Case 1 ' Select Key Pressed
        
            c6k.Write ("JOG000000" & Chr$(13))
            
            c6k.Write ("1INFNC1-1J:1INFNC2-1K:1INFNC3-2J:1INFNC4-2K" & Chr$(13)) 'Set inputs 1 & 2 for +/- Axis 1 ; Set inputs 3 & 4 for +/- Axis 2
            
            c6k.Write ("JOG110000" & Chr$(13)) 'Set Axes 2 & 5 to Jog Mode
            
            joyStr = "X-Z"
          
        Case 5 ' Toggle Flipped and Select Pressed
        
            c6k.Write ("JOG000000" & Chr$(13))
            
            c6k.Write ("1INFNC1-1J:1INFNC2-1K:1INFNC3-3J:1INFNC4-3K" & Chr$(13)) 'Set inputs 1 & 2 for +/- Axis 1 ; Set inputs 3 & 4 for +/- Axis 3
            
            c6k.Write ("JOG101000" & Chr$(13)) 'Set Axes 3 & 5 to Jog Mode
            
            joyStr = "X-Za"
        
        Case 2, 3, 6, 7 ' Release pressed in any combination
            
            'Return False to indicate that user input should end joystick mode
            joyFree = False
            mode_Last = 0
            Exit Function
            
        Case Else
            
            MsgBox "Error in JoyFree Handler - Case: " & mode
            joyFree = False
            mode_Last = 0
            Exit Function
        
    End Select
    
End If
    
joyFree = True

mode_Last = (mode + 1)

End Function

Private Sub startJoy()

'Immediately (!) disable jog mode for all axes
c6k.Write ("!JOG000000" & Chr$(13))

'Set motion parameters for Joystick Motion
c6k.Write ("JOG000000:JOGA4,5,5,1,5,5:JOGAD50,99,99,99,99,15:JOGVH8,8,10,2,5,3:JOGVL8,15,10,5,5,5" & Chr$(13))

joyActive = True

End Sub
Private Sub stopJoy()

'Disable jog mode on drives
c6k.Write ("JOG000000" & Chr$(13))

'Disable inputs
c6k.Write ("1INFC1-A:1INFC2-A:1INFC3-A:1INFC4-A")
joyActive = False

joyStr = ""

End Sub

Public Function getJoyActive() As Boolean
    getJoyActive = joyActive
End Function

Public Function getJoyStr() As String
    getJoyStr = joyStr
End Function

Public Function updDro()

'-- Set calibration Constants
Const Xcal = 26550
Const XcalV = 660

Const Ycal = 26550

Const Zcal = 500000

Const Ocal = 25000

Const Rcal = 4432


If Not fsUpToDate Then
    Call updFastStatus
End If

YMOTOR = (fsinfo.MotorPos(6))
    Let Text16.Text = Val(YMOTOR)
    Let Text16.Text = Format((Text16.Text / Ycal), "0.000")
                
XMOTOR = (fsinfo.MotorPos(2))
    Let Text17.Text = Val(XMOTOR)
    Let Text17.Text = Format((Text17.Text / Xcal), "0.000")
    
UDMOTOR = (fsinfo.MotorPos(3))
    Let Text14.Text = Val(UDMOTOR)
    Let Text14.Text = Format((Text14.Text / Zcal), "0.000")
    
OSSMOTOR = (fsinfo.MotorPos(6))
    Let Text15.Text = Val(OSSMOTOR)
    Let Text15.Text = Format((Text15.Text / Ocal), "0.000")
            
ROTATION = (fsinfo.MotorPos(5))
    Let Text20.Text = Val(ROTATION)
    Let Text20.Text = Format(((Text20.Text / Rcal)), "0.000")


XSpeedSHOW = (fsinfo.MotorVel(2))
    Let Text18.Text = Val(XSpeedSHOW)
    Let Text18.Text = Format((Text18.Text / XcalV), "0.000")
    
End Function

'============================================================ Drive Main Operation =================================================================================================

Public Sub strikeArc()

'Capture position of z-axis - further z inputs will be relative to strike (which should be surface of blade)
c6k.Write "PSETX,0,0,X:" & Chr$(13)
'Turn on inputs as necessary
Call setOutput("Argon", True)
Call setOutput("Airblade", True)

'Close contacts and Move main z-axis up by 0.9"
c6k.Write "A,10:AD,3:V,13:D,0.9:GO01:1OUT.16-1:WAIT(MOV=b000000):JOG100001" & Chr$(13)

End Sub

Public Sub runPass(passType As String)

Select Case passType
    Case "Blade"
        
    
    
End Select

End Sub



