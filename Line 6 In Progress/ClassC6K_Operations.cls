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




Option Explicit

'-- c6k Constants and Variables
Private FSinputState As Long        'Variable to store FastStatus input Data
Private FSoutputState As Long       'Variable to store FastStatus output Data

Private posCal(6) As Long                'Array to store position calibration data. These following constant values will be stored upon initialization,
Private Const posCalX = 26550 'counts/in        '26550      '       and will convert positional data from counts to appropriate units for each axis
Private Const posCalY = 38462 'counts/in        '62500
Private Const posCalZ = 253971 'counts/in       '39683
Private Const posCalZa = 255102 'counts/in      '510204
Private Const posCalO = 7987 'counts/in         '62500
Private Const posCalR = 2166 'counts/degree     '62500

Private revCal(6) As Long                'Array to store revolution calibration data. These following constant values will be stored upon initialization,
Private Const revCalX = 25000 'rev          '       and will convert motor revolutions to appropriate units for each axis. Current numbers are simple conversion to revolutions
Private Const revCalY = 25000 'rev
Private Const revCalZ = 25000 'rev
Private Const revCalZa = 25000 'rev
Private Const revCalO = 25000 'rev
Private Const revCalR = 25000 'rev
                                
'-- Critical input booleans and related variables
Private setE_Stop As Boolean    'Local E-stop variable

Private chkH2O As Boolean       'Local Water Sensor Boolean
Private countH2O As Integer     'Water Flow Counter
Private Const H2Ocount As Integer = 5   'Time between loss of signal and stopping machine, in seconds
Private Const H2Oconst As Integer = 20  'FSM Period in Hz


'-- Pass Type & Definition Variables
Private passWidth As Single     'Pass Width variable

Private passType As Integer

Private augerSet As Boolean         'Auger Active Bool
Private augerPitch As Double        'Auger Pitch, Diameter, & Angle
Private augerDiam As Double
Private augerAngle As Double
Private augerCoil As String

'-- Pass Speed Calculation Variables & Constants
Private passSpeed(3) As Single      'Array to store pass speeds at different pass witdths
Private Const passSpeed0 As Single = 0.125      ' - Defaults
Private Const passSpeed1 As Single = 0.119
Private Const passSpeed2 As Single = 0.1
Private Const passSpeed3 As Single = 0.075

Private oscSpeed(3) As Single   'Array to store oscillation speeds at different pass widths
Private Const oscSpeed0 As Single = 3.5
Private Const oscSpeed1 As Single = 3.4
Private Const oscSpeed2 As Single = 3#
Private Const oscSpeed3 As Single = 2.8

Private passThreshold(2) As Single 'Array to store pass witdth thresholds for determining base speeds
Private Const passThresh0 As Single = 1.25
Private Const passThresh1 As Single = 2#
Private Const passThresh2 As Single = 2.5

Private basePassSpeed As String     'Variables to store base speeds
Private baseOscSpeed As String
Private baseRotVel As String
Private baseVelY As String
Private baseVelZ As String

Private Const speedSetupPath As String = "C:\line 6\files\speedSetup6.asc"  '"F:\apps\exe\line 6\files\speedSetup6.asc"


'============================================================ Class Setup ==================================================================================================
Public Sub Enable()

' Ensure FSEnabled
c6k.FSEnabled = True

'Set up variables
FSinputState = 0
FSoutputState = 0

passType = passBlade

setE_Stop = False

Call loadPassSpeed

End Sub

Public Sub setPassWidth()

On Error GoTo ErrHand

Dim tempPassW As String

tempPassW = frmLine6.Text_Enter_Pass_Width.Text

If tempPassW = "" Then
    frmLine6.Text_Enter_Pass_Width.Text = "0.875"
    passWidth = 0.875
Else
    passWidth = CSng(tempPassW)
    
    'Apply limits to pass width
    If passWidth < 0.875 Then
        passWidth = 0.875
        frmLine6.Text_Enter_Pass_Width.Text = "0.875"
    ElseIf passWidth > 2.5 Then
        passWidth = 2.5
        frmLine6.Text_Enter_Pass_Width.Text = "2.500"
    End If
End If

Call setBaseSpeed

'Initialize Oscillator with new pass width
Call initOsc

Exit Sub

ErrHand:
    MsgBox "Error in pass width setting - String likely could not be entered as numerical value"
    End
Resume

End Sub

Public Sub defaultPassSpeed()

passSpeed(0) = passSpeed0
passSpeed(1) = passSpeed1
passSpeed(2) = passSpeed2
passSpeed(3) = passSpeed3

oscSpeed(0) = oscSpeed0
oscSpeed(1) = oscSpeed1
oscSpeed(2) = oscSpeed2
oscSpeed(3) = oscSpeed3

passThreshold(0) = passThresh0
passThreshold(1) = passThresh1
passThreshold(2) = passThresh2

' If prior user-input file exists, delete it
If Dir(speedSetupPath) <> "" Then Kill speedSetupPath
    
End Sub

Public Sub loadPassSpeed()

'If no prior user-defined speed profile is set, reset speeds to default
If Dir(speedSetupPath) = "" Then
    Call defaultPassSpeed
Else
    ' If it does exist, load the speeds
    Open speedSetupPath For Input As #1
    Input #1, passSpeed(0), passSpeed(1), passSpeed(2), passSpeed(3), oscSpeed(0), oscSpeed(1), oscSpeed(2), oscSpeed(3), passThreshold(0), passThreshold(1), passThreshold(2)
    Close #1
    
    End If

End Sub

Public Sub setPassSpeed(spdP0$, spdP1$, spdP2$, spdP3$, spdO0$, spdO1$, spdO2$, spdO3$, thr0$, thr1$, thr2$)

passSpeed(0) = spdP0$
passSpeed(1) = spdP1$
passSpeed(2) = spdP2$
passSpeed(3) = spdP3$

oscSpeed(0) = spdO0$
oscSpeed(1) = spdO1$
oscSpeed(2) = spdO2$
oscSpeed(3) = spdO3$

passThreshold(0) = thr0$
passThreshold(1) = thr1$
passThreshold(2) = thr2$

End Sub

Public Sub savePassSpeed()

Open speedSetupPath For Output As #1
Write #1, passSpeed(0), passSpeed(1), passSpeed(2), passSpeed(3), oscSpeed(0), oscSpeed(1), oscSpeed(2), oscSpeed(3), passThreshold(0), passThreshold(1), passThreshold(2)
Close #1

End Sub

Public Sub dispPassSpeed()

Dim i As Integer
For i = 0 To 3

    frmCalibrate.Text_Pass_Speed(i) = passSpeed(i)
    frmCalibrate.Text_Osc_Speed(i) = oscSpeed(i)
    
    If i <> 3 Then frmCalibrate.Text_Pass_Thrsh(i) = passThreshold(i)

Next i

End Sub

Private Sub setBaseSpeed()

'Set speed to match osc pass width based on set thresholds
If (passWidth = 0.875) Then                                                 'Pass width at base of 0.875, no oscillation
    basePassSpeed = passSpeed(0)
    baseOscSpeed = 0
ElseIf (passWidth > 0.875) And (passWidth < passThreshold(0)) Then
    basePassSpeed = passSpeed(0)
    baseOscSpeed = oscSpeed(0)
ElseIf (passWidth >= passThreshold(0)) And (passWidth < passThreshold(1)) Then
    basePassSpeed = passSpeed(1)
    baseOscSpeed = oscSpeed(1)
ElseIf (passWidth >= passThreshold(1)) And (passWidth < passThreshold(2)) Then
    basePassSpeed = passSpeed(2)
    baseOscSpeed = oscSpeed(2)
ElseIf passWidth >= passThreshold(2) Then
    basePassSpeed = passSpeed(3)
    baseOscSpeed = oscSpeed(3)
Else
    MsgBox ("Error in Speed Selection")
    Exit Sub
End If

'Reset speedmult to 1 (as this function will be called before a pass is started
Call Joy.resetSpeedMult

End Sub

Public Function setAugerParam(augerPitchIn As String, augerDiamIn As String, augerCoilDirection As String) As String

'Compute Auger Angle from Pitch and Diameter
augerAngle = Atn(augerPitchIn / (augerDiamIn * PI))

'Store Auger Pitch & Diameter
augerPitch = CDbl(augerPitchIn)
augerDiam = CDbl(augerDiamIn)

'Store Auger Coil Direction
If augerCoilDirection = "Right" Or augerCoilDirection = "Left" Then
    augerCoil = augerCoilDirection
Else
    MsgBox ("Invalid Auger Coil Direction")
End If

augerSet = True

'Return auger angle
setAugerParam = roundToString((augerAngle * 180 / PI), 2)

End Function

Public Sub clearAugerParam()
    augerAngle = 0
    augerPitch = 0
    
    augerSet = False
    
End Sub

Public Function getAugerDir() As String
    If augerSet Then getAugerDir = augerCoil Else MsgBox "Auger not setup!"
End Function

Public Function getAugerSet() As Boolean
    getAugerSet = augerSet
End Function

'============================================================ c6k Generic =================================================================================================

'-- updFastStatus() - Function to call Fast Status; only runs if up-to-date flag is false
Public Sub updFastStatus()

'Pull c6k Fast-Status Data
Dim temp() As Byte
temp = c6k.FastStatus
Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
    
FSinputState = fsinfo.ProgIn(1)
FSoutputState = fsinfo.ProgOut(1)

End Sub


Public Function getPassType() As Integer
    getPassType = passType
End Function

Public Sub setPassType()

If augerSet = False Then
    passType = passBlade
Else
    If frmLine6.Option_Auger_Direction(0).value = True Then
        passType = passAugFac
    Else
        passType = passAugEdg
    End If
End If

End Sub

Public Sub setUserSpeedBoost(b1 As Boolean, b2 As Boolean, b3 As Boolean, b4 As Boolean, b5 As Boolean, b6 As Boolean)

userSpeedBoost(1) = b1
userSpeedBoost(2) = b2
userSpeedBoost(3) = b3
userSpeedBoost(4) = b4
userSpeedBoost(5) = b5
userSpeedBoost(6) = b6

End Sub

'============================================================ Drive Setup & Testing =================================================================================================
'-- bootDrives() - boots up drives in c6k ------------------------------------------------------------------------------
Public Sub bootDrives()

' -- Load drive speed values

Dim tempPosCal As String
Dim tempRevCal As String


'Store defaults in calibration array
posCal(1) = posCalX
posCal(6) = posCalY
posCal(2) = posCalZ
posCal(3) = posCalZa
posCal(4) = posCalO
posCal(5) = posCalR

revCal(1) = revCalX
revCal(6) = revCalY
revCal(2) = revCalZ
revCal(3) = revCalZa
revCal(4) = revCalO
revCal(5) = revCalR

'Generate strings for c6k communication
Dim i As Integer

For i = 1 To 6
    
    If i = 1 Then
        tempPosCal = CStr(posCal(i))
        tempRevCal = CStr(revCal(i))
    Else
        tempPosCal = tempPosCal & "," & CStr(posCal(i))
        tempRevCal = tempRevCal & "," & CStr(revCal(i))
    End If

Next i


' -- Set Up & Enable Drives

'Set stop behavior - When stop command recieved, will decelerate to stop and clear command buffer
c6k.Write (":COMEXS0" & Chr$(13))

'Reset Motors - Denergizes (Drive0) drives 0-7, Sets Motor Polarity Direction to Standard, Disables Drive Fault Input Check
c6k.Write ("DRIVE0,0,0,0,0,0,0,0:CMDDIR000000:DRFEN00000000" & Chr$(13))

'Set Up Axis Defititions - Sets drives to Stepper mode, Sets Drive Resolution to 250000, Sets Distance Scales, Sets Acceleration Scale for All Axes,
'  Sets Velocity Scale for All Axes, Enables Scale Factors, Disables End-of-Travel Limit Input Checking, Sets Hard Limit Deceleration to 100
c6k.Write "AXSDEF00000000:@DRES250000:SCLD" & tempPosCal & ":@SCLA" & tempRevCal & ":@SCLV" & tempRevCal & ":SCALE1:LH0,0,0,0,0,0,0,0:LHAD100,100,100,100,100,100,100,100:" + Chr$(13)

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
c6k.Write ("WAIT(6AS=XXX1X):T.1:D,,,-1,,:GO0001000:VAR1=0" & Chr$(13))

End Sub

'============================================================ c6k General Input =========================================================================================
'-- getInputState() - Diagnotsic function to get full binary input state from c6k
Public Function getInputState() As Long

'Refresh input data

Call updFastStatus

getInputState = FSinputState
    
End Function

'-- inputStateStr() - Diagnostic function to allow for display of inputs in String format -----------------------------------
'   - Inputs: None
'   - Outputs: String - contains string of input data
Public Function getInputStateStr() As String

'Refresh input data

Call updFastStatus

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
Call updFastStatus

getOutputState = FSoutputState

End Function

Public Sub stopAllOut()
    c6k.Write "!S:!JOG000000:!1OUTALL9,16,0:!1OUTALL25,32,0:T2" & Chr$(13)
End Sub

' -- setOutput() - Function to set the state of an output
'   - Inputs:
'       - output - Integer corresponding to the output pin to be set. Valid values are 9-16 and 25-32
'       - value - boolean for the desired state of the output (true for on, false for off)
Public Sub setOutput(output As Integer, value As Boolean)

' If the output indicated is in either of the two output ranges, set according to the boolean input
If (output >= 9 And output <= 16) Or (output >= 25 And output <= 32) Then
    If value Then
        c6k.Write ("1OUT." & output & "-1" & Chr$(13))
    Else
        c6k.Write ("1OUT." & output & "-0" & Chr$(13))
    End If
    
Else
    MsgBox ("Error in Output Handler:" & Chr$(13) & "Invalid Input - " & output)
End If

End Sub

'============================================================ E-Stop & H2O Sensor Handler ==========================================================================================
'-- chkE_Stop() - Emergency Stop Check function - checks c6k inputs for E-Stop and sets flag accordingly
'   - Inputs: None
'   - Outputs: Boolean containing value of setE_Stop
Public Function chkE_Stop() As Boolean

'Refresh Input Data
Call updFastStatus

Dim isEstopPressed As Boolean

isEstopPressed = (FSinputState And bInEstop)


If (isEstopPressed And Not setE_Stop) Then          'If Estop is pressed and the was not pressed last time the function ran

    'Raise E-Stop Flag, disable jog mode on all axes, shut down all c6k outputs, and wait 2 seconds before accepting further commands
    setE_Stop = True
    Call activeEstop

    'Update form
    frmLine6.Label_Estop.Caption = "E-Stop Has Been Pressed!"
    frmLine6.Label_Estop.Visible = True
    frmLine6.Label_Estop.Refresh

    
ElseIf (Not isEstopPressed And setE_Stop) Then      'If Estop is not pressed but was pressed the last time the function ran

    'Lower E-Stop Flag and re-enable airblade and blower fan
    setE_Stop = False
    btnState (btnActive)
    Call setOutput(outAirblade, True)
    Call setOutput(outExhaust, True)
    frmLine6.Label_Estop.Visible = False
    frmLine6.Label_Estop.Refresh
    
End If
                                                    'If Estop is not pressed and was not pressed last time Or
                                                    'If Estop is pressed and was pressed last time
                                                    ' Then do nothing

chkE_Stop = setE_Stop

End Function

Public Function chkH2Oflow() As Boolean

'Refresh Input
Call updFastStatus

    

Dim isWaterFlowing As Boolean

isWaterFlowing = (FSinputState And bInH2O)

'E-Stop supersedes this algorithm - if the E-stop is pressed, H2O processing is irrelevant
If Not setE_Stop Then
    If isWaterFlowing Then                          'If Water is flowing, ensure count is 0
        countH2O = 0
    
    ElseIf (isWaterFlowing And Not chkH2O) Then         'If H2O not flowing and the H2O fault flag was not set last time the function ran
        
        'If the H2O counter has not reached the trigger point, increment the count
        If countH2O >= (H2Ocount * H2Oconst) Then
            countH2O = countH2O + 1
        Else
            'Raise H2O Fault Flag, disable jog mode on all axes, shut down all c6k outputs, and wait 2 seconds before accepting further commands
            chkH2O = True
            Call activeEstop
            frmLine6.Label_Estop.Caption = "Insufficient Water Flow!"
            frmLine6.Label_Estop.Visible = True
            frmLine6.Label_Estop.Refresh
        End If
        
    ElseIf (Not isWaterFlowing And chkH2O) Then        'If Estop is not pressed but was pressed the last time the function ran
    
        'Lower E-Stop Flag and re-enable airblade and blower fan
        chkH2O = False
        countH2O = 0
        btnState (btnActive)
        Call setOutput(outAirblade, True)
        Call setOutput(outExhaust, True)
        frmLine6.Label_Estop.Visible = False
        frmLine6.Label_Estop.Refresh
        
    End If
End If

chkH2Oflow = chkH2O
    

End Function

Public Sub activeEstop()

'Immediately stop everything (All axes, all outputs)
Call stopAllOut

btnState (btnEstop)

'Reset FSMs to return to WO Active State
fsmMain.State = 1
fsmRun.State = 0

'Call not-finished on WO
Call woMgr.notfinishWO

End Sub

Public Function updDro()

'update inputs
Call updFastStatus

'For each drive axis, pull position from fast status data and display on maintenance DRO
Dim temp As Double
Dim i As Integer
For i = 1 To 6
    temp = Val(fsinfo.MotorPos(i))
    frmMaintenance.Text_Pop_DRO(i - 1).Text = Format((temp / posCal(i)), "0.000")
Next i
    
End Function

'============================================================ Drive Main Operation =================================================================================================

Public Sub strikeArc()

'Capture position of z-axis - further z inputs will be relative to strike (which should be surface of blade)
c6k.Write "PSETX,0,0,X:" & Chr$(13)

'Capture current position based on pass type
Select Case passType
    Case passBlade
        c6k.Write "PSET0,X,X,X,X,0" & Chr$(13)  'Set X-Y to 0
    Case passAugFac
        c6k.Write "PSET0,X,X,X,0,0" & Chr$(13)  'Set X-Y-R to 0
    Case passAugEdg
    
        Dim tempSetPos As Double
        tempSetPos = (augerDiam / 2) - 1  'Strike position will be 1" from radial edge of auger
        
        If augerCoil = "Right" Then tempSetPos = -tempSetPos    'Facing Line6, a Right-Handed Auger will be impregnated on the Left Side such that the strike location will be -
            
        c6k.Write "PSET" & Format(tempSetPos, "0.000") & ",X,X,X,0,0" & Chr$(13)     'Set Y-R to 0 and X to 1" from radial edge. This will be used in speed calcs
    
    Case Else
        MsgBox "Error in Strike Arc Pset Routine"
End Select

'Turn on inputs as necessary
Call setOutput(outArgon, True)
Call setOutput(outAirblade, True)

'Close contacts and Move main z-axis up by 0.9"
c6k.Write "A,10:AD,3:V,13:D,0.9:GO01:1OUT.16-1:WAIT(MOV=b000000)" & Chr$(13)

End Sub

Public Sub initOsc()
    'Adjust pass with for base 0.875 pass width
    Dim tempPassWidth As Single
    tempPassWidth = passWidth - 0.875
    'Set Osc. Velocity
    c6k.Write ("V,,," & Str$(baseOscSpeed) & ",," & Chr(13))
    
    'Create oscillator program to check if motion has stopped and toggle direction
    c6k.Write ("VAR1=0:DEL OSCRUN:DEF OSCRUN" & Chr(13))
    'Check to see if the oscillator has stopped moving
    c6k.Write ("IF(MOV=bXXX0XX)" & Chr(13))
        c6k.Write ("IF(VAR1=0):D,,,-" & Str$(tempPassWidth) & ",,:VAR1=1" & Chr(13))             'If last pass was -, then set + distance & reset bool
        c6k.Write ("ELSE:D,,," & Str$(tempPassWidth) & ",,:VAR1=0:NIF" & Chr(13))      'If last pass was +, then set - distance & reset bool
    c6k.Write ("GOXXX1XX:NIF" & Chr(13))                                            'Initiate motion, increment counter
    c6k.Write ("END" & Chr(13))
    
End Sub

Public Sub runOsc()
    If baseOscSpeed > 0 Then c6k.Write ("GOSUB OSCRUN" & Chr(13))   'If oscillator is enabled, run osc.
End Sub

Public Sub returnOsc()
    'Set var3 to be negative the commanded position to return it to 0
    c6k.Write ("VAR3=-4PC:D,,,,VAR3,,:GOXXX1XX" & Chr(13))
End Sub

Public Sub RunPass(passRunState As String)

'-- Update Pass Velocity
Call setPassVel(passType)


'-- Set motor behavior
Select Case passRunState

    ' Start the motors in continuous move mode
    Case "Start"
        Select Case passType
            Case passBlade
                c6k.Write "MC100X00:D-,,,,,:GO100X00" & Chr(13)
                
            Case passAugEdg
                c6k.Write "MC000X11:D,,,,,:GO000X11" & Chr(13)
            
            Case passAugFac
                c6k.Write "MC010X11:D,,,,,:GO010X11" & Chr(13)
            
            Case Else
                MsgBox ("Error in pass type select")
                Exit Sub
            
        End Select
        
    ' Update velocity only
    Case "Update"
        Exit Sub
        
    ' Stop continuous move mode
    Case "Stop"
        c6k.Write "!S111011:MC000X00" & Chr(13)
        
    Case Else
        MsgBox ("Error in runPass function")
        
End Select

End Sub

Public Sub returnPass()
    c6k.Write ("MC000X00:D,6,,,,:GOX1XXXX:WAIT(MOVX0XXXX):VAR4=-1PC:VAR5=-5PC:VAR6=-6PC:DVAR4,,,,,VAR5,VAR6:GO1XXX11" & Chr(13))
End Sub

Private Sub setPassVel(passType As Integer)

'Create temporary variables to store pass speeds as strings
Dim vX As String
Dim vY As String
Dim vZ As String
Dim vZa As String
Dim vO As String
Dim vR As String

' Ensure Oscillator Speed is not affected by this sub
vO = ""

' Set base pass speed
Dim currentPassSpeed As Double
Dim currentRotSpeed As Double

currentPassSpeed = (basePassSpeed * Joy.getUserSpeedMult)

Dim tempSpeed1 As Double
Dim tempSpeed2 As Double
Dim tempSpeed3 As Double
tempSpeed1 = currentPassSpeed
tempSpeed2 = currentPassSpeed
tempSpeed3 = currentPassSpeed
    
Select Case passType
    Case passBlade
        
        
        'Check for speed boost
        If Joy.getUserSpeedBoost(driveX) = 1 Then tempSpeed1 = currentPassSpeed * (1 + Joy.getUserSpeedBoostMod)
        If Joy.getUserSpeedBoost(driveX) = 2 Then tempSpeed1 = currentPassSpeed * (1 - Joy.getUserSpeedBoostMod)
        
        'Set necessary Speeds
        vX = Format(tempSpeed1, "0.000")
        
        vY = ""
        vZ = ""
        vZa = ""
        vR = ""
        
    Case passAugEdg
    
        'Adjust base speed for pass type
        currentRotSpeed = currentPassSpeed / (PI * augerDiam)   'Rotational speed necessary for pass speed (linVel = (rotVel * (diam/2)) * (1rot/2PI rad)) in ROTATIONS/SEC
        currentPassSpeed = currentRotSpeed * augerPitch         'Linear speed to keep pace with the auger edge (linVel = rot/s * in/rot) in INCHES/SEC
        currentRotSpeed = currentRotSpeed * 360     'Convert to degrees per second
    
        'Check for speed boost
        If Joy.getUserSpeedBoost(driveY) = 1 Then tempSpeed1 = currentPassSpeed * (1 + Joy.getUserSpeedBoostMod)
        If Joy.getUserSpeedBoost(driveY) = 2 Then tempSpeed1 = currentPassSpeed * (1 - Joy.getUserSpeedBoostMod)
        
        If Joy.getUserSpeedBoost(driveR) = 1 Then tempSpeed3 = currentRotSpeed * (1 + Joy.getUserSpeedBoostMod)
        If Joy.getUserSpeedBoost(driveR) = 2 Then tempSpeed3 = currentRotSpeed * (1 - Joy.getUserSpeedBoostMod)
        
        'Set necessary Speeds
        vY = Format(tempSpeed1, "0.000")
        vR = Format(tempSpeed2, "0.000")
        
        vX = ""
        vZ = ""
        vZa = ""

    
    Case passAugFac
            
        Dim currentDiam As Single
            
        currentDiam = 1 * Val(fsinfo.MotorPos(1))
        If augerCoil = "Right" Then currentDiam = -currentDiam
            
        If currentDiam = 0 Then
            MsgBox "No Diameter Input in Pass Speed Calculation!"
            Exit Sub
        End If
            
        currentRotSpeed = currentPassSpeed / (PI * currentDiam)   'Rotational speed necessary for pass speed (linVel = (rotVel * (diam/2)) * (1rot/2PI rad)) in ROTATIONS/SEC
    
        Dim currentSpeedY As String
        Dim currentSpeedZ As String
        currentSpeedY = augerPitch * Sin(augerAngle) * currentRotSpeed
        currentSpeedZ = augerPitch * Cos(augerAngle) * currentRotSpeed
        
        currentRotSpeed = currentRotSpeed * 360     'DEGREES/SEC
        
        'Check for speed boost
        If Joy.getUserSpeedBoost(driveY) = 1 Then tempSpeed1 = currentSpeedY * (1 + Joy.getUserSpeedBoostMod)
        If Joy.getUserSpeedBoost(driveY) = 2 Then tempSpeed1 = currentSpeedY * (1 - Joy.getUserSpeedBoostMod)
      
        If Joy.getUserSpeedBoost(driveZ) = 1 Then tempSpeed2 = currentSpeedZ * (1 + Joy.getUserSpeedBoostMod)
        If Joy.getUserSpeedBoost(driveZ) = 2 Then tempSpeed2 = currentSpeedZ * (1 - Joy.getUserSpeedBoostMod)
      
        If Joy.getUserSpeedBoost(driveR) = 1 Then tempSpeed3 = currentRotSpeed * (1 + Joy.getUserSpeedBoostMod)
        If Joy.getUserSpeedBoost(driveR) = 2 Then tempSpeed3 = currentRotSpeed * (1 - Joy.getUserSpeedBoostMod)
        
        
        vY = Format(tempSpeed1, "0.000")
        vZ = Format(tempSpeed2, "0.000")
        vR = Format(tempSpeed3, "0.000")
        
        vX = ""
        vZa = ""
    
    Case Else
    
End Select

c6k.Write "V" & vX & "," & vZ & "," & vZa & "," & vO & "," & vR & "," & vY & Chr(13)

Select Case passType
    Case passBlade
        c6k.Write "GO1XXXXX" & Chr(13)
    Case passAugEdg
        c6k.Write "GOXXXX11" & Chr(13)
    Case passAugEdg
        c6k.Write "GOX1XX11" & Chr(13)
End Select

End Sub

