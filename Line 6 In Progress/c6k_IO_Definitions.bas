Attribute VB_Name = "c6k_Definitions"
'|====================================================================================================================|
'|
'|              ----- 6k Definitions Module -----
'|
'|
'|
'|====================================================================================================================|

'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- I/O Definitions ---===---===---===---===---===---===---===---===---===---===---===---===--


                    'Line 6 Compumotor I/O Chart:
                    '       - 1  - Input  - JoyLeft
                    '       - 2  - Input  - JoyRight
                    '       - 3  - Input  - JoyFront
                    '       - 4  - Input  - JoyBack
                    '       - 5  - Input  - JoyUp
                    '       - 6  - Input  - JoyDown
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

                    '       - 17 - Input  - Osc. Outside Limit Switch
                    '       - 18 - Input  - Osc. Home Limit Switch
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

'I/O Enum
Public Enum c6kIO
'----- Joystick Inputs ------------------------------------------------------------------------------------------------
    bjoyLeft = Input1
    joyLeft = 0
    bJoyRight = Input2
    joyRight = 1
    bJoyFront = Input3
    joyFront = 2
    bjoyBack = Input4
    joyBack = 3
    bJoyUp = Input5
    joyUp = 4
    bjoyDown = Input6
    joyDown = 5
    
    bJoyAxes = bjoyLeft And bJoyRight And bJoyFront And bjoyBack And bJoyUp And bjoyDown
    
    bJoySelect = Input22
    joySelect = 6
    bJoyRelease = Input23
    joyRelease = 7
    bJoyToggle = Input24
    joyToggle = 8
    
    bJoyButtons = bJoySelect And bJoyRelease And bJoyToggle
    
    bJoyAll = bJoyAxes And bJoyButtons
    
'----- Misc. Inputs ------------------------------------------------------------------------------------------------
    bInOscEOT = Input17
    bInOscHom = Input18
    bInH2O = Input19
    bInEstop = Input20
    
'----- Outputs ------------------------------------------------------------------------------------------------
    outAirblade = 9
    outLED = 10
    outLaser = 11
    outArgon = 12
    outExhaust = 13
    outTcFeed = 14
    outH2O = 15
    outWeldCt = 16
End Enum


'---===---===---===---===--- I/O Definitions ---===---===---===---===---===---===---===---===---===---===---===---===--
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=









'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Drive Definitions ---===---===---===---===---===---===---===---===---===---===---===---===


'-- Drive Definitions -------------------------------------------------------------------------------------------------------------------

                    'Line 6 c6k Drives:
                    '       - 1 - X-Axis
                    '       - 2 - Z-Axis
                    '       - 3 - Z-Alt Axis
                    '       - 4 - Oscillator
                    '       - 5 - UNUSED
                    '       - 6 - Y-Axis - 2 controllers ganged
                    '       - 7 - Auger Rotation
                    '       - 8 - UNUSED

Public Enum iDriveAxis          'Stores the integer Definitions of the 8 c6k Axes
    driveX = 1
    DriveZ
    driveZa
    driveO
    driveU1         'Unused Drive # 5
    DriveY
    DriveR
    driveU2         'Unused Drive # 8
End Enum

Public Enum iMotionAxes                 'Defines motion axes, which may differ from drive axes or include multiple drive axes
    axisX = driveX
    axisY = DriveY
    axisZ = DriveZ
    axisZa = driveZa
    axisR = DriveR
    
    axisrS = 9          'Run Speed Axis Definition
    
    AxisYR
    AxisZp
    AxisZpR
End Enum


'---===---===---===---===--- Drive Definitions ---===---===---===---===---===---===---===---===---===---===---===---===
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=






'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Variable Definitions ---===---===---===---===---===---===---===---===---===---===---===---






'Variables sent to the 6k for Oscillator Task
Private Const c6k_bOscTaskOn = "VARB1.5"
Private Const c6k_bOscOn = "VARB1.6"
Private Const c6k_bOscHold = "VARB1.8"
Private Const c6k_sOscVel = "VAR4"
Private Const c6k_sPassWd = "VAR12"




