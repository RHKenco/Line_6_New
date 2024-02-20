Attribute VB_Name = "c6k_Definitions"
'-- I/O Definitions -------------------------------------------------------------------------------------------------------------------

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

'Joystick Inputs
Public Const bJoyLeft As Long = Input1
Public Const joyLeft As Integer = 0
Public Const bJoyRight As Long = Input2
Public Const joyRight As Integer = 1
Public Const bJoyFront As Long = Input3
Public Const joyFront As Integer = 2
Public Const bJoyBack As Long = Input4
Public Const joyBack As Integer = 3
Public Const bJoyUp As Long = Input5
Public Const joyUp As Integer = 4
Public Const bJoyDown As Long = Input6
Public Const joyDown As Integer = 5

Public Const bJoySelect As Long = Input22
Public Const joySelect As Integer = 6
Public Const bJoyRelease As Long = Input23
Public Const joyRelease As Integer = 7
Public Const bJoyToggle As Long = Input24
Public Const joyToggle As Integer = 8

'Misc Inputs
Public Const bInOscEOT As Long = Input17
Public Const bInOscHom As Long = Input18
Public Const bInH2O As Long = Input19
Public Const bInEstop As Long = Input20

'Outputs
Public Const outAirblade As Integer = 9
Public Const outArgon As Integer = 12
Public Const outExhaust As Integer = 13
Public Const outTcFeed As Integer = 14
Public Const outH2O As Integer = 15
Public Const outWeldCt As Integer = 16


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
    driveZ
    driveZa
    driveO
    driveU1         'Unused Drive # 5
    driveY
    driveR
    driveU2         'Unused Drive # 8
End Enum




