VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTaskJoy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Enum iDriveAxis          'Stores the Definitions of the c6k Axes
    driveX = 1
    driveZ
    driveZa
    driveO
    driveU1         'Unused Drive # 5
    driveY
    driveR
    driveU2         'Unused Drive # 8
End Enum

Private Enum lDriveCalibration

    'Distance Scaling Value for Each Active Axis - COUNTS/IN
    dCalX = 26495
    dCalY = 39790
    dCalZ = 249896
    dCalZa = 249896
    dCalO_s = 7987
    dCalO_a = 7987
    dCalR = 5009
    
    'Velocity Scaling Value for Each Active Axis - COUNTS/IN
    vCalX = 26495
    vCalY = 39790
    vCalZ = 249896
    vCalZa = 249896
    vCalO_s = 7987
    vCalO_a = 7987
    vCalR = 5009

End Enum

Private Enum sDefaultSpeeds
    dJogX = 1
    dJogY = 1
    dJogZ = 1
    dJogZa = 1
    dJogR = 36
End Enum

Private Enum bDriveAxis         'Sets bits corresponding to each 6k axis
    bDriveX = (2 ^ (driveX - 1))
    bDriveY = (2 ^ (driveY - 1))
    bDriveZ = (2 ^ (driveZ - 1))
    bDriveZa = (2 ^ (driveZa - 1))
    bDriveO = (2 ^ (driveO - 1))
    bDriveR = (2 ^ (driveR - 1))
End Enum

Private Enum iMotionAxes                '
    axisrS
    axisX
    axisY
    axisZ
    axisZa
    axisR
    
    axisYR
    axisZp
    axisZpR
End Enum

Private Enum bMotionAxes                'Sets bits corresponding to the axes operated by the joystick
    aUNASSIGNED = (2 ^ (1 - 1))
    aINVERTED = (2 ^ (2 - 1))
    
    arS = (2 ^ (3 - 1))
    
    aX = bDriveX >> 4
    aY = bDriveY >> 4
    aZ = bDriveZ >> 4
    aZa = bDriveZa >> 4
    aR = bDriveR >> 4

    aYR = aYc Or aRc
    aZp = aYc Or aZc
    aZpR = aYc Or aZc Or aRc
End Enum

'Constants for strings associated with each axis
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

'Constants for Joystick Operation
Private joyFlag As Integer      'Joystick Flag Register

Private joyState$               'String describing joystick state

Private joybAxisFnc(2) As Integer   'Array containing the current function assigned to each joystick axis

Private passbAxisFnc As Integer     'Pass Function integer - See Enum "passType"

Private passAxisSpeed(8) As Single  'Pass Axis Speed array - contains the current velocity of each axis.







Dim tempInv As Boolean

If joybAxisFnc(i) And aX Then
    If joybAxisFnc(i) And aXi Then tempInv = True Else tempInv = False
    Call setAxisSpeed(driveX, tempInv)
End If

If joybAxisFnc(i) And aY Then
    If joybAxisFnc(i) And aYi Then tempInv = True Else tempInv = False
    Call setAxisSpeed(driveY, tempInv)
End If

If joybAxisFnc(i) And aZ Then
    If joybAxisFnc(i) And aZi Then tempInv = True Else tempInv = False
    Call setAxisSpeed(driveZ, tempInv)
End If

If joybAxisFnc(i) And aZa Then
    If joybAxisFnc(i) And aZai Then tempInv = True Else tempInv = False
    Call setAxisSpeed(driveZa, tempInv)
End If

If joybAxisFnc(i) And aR Then
    If joybAxisFnc(i) And aRi Then tempInv = True Else tempInv = False
    Call setAxisSpeed(driveR, tempInv)
End If

Call trimAxes

Call set6kVelVars

c6k.Write "2%SETVEL:"

End Sub

Private Sub setAxisSpeed(drive As Integer, inv As Boolean)
    If inv Then
        passAxisSpeed(drive) = passAxisSpeed(drive) - velDefault(drive)
    Else
        passAxisSpeed(drive) = passAxisSpeed(drive) + velDefault(drive)
    End If
End Sub

















'c6k Function to Set Pass Velocities based on Stored Variables
c6k.Write "DEL SETVEL:DEF SETVEL:"

    'For Axis 1 - X-Axis, set velocity if program velocity does not equal the variable value
    '   Note: - VAR1 is the VB6-Set Velocity; VAR6 is the Modified Velocity to deal with direction change
    c6k.Write "IF(1V<>VAR1):"
        'If the Set Direction and the Intended Velocity are in the same direction
        c6k.Write "IF(((1D > 0) & (VAR1 > 0)) | ((1D < 0) & (VAR1 < 0))):"
            'Simply set the new velocity
            c6k.Write "IF(VAR1 < 0):VAR6 = -1*VAR1:ELSE:VAR6 = VAR1:NIF:"
        'If they are not in the same direction
        c6k.Write "ELSE:"
            'If the velocity is not at zero
            c6k.Write "IF(1VEL <> 0):"
                'Set the desired velocity to zero to avoid a direction switch error
                c6k.Write "VAR6 = 0:"
            'If it has reached zero
            c6k.Write "ELSE:"
                'Set the new direction and desired velocity accordingly
                c6k.Write "IF(VAR1 > 0):D+,,,,,,:VAR6 = VAR1:ELSE:D-,,,,,,:VAR6 = -1*VAR1:NIF:"
            c6k.Write "NIF:"
        c6k.Write "NIF:"
    c6k.Write "NIF:"
    
    'For Axis 6 - Y-Axis, set velocity if program velocity does not equal the variable value
    '   Note: - VAR2 is the VB6-Set Velocity; VAR7 is the Modified Velocity to deal with direction change
    c6k.Write "IF(6V<>VAR2):"
        'If the Set Direction and the Intended Velocity are in the same direction
        c6k.Write "IF(((6D > 0) & (VAR2 > 0)) | ((6D < 0) & (VAR2 < 0))):"
            'Simply set the new velocity
            c6k.Write "IF(VAR2 < 0):VAR7 = -1*VAR2:ELSE:VAR7 = VAR2:NIF:"
        'If they are not in the same direction
        c6k.Write "ELSE:"
            'If the velocity is not at zero
            c6k.Write "IF(6VEL <> 0):"
                'Set the desired velocity to zero to avoid a direction switch error
                c6k.Write "VAR7 = 0:"
            'If it has reached zero
            c6k.Write "ELSE:"
                'Set the new direction and desired velocity accordingly
                c6k.Write "IF(VAR2 > 0):D,,,,,+,:VAR7 = VAR2:ELSE:D,,,,,-,:VAR7 = -1*VAR2:NIF:"
            c6k.Write "NIF:"
        c6k.Write "NIF:"
    c6k.Write "NIF:"
    
    'For Axis 2 - Z-Axis, set velocity if program velocity does not equal the variable value
    '   Note: - VAR3 is the VB6-Set Velocity; VAR8 is the Modified Velocity to deal with direction change
    c6k.Write "IF(2V<>VAR3):"
        'If the Set Direction and the Intended Velocity are in the same direction
        c6k.Write "IF(((2D > 0) & (VAR3 > 0)) | ((2D < 0) & (VAR3 < 0))):"
            'Simply set the new velocity
            c6k.Write "IF(VAR3 < 0):VAR8 = -1*VAR3:ELSE:VAR8 = VAR3:NIF:"
        'If they are not in the same direction
        c6k.Write "ELSE:"
            'If the velocity is not at zero
            c6k.Write "IF(2VEL <> 0):"
                'Set the desired velocity to zero to avoid a direction switch error
                c6k.Write "VAR8 = 0:"
            'If it has reached zero
            c6k.Write "ELSE:"
                'Set the new direction and desired velocity accordingly
                c6k.Write "IF(VAR3 > 0):D,+,,,,,:VAR8 = VAR3:ELSE:D,-,,,,,:VAR8 = -1*VAR3:NIF:"
            c6k.Write "NIF:"
        c6k.Write "NIF:"
    c6k.Write "NIF:"
    
    'For Axis 3 - Za-Axis, set velocity if program velocity does not equal the variable value
    '   Note: - VAR4 is the VB6-Set Velocity; VAR9 is the Modified Velocity to deal with direction change
    c6k.Write "IF(3V<>VAR4):"
        'If the Set Direction and the Intended Velocity are in the same direction
        c6k.Write "IF(((3D > 0) & (VAR4 > 0)) | ((3D < 0) & (VAR4 < 0))):"
            'Simply set the new velocity
            c6k.Write "IF(VAR4 < 0):VAR9 = -1*VAR4:ELSE:VAR9 = VAR4:NIF:"
        'If they are not in the same direction
        c6k.Write "ELSE:"
            'If the velocity is not at zero
            c6k.Write "IF(3VEL <> 0):"
                'Set the desired velocity to zero to avoid a direction switch error
                c6k.Write "VAR9 = 0:"
            'If it has reached zero
            c6k.Write "ELSE:"
                'Set the new direction and desired velocity accordingly
                c6k.Write "IF(VAR4 > 0):D,,+,,,,:VAR9 = VAR4:ELSE:D,,-,,,,:VAR9 = -1*VAR4:NIF:"
            c6k.Write "NIF:"
        c6k.Write "NIF:"
    c6k.Write "NIF:"
    
    'For Axis 7 - R-Axis, set velocity if program velocity does not equal the variable value
    '   Note: - VAR5 is the VB6-Set Velocity; VAR50 is the Modified Velocity to deal with direction change
    c6k.Write "IF(7V<>VAR5):"
        'If the Set Direction and the Intended Velocity are in the same direction
        c6k.Write "IF(((7D > 0) & (VAR5 > 0)) | ((7D < 0) & (VAR5 < 0))):"
            'Simply set the new velocity
            c6k.Write "IF(VAR5 < 0):VAR10 = -1*VAR5:ELSE:VAR10 = VAR5:NIF:"
        'If they are not in the same direction
        c6k.Write "ELSE:"
            'If the velocity is not at zero
            c6k.Write "IF(7VEL <> 0):"
                'Set the desired velocity to zero to avoid a direction switch error
                c6k.Write "VAR10 = 0:"
            'If it has reached zero
            c6k.Write "ELSE:"
                'Set the new direction and desired velocity accordingly
                c6k.Write "IF(VAR5 > 0):D,,,,,,+:VAR10 = VAR5:ELSE:D,,,,,,-:VAR10 = -1*VAR5:NIF:"
            c6k.Write "NIF:"
        c6k.Write "NIF:"
    c6k.Write "NIF:"
    
    'Set velocity & Initiate Motion
    c6k.Write "V(VAR6),(VAR8),(VAR9),,,(VAR7),(VAR10):GO111XX11:"

c6k.Write "END:"

