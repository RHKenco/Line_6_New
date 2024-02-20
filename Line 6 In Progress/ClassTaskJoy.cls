VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTaskDrives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



Private Sub setDriveVelocities()

    Call resetAxisSpeed

    'For each of the 3 joystick axes
    For i = 0 To 2
        
        'If the joystick axis is active
        If myJoy.AxisState(i) Then
            'Set the axis inversion flag based on the direction of the joystick
            Dim tempInv As Boolean
            If iJoyState(i) = 2 Then tempInv = True Else tempInv = False
            
            'For each axis, If the axis bit is set
            If myJoy.AxisFnc(i) And aX Then
                'If the individual axis's inversion flag is set, toggle the inversion bool
                If myJoy.AxisFnc(i) And aXi Then tempInv = Not tempInv
                'then add to the axis speed
                Call setDriveSpeed(driveX, tempInv, getVel(driveX, myJoy.AxisFnc(i)))
            End If

            If myJoy.AxisFnc(i) And aY Then
                If myJoy.AxisFnc(i) And aYi Then tempInv = Not tempInv
                Call setDriveSpeed(driveY, tempInv, getVel(driveX, myJoy.AxisFnc(i)))
            End If

            If myJoy.AxisFnc(i) And aZ Then
                If myJoy.AxisFnc(i) And aZi Then tempInv = Not tempInv
                Call setDriveSpeed(driveZ, tempInv, getVel(driveX, myJoy.AxisFnc(i)))
            End If

            If myJoy.AxisFnc(i) And aZa Then
                If myJoy.AxisFnc(i) And aZai Then tempInv = Not tempInv
                Call setDriveSpeed(driveZa, tempInv, getVel(driveX, myJoy.AxisFnc(i)))
            End If

            If myJoy.AxisFnc(i) And aR Then
                If myJoy.AxisFnc(i) And aRi Then tempInv = Not tempInv
                Call setDriveSpeed(driveR, tempInv, getVel(driveX, myJoy.AxisFnc(i)))
            End If
        End If
    Next i


    Call trimAxes

    Call set6kVelVars

    c6k.Write "2%SETVEL:"

End Sub


Private Function getVel(drive As Integer, joyAxisFnc As Integer)

    Dim desiredVel As Double

    Select Case drive
        Case driveX
            desiredVel = dJogX
        Case driveY
            desiredVel = dJogY
            If joyAxisFnc And axisZpR Then
                desiredVel = desiredVel * Auger.getyMult()
            ElseIf joyAxisFnc And axisYR Then
                desiredVel = desiredVel * Auger.getEdgVelMult()
            End If
        Case driveZ
            desiredVel = dJogZ
            If joyAxisFnc And axisZpR Then desiredVel = desiredVel * Auger.getzMult()
        Case driveZa
            desiredVel = dJogZa
        Case driveR
            desiredVel = dJogR
            If joyAxisFnc And axisYR Then desiredVel = desiredVel * Auger.getRotSpdMult()   'This will trip for both YR and ZpR
            
        Case Else
            MsgBox "Error in desiredVel Function"
            getVel = 0
            Exit Function
    End Select

    getVel = desiredVel

End Function


Private Sub setDriveSpeed(drive As Integer, inv As Boolean, vel As Double)
    'Either add to or subtract from the desired axis velocity based on the inversion input
    If inv Then
        newDriveSpeed(drive) = newDriveSpeed(drive) - vel
    Else
        newDriveSpeed(drive) = newDriveSpeed(drive) + vel
    End If
End Sub


Private Sub set6kVelVars()

End Sub


Private Sub resetAxisSpeed()

    Dim i As Integer
    'Clear each axis
    For i = 1 To 8
        newDriveSpeed(i) = 0
    Next i
    
    If passType <> passNotRunning Then
    'Set pass speeds
    End If

End Sub




Public Sub initDrives()


'c6k Function to Set Pass Velocities based on Stored Variables
c6k.Write "DEL SETVEL:DEF SETVEL:"

    'For Axis 1 - X-Axis, set velocity if program velocity does not equal the variable value
    '   Note: - VAR1 is the VB6-Set Velocity; VAR8 is the Modified Velocity to deal with direction change
    c6k.Write "IF(1V<>VAR1):"
        'If the Set Direction and the Intended Velocity are in the same direction
        c6k.Write "IF(((1D > 0) & (VAR1 > 0)) | ((1D < 0) & (VAR1 < 0))):"
            'Simply set the new velocity
            c6k.Write "IF(VAR1 < 0):VAR8 = -1*VAR1:ELSE:VAR8 = VAR1:NIF:"
        'If they are not in the same direction
        c6k.Write "ELSE:"
            'If the velocity is not at zero
            c6k.Write "IF(1VEL <> 0):"
                'Set the desired velocity to zero to avoid a direction switch error
                c6k.Write "VAR8 = 0:"
            'If it has reached zero
            c6k.Write "ELSE:"
                'Set the new direction and desired velocity accordingly
                c6k.Write "IF(VAR1 > 0):D+,,,,,,:VAR8 = VAR1:ELSE:D-,,,,,,:VAR8 = -1*VAR1:NIF:"
            c6k.Write "NIF:"
        c6k.Write "NIF:"
    c6k.Write "NIF:"
    
    'For Axis 6 - Y-Axis, set velocity if program velocity does not equal the variable value
    '   Note: - VAR2 is the VB6-Set Velocity; VAR9 is the Modified Velocity to deal with direction change
    c6k.Write "IF(6V<>VAR2):"
        'If the Set Direction and the Intended Velocity are in the same direction
        c6k.Write "IF(((6D > 0) & (VAR2 > 0)) | ((6D < 0) & (VAR2 < 0))):"
            'Simply set the new velocity
            c6k.Write "IF(VAR2 < 0):VAR9 = -1*VAR2:ELSE:VAR9 = VAR2:NIF:"
        'If they are not in the same direction
        c6k.Write "ELSE:"
            'If the velocity is not at zero
            c6k.Write "IF(6VEL <> 0):"
                'Set the desired velocity to zero to avoid a direction switch error
                c6k.Write "VAR9 = 0:"
            'If it has reached zero
            c6k.Write "ELSE:"
                'Set the new direction and desired velocity accordingly
                c6k.Write "IF(VAR2 > 0):D,,,,,+,:VAR9 = VAR2:ELSE:D,,,,,-,:VAR9 = -1*VAR2:NIF:"
            c6k.Write "NIF:"
        c6k.Write "NIF:"
    c6k.Write "NIF:"
    
    'For Axis 2 - Z-Axis, set velocity if program velocity does not equal the variable value
    '   Note: - VAR3 is the VB6-Set Velocity; VAR10 is the Modified Velocity to deal with direction change
    c6k.Write "IF(2V<>VAR3):"
        'If the Set Direction and the Intended Velocity are in the same direction
        c6k.Write "IF(((2D > 0) & (VAR3 > 0)) | ((2D < 0) & (VAR3 < 0))):"
            'Simply set the new velocity
            c6k.Write "IF(VAR3 < 0):VAR10 = -1*VAR3:ELSE:VAR10 = VAR3:NIF:"
        'If they are not in the same direction
        c6k.Write "ELSE:"
            'If the velocity is not at zero
            c6k.Write "IF(2VEL <> 0):"
                'Set the desired velocity to zero to avoid a direction switch error
                c6k.Write "VAR10 = 0:"
            'If it has reached zero
            c6k.Write "ELSE:"
                'Set the new direction and desired velocity accordingly
                c6k.Write "IF(VAR3 > 0):D,+,,,,,:VAR10 = VAR3:ELSE:D,-,,,,,:VAR10 = -1*VAR3:NIF:"
            c6k.Write "NIF:"
        c6k.Write "NIF:"
    c6k.Write "NIF:"
    
    'For Axis 3 - Za-Axis, set velocity if program velocity does not equal the variable value
    '   Note: - VAR4 is the VB6-Set Velocity; VAR11 is the Modified Velocity to deal with direction change
    c6k.Write "IF(3V<>VAR4):"
        'If the Set Direction and the Intended Velocity are in the same direction
        c6k.Write "IF(((3D > 0) & (VAR4 > 0)) | ((3D < 0) & (VAR4 < 0))):"
            'Simply set the new velocity
            c6k.Write "IF(VAR4 < 0):VAR11 = -1*VAR4:ELSE:VAR11 = VAR4:NIF:"
        'If they are not in the same direction
        c6k.Write "ELSE:"
            'If the velocity is not at zero
            c6k.Write "IF(3VEL <> 0):"
                'Set the desired velocity to zero to avoid a direction switch error
                c6k.Write "VAR11 = 0:"
            'If it has reached zero
            c6k.Write "ELSE:"
                'Set the new direction and desired velocity accordingly
                c6k.Write "IF(VAR4 > 0):D,,+,,,,:VAR11 = VAR4:ELSE:D,,-,,,,:VAR11 = -1*VAR4:NIF:"
            c6k.Write "NIF:"
        c6k.Write "NIF:"
    c6k.Write "NIF:"
    
    'For Axis 7 - R-Axis, set velocity if program velocity does not equal the variable value
    '   Note: - VAR5 is the VB6-Set Velocity; VAR50 is the Modified Velocity to deal with direction change
    c6k.Write "IF(7V<>VAR5):"
        'If the Set Direction and the Intended Velocity are in the same direction
        c6k.Write "IF(((7D > 0) & (VAR5 > 0)) | ((7D < 0) & (VAR5 < 0))):"
            'Simply set the new velocity
            c6k.Write "IF(VAR5 < 0):VAR12 = -1*VAR5:ELSE:VAR12 = VAR5:NIF:"
        'If they are not in the same direction
        c6k.Write "ELSE:"
            'If the velocity is not at zero
            c6k.Write "IF(7VEL <> 0):"
                'Set the desired velocity to zero to avoid a direction switch error
                c6k.Write "VAR12 = 0:"
            'If it has reached zero
            c6k.Write "ELSE:"
                'Set the new direction and desired velocity accordingly
                c6k.Write "IF(VAR5 > 0):D,,,,,,+:VAR12 = VAR5:ELSE:D,,,,,,-:VAR12 = -1*VAR5:NIF:"
            c6k.Write "NIF:"
        c6k.Write "NIF:"
    c6k.Write "NIF:"
    
    'Set velocity & Initiate Motion
    c6k.Write "V(VAR8),(VAR10),(VAR11),,,(VAR9),(VAR12):GO111XX11:"

c6k.Write "END:"

'c6k Function to return pass to origin
c6k.Write "DEL RTNPASS:DEF RTNPASS:"
    'Disable MC to move based on displacement instead of velocity
    c6k.Write "MC000XX00:"
    'Lift up the head to avoid clamps
    c6k.Write "D0,3,,,,0,0:GOX1XXXXX:WAIT(MOV = bX0XXXXX):"
    'Return the head to the starting location
    c6k.Write "GO1XXXX11:WAIT(MOV = b0XXXX00):"
    'Drop the head back to the pass height
    c6k.Write "D,0,,,,,:GOX1XXXXX:WAIT(MOV = bX0XXXXX):"
    'Return to continuous motion mode
    c6k.Write "MC111XX11:"
c6k.Write "End:"

End Sub
