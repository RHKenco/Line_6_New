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


Enum passRunModes
    passNone
    passStrike
    passRunning
    passPaused
    passEnd
    passReset

End Enum

Public RunMode As Integer

Private userSpeedMult As Single

Private Sub getPassSpeed()

    If RunMode = passRunning Then
        Dim passType As Integer
        passType = c6kOps.getPassType
        
        Select Case passType
            Case passBlade
                Call setDriveSpeed(driveX, False, (basePassSpeed * userSpeedMult))
            Case passAugEdg
                Call setDriveSpeed(driveY, False, (basePassSpeed * userSpeedMult * Auger.getEdgeVelMult))
                Call setDriveSpeed(driveR, False, (basePassSpeed * userSpeedMult * Auger.getRotationMult))
            Case passAugFac
                Call setDriveSpeed(driveY, False, (basePassSpeed * userSpeedMult * Auger.getyMult))
                Call setDriveSpeed(driveZ, False, (basePassSpeed * userSpeedMult * Auger.getzMult))
                Call setDriveSpeed(driveR, False, (basePassSpeed * userSpeedMult * Auger.getRotationMult))
        End Select
    End If

End Sub

Public Sub setPassSpeed()

    Call getPassSpeed
    
    Call runPassVel

End Sub


Public Sub setDriveVelocities()

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
            
            'Check if the axis is in run speed mode
            If iJoyState(i) And arS Then
                If iJoyState(i) Then
                    If iJoyState(i) = 1 Then
                        userSpeedMult = userSpeedMult + speedMultInc
                        If userSpeedMult > speedMultMax Then userSpeedMult = speedMultMax
                    Else
                        userSpeedMult = userSpeedMult - speedMultInc
                        If userSpeedMult < speedMultMin Then userSpeedMult = speedMultMin
                    End If
                 End If
            End If
        
        End If
    Next i
    
    Call getPassSpeed
    
    Call runPassVel


End Sub


Private Sub runPassVel()

    Call trimAxes

    Call set6kVelVars

    c6k.Write "2%SETVEL:"

End Sub

Private Sub trimAxes()
    If Abs(newDriveSpeed(1)) > xCalX Then If newDriveSpeed(1) > 0 Then newDriveSpeed(1) = xCalX Else newDriveSpeed(1) = -xCalX
    If Abs(newDriveSpeed(2)) > xCalZ Then If newDriveSpeed(2) > 0 Then newDriveSpeed(2) = xCalZ Else newDriveSpeed(2) = -xCalZ
    If Abs(newDriveSpeed(3)) > xCalZa Then If newDriveSpeed(3) > 0 Then newDriveSpeed(3) = xCalZa Else newDriveSpeed(3) = -xCalZa
    If Abs(newDriveSpeed(6)) > xCalY Then If newDriveSpeed(6) > 0 Then newDriveSpeed(6) = xCalY Else newDriveSpeed(6) = -xCalY
    If Abs(newDriveSpeed(7)) > xCalR Then If newDriveSpeed(7) > 0 Then newDriveSpeed(7) = xCalR Else newDriveSpeed(7) = -xCalR
End Sub

Private Sub set6kVelVars()
    For i = 1 To 8
        If i <> 4 Then Call c6kOps.c6kVarSet(("VAR" & i), Format(newDriveSpeed(i), "X0.000"))
    Next i
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
    '   Note: - VAR1 is the VB6-Set Velocity; VAR9 is the Modified Velocity to deal with direction change
    c6k.Write "IF(1V<>VAR1):"
        'If the Set Direction and the Intended Velocity are in the same direction
        c6k.Write "IF(((1D > 0) & (VAR1 > 0)) | ((1D < 0) & (VAR1 < 0))):"
            'Simply set the new velocity
            c6k.Write "IF(VAR1 < 0):VAR9 = -1*VAR1:ELSE:VAR9 = VAR1:NIF:"
        'If they are not in the same direction
        c6k.Write "ELSE:"
            'If the velocity is not at zero
            c6k.Write "IF(1VEL <> 0):"
                'Set the desired velocity to zero to avoid a direction switch error
                c6k.Write "VAR9 = 0:"
            'If it has reached zero
            c6k.Write "ELSE:"
                'Set the new direction and desired velocity accordingly
                c6k.Write "IF(VAR1 > 0):D+,,,,,,:VAR9 = VAR1:ELSE:D-,,,,,,:VAR9 = -1*VAR1:NIF:"
            c6k.Write "NIF:"
        c6k.Write "NIF:"
    c6k.Write "NIF:"
    
    'For Axis 6 - Y-Axis, set velocity if program velocity does not equal the variable value
    '   Note: - VAR6 is the VB6-Set Velocity; VAR14 is the Modified Velocity to deal with direction change
    c6k.Write "IF(6V<>VAR6):"
        'If the Set Direction and the Intended Velocity are in the same direction
        c6k.Write "IF(((6D > 0) & (VAR6 > 0)) | ((6D < 0) & (VAR6 < 0))):"
            'Simply set the new velocity
            c6k.Write "IF(VAR6 < 0):VAR14 = -1*VAR6:ELSE:VAR14 = VAR6:NIF:"
        'If they are not in the same direction
        c6k.Write "ELSE:"
            'If the velocity is not at zero
            c6k.Write "IF(6VEL <> 0):"
                'Set the desired velocity to zero to avoid a direction switch error
                c6k.Write "VAR14 = 0:"
            'If it has reached zero
            c6k.Write "ELSE:"
                'Set the new direction and desired velocity accordingly
                c6k.Write "IF(VAR6 > 0):D,,,,,+,:VAR14 = VAR6:ELSE:D,,,,,-,:VAR14 = -1*VAR6:NIF:"
            c6k.Write "NIF:"
        c6k.Write "NIF:"
    c6k.Write "NIF:"
    
    'For Axis 2 - Z-Axis, set velocity if program velocity does not equal the variable value
    '   Note: - VAR2 is the VB6-Set Velocity; VAR10 is the Modified Velocity to deal with direction change
    c6k.Write "IF(2V<>VAR2):"
        'If the Set Direction and the Intended Velocity are in the same direction
        c6k.Write "IF(((2D > 0) & (VAR2 > 0)) | ((2D < 0) & (VAR2 < 0))):"
            'Simply set the new velocity
            c6k.Write "IF(VAR2 < 0):VAR10 = -1*VAR2:ELSE:VAR10 = VAR2:NIF:"
        'If they are not in the same direction
        c6k.Write "ELSE:"
            'If the velocity is not at zero
            c6k.Write "IF(2VEL <> 0):"
                'Set the desired velocity to zero to avoid a direction switch error
                c6k.Write "VAR10 = 0:"
            'If it has reached zero
            c6k.Write "ELSE:"
                'Set the new direction and desired velocity accordingly
                c6k.Write "IF(VAR2 > 0):D,+,,,,,:VAR10 = VAR2:ELSE:D,-,,,,,:VAR10 = -1*VAR2:NIF:"
            c6k.Write "NIF:"
        c6k.Write "NIF:"
    c6k.Write "NIF:"
    
    'For Axis 3 - Za-Axis, set velocity if program velocity does not equal the variable value
    '   Note: - VAR3 is the VB6-Set Velocity; VAR11 is the Modified Velocity to deal with direction change
    c6k.Write "IF(3V<>VAR3):"
        'If the Set Direction and the Intended Velocity are in the same direction
        c6k.Write "IF(((3D > 0) & (VAR3 > 0)) | ((3D < 0) & (VAR3 < 0))):"
            'Simply set the new velocity
            c6k.Write "IF(VAR3 < 0):VAR11 = -1*VAR3:ELSE:VAR11 = VAR3:NIF:"
        'If they are not in the same direction
        c6k.Write "ELSE:"
            'If the velocity is not at zero
            c6k.Write "IF(3VEL <> 0):"
                'Set the desired velocity to zero to avoid a direction switch error
                c6k.Write "VAR11 = 0:"
            'If it has reached zero
            c6k.Write "ELSE:"
                'Set the new direction and desired velocity accordingly
                c6k.Write "IF(VAR3 > 0):D,,+,,,,:VAR11 = VAR3:ELSE:D,,-,,,,:VAR11 = -1*VAR3:NIF:"
            c6k.Write "NIF:"
        c6k.Write "NIF:"
    c6k.Write "NIF:"
    
    'For Axis 7 - R-Axis, set velocity if program velocity does not equal the variable value
    '   Note: - VAR7 is the VB6-Set Velocity; VAR70 is the Modified Velocity to deal with direction change
    c6k.Write "IF(7V<>VAR7):"
        'If the Set Direction and the Intended Velocity are in the same direction
        c6k.Write "IF(((7D > 0) & (VAR7 > 0)) | ((7D < 0) & (VAR7 < 0))):"
            'Simply set the new velocity
            c6k.Write "IF(VAR7 < 0):VAR15 = -1*VAR7:ELSE:VAR15 = VAR7:NIF:"
        'If they are not in the same direction
        c6k.Write "ELSE:"
            'If the velocity is not at zero
            c6k.Write "IF(7VEL <> 0):"
                'Set the desired velocity to zero to avoid a direction switch error
                c6k.Write "VAR15 = 0:"
            'If it has reached zero
            c6k.Write "ELSE:"
                'Set the new direction and desired velocity accordingly
                c6k.Write "IF(VAR7 > 0):D,,,,,,+:VAR15 = VAR7:ELSE:D,,,,,,-:VAR15 = -1*VAR7:NIF:"
            c6k.Write "NIF:"
        c6k.Write "NIF:"
    c6k.Write "NIF:"
    
    'Set velocity & Initiate Motion
    c6k.Write "V(VAR8),(VAR10),(VAR11),,,(VAR9),(VAR15):GO111XX11:"

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
