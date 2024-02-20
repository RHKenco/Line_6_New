VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAuger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Enum augerFlight
    augerRight = 1
    augerLeft
End Enum

Private Active As Boolean           'Auger active / inactive boolean
    
Private Pitch As Single             'Auger Pitch in IN
Private Diameter As Single          'Auger Diameter in IN
    
Private Angle As Single             'Auger Face Angle in RADIANS
    
Private Flight As augerFlight           'Auger Flight Direction (Right Flight or Left Flight
    
Private rotationMult As Double        'Multiplier applied to rotational Velocity for Auger Edge   in DEG/IN
Private edgeVelMult As Double        'Multiplier applied to Linear Velocity for Auger Edge       in IN/IN
    
Private fieldVelMult As Double        'Linear Multiplier with Diameter removed, to allow for varying speed with diameter position
Private yMult As Double             'Y and Z multipliers for Z' Axis
Private zMult As Double

Public Sub initAuger()

    'Set all multipliers to 1
    rotationMult = 1
    edgeVelMult = 1
    fieldVelMult = 1
    yMult = 1
    zMult = 1

End Sub


Public Function setAugerParam(augerPitchIn As String, augerDiamIn As String, augerFlightDirection As Integer) As String

Active = True

'Store auger pitch, diameter, and Flight
Pitch = augerPitchIn
Diameter = augerDiamIn
If augerFligthDirection = 1 Then
    Flight = augerRight
ElseIf augerFlightDirection = 2 Then
    Flight = augerLeft
Else
    GoTo AugerError
End If

'Compute Auger Angle from Pitch and Diameter
Angle = Atn(augerPitchIn / (augerDiamIn * PI))
angleDeg = (Angle * 180 / PI)

'Compute mult factors for pass velocities
rotationMult = 360 / (PI * Diameter)
edgeVelMult = Pitch * rotationMult / 360

fieldVelMult = 360 / PI
yMult = Pitch * rotationMult * Sin(Angle) / 360
zMult = Pitch * rotationMult * Cos(Angle) / 360

'Return auger angle
setAugerParam = Format(angleDeg, "0.00")

AugerError:
Call clearAugerParam
MsgBox "Error in Auger Parameter Setup. Auger has been cleared!"

End Function


Public Sub clearAugerParam()

    Active = False
    Pitch = 0
    Diameter = 0
    Flight = 0
    Angle = 0
    angleDeg = 0
    rotationMult = 1
    edgeVelMult = 1
    fieldVelMult = 1
    yMult = 1
    zMult = 1

    
End Sub

Public Function getActive() As Boolean
    getActive = Active
End Function

Public Function getPitch() As Double
    If Active Then getPitch = Pitch Else MsgBox "Auger not setup!"
End Function

Public Function getDiam() As Double
    If Active Then getDiam = Diameter Else MsgBox "Auger not setup!"
End Function

Public Function getFlight() As Integer
    If Active Then getFlight = Flight Else MsgBox "Auger not setup!"
End Function

Public Function getAngle(Optional angleDeg As Boolean) As Double
    If Active Then getAngle = Angle Else MsgBox "Auger not setup!"
    If angleDeg Then Angle = 180 * Angle / PI
End Function


Public Function getRotationMult() As Double
    getRotationMult = rotationMult
End Function

Public Function getEdgeVelMult() As Double
    getEdgeVelMult = edgeVelMult
End Function

Public Function getFieldVelMult() As Double
    getFieldVelMult = fieldVelMult
End Function

Public Function getyMult() As Double
    getyMult = yMult
End Function

Public Function getzMult() As Double
    getzMult = zMult
End Function

