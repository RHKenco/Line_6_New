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



Public Sub initAuger()

Private myAuger As augerParam

    'Set all multipliers to 1
    myAuger.rotSpdMlt = 1
    myAuger.edgVelMlt = 1
    myAuger.fldVelMlt = 1
    myAuger.yMult = 1
    myAuger.zMult = 1

End Sub


Public Function setAugerParam(augerPitchIn As String, augerDiamIn As String, augerCoilDirection As String) As String

'Store auger pitch, diameter, and hand
myAuger.Pitch = augerPitchIn
myAuger.Diameter = augerDiamIn
If augerCoilDirection = "Right" Then myAuger.Hand = augerRight Else myAuger.Hand = augerLeft

'Compute Auger Angle from Pitch and Diameter
myAuger.Angle = Atn(augerPitchIn / (augerDiamIn * PI))
myAuger.AngleDeg = (myAuger.Angle * 180 / PI)

myAuger.Enabled = True

'Compute mult factors for pass velocities
myAuger.rotSpdMlt = 360 / (PI * myAuger.Diameter)
myAuger.edgVelMlt = myAuger.Pitch * myAuger.rotSpdMlt / 360

myAuger.fldVelMlt = 360 / PI
myAuger.yMult = myAuger.Pitch * myAuger.rotSpdMult * Sin(myAuger.Angle) / 360
myAuger.zMult = myAuger.Pitch * myAuger.rotSpdMult * Cos(myAuger.Angle) / 360

'Return auger angle
setAugerParam = roundToString(myAuger.AngleDeg, 2)

End Function


Public Sub clearAugerParam()

    myAuger.Enabled = False
    myAuger.Pitch = 0
    myAuger.Diameter = 0
    myAuger.Hand = 0
    myAuger.Angle = 0
    myAuger.AngleDeg = 0
    myAuger.rotSpdMlt = 1
    myAuger.edgVelMlt = 1
    myAuger.fldVelMlt = 1
    myAuger.yMult = 1
    myAuger.zMult = 1

    
End Sub

Public Function getActive() As Boolean
    If myAuger.Active Then getSet = myAuger.Enabled Else MsgBox "Auger not setup!"
End Function

Public Function getPitch() As Double
    If myAuger.Active Then getPitch = myAuger.Pitch Else MsgBox "Auger not setup!"
End Function

Public Function getDiam() As Double
    If myAuger.Active Then getDiam = myAuger.Diameter Else MsgBox "Auger not setup!"
End Function

Public Function getHand() As Integer
    If myAuger.Active Then getHand = myAuger.Hand Else MsgBox "Auger not setup!"
End Function

Public Function getAngle() As Double
    If myAuger.Active Then getAngle = myAuger.Angle Else MsgBox "Auger not setup!"
End Function

Public Function getAngleDeg() As Double
    If myAuger.Active Then getAngleDeg = myAuger.AngleDeg Else MsgBox "Auger not setup!"
End Function

Public Function getRotSpdMlt() As Double
    getRotSpdMlt = myAuger.rotSpdMlt
End Function

Public Function getEdgVelMlt() As Double
    getRotSpdMlt = myAuger.edgVelMlt
End Function

Public Function getFldVelMlt() As Double
    getRotSpdMlt = myAuger.fldVelMlt
End Function

Public Function getyMult() As Double
    getRotSpdMlt = myAuger.yMult
End Function

Public Function getzMult() As Double
    getRotSpdMlt = myAuger.rotSpdMlt
End Function

