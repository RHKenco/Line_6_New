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

Private myAuger As augerParam

Public Sub initAuger()

    'Set all multipliers to 1
    myAuger.rotSpdMult = 1
    myAuger.edgVelMult = 1
    myAuger.fldVelMult = 1
    myAuger.yMult = 1
    myAuger.zMult = 1

End Sub


Public Function setAugerParam(augerPitchIn As String, augerDiamIn As String, augerCoilDirection As String) As String

myAuger.Active = True

'Store auger pitch, diameter, and hand
myAuger.Pitch = augerPitchIn
myAuger.Diameter = augerDiamIn
If augerCoilDirection = "Right" Then myAuger.Hand = augerRight Else myAuger.Hand = augerLeft

'Compute Auger Angle from Pitch and Diameter
myAuger.Angle = Atn(augerPitchIn / (augerDiamIn * PI))
myAuger.AngleDeg = (myAuger.Angle * 180 / PI)

'Compute mult factors for pass velocities
myAuger.rotSpdMult = 360 / (PI * myAuger.Diameter)
myAuger.edgVelMult = myAuger.Pitch * myAuger.rotSpdMult / 360

myAuger.fldVelMult = 360 / PI
myAuger.yMult = myAuger.Pitch * myAuger.rotSpdMult * Sin(myAuger.Angle) / 360
myAuger.zMult = myAuger.Pitch * myAuger.rotSpdMult * Cos(myAuger.Angle) / 360

'Return auger angle
setAugerParam = Format(myAuger.AngleDeg, "0.00")

End Function


Public Sub clearAugerParam()

    myAuger.Active = False
    myAuger.Pitch = 0
    myAuger.Diameter = 0
    myAuger.Hand = 0
    myAuger.Angle = 0
    myAuger.AngleDeg = 0
    myAuger.rotSpdMult = 1
    myAuger.edgVelMult = 1
    myAuger.fldVelMult = 1
    myAuger.yMult = 1
    myAuger.zMult = 1

    
End Sub

Public Function getActive() As Boolean
    getActive = myAuger.Active
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

Public Function getRotSpdMult() As Double
    getRotSpdMult = myAuger.rotSpdMult
End Function

Public Function getEdgVelMult() As Double
    getEdgVelMult = myAuger.edgVelMult
End Function

Public Function getFldVelMult() As Double
    getFldVelMult = myAuger.fldVelMult
End Function

Public Function getyMult() As Double
    getyMult = myAuger.yMult
End Function

Public Function getzMult() As Double
    getzMult = myAuger.rotSpdMult
End Function

