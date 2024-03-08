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
'|====================================================================================================================|
'|
'|              ----- Auger Class -----
'|
'|
'|
'|====================================================================================================================|

Option Explicit


'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Setup - Constants & Variables ---===---===---===---===---===---===---===---===---===---===

'--------------------------- Class Enumerations & Constants -----------------------------------------------------------

    Public Enum augerFlight
        augerNull = 0
        augerRight
        augerLeft
    End Enum
    
'--------------------------- Class Variables --------------------------------------------------------------------------

    'Public Variables
    Public Active As Boolean            'Auger active / inactive boolean
        
    Public Pitch As Single              'Auger Pitch in IN
    Public Diameter As Single           'Auger Diameter in IN

    Public Flight As Integer            'Auger Flight Direction (Right Flight or Left Flight

    'Read-Only Variables
    Private Angle As Single             'Auger Face Angle in RADIANS
        
    'Private Variables
    Private rotationMult As Double      'Multiplier applied to rotational Velocity for Auger Edge   in DEG/IN
    Private edgeVelMult As Double       'Multiplier applied to Linear Velocity for Auger Edge       in IN/IN
        
    Private fieldVelMult As Double      'Linear Multiplier with Diameter removed, to allow for varying speed with diameter position
    Private yMult As Double             'Y and Z multipliers for Z' Axis
    Private zMult As Double

'---===---===---===---===--- Setup - Constants & Variables ---===---===---===---===---===---===---===---===---===---===
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=









'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--

'--------------------------- Init & Clear Functions -------------------------------------------------------------------
    Public Sub initAuger()

        'Set all multipliers to 0
        rotationMult = 0
        edgeVelMult = 0
        fieldVelMult = 0
        yMult = 0
        zMult = 0

    End Sub


    Public Sub clearAugerParam()

        Active = False
        Pitch = 0
        Diameter = 0
        Flight = 0
        Angle = 0
        angleDeg = 0
        rotationMult = 0
        edgeVelMult = 0
        fieldVelMult = 0
        yMult = 0
        zMult = 0
        
    End Sub

'---------------------------- Read-Only Set & Returns -----------------------------------------------------------------

    Public Function getAngle(Optional angleDeg As Boolean) As Double
        If Active Then getAngle = Angle Else MsgBox "Auger not setup!"
        If angleDeg Then Angle = 180 * Angle / PI
    End Function

'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=








'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Advanced Functions - Auger Computations ---===---===---===---===---===---===---===---===--

    Public Sub setAugerParam(augerPitchIn As String, augerDiamIn As String, augerFlightDirection As Integer)

    Active = True

    'Store auger pitch, diameter, and Flight
    Pitch = augerPitchIn
    Diameter = augerDiamIn
    If augerFlightDirection = 1 Then
        Flight = augerRight
    ElseIf augerFlightDirection = 2 Then
        Flight = augerLeft
    Else
        GoTo AugerError
    End If

    'Compute Auger Angle from Pitch and Diameter
    Angle = Atn(augerPitchIn / (augerDiamIn * PI))

    'Compute mult factors for pass velocities
    rotationMult = 360 / (PI * Diameter)
    edgeVelMult = Pitch * rotationMult / 360

    fieldVelMult = 360 / PI
    yMult = Pitch * rotationMult * Sin(Angle) / 360
    zMult = Pitch * rotationMult * Cos(Angle) / 360

AugerError:
    Call clearAugerParam
    MsgBox "Error in Auger Parameter Setup. Auger has been cleared!"

    End Function


    Public Function augerVel(Axis As Integer) As DriveAxes
    
        Dim myAxes As DriveAxes
        
        Select Case Axis
            Case AxisYR
                myAxes.Axis(DriveR) = rotationMult
                myAxes.Axis(DriveY) = edgeVelMult
            Case AxisZp
                myAxes.Axis(DriveY) = yMult
                myAxes.Axis(DriveZ) = zMult
            Case AxisZpR
                myAxes.Axis(DriveR) = rotationMult
                myAxes.Axis(DriveY) = yMult
                myAxes.Axis(DriveZ) = zMult
            
            Case Else
                MsgBox "Invalid Axis Input to Auger Computations"
        End Select
    
        augerVel = myAxes
    
    End Function
    
    Public Function augerField(Diameter As Single) As Single
    
        augerField = Diameter * fieldVelMult
    
    End Function


'---===---===---===---===--- Advanced Functions - Auger Computations ---===---===---===---===---===---===---===---===--
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
