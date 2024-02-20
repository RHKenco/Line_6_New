VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassJoystick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'|====================================================================================================================|
'|
'|              ----- Joystick Class -----
'|
'|
'|
'|====================================================================================================================|


Option Explicit

'-- Class Constants & Variables ---------------------------------------------------------------------------------------

Public Enum iMotionAxes                 'Integer Values for Motion Axes, for entering data into arrays
    axisX = driveX
    axisY = driveY
    axisZ = driveZ
    axisZa = driveZa
    axisR = driveR
    
    axisrS = 9
    
    axisYR
    axisZp
    axisZpR
End Enum

Private Enum bMotionAxes                'Sets bits corresponding to the axes operated by the joystick
    'The first byte of any axis status variable will correspond to the 8 axes on the 6k
    aX = (2 ^ (driveX - 1))
    aY = (2 ^ (driveY - 1))
    aZ = (2 ^ (driveZ - 1))
    aZa = (2 ^ (driveZa - 1))
    aR = (2 ^ (driveR - 1))
    aO = (2 ^ (driveO - 1))
    
    'The Second byte of any axis status variable will correspond to the inversion state of each axis
    aXi = (2 ^ (driveX + 7))
    aYi = (2 ^ (driveY + 7))
    aZi = (2 ^ (driveZ + 7))
    aZai = (2 ^ (driveZa + 7))
    aRi = (2 ^ (driveR + 7))
    aOi = (2 ^ (driveO + 7))
    
    'The first bit of the Third byte indicates if the axis is in speed-adjust mode
    arS = (2 ^ (17 - 1))
    
    'The remainder of the Third byte contains status information
    aUNASSIGNED = (2 ^ (18 - 1))
    aINVERTED = (2 ^ (19 - 1))
    
    
    
    'Multiple drive-axis virtual axes will contain all necessary drive axes
    aYR = aYc Or aRc
    aZp = aYc Or aZc
    aZpR = aYc Or aZc Or aRc
End Enum

'Constants for strings associated with each axis
Private Const arSstr = "Run Speed"
Private Const aXstr = "X"
Private Const aYstr = "Y"
Private Const aZstr = "Z"
Private Const aZastr = "Za"
Private Const aRstr = "R"
Private Const aYRstr = "YR"
Private Const aZpstr = "Z'"
Private Const aZpRstr = "Z'R"

Public Enum joyAxisState
    joyOff = 0
    joyPos
    joyNeg
End Enum

Active As Boolean       'Active boolean - true when joystick is doing stuff

State As String      'State String - To display currently active axes

AxisFnc(2) As Long   'Axis Function Array - contains data pairing joystick axes with c6k drives

AxisState(2) As Integer 'Axis State Array - See enum joyAxisState for options

speed(8) As Double      'Array containing the speeds the joystick intends to send to the 6k


Public Function getActive() As Boolean
    getActive = Active
End Function

Public Function getState() As String
    getState = State
End Function

Private Function getStr(axisIn As Long) As String

    Select Case axisIn
        Case axisX
            getStr = aXstr
    
        Case axisY
            getStr = aYstr

        Case axisZ
            getStr = aZstr
    
        Case axisR
            getStr = aRstr

        Case axisYR
            getStr = aYRstr

        Case axisZp
            getStr = aZpstr
    
        Case axisZpR
            getStr = aZpRstr
    
        Case axisrS
            getStr = arSstr
        
        Case Else
            MsgBox "Error: " & axisIn & " does not correspond with a valid axis"
    End Select

End Function

Public Sub updateJoy()

End Sub








