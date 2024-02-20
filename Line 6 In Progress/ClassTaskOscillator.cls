VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTaskOscillator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Variables sent to the 6k for Oscillator Task
Private Const c6k_bOscTaskOn = "VARB1.5"
Private Const c6k_bOscOn = "VARB1.6"
Private Const c6k_bOscHold = "VARB1.8"
Private Const c6k_sOscVel = "VAR15"
Private Const c6k_sPassWd = "VAR16"

'Oscillator States used in the oscOps command
Enum OscStates
    oscHoldOn = 1
    oscHoldOff
    oscStart
    oscStop
End Enum

Public Sub initOsc()

    'Create Oscillator Home Program
    c6k.Write "DEL OSCHOME:DEF OSCHOME:"
        'Set busy flag
        c6k.Write "VARB1.1 = b1:"
        'Set Parameters for Homing Oscillator
        c6k.Write "1INFNC18-4T:A,,,10,,:AD,,,10,,:V,,,4,,:HOMA,,1,,:HOMAD,,,10,,:HOMV,,,5,,:HOMVF,,,0.5,,:HOMBAC.4 = b1:HOMEDG.4 = b1:HOMDF.4 = b1:"
        'Back up Osc before homing
        c6k.Write "D,,,-0.3,,:GOxxx1xxx = b1:WAIT(MOV = bxxx0xxx):"
        'Home oscillator
        c6k.Write "HOMxxx1xxx:WAIT(MOV = bxxx0xxx):T0.1:"
        'Back up osc to center and set 0
        c6k.Write "D,,,-1.5,,:GOxxx1xxx:WAIT(MOV = bxxx0xxx):PSET,,,0,,:"
        'Set hardware travel limits
        c6k.Write "LH,,,3,,:1INFNC18-4R:1INFNC17-4S:1INLVL.17 = b1:"
        'Clear busy flag
        c6k.Write "VARB1.1 = b0:"
    c6k.Write "END:"

    'Create Oscillator Run Program
    c6k.Write "DEL OSCRUN:DEF OSCRUN:"
        'While OscTaskOn bit is enabled
        c6k.Write "WHILE(VARB1.5 = b1):"
            'If the OscOn or OscRunning bits are set
            c6k.Write "IF(VARB1.6 = b1 | VARB1.7 = b1):"
                'If both the OscOn and OscRunning bits are set
                c6k.Write "IF(VARB1.6 = b1 & VARB1.7 = b1):"
                    'Wait until motion has stopped, then reverse oscilator and resume motion
                    c6k.Write "WAIT(MOV = bxxx0xxx):VAR16 = -1*VAR16:D,,,VAR16,,:GOxxx1xxx:"
                'If the Osc On bit is not set (stopping osc) or the OscRunning bit is not set (first loop)
                c6k.Write "ELSE:"
                    'Stop the oscillator, then go to oscillation inside edge (closest to operator) and wait for hold bit to drop
                    c6k.Write "!S.4 = b0:WAIT(MOV = bxxx0xxx):D,,,VAR15,,:GOxxx1xxx:WAIT(VARB1.8 = b0):V,,,VAR6,,:VAR15 = -1*VAR16:VARB1.7 = ~(VARB1.7):"
                c6k.Write "NIF:"
            'Delay 5ms between each loop
            c6k.Write "NIF:T0.005:"
        c6k.Write "NWHILE:"
    c6k.Write "END:"
            
    'Initialize Oscillator Task in Controller
    c6k.Write "1%TSKAX4,4:"
    c6k.Write "1%OSCHOME:"

End Sub

Public Sub c6kVarSet(var As String, value As String)

c6k.Write var & " = " & value & ":"

End Sub

Public Sub setOscVel()

    'Adjust pass width for oscillator motor distance
    Dim tempPassWidth As Single
    tempPassWidth = 0.5 * (passWidth - 0.875)

    'Set Oscillator pass width and speed
    Call c6kVarSet(c6k_sPassWd, Format(tempPassWidth, "#0.000"))
    Call c6kVarSet(c6k_sOscVel, Format(baseOscSpeed, "#0.000"))
    
    'Enable Oscillator task
    Call c6kVarSet(c6k_bOscTaskOn, "b1")
    c6k.Write "1%OSCRUN:"
    
End Sub

Public Sub oscOps(state As Integer)
    Select Case state
        Case oscStart
            Call c6kVarSet(c6k_bOscOn, "b1")
        Case oscStop
            Call c6kVarSet(c6k_bOscOn, "b0")
        Case oscHoldOn
            Call c6kVarSet(c6k_bOscHold, "b1")
            Call c6kVarSet(c6k_bOscRun, "b0")
        Case oscHoldOff
            Call c6kVarSet(c6k_bOscHold, "b0")
    End Select
End Sub

Public Sub stopOsc()
    Call c6kVarSet(c6k_bOscTaskOn, "b0")
End Sub

