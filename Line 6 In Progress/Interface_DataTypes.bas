Attribute VB_Name = "Generic_Enums"
Enum passTypeLabels
    passBlade = 1
    passAugEdg
    passAugFac
End Enum


Public Function Trim(myInput As Double, min As Double, max As Double) As Double

    If myInput < min Then myInput = min
    If myInput > max Then myInput = max
    
    Trim = myInput
    
End Function


Public Type DriveAxes

    Axis(8) As Single

End Type
    
