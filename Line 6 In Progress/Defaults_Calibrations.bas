Attribute VB_Name = "Defaults_Calibrations"


Public Enum lDriveCalibration

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

Public Enum iDriveLims
    
    'Minimum speeds for each drive, accounting for scaling
    nCalX
    nCalY
    nCalZ
    nCalZa
    nCalR
    
    'Maximum speeds for each drive, accounting for scaling
    xCalX
    xCalY
    xCalZ
    xCalZa
    xCalR
    
End Enum

Public Enum sDefaultSpeeds         'Default jog speeds for primary axes
    dJogX = 1
    dJogY = 1
    dJogZ = 1
    dJogZa = 1
    dJogR = 36
End Enum
