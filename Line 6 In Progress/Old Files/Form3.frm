VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8205
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form3"
   ScaleHeight     =   8205
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   3960
      TabIndex        =   30
      Top             =   840
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1800
      Top             =   120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3960
      TabIndex        =   29
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   4200
      TabIndex        =   28
      Text            =   "Text13"
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   3000
      TabIndex        =   27
      Text            =   "Text12"
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text11 
      Height          =   372
      Left            =   4200
      TabIndex        =   22
      Text            =   "Text11"
      Top             =   5400
      Width           =   732
   End
   Begin VB.TextBox Text10 
      Height          =   372
      Left            =   3000
      TabIndex        =   21
      Text            =   "Text10"
      Top             =   5400
      Width           =   732
   End
   Begin VB.TextBox Text9 
      Height          =   372
      Left            =   1800
      TabIndex        =   20
      Text            =   "Text9"
      Top             =   5400
      Width           =   732
   End
   Begin VB.TextBox Text8 
      Height          =   372
      Left            =   4200
      TabIndex        =   17
      Text            =   "Text8"
      Top             =   4680
      Width           =   732
   End
   Begin VB.TextBox Text7 
      Height          =   372
      Left            =   3000
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   4680
      Width           =   732
   End
   Begin VB.TextBox Text6 
      Height          =   372
      Left            =   4200
      TabIndex        =   13
      Text            =   "Text6"
      Top             =   3960
      Width           =   732
   End
   Begin VB.TextBox Text5 
      Height          =   372
      Left            =   3000
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   3960
      Width           =   732
   End
   Begin VB.TextBox Text4 
      Height          =   372
      Left            =   4200
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   2760
      Width           =   732
   End
   Begin VB.TextBox Text3 
      Height          =   372
      Left            =   3000
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   2760
      Width           =   732
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   2760
      Width           =   732
   End
   Begin VB.CommandButton Command2 
      Caption         =   "degrees"
      Height          =   252
      Left            =   480
      TabIndex        =   4
      Top             =   2880
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   1920
      TabIndex        =   2
      Text            =   "5"
      Top             =   1440
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1212
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5772
      Left            =   6000
      TabIndex        =   0
      Top             =   720
      Width           =   9132
      _ExtentX        =   16113
      _ExtentY        =   10186
      _Version        =   393216
      Rows            =   190
      Cols            =   8
      FixedCols       =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   26
      Top             =   6480
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Y Speed"
      Height          =   252
      Index           =   11
      Left            =   4200
      TabIndex        =   25
      Top             =   5160
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "X Speed"
      Height          =   252
      Index           =   10
      Left            =   3000
      TabIndex        =   24
      Top             =   5160
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "RO Speed"
      Height          =   252
      Index           =   9
      Left            =   1800
      TabIndex        =   23
      Top             =   5160
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "Y DIFF"
      Height          =   252
      Index           =   7
      Left            =   4200
      TabIndex        =   19
      Top             =   4440
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "X DIFF"
      Height          =   252
      Index           =   8
      Left            =   3000
      TabIndex        =   18
      Top             =   4440
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "Y Axis Pos"
      Height          =   252
      Index           =   6
      Left            =   4200
      TabIndex        =   15
      Top             =   3720
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "X Axis Pos"
      Height          =   252
      Index           =   4
      Left            =   3000
      TabIndex        =   14
      Top             =   3720
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "ACTUAL POSITION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   480
      TabIndex        =   11
      Top             =   3600
      Width           =   1932
   End
   Begin VB.Label Label1 
      Caption         =   "Y Axis Pos"
      Height          =   252
      Index           =   3
      Left            =   4200
      TabIndex        =   10
      Top             =   2520
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "X Axis Pos"
      Height          =   252
      Index           =   2
      Left            =   3000
      TabIndex        =   9
      Top             =   2520
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "Degrees"
      Height          =   252
      Index           =   1
      Left            =   1800
      TabIndex        =   8
      Top             =   2520
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Radius"
      Height          =   252
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   1812
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Rad As Single
Dim T As Integer
MSFlexGrid1.col = 0
MSFlexGrid1.row = T
For T = 1 To 18
Rad = Text1.Text
    ' Z axis
    MSFlexGrid1.col = 0
    MSFlexGrid1.row = T
    
  Zaxis = T * 10
    MSFlexGrid1.Text = Format(Zaxis, "#0.000")
  RZaxis = T * 10 * 0.0174533
    
    
    
    
    'Time X Axis
     MSFlexGrid1.col = 1
     MSFlexGrid1.row = T
     MSFlexGrid1.Text = Format(Rad - (Rad * (Cos(RZaxis))), "#0.000")
     
      
     
    'Time X Axis
     MSFlexGrid1.col = 2
     MSFlexGrid1.row = T
     MSFlexGrid1.Text = Format((Rad * (Sin(RZaxis))), "#0.000")
     If T = 1 Then
        XSpeed = Format(MSFlexGrid1.TextMatrix(1, 1) / MSFlexGrid1.TextMatrix(1, 2), "#0.000")
        YSpeed = Format(MSFlexGrid1.TextMatrix(1, 2) / MSFlexGrid1.TextMatrix(1, 1), "#0.000")
        MSFlexGrid1.TextMatrix(1, 6) = XSpeed
        MSFlexGrid1.TextMatrix(1, 7) = YSpeed
    End If
    If T > 1 And T < 10 Then
        XSpeed = Format((MSFlexGrid1.TextMatrix(T, 1) - MSFlexGrid1.TextMatrix(T - 1, 1)) / (MSFlexGrid1.TextMatrix(T, 2) - MSFlexGrid1.TextMatrix(T - 1, 2)), "#0.000")
        YSpeed = Format((MSFlexGrid1.TextMatrix(T, 2) - MSFlexGrid1.TextMatrix(T - 1, 2)) / (MSFlexGrid1.TextMatrix(T, 1) - MSFlexGrid1.TextMatrix(T - 1, 1)), "#0.000")
         MSFlexGrid1.TextMatrix(T, 4) = Format((MSFlexGrid1.TextMatrix(T, 1) - MSFlexGrid1.TextMatrix(T - 1, 1)), "#0.000")
        MSFlexGrid1.TextMatrix(T, 5) = Format((MSFlexGrid1.TextMatrix(T, 2) - MSFlexGrid1.TextMatrix(T - 1, 2)), "#0.000")
        
        MSFlexGrid1.TextMatrix(T, 6) = XSpeed
        MSFlexGrid1.TextMatrix(T, 7) = YSpeed
    End If
    If T > 9 Then
        XSpeed = Format((MSFlexGrid1.TextMatrix(T, 1) - MSFlexGrid1.TextMatrix(T - 1, 1)) / (MSFlexGrid1.TextMatrix(T - 1, 2) - MSFlexGrid1.TextMatrix(T, 2)), "#0.000")
        YSpeed = Format((MSFlexGrid1.TextMatrix(T - 1, 2) - MSFlexGrid1.TextMatrix(T, 2)) / (MSFlexGrid1.TextMatrix(T, 1) - MSFlexGrid1.TextMatrix(T - 1, 1)), "#0.000")
        
          MSFlexGrid1.TextMatrix(T, 4) = Format((MSFlexGrid1.TextMatrix(T - 1, 1) - MSFlexGrid1.TextMatrix(T, 1)), "#0.000")
        MSFlexGrid1.TextMatrix(T, 5) = Format((MSFlexGrid1.TextMatrix(T - 1, 2) - MSFlexGrid1.TextMatrix(T, 2)), "#0.000")
                
        MSFlexGrid1.TextMatrix(T, 6) = XSpeed
        MSFlexGrid1.TextMatrix(T, 7) = YSpeed
    End If
Next T
'  MotorA = (fsinfo.MotorPos(1))
'                Let APT.Text1(12).Text = Val(MotorA)
'                Let APT.Text1(12).Text = APT.Text1(12).Text / 25000
'                APT.Text1(12).Text = Format(APT.Text1(12).Text, "0.000")
'                APT.Text1(12).Refresh


'    c6k.Write ("MA11:@A10:@AD10:@V.5:D" + Str(Yaxis) + "," + Str(Xaxis) + ",,,,:GO110000:WAIT(MOV=00):" & Chr$(13))











End Sub

Private Sub Command2_Click()
Dim EStop
Dim temp() As Byte
Timer1.Enabled = False
c6k.FSEnabled = True                'enable fast status
temp = c6k.FastStatus                  'get fast status information
Call CopyMemory(fsinfo, temp(0), 280)
 XSpeed = 0
 YSpeed = 0
c6k.Write ("MC000000:MA000000:COMEXC0:1INFNC20-D:INENXXXX1" & Chr$(13)) ' 20= E-Stop
c6k.Write ("JOG000000:1OUT.16-0:1OUT.13-1:T1:1OUT.15-1:" & Chr$(13))
c6k.Write ("COMEXS0:COMEXL0:SCALE1:LH0,0,0,0,0,0:SCLD39683,39683,510204,,62500,62500:" & Chr$(13))
'HOME ROT HEAD
'c6k.Write ("1INFNC21-5T:" & Chr$(13)) 'home limit
'c6k.Write ("1INFNC18-6T:" & Chr$(13)) ' home limit
'c6k.Write ("@A10:@AD10:@V4:D,,,,-4,.3:GO000011:" & Chr$(13))
'c6k.Write ("HOMA,,,,2,2:HOMAD,,,,50,50:@HOMZ0:HOMV,,,,3,2:HOMVF,,,,1,1:" & Chr$(13))
'c6k.Write ("HOMBAC111111:HOMEDG111100:HOMDF000011:HOM,,,,0,1:" & Chr$(13))
'c6k.Write ("WAIT(6AS=XXXX1 AND 5AS=XXXX1):T.1:D,,,,.37,1.375:GO000011:" & Chr$(13))
'c6k.Write ("WAIT(MOV=000000):PSET0,0,0,0,0,0:1OUT.12-0" & Chr$(13))
'c6k.Write ("MA11001000:COMEXC1:@A10:@AD10:GO110011:" & Chr$(13))
  
    'JOYSTICK ON

c6k.Write "1INFNC22-N:1INFNC23-M:JOYAXL1-1,1-2,1-0,1-0:JOYAXH1-0,1-0,1-1,,1-2:JOYVH13,13,13,2,1:JOYVL13,13,13,5,1:JOYA10,10,10,10,2:JOYAD30,100,100,10,10:" & Chr$(13)
c6k.Write "1JOYZ.3=1:1JOYZ.2=1:1JOYEDB.3=1.18:1JOYEDB.2=1.18:1JOYCDB.2=.5:1JOYCDB.3=.5:" & Chr$(13)
c6k.Write "JOY11101" & Chr$(13)
Label2.Caption = "Set 0,0    Press JoyStick Release"
Label2.Refresh
tom = 0
Do
    temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
    If CStr((fsinfo.ProgIn(1) And Input20) / Input20) = 0 Then
        If (Last_Pcut_State = 0) Then
            c6k.Write "!COMEXS0:" & Chr$(13)
            c6k.Write "1OUTALL10,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1" & Chr$(13)
            Last_Pcut_State = 1
            Label2.Caption = "E Stop!!!!!"
            Label2.Refresh
           ' Timer2.Enabled = True
            EStopPos = ""
            Exit Sub
        End If
    Else
        If (Last_Pcut_State = 1) Then
            Last_Pcut_State = 0
            Label2.Caption = ""
            Label7.Refresh
        End If
    End If
 
    If tom = 0 Then
        temp = c6k.FastStatus
        Call CopyMemory(fsinfo, temp(0), 280)
        If CStr((fsinfo.ProgIn(1) And Input17) / Input17) > 0 Then
            c6k.Write "!JOG000000" & Chr$(13)
            tom = 1
            Label7.Caption = "Limit Switch On "
            Label7.Refresh
        End If
    End If
 
    temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)
    If CStr((fsinfo.ProgIn(1) And Input23) / Input23) > 0 Then
       Label2.Caption = "joystick on"
       Label2.Refresh
            
    Else
        Label2.Caption = "joystick off"
        Label2.Refresh
        Exit Do
    End If
Loop ' UNTIL
c6k.Write "JOG000000:PSET0,0,0,0,0,0" & Chr$(13)
c6k.Write ("WAIT(MOV=000000):PSET0,0,0,0,0,0:1OUT.12-0" & Chr$(13))

'first 90 degrees or rotation
MasterSpeed = 0.25
c6k.Write "VAR10=" + Str(MasterSpeed) + ":" & Chr$(13)
'XSpeed = Format(MasterSpeed * MSFlexGrid1.TextMatrix(1, 6), "#0.000")
'YSpeed = Format(MasterSpeed * MSFlexGrid1.TextMatrix(1, 7), "#0.000")
CirRad = Text1.Text

c6k.Write ("MA11001000:COMEXC1:@A10:@AD10:V0,0,,," + Str(MasterSpeed) + ":" & Chr$(13))
c6k.Write ("D" + Str(CirRad) + "," + Str(CirRad * -1) + ",,,90:GO110010:" & Chr$(13))
  
Label2.Caption = 0
Do
   Label2.Caption = Label2.Caption + 1
   Label2.Refresh
    temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
    If CStr((fsinfo.ProgIn(1) And Input20) / Input20) = 0 Then
        If (Last_Pcut_State = 0) Then
            c6k.Write "!COMEXS0:" & Chr$(13)
            c6k.Write "1OUTALL10,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1" & Chr$(13)
            Last_Pcut_State = 1
            Label2.Caption = "E Stop!!!!!"
            Label2.Refresh
           ' Timer2.Enabled = True
            EStopPos = ""
            Exit Sub
        End If
    Else
        If (Last_Pcut_State = 1) Then
            Last_Pcut_State = 0
            Label2.Caption = ""
            Label2.Refresh
        End If
    End If
 '' ******* Update Motor Position ****************
    temp = c6k.FastStatus                  'get fast status information
    Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
        
    ROTATION = (fsinfo.MotorPos(5))
    Let Text2.Text = Val(ROTATION)
    Let Text2.Text = Format(((Text2.Text / 4432) / 1), "0.000")

    YMOTOR = (fsinfo.MotorPos(1))
    Let Text6.Text = Val(YMOTOR)
    Let Text6.Text = Format((Text6.Text / 39683), "0.000")
                        
    XMOTOR = (fsinfo.MotorPos(2))
    Let Text5.Text = Val(XMOTOR)
    Let Text5.Text = Format((Text5.Text / 39683), "0.000")
            
    RoSpeedSHOW = (fsinfo.MotorVel(5))
    Let Text9.Text = Val(RoSpeedSHOW)
    Let Text9.Text = Format((Text9.Text / 660), "0.000")
            
    YSpeedSHOW = (fsinfo.MotorVel(1))
    Let Text11.Text = Val(YSpeedSHOW)
    Let Text11.Text = Format((Text11.Text / 660), "0.000")
           
    XSpeedSHOW = (fsinfo.MotorVel(2))
    Let Text10.Text = Val(XSpeedSHOW)
    Let Text10.Text = Format((Text10.Text / 660), "0.000")
            
        
    RZaxis = Val(Text2.Text) * 0.0174533
    Rad = Val(Text1.Text)
    Text3.Text = Format(((Rad - (Rad * (Cos(RZaxis)))) * -1), "#0.000") 'X AXIS
    Text4.Text = Format((Rad * (Sin(RZaxis))), "#0.000") 'Y AXIS
          
    Text2.Refresh
    Text3.Refresh
    Text4.Refresh
    Text5.Refresh
    Text6.Refresh
    Text7.Refresh
    Text8.Refresh
    Text9.Refresh
    Text10.Refresh
    Text11.Refresh
    If Label2.Caption = 1 Then
    tom = 2
    End If
    
    XDiff = Val(Text3.Text) - Val(Text5.Text)
    Let Text7.Text = XDiff
            
    YDiff = Val(Text4.Text) - Val(Text6.Text)
    Let Text8.Text = YDiff
   
    '***********adj speed
 '   If CStr((fsinfo.ProgIn(1) And Input24) / Input24) = 0 Then 'pause
        If XDiff > 0 Then
            XSpeed = XSpeed - 0.01
        ElseIf XDiff < 0 Then
            XSpeed = XSpeed + 0.01
        End If
        If XSpeed < 0 Then XSpeed = 0
        If YDiff > 0 Then
            YSpeed = YSpeed + 0.01
        ElseIf YDiff < 0 Then
            YSpeed = YSpeed - 0.01
        End If
         If YSpeed < 0 Then YSpeed = 0
         If XSpeed = "" Then XSpeed = 0
 '   End If
 
 
 
 
 c6k.Write ("!V" + Str(YSpeed) + "," + Str(XSpeed) + ",,," + Str(MasterSpeed) + ":!GO110010:" & Chr$(13))
  '************* PAUSE
 '   c6k.Write ("IF(1IN=bXXXXXXXXXXXXXXXXXXXXXXX0 AND VAR12=1):!V" + Str(YSpeed) + "," + Str(XSpeed) + ",,," + Str(MasterSpeed) + ":!GO110010:VAR12=0:NIF:" & Chr$(13))
 '   c6k.Write ("IF(1IN=bXXXXXXXXXXXXXXXXXXXXXXX1) :!V0,0,,,0:!GO11XX1:VAR12=1:NIF " & Chr$(13))

'*************** X AXIS SPEED CONTROL
'    c6k.Write "IF(1IN=bXXXXXXXXXXXXXXXXXXXXX1  AND 1ANI.2>5):VAR10=VAR10-.002:V,,,,(VAR10):GOXXXX1:NIF " & Chr$(13)
'    c6k.Write "IF(1IN=bXXXXXXXXXXXXXXXXXXXXX1  AND 1ANI.2<1):VAR10=VAR10+.002:V,,,,(VAR10):GOXXXX1:NIF " & Chr$(13)
 
 Text12.Text = XSpeed
 Text13.Text = YSpeed
 Text12.Refresh
 Text13.Refresh
 
 For i = 1 To 10000000
 Next i
 
 Loop Until Text2.Text >= 90
    
  
 '***************** SECOND 90 degrees or rotation

'CirRad = Text1.Text


'c6k.Write ("D0," + Str(CirRad * 2) + ",,,180:GO110010:" & Chr$(13))
  

'Do
'    temp = c6k.FastStatus
'    Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
'    If CStr((fsinfo.ProgIn(1) And Input20) / Input20) = 0 Then
'        If (Last_Pcut_State = 0) Then
'            c6k.Write "!COMEXS0:" & Chr$(13)
'            c6k.Write "1OUTALL10,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1" & Chr$(13)
'            Last_Pcut_State = 1
'            Label2.Caption = "E Stop!!!!!"
'            Label2.Refresh
'            'Timer2.Enabled = True
'            EStopPos = ""
'            Exit Sub
'        End If
'    Else
'        If (Last_Pcut_State = 1) Then
'            Last_Pcut_State = 0
'            Label2.Caption = ""
'            Label7.Refresh
'        End If
'    End If
' '' ******* Update Motor Position ****************
'    temp = c6k.FastStatus                  'get fast status information
'    Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
'
'    ROTATION = (fsinfo.MotorPos(5))
'    Let Text2.Text = Val(ROTATION)
'    Let Text2.Text = Format(((Text2.Text / 4432) / 1), "0.000")'
'
'    YMOTOR = (fsinfo.MotorPos(1))
'    Let Text6.Text = Val(YMOTOR)
'    Let Text6.Text = Format((Text6.Text / 25000), "0.000")
'
'    XMOTOR = (fsinfo.MotorPos(2))
'    Let Text5.Text = Val(XMOTOR)
'    Let Text5.Text = Format((Text5.Text / 39683), "0.000")
'
'    RoSpeedSHOW = (fsinfo.MotorVel(5))
'    Let Text9.Text = Val(RoSpeedSHOW)
'    Let Text9.Text = Format((Text9.Text / 660), "0.000")
'
'    YSpeedSHOW = (fsinfo.MotorVel(1))
'    Let Text11.Text = Val(YSpeedSHOW)
'    Let Text11.Text = Format((Text11.Text / 660), "0.000")
'
'    XSpeedSHOW = (fsinfo.MotorVel(2))
'    Let Text10.Text = Val(XSpeedSHOW)
'    Let Text10.Text = Format((Text10.Text / 660), "0.000")
'
'
'    RZaxis = Text1.Text * 0.0174533
'    Rad = Val(Text1.Text)
'    Text3.Text = Format(Rad - (Rad * (Cos(RZaxis))), "#0.000")
'    Text4.Text = Format((Rad * (Sin(RZaxis))), "#0.000")
'
'    Text2.Refresh
'    Text3.Refresh
'    Text4.Refresh
'    Text5.Refresh
'    Text6.Refresh
'    Text7.Refresh
'    Text8.Refresh
'    Text9.Refresh
'    Text10.Refresh
'    Text11.Refresh
'
'    XDiff = Val(Text3.Text) - Val(Text5.Text)
'    Let Text7.Text = XDiff
'
'    YDiff = Val(Text4.Text) - Val(Text6.Text)
'    Let Text8.Text = YDiff
'
'    '***********adj speed
'    If CStr((fsinfo.ProgIn(1) And Input24) / Input24) = 0 Then 'pause
'        If XDiff > 0 Then
'            XSpeed = XSpeed + 0.001
'        ElseIf XDiff < 0 Then
'            XSpeed = XSpeed - 0.001
'        End If
'
'        If YDiff > 0 Then
'            YSpeed = YSpeed + 0.001
'        ElseIf YDiff < 0 Then
'            YSpeed = YSpeed - 0.001
'        End If
'    End If
'  '************* PAUSE
'    c6k.Write ("IF(1IN=bXXXXXXXXXXXXXXXXXXXXXXX0 AND VAR12=1):!V" + Str(YSpeed) + "," + Str(XSpeed) + ",,," + Str(MasterSpeed) + ":!GO110010:VAR12=0:NIF:" & Chr$(13))
'    c6k.Write ("IF(1IN=bXXXXXXXXXXXXXXXXXXXXXXX1) :!V0,0,,,0:!GO11XX1:VAR12=1:NIF " & Chr$(13))'
'
''*************** X AXIS SPEED CONTROL
'    c6k.Write "IF(1IN=bXXXXXXXXXXXXXXXXXXXXX1  AND 1ANI.2>5):VAR10=VAR10-.002:V,,,,(VAR10):GOXXXX1:NIF " & Chr$(13)
'    c6k.Write "IF(1IN=bXXXXXXXXXXXXXXXXXXXXX1  AND 1ANI.2<1):VAR10=VAR10+.002:V,,,,(VAR10):GOXXXX1:NIF " & Chr$(13)
'
' For i = 1 To 100000
' Next i
 
 
 
' Loop Until Text2.Text = 90
End Sub

Private Sub Command3_Click()
Timer1.Enabled = True

c6k.Write "1INFNC22-N:1INFNC23-M:JOYAXL1-1,1-2,1-0,1-0:JOYAXH1-0,1-0,1-1,,1-2:JOYVH15,12,14,2,4:JOYVL15,12,14,5,4:JOYA10,10,10,10,2:JOYAD30,100,100,10,10:" & Chr$(13)
c6k.Write "1JOYZ.3=1:1JOYZ.2=1:1JOYEDB.3=1.18:1JOYEDB.2=1.18:1JOYCDB.2=.5:1JOYCDB.3=.5:" & Chr$(13)
c6k.Write "JOY11101" & Chr$(13)
End Sub


Private Sub Command4_Click()
c6k.Write "JOG000000:PSET0,0,0,0,0,0" & Chr$(13)
End Sub

Private Sub Form_Load()
MSFlexGrid1.col = 0
MSFlexGrid1.row = 0

MSFlexGrid1.col = 1
MSFlexGrid1.row = 0
MSFlexGrid1.Text = "X Axis"
MSFlexGrid1.col = 2
MSFlexGrid1.Text = "Y Axis"
MSFlexGrid1.col = 3
MSFlexGrid1.Text = ""
MSFlexGrid1.col = 4
MSFlexGrid1.Text = ""
MSFlexGrid1.col = 5
MSFlexGrid1.Text = ""
MSFlexGrid1.col = 6
MSFlexGrid1.Text = "X Speed"
MSFlexGrid1.col = 7
MSFlexGrid1.Text = "Y Speed"

End Sub

Private Sub Timer1_Timer()
Dim temp() As Byte

   c6k.FSEnabled = True                'enable fast status






'Dim temp() As Byte                      'create temporary byte array
temp = c6k.FastStatus                  'get fast status information
Call CopyMemory(fsinfo, temp(0), 280)
 
 'Label2.Caption = Label2.Caption + 1
'   Label2.Refresh
'    temp = c6k.FastStatus
'    Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
'    If CStr((fsinfo.ProgIn(1) And Input20) / Input20) = 0 Then
'        If (Last_Pcut_State = 0) Then
'            c6k.Write "!COMEXS0:" & Chr$(13)
'            c6k.Write "1OUTALL10,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1" & Chr$(13)
'            Last_Pcut_State = 1
'            Label2.Caption = "E Stop!!!!!"
'            Label2.Refresh
'           ' Timer2.Enabled = True
'            EStopPos = ""
'            Exit Sub
'        End If
'    Else
'        If (Last_Pcut_State = 1) Then
'            Last_Pcut_State = 0
'            Label2.Caption = ""
'            Label2.Refresh
'        End If
'    End If
 '' ******* Update Motor Position ****************
    temp = c6k.FastStatus                  'get fast status information
    Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
        
    ROTATION = (fsinfo.MotorPos(5))
    Let Text2.Text = Val(ROTATION)
    Let Text2.Text = Format(((Text2.Text / 4432) / 1), "0.000")

    YMOTOR = (fsinfo.MotorPos(1))
    Let Text6.Text = Val(YMOTOR)
    Let Text6.Text = Format((Text6.Text / 39683), "0.000")
                        
    XMOTOR = (fsinfo.MotorPos(2))
    Let Text5.Text = Val(XMOTOR)
    Let Text5.Text = Format((Text5.Text / 39975), "0.000")
            
    'RoSpeedSHOW = (fsinfo.MotorVel(5))
    'Let Text9.Text = Val(RoSpeedSHOW)
    'Let Text9.Text = Format((Text9.Text / 660), "0.000")
            
    'YSpeedSHOW = (fsinfo.MotorVel(1))
    'Let Text11.Text = Val(YSpeedSHOW)
    'Let Text11.Text = Format((Text11.Text / 660), "0.000")
           
    'XSpeedSHOW = (fsinfo.MotorVel(2))
    'Let Text10.Text = Val(XSpeedSHOW)
    'Let Text10.Text = Format((Text10.Text / 660), "0.000")
            
        
    'RZaxis = Val(Text2.Text) * 0.0174533
    'Rad = Val(Text1.Text)
    'Text3.Text = Format(((Rad - (Rad * (Cos(RZaxis)))) * -1), "#0.000") 'X AXIS
    'Text4.Text = Format((Rad * (Sin(RZaxis))), "#0.000") 'Y AXIS
          
    Text2.Refresh
    'Text3.Refresh
    'Text4.Refresh
    Text5.Refresh
    Text6.Refresh
    'Text7.Refresh
    'Text8.Refresh
    'Text9.Refresh
    'Text10.Refresh
    'Text11.Refresh
End Sub


