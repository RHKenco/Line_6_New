VERSION 5.00
Begin VB.Form frmFastStatus 
   Caption         =   "FastStatus Information - Ethernet Only"
   ClientHeight    =   8640
   ClientLeft      =   2385
   ClientTop       =   1695
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   9210
   Begin VB.CommandButton Command5 
      Caption         =   "E STOP"
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   8280
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3840
      TabIndex        =   10
      Text            =   "Text5"
      Top             =   8280
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   7440
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   8280
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   8280
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   7800
      Width           =   8295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "JOYSTICK"
      Height          =   255
      Left            =   7560
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   7560
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2895
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4695
      Width           =   8970
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Disable FastStatus"
      Height          =   405
      Left            =   7395
      TabIndex        =   3
      Top             =   195
      Width           =   1650
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   7290
      TabIndex        =   2
      Top             =   2385
      Width           =   1800
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7920
      Top             =   750
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7440
      Top             =   765
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   135
      ScaleHeight     =   2115
      ScaleWidth      =   6960
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   7020
   End
   Begin VB.PictureBox Grid2 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   135
      ScaleHeight     =   2115
      ScaleWidth      =   6960
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2400
      Width           =   7020
   End
End
Attribute VB_Name = "frmFastStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






















'Option Explicit


Sub CellText(grd As Control, row%, col%, msg)
    grd.row = row
    grd.col = col
    grd.Text = CStr(msg)
End Sub

Private Sub Command1_Click()
    If Timer1.Enabled Then
        Timer1.Enabled = False
        DoEvents
        c6k.FSEnabled = False
        Command1.Caption = "Enable FastStatus"
    Else
        c6k.FSUpdateRate = 100                       'set fast status update rate to 100ms
        c6k.FSEnabled = True                        'enable fast status polling
        Timer1.Enabled = True                       'enable fast status updates
        Command1.Caption = "Disable FastStatus"
    End If
End Sub


Private Sub Command2_Click()
' buffer = buffer & Chr$(13)      'append the CR
                Do
                Timer1.Enabled = False          'disable response polling to avoid simultaneous read/write
                c6k.Write ("WAIT(1IN=b1):GO" & Chr$(13))
                c6k.Write ("WAIT(1IN=b0):GO" & Chr$(13)) 'send commands to 6k
                Timer1.Enabled = True
                Loop
                
End Sub

Private Sub Command3_Click()
Timer1.Enabled = False          'disable response polling to avoid simultaneous read/write
'c6k.write "1ANIRNG.3=3:1INFNC23-O:1INFNC24-M:JOYAXH1-27:JOYVH50,:JOYVL10,:JOYA100:" & Chr$(13)
'c6k.write "1JOYEDB.3=2:1JOYZ.3=1:JOY1" & Chr$(13) 'send commands to 6k
Timer1.Enabled = True
End Sub


Private Sub Command5_Click()
'working on E-STOP ***************************
      'If (status_high% And PCUT_MASK) > 0 Then
      '  If (Last_Pcut_State = 0) Then
      '      Last_Pcut_State = 1
      '      Text9.Text = ""
      '  End If
    'Else
    '    If (Last_Pcut_State = 1) Then
    '        Last_Pcut_State = 0
    '        cmd$ = "OUT.3-0:OUT.4-0:OUT.5-0:OUT.6-0:" + Chr$(ENTER)
    '         tmp% = SendAT6400Block(Device_Address%, cmd$, 0)
    '         Text9.Text = "E-Stop"
     
     Text5.Text = CStr((fsinfo.AxisStatus(1) And Drive1) / Drive1)
     'Text5.Text = ((fsinfo.AxisStatus(1) And Drive1) / Drive1)
     If ((fsinfo.AxisStatus(1) And Drive1) / Drive1) > 0 Then
        'If (Last_Pcut_State = 0) Then
        '    Last_Pcut_State = 1
         c6k.Write ("1OUT.12-0" & Chr$(13))
        'End If
    Else
    c6k.Write ("1OUT.12-1:T1:DRIVE1,1" & Chr$(13))
        'If (Last_Pcut_State = 1) Then
        '    Last_Pcut_State = 0
            
        'End If
    End If
'END E STOP
End Sub

Private Sub Form_Load()
On Error GoTo loaderr
    
    Dim n%
    Dim desc As String
'    ' setup the fixed labels and alignment of the grid
'    Grid1.ColWidth(0) = 2800
'    Grid1.ColWidth(1) = 3920
'    Grid1.FixedAlignment(0) = 2     'center fixed labels
'    Grid1.FixedAlignment(1) = 2
'    Grid1.ColAlignment(0) = 0       'left aligned
'    Grid1.ColAlignment(1) = 0
'    Call CellText(Grid1, 0, 0, "Description")
'    Call CellText(Grid1, 0, 1, "Value")
'
'    For n = 1 To 8
'        Call CellText(Grid1, n, 0, " Axis " & CStr(n) & " - Motor Position")
'    Next 'n'
'
'    For n = 1 To 8
'        Call CellText(Grid1, n + 9, 0, " Axis " & CStr(n) & " - Motor Velocity")
'    Next 'n'

'    For n = 1 To 8
'        Call CellText(Grid1, n + 18, 0, " Axis " & CStr(n) & " - Encoder Position")
'    Next 'n
'
'    For n = 1 To 8
'        desc = " Axis " & CStr(n) & " - Axis Status"
'        Call CellText(Grid1, n + 27, 0, desc)
'        List1.AddItem desc
'    Next 'n
'    List1.ListIndex = 0
'
'    Call CellText(Grid1, 37, 0, " System Status")
'    Call CellText(Grid1, 38, 0, " Error Status")
'    Call CellText(Grid1, 39, 0, " User Status")
'    List1.AddItem "System Status"
'    List1.AddItem "Error Status"
'    List1.AddItem "User Status"
'    List1.AddItem "Trigger Status"
'    List1.AddItem "Hardware Limits"
'    List1.AddItem "Alarm Status"
'
'    For n = 0 To 3
'        Call CellText(Grid1, n + 41, 0, " Programmable Inputs - Brick " & CStr(n))
'    Next 'n
'
'    For n = 0 To 3
'        Call CellText(Grid1, n + 46, 0, " Programmable Outputs - Brick " & CStr(n))
'    Next 'n
'
'    Call CellText(Grid1, 51, 0, " Trigger Status")
'    Call CellText(Grid1, 52, 0, " Hardware Limits")
'    Call CellText(Grid1, 53, 0, " Analog Input 1")
'    Call CellText(Grid1, 54, 0, " Analog Input 2")
    
'    For n = 1 To 10
'        Call CellText(Grid1, n + 55, 0, " Integer Variables - VARI" & CStr(n))
'    Next 'n
    
'    For n = 1 To 10
'        Call CellText(Grid1, n + 66, 0, " Binary Variables - VARB" & CStr(n))
'    Next 'n
    
'    Call CellText(Grid1, 78, 0, " Timer Frame Counter")
'    Call CellText(Grid1, 79, 0, " Command Count")
'    Call CellText(Grid1, 80, 0, " Timer")
'    Call CellText(Grid1, 81, 0, " Alarm Status")
    
    'detail grid
'    Grid2.ColWidth(0) = 6100
'    Grid2.ColWidth(1) = 620
'    Grid2.FixedAlignment(0) = 2     'center fixed labels
'    Grid2.FixedAlignment(1) = 2
'    Grid2.ColAlignment(0) = 0       'left aligned
'    Grid2.ColAlignment(1) = 2
'    Call CellText(Grid2, 0, 0, "Description")
'    Call CellText(Grid2, 0, 1, "Value")
'
    frmMain!Terminal_Timer.Enabled = False      'the terminal's timer must be disabled to avoid window bug
    c6k.FSEnabled = True                'enable fast status
    Timer1.Enabled = True               'enable polling
    Timer2.Enabled = True               'enable detail view
Timer1.Enabled = False
Timer2.Enabled = False
frmFastStatus.Visible = False
Form10.Show
    
    Exit Sub
    
loaderr:
    Disconnect
End Sub


Private Sub Form_Resize()
    Dim nh%, nw%
    
    nw = Me.Width - 360
    nh = Me.Height - 5205
    
    If nw < 8970 Then nw = 8970
    If nh < 2895 Then nh = 2895

    Text1.Width = nw
    Text1.Height = nh
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Timer1.Enabled Then
        Timer1.Enabled = False
        c6k.FSEnabled = False
    End If
    frmMain.Show
End Sub


Private Sub List1_Click()
On Error GoTo listerr
    Dim n%
    
    Select Case List1.ListIndex
        Case 0 To 7
            Call CellText(Grid2, 0, 0, "Axis Status - Description")
            Call CellText(Grid2, 1, 0, "Bit 1 - Moving/Not Moving")
            Call CellText(Grid2, 2, 0, "Bit 2 - Negative/Positive Direction")
            Call CellText(Grid2, 3, 0, "Bit 3 - Accelerating")
            Call CellText(Grid2, 4, 0, "Bit 4 - At Velocity")
            Call CellText(Grid2, 5, 0, "Bit 5 - Home Successful")
            Call CellText(Grid2, 6, 0, "Bit 6 - Absolute/Incremental")
            Call CellText(Grid2, 7, 0, "Bit 7 - Continuous/Preset")
            Call CellText(Grid2, 8, 0, "Bit 8 - Jog Mode")
            Call CellText(Grid2, 9, 0, "Bit 9 - Joystick Mode")
            Call CellText(Grid2, 10, 0, "Bit 10 - Reserved")
            Call CellText(Grid2, 11, 0, "Bit 11 - Reserved")
            Call CellText(Grid2, 12, 0, "Bit 12 - Stall Detected")
            Call CellText(Grid2, 13, 0, "Bit 13 - Drive Shutdown")
            Call CellText(Grid2, 14, 0, "Bit 14 - Drive Fault Occured")
            Call CellText(Grid2, 15, 0, "Bit 15 - POS Hardware Limit")
            Call CellText(Grid2, 16, 0, "Bit 16 - NEG Hardware Limit")
            Call CellText(Grid2, 17, 0, "Bit 17 - POS Software Limit")
            Call CellText(Grid2, 18, 0, "Bit 18 - NEG Software Limit")
            Call CellText(Grid2, 19, 0, "Bit 19 - Reserved")
            Call CellText(Grid2, 20, 0, "Bit 20 - Reserved")
            Call CellText(Grid2, 21, 0, "Bit 21 - Reserved")
            Call CellText(Grid2, 22, 0, "Bit 22 - Reserved")
            Call CellText(Grid2, 23, 0, "Bit 23 - Position Error Exceeded")
            Call CellText(Grid2, 24, 0, "Bit 24 - In Target Zone")
            Call CellText(Grid2, 25, 0, "Bit 25 - Target Zone Timeout")
            Call CellText(Grid2, 26, 0, "Bit 26 - Pending GOWHEN")
            Call CellText(Grid2, 27, 0, "Bit 27 - Reserved")
            Call CellText(Grid2, 28, 0, "Bit 28 - Registration Triggered")
            Call CellText(Grid2, 29, 0, "Bit 29 - Reserved")
            Call CellText(Grid2, 30, 0, "Bit 30 - OTF or REG Move not Possible")
            Call CellText(Grid2, 31, 0, "Bit 31 - Reserved")
            Call CellText(Grid2, 32, 0, "Bit 32 - Reserved")
        
        Case 8
            Call CellText(Grid2, 0, 0, "System Status - Description")
            Call CellText(Grid2, 1, 0, "Bit 1 - System Ready")
            Call CellText(Grid2, 2, 0, "Bit 2 - Reserved")
            Call CellText(Grid2, 3, 0, "Bit 3 - Executing a Program")
            Call CellText(Grid2, 4, 0, "Bit 4 - Immediate Command")
            Call CellText(Grid2, 5, 0, "Bit 5 - In ASCII Mode")
            Call CellText(Grid2, 6, 0, "Bit 6 - In Echo Mode")
            Call CellText(Grid2, 7, 0, "Bit 7 - Defining a Program")
            Call CellText(Grid2, 8, 0, "Bit 8 - In Trace Mode")
            Call CellText(Grid2, 9, 0, "Bit 9 - In Step Mode")
            Call CellText(Grid2, 10, 0, "Bit 10 - In Translation Mode")
            Call CellText(Grid2, 11, 0, "Bit 11 - Command Error Occured")
            Call CellText(Grid2, 12, 0, "Bit 12 - Break Point Active")
            Call CellText(Grid2, 13, 0, "Bit 13 - Pause Active")
            Call CellText(Grid2, 14, 0, "Bit 14 - Wait Active")
            Call CellText(Grid2, 15, 0, "Bit 15 - Monitoring On Condition")
            Call CellText(Grid2, 16, 0, "Bit 16 - Waiting for Data (READ)")
            Call CellText(Grid2, 17, 0, "Bit 17 - Loading Thumbwheel Data")
            Call CellText(Grid2, 18, 0, "Bit 18 - External Program Select Mode")
            Call CellText(Grid2, 19, 0, "Bit 19 - Dwell in Progress (T command)")
            Call CellText(Grid2, 20, 0, "Bit 20 - Waiting for RP240 Data")
            Call CellText(Grid2, 21, 0, "Bit 21 - RP240 Connected")
            Call CellText(Grid2, 22, 0, "Bit 22 - Non-volatile Memory Error")
            Call CellText(Grid2, 23, 0, "Bit 23 - Servo data gathering in progress.")
            Call CellText(Grid2, 24, 0, "Bit 24 - Reserved")
            Call CellText(Grid2, 25, 0, "Bit 25 - Reserved")
            Call CellText(Grid2, 26, 0, "Bit 26 - Reserved")
            Call CellText(Grid2, 27, 0, "Bit 27 - Reserved")
            Call CellText(Grid2, 28, 0, "Bit 28 - Reserved")
            Call CellText(Grid2, 29, 0, "Bit 29 -Compiled Memory is 75% full")
            Call CellText(Grid2, 30, 0, "Bit 30 - Compiled Memory is 100% full.")
            Call CellText(Grid2, 31, 0, "Bit 31 - Compiled operation failed.")
            Call CellText(Grid2, 32, 0, "Bit 32 - Reserved")
        
        Case 9 'error status
            Call CellText(Grid2, 0, 0, "Error Status - Description")
            Call CellText(Grid2, 1, 0, "Bit 1 - Stall Detected")
            Call CellText(Grid2, 2, 0, "Bit 2 - Hard Limit Hit")
            Call CellText(Grid2, 3, 0, "Bit 3 - Soft Limit Hit")
            Call CellText(Grid2, 4, 0, "Bit 4 - Drive Fault")
            Call CellText(Grid2, 5, 0, "Bit 5 - Reserved")
            Call CellText(Grid2, 6, 0, "Bit 6 - Kill Input")
            Call CellText(Grid2, 7, 0, "Bit 7 - User Fault Input")
            Call CellText(Grid2, 8, 0, "Bit 8 - Stop Input")
            Call CellText(Grid2, 9, 0, "Bit 9 - Enable input is activated")
            Call CellText(Grid2, 10, 0, "Bit 10 - Pre-emptive or registration move profile not possible.")
            Call CellText(Grid2, 11, 0, "Bit 11 - Target zone setting timeout period exceeded.")
            Call CellText(Grid2, 12, 0, "Bit 12 - Maximum position error exceeded.")
            Call CellText(Grid2, 13, 0, "Bit 13 - Reserved")
            Call CellText(Grid2, 14, 0, "Bit 14 - GOWHEN position condition already true when move or shift was executed.")
            Call CellText(Grid2, 15, 0, "Bit 15 - Reserved")
            Call CellText(Grid2, 16, 0, "Bit 16 - Bad command detected (bit is cleared with TCMDER)")
            Call CellText(Grid2, 17, 0, "Bit 17 - Encoder failure")
            Call CellText(Grid2, 18, 0, "Bit 18 - Cable to expansion I/O brick is disconnected or power to I/O brick is lost.")
            
            For n = 19 To 32
                Call CellText(Grid2, n, 0, "Bit " & CStr(n) & " - Reserved")
            Next 'n
            
        Case 10 'user status
            For n = 1 To 16
                Call CellText(Grid2, n, 0, "Bit " & CStr(n) & " - User Status")
            Next 'n
            For n = 17 To 32
                Call CellText(Grid2, n, 0, "Bit " & CStr(n) & " - Reserved")
            Next 'n
        
        Case 11 'Triggers
            Call CellText(Grid2, 0, 0, "Trigger Status - Description")
            Call CellText(Grid2, 1, 0, "Bit 1 - Axis 1 - Trigger A")
            Call CellText(Grid2, 2, 0, "Bit 2 - Axis 1 - Trigger B")
            Call CellText(Grid2, 3, 0, "")
            Call CellText(Grid2, 4, 0, "Bit 3 - Axis 2 - Trigger A")
            Call CellText(Grid2, 5, 0, "Bit 4 - Axis 2 - Trigger B")
            Call CellText(Grid2, 6, 0, "")
            Call CellText(Grid2, 7, 0, "Bit 5 - Axis 3 - Trigger A")
            Call CellText(Grid2, 8, 0, "Bit 6 - Axis 3 - Trigger B")
            Call CellText(Grid2, 9, 0, "")
            Call CellText(Grid2, 10, 0, "Bit 7 - Axis 4 - Trigger A")
            Call CellText(Grid2, 11, 0, "Bit 8 - Axis 4 - Trigger B")
            Call CellText(Grid2, 12, 0, "")
            Call CellText(Grid2, 13, 0, "Bit 9 - Axis 5 - Trigger A")
            Call CellText(Grid2, 14, 0, "Bit 10 - Axis 5 - Trigger B")
            Call CellText(Grid2, 15, 0, "")
            Call CellText(Grid2, 16, 0, "Bit 11 - Axis 6 - Trigger A")
            Call CellText(Grid2, 17, 0, "Bit 12 - Axis 6 - Trigger B")
            Call CellText(Grid2, 18, 0, "")
            Call CellText(Grid2, 19, 0, "Bit 13 - Axis 7 - Trigger A")
            Call CellText(Grid2, 20, 0, "Bit 14 - Axis 7 - Trigger B")
            Call CellText(Grid2, 21, 0, "")
            Call CellText(Grid2, 22, 0, "Bit 15 - Axis 8 - Trigger A")
            Call CellText(Grid2, 23, 0, "Bit 16 - Axis 8 - Trigger B")
            Call CellText(Grid2, 24, 0, "")
            Call CellText(Grid2, 25, 0, "Bit 17 - Master - Trigger M")
        
            For n = 26 To 32
                Call CellText(Grid2, n, 0, "")
            Next 'n
        
        Case 12 'hardware limits
            Call CellText(Grid2, 0, 0, "Limit Status - Description")
            Call CellText(Grid2, 1, 0, "Bit 1 - Axis 1 - POS Limit")
            Call CellText(Grid2, 2, 0, "Bit 2 - Axis 1 - NEG Limit")
            Call CellText(Grid2, 3, 0, "Bit 3 - Axis 1 - HOME Limit")
            Call CellText(Grid2, 4, 0, "")
            Call CellText(Grid2, 5, 0, "Bit 4 - Axis 2 - POS Limit")
            Call CellText(Grid2, 6, 0, "Bit 5 - Axis 2 - NEG Limit")
            Call CellText(Grid2, 7, 0, "Bit 6 - Axis 2 - HOME Limit")
            Call CellText(Grid2, 8, 0, "")
            Call CellText(Grid2, 9, 0, "Bit 7 - Axis 3 - POS Limit")
            Call CellText(Grid2, 10, 0, "Bit 8 - Axis 3 - NEG Limit")
            Call CellText(Grid2, 11, 0, "Bit 9 - Axis 3 - HOME Limit")
            Call CellText(Grid2, 12, 0, "")
            Call CellText(Grid2, 13, 0, "Bit 10 - Axis 4 - POS Limit")
            Call CellText(Grid2, 14, 0, "Bit 11 - Axis 4 - NEG Limit")
            Call CellText(Grid2, 15, 0, "Bit 12 - Axis 4 - HOME Limit")
            Call CellText(Grid2, 16, 0, "")
            Call CellText(Grid2, 17, 0, "Bit 13 - Axis 5 - POS Limit")
            Call CellText(Grid2, 18, 0, "Bit 14 - Axis 5 - NEG Limit")
            Call CellText(Grid2, 19, 0, "Bit 15 - Axis 5 - HOME Limit")
            Call CellText(Grid2, 20, 0, "")
            Call CellText(Grid2, 21, 0, "Bit 16 - Axis 6 - POS Limit")
            Call CellText(Grid2, 22, 0, "Bit 17 - Axis 6 - NEG Limit")
            Call CellText(Grid2, 23, 0, "Bit 18 - Axis 6 - HOME Limit")
            Call CellText(Grid2, 24, 0, "")
            Call CellText(Grid2, 25, 0, "Bit 19 - Axis 7 - POS Limit")
            Call CellText(Grid2, 26, 0, "Bit 20 - Axis 7 - NEG Limit")
            Call CellText(Grid2, 27, 0, "Bit 21 - Axis 7 - HOME Limit")
            Call CellText(Grid2, 28, 0, "")
            Call CellText(Grid2, 29, 0, "Bit 22 - Axis 8 - POS Limit")
            Call CellText(Grid2, 30, 0, "Bit 23 - Axis 8 - NEG Limit")
            Call CellText(Grid2, 31, 0, "Bit 24 - Axis 8 - HOME Limit")
            Call CellText(Grid2, 32, 0, "")
    
        Case 13 'alarm limits
            Call CellText(Grid2, 0, 0, "Alarm Status - Description")
            For n = 1 To 12
                Call CellText(Grid2, n, 0, "Bit " & CStr(n) & " Software Alarm #" & CStr(n))
            Next 'n
            Call CellText(Grid2, 13, 0, "Bit 13 - Command Buffer Full")
            Call CellText(Grid2, 14, 0, "Bit 14 - ENABLE Input Activated")
            Call CellText(Grid2, 15, 0, "Bit 15 - Program Complete")
            Call CellText(Grid2, 16, 0, "Bit 16 - Drive Fault on any Axis")
            Call CellText(Grid2, 17, 0, "Bit 17 - Reserved")
            Call CellText(Grid2, 18, 0, "Bit 18 - Reserved")
            Call CellText(Grid2, 19, 0, "Bit 19 - Limit Hit - hard/soft, any axis")
            Call CellText(Grid2, 20, 0, "Bit 20 - Stall Detected (stepper) or Position Error (servo)")
            Call CellText(Grid2, 21, 0, "Bit 21 - Timer (TIMINT)")
            Call CellText(Grid2, 22, 0, "Bit 22 - Reserved")
            Call CellText(Grid2, 23, 0, "Bit 23 - Input (any defined as INFNCi-I or LIMFNCi-I)")
            Call CellText(Grid2, 24, 0, "Bit 24 - Command Error")
            For n = 25 To 32
                Call CellText(Grid2, n, 0, "Bit " & CStr(n) & " Motion Complete on Axis " & CStr(n - 24))
            Next 'n
            
    End Select
        
    Exit Sub
    
listerr:
    Disconnect
End Sub


Private Sub Text1_Change()
    'the text box has a finite buffer so
    'make sure it doesn't overflow
    
    If Len(Text1.Text) > 16000 Then
        Text1.Text = Right$(Text1.Text, 500)    'buffer just the last 500 characters
    End If

End Sub

Private Sub Text1_DblClick()
    Text1.Text = ""     'clear the terminal display
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
'this routine processes the terminal's key presses
On Error GoTo text1keypress_error

    'exit if not connected
    If Not connected Then
        KeyAscii = 0
    End If
    
    Dim temp%
    Static buffer$      'local command buffer
    
    'perform action based on value of key being pressed
    Select Case KeyAscii
        'backspace
        Case 8
            If Len(buffer) > 0 Then buffer = Left$(buffer, Len(buffer) - 1) 'erase one char from buffer
            
        
        'CR or colon - 6000 command delimeter
        Case 13, Asc(":")
            If Format$(buffer, ">") = "CLS" Then      'internal clear screen command
                Text1.Text = ""
                KeyAscii = 0
            Else
                buffer = buffer & Chr$(13)      'append the CR
                Timer1.Enabled = False          'disable response polling to avoid simultaneous read/write
                temp = c6k.Write(buffer)        'send commands to 6k
                Timer1.Enabled = True           'enable response polling
            End If
            buffer = ""                         'empty the command local buffer
        
        
        'anything else just add to the buffer
        Case Else
            buffer = buffer & Chr$(KeyAscii)    'append char to the local command buffer
            
    End Select
    Exit Sub
    
text1keypress_error:
    Disconnect
End Sub

Private Sub Timer1_Timer()
On Error GoTo timer1err


Dim EStop
    'terminal display
   Dim A
   Dim tom$
   Dim DICK$
   Dim Harry$
    Dim msg$
'Input23Mask = 8388608
    'c6k.write ("TINO" & Chr$(13))
   ' msg = c6k.read()                           'get response
   ' If Len(msg) Then Text1.SelText = msg      'if not empty then display in the text box
    
    
    
    'fast status display
    Dim n%
    Dim temp() As Byte                      'create temporary byte array
    
    temp = c6k.FastStatus                  'get fast status information
    Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
 
     c6k.Write ("TINO" & Chr$(13))
    EStop = c6k.Read()                           'get response
    EStop = Mid(EStop, 12, 1)
 If EStop = 0 Then
     'Text1.SelText = "ON"
     If (Last_Pcut_State = 0) Then
            Last_Pcut_State = 1
            c6k.Write ("1OUT.12-0" & Chr$(13))
     End If
 Else
     'Text1.SelText = "OFF"
     If (Last_Pcut_State = 1) Then
            Last_Pcut_State = 0
            'c6k.write ("1OUT.12-0" & Chr$(13))
        c6k.Write ("DRIVE1,1:1OUT.12-1" & Chr$(13))
     End If
 
 End If
     
     Text5.Refresh
 'END E STOP
    
    
    
    
    
    
    
    'Let A = fsinfo
    tom = fsinfo.AxisStatus(1)
    DICK = BitText32(fsinfo.ProgIn(1))
 
    Text5.Text = CStr((fsinfo.ProgIn(1) And Input11) / Input11)
'    Text4.Text = CStr((fsinfo.ProgIn(1) And Input23) / Input23)
 '   Text4.Text = Text4.Text & "   " & CStr((fsinfo.ProgIn(1) And Input24) / Input24)
    'Harry = (fsinfo.ProgIn(1) And Input23Mask) / Input23Mask
    '(v And 8388608) \ 8388608
    'update the grid with the new data
    For n = 1 To 8
        Call CellText(Grid1, n, 1, fsinfo.MotorPos(n))
    Next 'n
Text2.Text = tom / 250000 & "   " & DICK
Text3.Text = Input23Mask


    For n = 1 To 8
        Call CellText(Grid1, n + 9, 1, fsinfo.MotorVel(n))
    Next 'n

    For n = 1 To 8
        Call CellText(Grid1, n + 18, 1, fsinfo.EncoderPos(n))
    Next 'n

    For n = 1 To 8
        Call CellText(Grid1, n + 27, 1, BitText32(fsinfo.AxisStatus(n)))
    Next 'n

    Call CellText(Grid1, 37, 1, BitText32(fsinfo.SysStatus))
    Call CellText(Grid1, 38, 1, BitText32(fsinfo.ErrorStatus))
    Call CellText(Grid1, 39, 1, BitText32(fsinfo.UserStatus))
    
    For n = 1 To 3
        Call CellText(Grid1, n + 41, 1, BitText32(fsinfo.ProgIn(n)))
    Next 'n
    
    For n = 0 To 3
        Call CellText(Grid1, n + 46, 1, BitText32(fsinfo.ProgOut(n)))
    Next 'n
   
    Call CellText(Grid1, 51, 1, BitText32(fsinfo.Triggers))
    Call CellText(Grid1, 52, 1, BitText32(fsinfo.Limits))
    Call CellText(Grid1, 53, 1, fsinfo.Analog(1))
    Call CellText(Grid1, 54, 1, fsinfo.Analog(2))
    
    For n = 1 To 10
        Call CellText(Grid1, n + 55, 1, fsinfo.VarI(n))
    Next 'n
    
    For n = 1 To 10
        Call CellText(Grid1, n + 66, 1, BitText32(fsinfo.VarB(n)))
    Next 'n
    
    Call CellText(Grid1, 78, 1, fsinfo.Counter)
    Call CellText(Grid1, 79, 1, fsinfo.CmdCount)
    Call CellText(Grid1, 80, 1, fsinfo.Timer)
    
    alarms = c6k.AlarmStatus(0)
    Call CellText(Grid1, 81, 1, BitText32(alarms))


Exit Sub
    
timer1err:
    Disconnect
 End Sub


Private Sub Timer2_Timer()
On Error GoTo timer2err
    Dim n%, row%
    
    'display detailed status information
    Select Case List1.ListIndex
        Case 0 To 7    ' axis status
            For n = 0 To 30
                If ((fsinfo.AxisStatus(List1.ListIndex + 1) And (2 ^ n)) > 0) Then
                    Call CellText(Grid2, n + 1, 1, "Yes")
                Else
                    Call CellText(Grid2, n + 1, 1, "No")
                End If
            Next 'n
            
            If fsinfo.AxisStatus(List1.ListIndex + 1) > 2147483647 Then
                Call CellText(Grid2, 32, 1, "Yes")
            Else
                Call CellText(Grid2, 32, 1, "No")
            End If
        
        Case 8 'system status
            For n = 0 To 30
                If ((fsinfo.SysStatus And (2 ^ n)) > 0) Then
                    Call CellText(Grid2, n + 1, 1, "Yes")
                Else
                    Call CellText(Grid2, n + 1, 1, "No")
                End If
             Next 'n
            
            If fsinfo.SysStatus > 2147483647 Then
                Call CellText(Grid2, 32, 1, "Yes")
            Else
                Call CellText(Grid2, 32, 1, "No")
            End If
        
        Case 9 'error status
            For n = 0 To 30
                If ((fsinfo.ErrorStatus And (2 ^ n)) > 0) Then
                    Call CellText(Grid2, n + 1, 1, "Yes")
                Else
                    Call CellText(Grid2, n + 1, 1, "No")
                End If
            Next 'n
            
            If fsinfo.ErrorStatus > 2147483647 Then
                Call CellText(Grid2, 32, 1, "Yes")
            Else
                Call CellText(Grid2, 32, 1, "No")
            End If
            
        Case 10 'user status
            For n = 0 To 30
                If ((fsinfo.UserStatus And (2 ^ n)) > 0) Then
                    Call CellText(Grid2, n + 1, 1, "Yes")
                Else
                    Call CellText(Grid2, n + 1, 1, "No")
                End If
            Next 'n
            
            If fsinfo.UserStatus > 2147483647 Then
                Call CellText(Grid2, 32, 1, "Yes")
            Else
                Call CellText(Grid2, 32, 1, "No")
            End If
        
        Case 11 'Triggers
            row = 0
            For n = 0 To 16
                If ((n Mod 2) = 0 And n <> 0) Then
                    Call CellText(Grid2, row + 1, 1, "")
                    row = row + 2
                Else
                    row = row + 1
                End If
                If ((fsinfo.Triggers And (2 ^ n)) > 0) Then
                    Call CellText(Grid2, row, 1, "On")
                Else
                    Call CellText(Grid2, row, 1, "Off")
                End If
            Next 'n
        
            For n = 26 To 32
                Call CellText(Grid2, n, 1, "")
            Next 'n
        
        Case 12 'hardware limits
            row = 0
            For n = 0 To 23
                If ((n Mod 3) = 0 And n <> 0) Then
                    Call CellText(Grid2, row + 1, 1, "")
                    row = row + 2
                Else
                    row = row + 1
                End If
                If ((fsinfo.Limits And (2 ^ n)) > 0) Then
                    Call CellText(Grid2, row, 1, "On")
                Else
                    Call CellText(Grid2, row, 1, "Off")
                End If
            Next 'n
            Call CellText(Grid2, 32, 1, "")
                
        Case 13 'alarm status limits
            For n = 0 To 30
                If ((alarms And (2 ^ n)) > 0) Then
                    Call CellText(Grid2, n + 1, 1, "Yes")
                Else
                    Call CellText(Grid2, n + 1, 1, "No")
                End If
            Next 'n
            
            If alarms > 2147483647 Then
                Call CellText(Grid2, 32, 1, "Yes")
            Else
                Call CellText(Grid2, 32, 1, "No")
            End If
    End Select
    Exit Sub
    
timer2err:
    Disconnect
End Sub


