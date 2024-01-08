VERSION 5.00
Begin VB.Form frmLine6 
   BackColor       =   &H00C00000&
   Caption         =   "Line 6"
   ClientHeight    =   6660
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   17385
   LinkTopic       =   "Form6"
   ScaleHeight     =   6660
   ScaleWidth      =   17385
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option_Auger_Direction 
      BackColor       =   &H00C00000&
      Caption         =   "Run Edge"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   22
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Frame Frame_Auger_Switch 
      BackColor       =   &H00C00000&
      Caption         =   "Auger Direction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   5400
      TabIndex        =   20
      Top             =   5040
      Width           =   2175
      Begin VB.OptionButton Option_Auger_Direction 
         BackColor       =   &H00C00000&
         Caption         =   "Run Face"
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox Text_Enter_Pass_Width 
      Height          =   375
      Left            =   5400
      TabIndex        =   18
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Button_Go 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   600
      TabIndex        =   17
      Top             =   4320
      Width           =   4575
   End
   Begin VB.CommandButton Button_Fin 
      Caption         =   "Finish"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   16
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Button_NF 
      Caption         =   "N / F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   15
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Button_Start 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3720
      TabIndex        =   14
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text_Pop_Total_Qty 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text_Pop_Part_Num 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Timer Timer_Blade_Runtime 
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer_FSM 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Button_WO_Enter_Clear 
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text_Pop_Due_Date 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text_Pop_Dwg_Num 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox Text_Enter_WO 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Text            =   "Enter Work Order Number Here"
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label Label_Estop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E-Stop Enabled"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000CF&
      Height          =   495
      Left            =   6360
      TabIndex        =   24
      Top             =   5520
      Visible         =   0   'False
      Width           =   12615
   End
   Begin VB.Label Var_Label_Joystick_Status 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "Joystick Enabled:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3495
      Left            =   5400
      TabIndex        =   23
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label_Pass_Width 
      BackColor       =   &H00C00000&
      Caption         =   "Pass Width"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5400
      TabIndex        =   19
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Var_Label_System_Status 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ready to Run Blade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6120
      TabIndex        =   13
      Top             =   960
      Width           =   12615
   End
   Begin VB.Label Var_Label_WO_Active 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No Active Work Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6120
      TabIndex        =   12
      Top             =   480
      Width           =   12615
   End
   Begin VB.Label Label_Total_Qty 
      BackColor       =   &H00C00000&
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label_Part_Num 
      BackColor       =   &H00C00000&
      Caption         =   "Part Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label_WO_Info 
      BackColor       =   &H00C00000&
      Caption         =   "Work Order Information:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label_Due_Date 
      BackColor       =   &H00C00000&
      Caption         =   "Due Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label_Dwg_Num 
      BackColor       =   &H00C00000&
      Caption         =   "Dwg Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label_WO 
      BackStyle       =   0  'Transparent
      Caption         =   "Work_Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Menu Topbar_Set_0 
      Caption         =   "Set 0"
   End
   Begin VB.Menu Topbar_Joystick 
      Caption         =   "Joystick"
   End
   Begin VB.Menu Topbar_Maintenance 
      Caption         =   "Maintenance"
   End
   Begin VB.Menu Topbar_Test_Dropdown 
      Caption         =   "Test"
      Begin VB.Menu topbar_test_1 
         Caption         =   "Simple Move Motors"
      End
      Begin VB.Menu topbar_test_2 
         Caption         =   "Check Inputs (Generic)"
      End
      Begin VB.Menu topbar_test_3 
         Caption         =   "Check Inputs (Specific)"
      End
      Begin VB.Menu topbar_test_4 
         Caption         =   "Basic Joystick"
      End
   End
   Begin VB.Menu Topbar_Calibration 
      Caption         =   "Calibration"
   End
   Begin VB.Menu Topbar_Reset 
      Caption         =   "Reset"
   End
End
Attribute VB_Name = "frmLine6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button_Go_Click()

FSM_Line6.State = 2

End Sub

Private Sub Button_Start_Click()

'Start Punch
Call woMgr.startWO

'Enable Timeout Timer
Call woMgr.woTimer("Start")
    
End Sub

Private Sub Button_WO_Enter_Clear_Click()

If Not woMgr.isWOactive() Then
    Call woMgr.loadWO
Else
   Call woMgr.clearWO
End If

End Sub

Private Sub Command1_Click()
    Call c6kOps.testDrives
End Sub

Private Sub Form_Activate()

'When the form returns to focus, activate the FSM
Timer_FSM.Enabled = True

End Sub

Private Sub Form_Deactivate()

'When the form loses focus, deactivate the FSM and reset the FastStatus up-to-date flag
Timer_FSM.Enabled = False
Call c6kOps.resetFSupd

End Sub

Private Sub Form_Load()
'When the form is loaded:

'-- Set Up Defaults for Text
Text_Enter_WO.Tag = "Enter Work Order Number Here"
Text_Enter_WO.Text = Text_Enter_WO.Tag

Var_Label_WO_Active.Tag = "No Active Work Order"
Var_Label_WO_Active.Caption = Var_Label_WO_Active.Tag
Var_Label_WO_Active.Visible = True

Var_Label_System_Status = "Ready to Run Blade"
Var_Label_System_Status.Visible = False

Var_Label_Joystick_Status.Visible = False

Label_Estop.Visible = False

'Initialize Form Buttons
woMgr.btnState ("Inactive")

'--
Call c6kOps.Enable

'-- Run motor setup subroutine
Call c6kOps.bootDrives

'-- Check for previously active WO
Call woMgr.chkActiveWO

'-- Start Airblade & Exhaust Fan
Call c6kOps.setOutput("Airblade", True)
Call c6kOps.setOutput("Exhaust", True)

'-- Enable FSM
Timer_FSM.Enabled = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If User hits X
    If UnloadMode = 0 Then
        If MsgBox("Are you sure you want to close?", vbYesNo Or vbQuestion) = vbNo Then Cancel = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call FL6_End_Program
    
End Sub

Private Sub FL6_End_Program()

    'Shut Down c6k - Immediate jog off, disable outputs, disable drives, clear command buffer
    c6k.Write "!JOG00000000:1OUTALL9,16,0:1OUTALL25,32,0:DRIVE0,0,0,0,0,0,0,0:S"
    
    'Reset c6k Controller to avoid retained settings & commands in next boot. Will disconnect ethernet
    c6k.Write ("RESET")
    
    'Unload all Forms except frmLine6 and frmMain
    For Each tmpForm In Forms
        If (tmpForm.Name <> "frmLine6") Or (tmpForm.Name <> "frmMain") Then
            Unload tmpForm
            Set tmpForm = Nothing
        End If
    Next
        
    'Show frmMain
    frmMain.Show
End Sub

Private Sub Text_Enter_WO_GotFocus()

If Text_Enter_WO.Text = Text_Enter_WO.Tag Then Text_Enter_WO = vbNullString

End Sub

Private Sub Text_Enter_WO_LostFocus()

If Trim(Text_Enter_WO) = vbNullString Then Text_Enter_WO = Text_Enter_WO.Tag
    
End Sub

Private Sub Text_Pop_Mat_Change()

End Sub

'-- Main FSM Timer - Operates the non-UI elements of the program ------------------------------------
Private Sub Timer_FSM_Timer()

Call fsmMain.Run

End Sub

Private Sub Topbar_Calibration_Click()

frmCalibrate.Show

End Sub

Private Sub Topbar_Joystick_Click()

' Toggle joystick active boolean
If Not c6kOps.getJoyActive Then
    c6kOps.runJoy ("Enable")
Else
    c6kOps.runJoy ("Disable")
    Var_Label_Joystick_Status.Visible = False
End If

End Sub

Private Sub Topbar_Maintenance_Click()

'Show Maintenance Form
frmMaintenance.Show

End Sub

Private Sub Topbar_Reset_Click()

'Display Verification Prompt
temp = MsgBox("Reset Drives?", 1, "Reset")

If temp = 1 Then
    'Send Reset command to 6k - simulates hard reboot
    c6k.Write ("!RESET" & Chr$(13))
    
    'Unload the main form, as the reset command will disconnect ethernet
    frmLine6.Unload
End If
    
End Sub

Private Sub Topbar_Set_0_Click()

'Display Verification Prompt
temp = MsgBox("Set Machine Home?", 1, "Set 0")

If temp = 0 Then
    'Send command to 6k to set 0 position on all axes
    Call c6kOps.setMachineHome
End If

End Sub

Private Sub topbar_test_1_Click()
    Call c6kOps.testDrives
End Sub

Private Sub topbar_test_2_Click()
    MsgBox ("Inputs Are: " & Chr$(13) & c6kOps.getInputStateStr)
End Sub

Private Sub topbar_test_4_Click()

c6k.Write ("!JOG000000:" & Chr$(13))
c6k.Write ("JOG000000:1INFNC1-5J:1INFNC2-5K:1INFNC3-2K:1INFNC4-2J:JOGA4,5,5,1,5,5:JOGAD50,99,99,99,99,15:JOGVH8,8,10,2,5,3:JOGVL8,15,10,5,5,5:JOG01010" & Chr$(13))

End Sub
