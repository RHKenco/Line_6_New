VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassWO_Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'|====================================================================================================================|
'|
'|              ----- WO Manager Class -----
'|
'|
'|
'|====================================================================================================================|

Option Explicit

' Constant Values
Private Const woTimeoutTime = 300   'Work Order Timeout Clock Time in Seconds
Private Const woTimeoutCalib = 20   'Calibration constant - cycle of timer in Hz


' Filepaths utilized in multiple locations
Private Const woFormASCpath As String = "F:\MFG\FORM.ASC"
Private Const woBarDataPath As String = "F:\MFG\Bar_Data.dta"
Private Const woGrossPITpath As String = "F:\MFG\GROSSPIT.GPA"
Private Const woGaryPath As String = "F:\mfg\GARY.GPA"
Private Const woGoalPath As String = "F:\mfg\GOAL.GPA"
Private Const woTempMain As String = "F:\BARCODE\temp6.temp"
Private Const woTempDir As String = "F:\BARCODE\"

Private woTempData As String        ' Varible to store path to WO temp data

'Const Strings for File Write
Private Const machineNo = 6


'Internal Variables
Private woActive As Boolean         ' Work Order Active flag
Private woCountdownTime As Integer  ' Countdown timer variable for timeout time

Private woNum As String             'Form.ASC Data: - Work Order Number
Private woPartNum As String         '               - Part Number
Private woTqty As Single            '               - Total WO Quantity
Private woPrice As Single           '               - Price of Item
Private woTCinch As Single          '               - TC Inches on part
Private woComments As String        '               - Comments in Form.ASC
Private woDate As String            '               - Due Date on WO
Private woDWGno As String           '               - DWG Number for item
'Private woCScost As String          '               - Unused Column
'Private woSBTavg As String          '               - Unused Column

Private woLastStartTime As String       ' Start time of WO
Private woLastEndTime As String         ' Not-Finish Time of WO
Private woTotalTime As String       ' Total time since last Finish



Public Function isWOactive() As Boolean
    isWOactive = woActive
End Function

Public Function getWOnum() As String
    getWOnum = woNum
End Function

Public Sub loadWO()

'-- Load FORM.ASC Data & Populate Necessary Fields ------------------------------------------
If frmLine6.Text_Enter_WO.Text <> frmLine6.Text_Enter_WO.Tag And frmLine6.Text_Enter_WO.Text <> "" Then

    'Pull WO Number entered in text box
    WORK_ORDER$ = frmLine6.Text_Enter_WO.Text
    
    'Load data from FORM.ASC; if successful, then
    If pullFormASC(WORK_ORDER$) Then
        
        'Raise WO Active flag
        woActive = True
        
        'Populate Part # Field
        frmLine6.Text_Pop_Part_Num.Text = woPartNum
        'Populate Dwg # Field
        frmLine6.Text_Pop_Dwg_Num.Text = woDWGno
        'Check Due Date; if overdue change background to red
        If Date >= CDate(woDate) Then frmLine6.Text_Pop_Due_Date.BackColor = QBColor(4) Else frmLine6.Text_Pop_Due_Date.BackColor = &HFFFFFF
        'Populate Due Date
        frmLine6.Text_Pop_Due_Date.Text = woDate
        
        'Update Work Order Status & System Status Labels
        frmLine6.Var_Label_WO_Active.Caption = "Work Order " + woNum + " Active"
        frmLine6.Var_Label_WO_Active.Visible = True
        
        frmLine6.Var_Label_System_Status.Caption = "Ready to Run Blade"
        frmLine6.Var_Label_System_Status.Visible = True
        
        'Update Buttons
        btnState ("Active")
        
        frmLine6.Refresh
        
        'Set up temporary file to save current active work-order (in case of program crashing)
        Open woTempMain For Output As #1
        Let Date1$ = Date$
        Write #1, InString$, Date1$
        Close #1
        
        'Create temp file to store wo time data
        Let woTempData = woTempDirectory & woNum & "6.tmp"
        
        If Dir(woTempData) = "" Then
            Call updTemp("New")
        Else
            Call updTemp("Load")
        End If

End If

End Sub

Public Sub clearWO()
    '-- Clear All Fields and Reset Button to Enter ------------------------------------------------
    frmLine6.Text_Enter_WO.Text = Text_Enter_WO.Tag
    frmLine6.Text_Enter_WO.Locked = False
    
    frmLine6.Text_Pop_Part_Num.Text = ""
    frmLine6.Text_Pop_Dwg_Num.Text = ""
    frmLine6.Text_Pop_Due_Date.Text = ""

    'Update Work Order Status & System Status Labels
    frmLine6.Var_Label_WO_Active.Caption = Var_Label_WO_Active.Tag
    frmLine6.Var_Label_WO_Active.Visible = True
    
    frmLine6.Var_Label_System_Status.Visible = False
    
    'Clear Stored FORM.ASC Variables
    woNum = ""
    woPartNum = ""
    woTqty = 0
    woPrice = 0
    woTCinch = 0
    woComments = ""
    woDate = ""
    woDWGno = ""
    woActive = False
    
    'Update Buttons
    btnState ("Inactive")
    
    frmLine6.Refresh
    
    'Kill WO temp files (if they exist)
    If Dir(woTempMain) <> "" Then Kill woTempMain
    If Dir(woTempData) <> "" Then
        Kill woTempData
        woTempData = ""
    End If
    
End Sub

' -- chkActiveWO - Function to be called when loading program to check for a temporary file containing the currently active work order. ------------------------------------------------
'                - Allows for work-order to be maintained if program is shut down improperly
Public Sub chkActiveWO()

' Check to see if temp file exists. If it does not, exit subroutine. Else,
 If Dir(woTempMain) = "" Then
   chkActiveWO() = ""
   Exit Sub
Else
    ' Load temp file to check WOnum (temp6) and Date
    Open (woTempMain) For Input As #2
    Input #2, temp6, DAT
    Close #2
    
    ' Load backup data temp file & store in mgr variables
    
    
    
    
    
    
    
    
    ' ***** ADD CODE HERE ***** Pseudocode:
        ' Loads temporary file with necessary WO start time data
        ' Checks current time - if the current time is more than timeout time different from the temp file,
        ' Punch a not-finish on the WO with time based on the timeout period & Exit Sub
        
        
        
        
        
        
        
    
    ' Update Form text and call function to load WO
    frmLine6.Text_Enter_WO.Text = temp6
    Call loadWO
   
End Sub

'-- btnState() - Function to handle visibility and enabling of Enter, Go, Start, N/F, and Finish Buttons -------------------------------------
'   - Inputs: btn_State - string containing desired button state
Public Sub btnState(btn_state As String)
    
    If btn_state = "Inactive" Then
        
        frmLine6.Button_WO_Enter_Clear.Enabled = True
        frmLine6.Button_WO_Enter_Clear.Caption = "Enter"
        frmLine6.Text_Enter_WO.Locked = False
        
        frmLine6.Button_Start.Enabled = False
        frmLine6.Button_Start.Visible = True
        frmLine6.Button_NF.Enabled = False
        frmLine6.Button_NF.Visible = False
        frmLine6.Button_Fin.Enabled = False
        frmLine6.Button_Fin.Visible = False
        
        frmLine6.Button_Go.Enabled = False
        frmLine6.Button_Go.Visible = True
        
    ElseIf btn_state = "Active" Then
            
        frmLine6.Button_WO_Enter_Clear.Enabled = True
        frmLine6.Button_WO_Enter_Clear.Caption = "Clear"
        frmLine6.Text_Enter_WO.Locked = True
                    
        frmLine6.Button_Start.Enabled = True
        frmLine6.Button_Start.Visible = True
        frmLine6.Button_NF.Enabled = False
        frmLine6.Button_NF.Visible = False
        frmLine6.Button_Fin.Enabled = False
        frmLine6.Button_Fin.Visible = False
        
        frmLine6.Button_Go.Enabled = False
        
    ElseIf btn_state = "Started" Then
    
        frmLine6.Button_WO_Enter_Clear.Enabled = False
            
        frmLine6.Button_Start.Enabled = False
        frmLine6.Button_Start.Visible = False
        frmLine6.Button_NF.Enabled = True
        frmLine6.Button_NF.Visible = True
        frmLine6.Button_Fin.Enabled = True
        frmLine6.Button_Fin.Visible = True
        
        frmLine6.Button_Go.Enabled = True
        
    ElseIf btn_state = "Running" Then
    
        frmLine6.Button_WO_Enter_Clear.Enabled = False
            
        frmLine6.Button_Start.Enabled = False
        frmLine6.Button_Start.Visible = False
        frmLine6.Button_NF.Enabled = False
        frmLine6.Button_NF.Visible = True
        frmLine6.Button_Fin.Enabled = False
        frmLine6.Button_Fin.Visible = True
        
        frmLine6.Button_Go.Enabled = False
        
    ElseIf btn_state = "Returned" Then
    
        frmLine6.Button_WO_Enter_Clear.Enabled = False
            
        frmLine6.Button_Start.Enabled = False
        frmLine6.Button_Start.Visible = False
        frmLine6.Button_NF.Enabled = True
        frmLine6.Button_NF.Visible = True
        frmLine6.Button_Fin.Enabled = True
        frmLine6.Button_Fin.Visible = True
        
        frmLine6.Button_Go.Enabled = False
        
    Else
        MsgBox ("Error in Button State Function")
        
        Exit Sub
        
    End If
    
    frmLine6.Refresh

End Sub

Public Sub startWO()
    Call updTemp("Start")
    Call writeData("Start")
End Sub

Public Sub notfinishWO()
    Call updTemp("End")
    Call writeData("N/F")
End Sub

Public Sub finishWO()
    Call updTemp("End")
    Call writeData("Finish")
End Sub

Public Sub timeoutWO()
    Call updTemp("End")
    Call writeData("TimeOut")
End Sub

Private Function pullFormASC(WORK_ORDER As String) As Boolean

' Load FORM.ASC
Open woFormASCpath For Input As #1
'
Do While WORK_ORDER <> WOK$
     Input #1, WOK$, PartNumber, TQTY!, PRICE!, Tc!, COMM$, DAT$, DWG$, UNUSED_1!, UNUSED_2!
        If WORK_ORDER = WOK$ Then
            
            Close #1
            
            woNum = WOK$
            woPartNum = PartNumber
            woTqty = TQTY!
            woPrice = PRICE!
            woTCinch = Tc!
            woComments = COMM$
            woDate = DAT$
            woDWGno = DWG$
            
            pullFormASC = True
            
            'Exit Loop
            Exit Do
        End If
        
        ' Check for End-Of-File
        If EOF(1) Then
        
            'If EOF, Close FORM.ASC, Throw Error Message, and Exit Sub.
            Close #1
            MsgBox "WORK ORDER NOT FOUND!"
            
            pullFormASC = False
            
            'Exit Loop
            Exit Do
            
        End If
        
Loop

'Close FORM.ASC
Close #1

End Function

'-- writeData() - Function to write data to the various files necessary for operations.
'   - Inputs: Punch - string describing the type of punch delivered, from which data will be directed to the appropriate files
Private Sub writeData(Punch As String)

'Load Bar_Data
Open woBarDataPath For Append As #1

Select Case Punch
    Case "Start"
        
        'Write Start punch to Bar_Data
        Write #1, woNum, machineNo, "Start", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss")
        
        Close #1
        
    Case "N/F"
    
        'Write Not-Finished punch to Bar_Data
        Write #1, woNum, machineNo, "N\F", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss")
        
        Close #1
    
    Case "TimeOut"
        
        'Write Time-Out Punch to Bar_Data
        Write #1, woNum, machineNo, "Timed Out", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss")
        
        Close #1

    Case "Finished"
        
        'Prompt User for Quantity Finished
        Dim FinishedQty As Integer
        FinishedQty = Val(InputBox("ENTER QTY FINISHED", "Qty", 0, 4000, 1500))
        
        If FinishedQty < 1 Then
            MsgBox ("Finish Punch Canceled")
            Exit Sub
        End If
        
        Call runCalcs
        
        'Write Finished Punch & Data to Bar_Data
        Write #1, woNum, machineNo, "Finish", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss"), woPartNum, FinishedQty, AvgTime!, STLCOST!, GROSSPROFIT! & "%", COM$
        
        Close #1
        
        'Output to GROSSPIT
        Open woGrossPITpath For Append As #1
        Write #1, woNum, PartNumber, woTqty, FinishedQty, Sta$, Date$, PRICE!, Tc!, STLCOST!, AvgTime!, TCCOST!, CONVCOST!, UNITCOST!, UNITPROFIT!, GROSSPROFIT!, COM$, TTime!, TSALES!, TTC!, TTCCOST!, TSTLCOST!, TCONVCOST!, TOTALCOST!, TOTALPROFIT!, DAT$, DWG$, CSCOST!
        Close #1
        
        'Output to GARY
        Open woGaryPath For Append As #1
        Write #1, woNum, PartNumber, woTqty, FinishedQty, Sta$, Date$, PRICE!, Tc!, STLCOST!, AvgTime!
        Close #1
        
        'Output to GOAL
        Open woGoalPath For Append As #1
        Write #1, woNum, PartNumber, woTqty, FinishedQty, Sta$, Date$, PRICE!, Tc!, STLCOST!, AvgTime!
        Close #1
        
    Case Else
        MsgBox ("Error in Punch Handler")
        
End Select

End Sub

' -- updTemp() - updates temp file to track work-order progress between Start and Finish punches
Private Sub updTemp(cmd As String)

Select Case cmd
    Case "New"
    
        Open woTempData For Output As #2
        Write #2, woNum, woLastStartTime, 0, 0  ' Write to Temp file WO number, WO start time, Elapsed time (0 for first punch), Total Time (0 for first punch)
        
        Close #2
        
    Case "Load"
    
        Open woTempData For Input As #2
        Input #2, woNumTemp, woLastStartTimeTemp, woLastEndTimeTemp, woTotalTimeTemp  ' Write to Temp file WO number, WO start time, Elapsed time (0 for first punch), Total Time (0 for first punch)
        
        woLastStartTime = woLastStartTimeTemp
        woLastEndTime = woLastEndTimeTemp
        woTotalTime = woTotalTimeTemp
        
        Close #2
        
    Case "Start"
        
        ' Load last Start punch timestamp from woTempData
        Open woTempData For Input As #2
        Input #2, woNumTemp, woLastStartTimeTemp, woLastEndTimeTemp, woTotalTimeTemp  ' Write to Temp file WO number, WO start time, Elapsed time (0 for first punch), Total Time (0 for first punch)
        
        ' If the timestamp does not match the one stored in the class, there is probably an error
        If woLastStartTime <> woLastStartTimeTemp Then MsgBox ("Start time match error")
        
        ' Get current time as Start time
        woLastStartTime = Now()
        
        Close #2
        
        ' Store values back in temp file
        Open woTempData For Output As #2
        Write #2, woNum, woLastStartTime, woLastEndTime, woTotalTime  ' Write to Temp file WO number, WO start time, Elapsed time (0 for first punch), Total Time (0 for first punch)
    
        Close #2
        
    Case "End"
    
        ' Load last Start punch timestamp from woTempData
        Open woTempData For Input As #2
        Input #2, woNumTemp, woLastStartTimeTemp, woLastEndTimeTemp, woTotalTimeTemp  ' Write to Temp file WO number, WO start time, Elapsed time (0 for first punch), Total Time (0 for first punch)
        
        ' If the timestamp does not match the one stored in the class, there is probably an error
        If woLastStartTime <> woLastStartTimeTemp Then MsgBox ("Start time match error")
        
        ' Get current time as N/F time
        woLastEndTime = Now()
        
        ' Generate new total time based on last start time, current N/F time, and previous total time
        woTotalTime = DateDiff("s", woLastEndTime, woLastStartTime) + woTotalTimeTemp
        
        Close #2
        
        ' Store values back in temp file
        Open woTempData For Output As #2
        Write #2, woNum, woLastStartTime, woLastEndTime, woTotalTime  ' Write to Temp file WO number, WO start time, Elapsed time (0 for first punch), Total Time (0 for first punch)
    
        Close #2
        
    Case Else
        MsgBox ("Error in Temp File Handler")
End Select
    
    
End Sub

' -- runCalcs() - Runs calculations necessary to push WO data out when finish is pressed.
Private Sub runCalcs()
    
Let STYPE$ = "0"


'*********************CALC
'T.C COST/LB        4.81
'INCHES/LB         28.00
'T.C. COST/SQ"     $0.172
'MACH.P/DAY     $1344.00
'LABOR P/DAY    $1888.00
'TOTAL P/DAY    $3232.00
'CAP. UTIL%        80.00
'CAP.DAILY COST $3232.00
'PLANT COST/HR    404.00
'*************************************************************************

Let TCCOST! = 0.171821
Let HOCOST! = 404!
Let TTime! = woTotalTime / 60
Let AvgTime! = Format$((TTime! / Qty!), "CURRENCY")
Let TCCOST! = Format$((TCCOST! * Tc!), "CURRENCY")
Let HORATE! = 0.5


Let CONVCOST! = Format$((HOCOST! * AvgTime! * HORATE!), "CURRENCY")
Let UNITCOST! = Format$((TCCOST! + CONVCOST! + STLCOST!), "CURRENCY")
Let UNITPROFIT! = Format$((PRICE! - UNITCOST!), "CURRENCY")
'Let GROSSPROFIT! = Format$(((UNITPROFIT! / PRICE!) * 100), "CURRENCY")
Let TSALES! = Qty! * PRICE!
Let TTC! = Tc! * Qty!
Let TTCCOST! = TCCOST! * Qty!
Let TSTLCOST! = STLCOST! * Qty!
Let TCONVCOST! = CONVCOST! * Qty!
Let TOTALCOST! = UNITCOST! * Qty!
Let TOTALPROFIT! = UNITPROFIT! * Qty!
Let COM$ = COM1$ + Com2$ + Com3$ + COM4$ + COM5$ + COM6$ + COM7$

Let diff = SBT_AvgTime! - AvgTime!
If AvgTime! < SBT_AvgTime! Then Let diffcom = " Faster  "
If AvgTime! > SBT_AvgTime! Then Let diffcom = " Slower  "


Let TTime! = 0: Let ATSALES! = 0: Let ATTC! = 0: Let BTTIME! = 0: Let BTSALES! = 0: Let BTTC! = 0: Let CTTIME! = 0: Let CTSALES! = 0: Let CTTC! = 0: Let TUNITPROFIT! = 0
Let TTC! = 0
Let COM1$ = "": Let Com2$ = "": Let Com3$ = "": Let COM4$ = "": Let COM5$ = "": Let COM6$ = "": Let COM7$ = ""
Let STIME! = 0: Let ETIME! = 0: Let Qty! = 0: Let TTime! = 0: Let C$ = "": Let B$ = "": Let A$ = ""
Let WOK$ = ""
Let InString1 = ""


End Sub

Public Function woTimer(timerCmd As String) As Integer

Select Case timerCmd
    Case "Start", "Reset"
        woCountdownTime = woTimeoutTime * woTimeoutCalib
    Case "Decrement"
        If woCountdownTime > 0 Then
            woCountdownTime = woCountdownTime - 1
        End If
    Case "Get"
    
    Case Else
        MsgBox ("Inproper Input in woTimer Function")
    
End Select

woTimer = woCountdownTime

End Function
