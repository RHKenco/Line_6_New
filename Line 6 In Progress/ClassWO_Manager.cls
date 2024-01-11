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
Private Const woFormASCpath As String = "C:\line 6\mfg\testFORM.ASC" '"F:\MFG\FORM.ASC"
Private Const woBarDataPath As String = "C:\line 6\mfg\testBar_Data.dta" '"F:\MFG\Bar_Data.dta"
Private Const woGrossPITpath As String = "C:\line 6\mfg\testGROSSPIT.GPA" '"F:\MFG\GROSSPIT.GPA"
Private Const woGaryPath As String = "C:\line 6\mfg\testGARY.GPA" '"F:\mfg\GARY.GPA"
Private Const woGoalPath As String = "C:\line 6\mfg\testGOAL.GPA" '"F:\mfg\GOAL.GPA"
Private Const woTempMain As String = "C:\line 6\files\temp6.temp" '"F:\BARCODE\temp6.temp"
Private Const woTempDir As String = "C:\line 6\files\" '"F:\BARCODE\"

Private woTempData As String        ' Varible to store path to WO temp data

'Const Strings for File Write
Private Const machineNo = 6

Private woLastStartTime As String       ' Start time of WO
Private woLastEndTime As String         ' Not-Finish Time of WO
Private woTotalTime As Single       ' Total time since last Finish

'Internal Variables
Private woActive As Boolean         ' Work Order Active flag
Private woStarted As Boolean        ' Work Order Started Flag
Private woCountdownTime As Integer  ' Countdown timer variable for timeout time

Private woNum As String             'Form.ASC Data: - Work Order Number
Private woPartNum As String         '               - Part Number
Private woTqty As Single            '               - Total WO Quantity
Private woPrice As Single           '               - Price of Item
Private woTCinch As Single          '               - TC Inches on part
Private woFormASCcom As String        '             - Comments in Form.ASC
Private woDate As String            '               - Due Date on WO
Private woDWGno As String           '               - DWG Number for item
Private woCScost As String          '               - Unused Column
Private woSBTavg As String          '               - Unused Column


Private Const cTCcostLB As Single = 4.81        'Calculation Constants: - T.C COST/LB        4.81
Private Const cTCinLB As Single = 28#           '                       - INCHES/LB         28.00
Private Const cTCcostSQ As Single = 0.171821    '                       - T.C. COST/SQ"     $0.172
Private Const cCapUtil As Single = 80#          '                       - CAP. UTIL%        80.00
Private Const cDayCost As Single = 3232#        '                       - CAP.DAILY COST $3232.00  - 'MACH.P/DAY     $1344.00, 'LABOR P/DAY    $1888.00, 'TOTAL P/DAY    $3232.00
Private Const cHrCost As Single = 404           '                       - PLANT COST/HR    404.00
Private Const cHrRate As Single = 0.5           '                       - ?



Private TTime!
Private AvgTime!
Private woTCcostSQ

Private CONVCOST!
Private UNITCOST!
Private UNITPROFIT!
Private GROSSPROFIT!
Private TSALES!
Private TTC!
Private TTCCOST!
Private TSTLCOST!
Private TCONVCOST!
Private TOTALCOST!
Private TOTALPROFIT!




Public Function isWOactive() As Boolean
    isWOactive = woActive
End Function

Public Function isWOstarted() As Boolean
    isWOstarted = woStarted
End Function

Public Function getWOnum() As String
    getWOnum = woNum
End Function

Public Sub loadWO()

'-- Load FORM.ASC Data & Populate Necessary Fields ------------------------------------------
If frmLine6.Text_Enter_WO.Text <> frmLine6.Text_Enter_WO.Tag And frmLine6.Text_Enter_WO.Text <> "" Then

    'Pull WO Number entered in text box
    woNum = frmLine6.Text_Enter_WO.Text
    
    'Load data from FORM.ASC; if successful, then
    If pullFormASC(woNum) Then
        
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
        
        Call statusMsg("Active Blade", "")
        
        'Update Buttons
        Call btnState("Active Blade")
        
        frmLine6.Refresh
        
        'Set up temporary file to save current active work-order (in case of program crashing)
        Open woTempMain For Output As #1
        Write #1, woNum, Now
        Close #1
        
        'Create temp file to store wo time data
        Let woTempData = woTempDir & woNum & "6.tmp"
        
        If Dir(woTempData) = "" Then
            Call updTemp("New")
        Else
            Call updTemp("Load")
        End If
    End If
End If

End Sub

Public Sub clearWO()
    '-- Clear All Fields and Reset Button to Enter ------------------------------------------------
    frmLine6.Text_Enter_WO.Text = frmLine6.Text_Enter_WO.Tag
    frmLine6.Text_Enter_WO.Locked = False
    
    frmLine6.Text_Pop_Part_Num.Text = ""
    frmLine6.Text_Pop_Dwg_Num.Text = ""
    frmLine6.Text_Pop_Due_Date.Text = ""

    'Update Work Order Status & System Status Labels
    frmLine6.Var_Label_WO_Active.Caption = frmLine6.Var_Label_WO_Active.Tag
    frmLine6.Var_Label_WO_Active.Visible = True
    
    Call statusMsg("Inactive", "")
    
    'Clear Stored FORM.ASC Variables
    woNum = ""
    woPartNum = ""
    woTqty = 0
    woPrice = 0
    woTCinch = 0
    woFormASCcom = ""
    woDate = ""
    woDWGno = ""
    woActive = False
    
    'Clear Auger Data
    Call c6kOps.clearAugerParam
    
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
   Exit Sub
Else
    
    ' Prompt the user to check if they would like to load the temp file
    If MsgBox("Previously Active WO Detected. Would you like to load?", vbYesNo) = vbNo Then Exit Sub
    
    Dim DAT As String
    
    ' Load temp file to check WOnum (temp6) and Date
    Open (woTempMain) For Input As #2
    Input #2, woNum, DAT
    Close #2
    
    'Define temp file directory path
    Let woTempData = woTempDir & woNum & "6.tmp"
    
    ' Load backup data temp file & store in mgr variables
    Call updTemp("Load")
    
    
    ' ***** ADD CODE HERE ***** Pseudocode:
        ' Checks current time - if the current time is more than timeout time different from the temp file,
        ' Punch a not-finish on the WO with time based on the timeout period & Exit Sub
        
    
    ' Update Form text and call function to load WO
    frmLine6.Text_Enter_WO.Text = woNum
    
    Call loadWO
    
End If

End Sub



Public Sub startWO()
    Call updTemp("Start")
    Call writeData("Start")
    Call btnState("Started")
    
    'Start Timeout Timer
    Call woTimer("Start")
    woStarted = True
            
End Sub

Public Sub notfinishWO()
    Call updTemp("End")
    Call writeData("N/F")
    Call btnState("Active")
    
    'Stop Timeout Timer
    Call woTimer("Reset")
    woStarted = False
    
    Call statusMsg("Not-Finish", "")

End Sub

Public Sub finishWO()
    Call updTemp("End")
    Call writeData("Finish")
    Call btnState("Active")

    'Stop Timeout Timer
    Call woTimer("Reset")
    woStarted = False
    
End Sub

Private Sub timeoutWO()
    Call writeData("TimeOut")
    Call btnState("Active")
    
    Call systemMsg("Timeout")

End Sub

Private Function pullFormASC(WORK_ORDER As String) As Boolean

' Load FORM.ASC
Open woFormASCpath For Input As #1

Dim WOK$
Dim UNUSED_1!
Dim UNUSED_2!

Do While woNum <> WOK$
     Input #1, WOK$, woPartNum, woTqty, woPrice, woTCinch, woFormASCcom, woDate, woDWGno, UNUSED_1!, UNUSED_2!
        If woNum = WOK$ Then
            
            Close #1
            
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

    Case "Finish"
        
        'Prompt User for Quantity Finished
        Dim FinishedQty As Integer
        FinishedQty = Val(InputBox("ENTER QTY FINISHED", "Qty", 0, 4000, 1500))
        
        If FinishedQty < 1 Then
            MsgBox ("Finish Punch Canceled")
            Exit Sub
        End If
        
        Call runCalcs(FinishedQty)
        
        'Sta$ appears to indicate the line on which a part was impregnated
        'Alternative values could occur if the temp file comment is set to specific values, but the comment is never set in this program
        Dim Sta$
        Let Sta$ = "6"
        
        'STLCost! appears in some of the grosspit calculations. Looking at the data in grosspit, however, STLCost appears to only be set in the Line 2 and Line 3 programs.
        Dim STLCOST!
        Let STLCOST! = 0
        
        'Write Finished Punch & Data to Bar_Data
        Write #1, woNum, machineNo, "Finish", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss"), woPartNum, FinishedQty, AvgTime!, STLCOST!, GROSSPROFIT! & "%", woFormASCcom
        Close #1
        
        'Output to GROSSPIT
        Open woGrossPITpath For Append As #1
        Write #1, woNum, PartNumber, woTqty, FinishedQty, Sta$, Date$, woPrice, woTCinch, STLCOST!, AvgTime!, cTCcostSQ, CONVCOST!, UNITCOST!, UNITPROFIT!, GROSSPROFIT!, woFormASCcom, TTime!, TSALES!, TTC!, TTCCOST!, TSTLCOST!, TCONVCOST!, TOTALCOST!, TOTALPROFIT!, woDate, woDWGno, woCScost
        Close #1
        
        'Output to GARY
        Open woGaryPath For Append As #1
        Write #1, woNum, PartNumber, woTqty, FinishedQty, Sta$, Date$, woPrice, woTCinch, STLCOST!, AvgTime!
        Close #1
        
        'Output to GOAL
        Open woGoalPath For Append As #1
        Write #1, woNum, PartNumber, woTqty, FinishedQty, Sta$, Date$, woPrice, woTCinch, STLCOST!, AvgTime!
        Close #1
        
        'Write status message to Form
        Call statusMsg("Finish", CStr(FinishedQty))
        
    Case Else
        MsgBox ("Error in Punch Handler")
        
End Select

End Sub

' -- updTemp() - updates temp file to track work-order progress between Start and Finish punches
Private Sub updTemp(cmd As String)

Dim woNumTemp$
Dim woLastStartTimeTemp$
Dim woLastEndTimeTemp$
Dim woTotalTimeTemp!
Dim woIsTimedout As Boolean
woIsTimedout = False

Select Case cmd
    Case "New"
    
        Open woTempData For Output As #2
        Write #2, woNum, woLastStartTime$, 0, 0  ' Write to Temp file WO number, WO start time, Elapsed time (0 for first punch), Total Time (0 for first punch)
        
        Close #2
        
    Case "Load"
        
        'If the file does not exist to be loaded
        If Dir(woTempData) = "" Then
            MsgBox ("Temp File does not exist. The WO will be activated in the form, but no prior data will be recovered.")
            Exit Sub
        End If
        
        Open woTempData For Input As #2
        Input #2, woNumTemp$, woLastStartTime$, woLastEndTime$, woTotalTime  ' Write to Temp file WO number, WO start time, Elapsed time (0 for first punch), Total Time (0 for first punch)
        
        Close #2
        
    Case "Start"
        
        ' Load last Start punch timestamp from woTempData
        Open woTempData For Input As #2
        Input #2, woNumTemp$, woLastStartTimeTemp$, woLastEndTimeTemp$, woTotalTimeTemp!  ' Write to Temp file WO number, WO start time, Elapsed time (0 for first punch), Total Time (0 for first punch)
        
        ' If the timestamp does not match the one stored in the class, there is probably an error
        If woLastStartTime <> woLastStartTimeTemp$ Then MsgBox ("Start time match error")
        
        ' Get current time as Start time
        woLastStartTime$ = Now()
        
        Close #2
        
        ' Store values back in temp file
        Open woTempData For Output As #2
        Write #2, woNum, woLastStartTime$, woLastEndTime$, woTotalTime  ' Write to Temp file WO number, WO start time, Elapsed time (0 for first punch), Total Time (0 for first punch)
    
        Close #2
        
    Case "End"
        woIsTimedout = False
        
        GoTo SaveData:
        
    Case "Timeout"
        woIsTimedout = True
        
        GoTo SaveData:
        
    Case Else
        MsgBox ("Error in Temp File Handler")
End Select
    
Exit Sub

SaveData:
    
        ' Load last Start punch timestamp from woTempData
        Open woTempData For Input As #2
        Input #2, woNumTemp$, woLastStartTimeTemp$, woLastEndTimeTemp$, woTotalTimeTemp!  ' Write to Temp file WO number, WO start time, Elapsed time (0 for first punch), Total Time (0 for first punch)
        
        ' If the timestamp does not match the one stored in the class, there is probably an error
        If woLastStartTime <> woLastStartTimeTemp$ Then MsgBox ("Start time match error")
        
        ' Get current time as N/F time
        woLastEndTime$ = Now()
        
        ' Generate new total time based on last start time, current N/F time, and previous total time
        woTotalTime = DateDiff("s", woLastEndTime$, woLastStartTime$) + woTotalTimeTemp!
        
        If woIsTimedout Then woTotalTime = woTotalTime - (woTimeoutTime)
        
        Close #2
        
        ' Store values back in temp file
        Open woTempData For Output As #2
        Write #2, woNum, woLastStartTime$, woLastEndTime$, woTotalTime  ' Write to Temp file WO number, WO start time, Elapsed time (0 for first punch), Total Time (0 for first punch)
    
        Close #2
    
End Sub

' -- runCalcs() - Runs calculations necessary to push WO data out when finish is pressed.
Private Sub runCalcs(finQt As Integer)

'Unused Value - See writeData function for more details.
Dim STLCOST!
STLCOST! = 0

Let TTime! = woTotalTime / 3600                                 'Total time is in seconds, divide by 3600 to get to hours.
Let AvgTime! = Format$((TTime! / finQt), "CURRENCY")
Let woTCcostSQ = Format$((cTCcostSQ * woTCinch), "CURRENCY")


Let CONVCOST! = Format$((cHrCost * AvgTime! * cHrRate), "CURRENCY")
Let UNITCOST! = Format$((cTCcostSQ + CONVCOST! + STLCOST!), "CURRENCY")
Let UNITPROFIT! = Format$((woPrice - UNITCOST!), "CURRENCY")
Let GROSSPROFIT! = Format$(((UNITPROFIT! / woPrice) * 100), "CURRENCY")
Let TSALES! = finQt * woPrice
Let TTC! = Tc! * finQt
Let TTCCOST! = cTCcostSQ * finQt
Let TSTLCOST! = STLCOST! * finQt
Let TCONVCOST! = CONVCOST! * finQt
Let TOTALCOST! = UNITCOST! * finQt
Let TOTALPROFIT! = UNITPROFIT! * finQt

End Sub

Public Function woTimer(timerCmd As String) As Integer

Select Case timerCmd
    Case "Start", "Reset"
        woCountdownTime = woTimeoutTime * woTimeoutCalib
    Case "Count"
        If woCountdownTime > 0 Then
            woCountdownTime = woCountdownTime - 1
            
            Dim tempMin As Integer
            Dim tempSec As Integer
            Dim tempStr As String
            tempMin = CInt((woCountdownTime / woTimeoutCalib) \ 60)
            tempSec = CInt((woCountdownTime / woTimeoutCalib) - tempMin * 60)
            
            tempStr = CStr(tempMin) & ":" & Format(tempSec, "00")
            
            Call statusMsg("Started", tempStr)
            
        Else
            
            Call timeoutWO
            
        End If
    
    Case Else
        MsgBox ("Inproper Input in woTimer Function")
    
End Select

woTimer = woCountdownTime

End Function
