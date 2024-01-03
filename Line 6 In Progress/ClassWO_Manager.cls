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

Private Const woTimeoutTime = 300   'Work Order Timeout Clock Time in Seconds
Private Const woTimeoutCalib = 20   'Calibration constant - cycle of timer in Hz

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

Private woStartTime As String       ' Start time of WO
Private woEndTime As String         ' Finish / Not-Finish Time of WO



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
End Sub

'-- btnState - Function to handle visibility and enabling of Enter, Go, Start, N/F, and Finish Buttons -------------------------------------
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

Private Sub notfinishWO()

End Sub

Private Sub finishWO()

Dim DAT As String

'-- Check to make sure a WO has been loaded
If Dir("F:\BARCODE\temp6.tmp") = "" Then
    'FileExists = False
     MsgBox ("No Active Work Order "), , ("ERROR")
     Exit Sub
End If

'-- Ask user quantity finished
Let Qty! = Val(InputBox("ENTER QTY FINISHED", "Qty", 0, 4000, 1500))
If Qty! = 0 Then Exit Sub
    
    
    
Open "f:\mfg\FORM.ASC" For Input As #5
Do Until EOF(5)
Input #5, WOK$, PartNumber, TQTY!, PRICE!, Tc!, COMM$, DAT$, DWG$, CSCOST!, SBT_AvgTime!
   If InString$ = WOK$ Then
      If TQTY! < Qty! Then
        MsgBox ("You Entered More Than The Work Order QTY"), , ("ERROR")
      Close #5
      Exit Sub
     End If
    Exit Do
    Else
    Let WOK$ = "": PartNumber = "": TQTY! = 0: PRICE! = 0.00001: Tc! = 0: COMM$ = "": DAT$ = "": DWG$ = "no_dwg": CSCOST! = 0
   End If
 Loop
 Close #5
If PartNumber = "" Then Let PartNumber = InString$ & " Override"
Open "F:\BARCODE\temp6.tmp" For Input As #2
Input #2, temp6, DAT
Close #2

Open "F:\BARCODE\" + temp6 + "6.tmp" For Input As #2
Input #2, InString1, STIME!, ETIME!, TTime!, Com2$
Close #2
            
'Compute Elapsed Time
Let ETIME! = Timer: Let TTime! = ((ETIME! - STIME!) / 60) + TTime!

Kill "F:\BARCODE\" + temp6 + "6.tmp"
Let STIME! = STIME! / 3600: Let ETIME! = ETIME! / 3600:
Label21.Caption = "RunMode - OFF"
Label17.Caption = Format(Now, "SHORT TIME")
            
            
Kill "F:\BARCODE\temp6.tmp"
'********** relay off **********
' ULStat% = cbDBitOut%(0, 10, 2, 0)
'   If ULStat% <> 0 Then Stop
 Timer2.Enabled = False
'Text4.Visible = False
'Label19.Visible = False
 Label21.Caption = "Press Start  or  Ganged  "


Command6.Enabled = True
Command8.Enabled = True
Command13.Enabled = True




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
Let TTime! = TTime! / 60
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
MSFlexGrid1.AddItem InString$ & Chr(9) & PartNumber & Chr(9) & "4" & Chr(9) & Qty! & Chr(9) & SBT_AvgTime! & Chr(9) & AvgTime! & Chr(9) & Format((diff * 60), "####0.00") & diffcom & Chr(9) & COM$, 1
Open "F:\MFG\Bar_Data.dta" For Append As #1
If RunCondition5 = "TimeOut" Then
    RunCondition5 = ""
    Write #1, InString$, "6", "Finish", Format(Now, "yy-mm-dd"), "Timed Out", PartNumber, Qty!, AvgTime!, STLCOST!, GROSSPROFIT! & "%", COM$
Else
    Write #1, InString$, "6", "Finish", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss"), PartNumber, Qty!, AvgTime!, STLCOST!, GROSSPROFIT! & "%", COM$
End If
Close #1

Let Sta$ = "6"

19495 Open "F:\MFG\GROSSPIT.GPA" For Append As #1
Write #1, WOK$, PartNumber, TQTY!, Qty!, Sta$, Date$, PRICE!, Tc!, STLCOST!, AvgTime!, TCCOST!, CONVCOST!, UNITCOST!, UNITPROFIT!, GROSSPROFIT!, COM$, TTime!, TSALES!, TTC!, TTCCOST!, TSTLCOST!, TCONVCOST!, TOTALCOST!, TOTALPROFIT!, DAT$, DWG$, CSCOST!
Close #1

Open "F:\mfg\GARY.GPA" For Append As #1
Write #1, WOK$, PartNumber, TQTY!, Qty!, Sta$, Date$, PRICE!, Tc!, STLCOST!, AvgTime!
Close #1
Open "F:\mfg\GOAL.GPA" For Append As #1
Write #1, WOK$, PartNumber, TQTY!, Qty!, Sta$, Date$, PRICE!, Tc!, STLCOST!, AvgTime!
Close #1

Let TTime! = 0: Let ATSALES! = 0: Let ATTC! = 0: Let BTTIME! = 0: Let BTSALES! = 0: Let BTTC! = 0: Let CTTIME! = 0: Let CTSALES! = 0: Let CTTC! = 0: Let TUNITPROFIT! = 0
Let TTC! = 0
Let COM1$ = "": Let Com2$ = "": Let Com3$ = "": Let COM4$ = "": Let COM5$ = "": Let COM6$ = "": Let COM7$ = ""
Let STIME! = 0: Let ETIME! = 0: Let Qty! = 0: Let TTime! = 0: Let C$ = "": Let B$ = "": Let A$ = ""
Let WOK$ = ""
Let InString1 = ""

Exit Sub
errorhand:
If Err.Number = 70 Then
Form7.Show 1
Resume 0
Else
Resume Next
End If

End Sub

Private Function pullFormASC(WORK_ORDER As String) As Boolean

' Load FORM.ASC
Open "F:\MFG\FORM.ASC" For Input As #1
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
