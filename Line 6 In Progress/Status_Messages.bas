Attribute VB_Name = "Status_Messages"

'Message Definitions
Public Const msgInactive As Integer = 1
Public Const msgActive As Integer = 2
Public Const msgActiveBlade As Integer = 3
Public Const msgActiveAugFace As Integer = 4
Public Const msgActiveAugEdge As Integer = 5

Public Const msgStarted As Integer = 6
Public Const msgStrike As Integer = 7
Public Const msgToggle As Integer = 8
Public Const msgPaused As Integer = 9
Public Const msgRunning As Integer = 10
Public Const msgCompleted As Integer = 11
Public Const msgReturning As Integer = 12
Public Const msgNextPass As Integer = 13

Public Const msgNotFinished As Integer = 14
Public Const msgFinished As Integer = 15
Public Const msgTimeout As Integer = 16

Public Sub statusMsg(newStatus As Integer, Optional displayStr As String)

'Exit sub if Msg hasn't changed
Static currentMsg
Static currentStr
If (currentMsg = newStatus) And (currentStr = displayStr) Then Exit Sub

' If input is "Active" without discrimination for auger setup, adjust state accordingly
If newStatus = msgActive Then
    If Not c6kOps.getAugerSet() Then
        newStatus = msgActiveBlade
    Else
        If frmLine6.Option_Auger_Direction(0) Then newStatus = msgActiveAugFace Else newStatus = msgActiveAugEdge
    End If
End If

Dim tempLabel As String
    
Select Case newStatus
    Case msgInactive
        frmLine6.Var_Label_System_Status.Caption = ""
        frmLine6.Var_Label_System_Status.Visible = False
        frmLine6.Var_Label_System_Status.Refresh
        Exit Sub
    Case msgActiveBlade
        tempLabel = "Ready to Start Blade"
    Case msgActiveAugEdge
        tempLabel = "Ready to Start Auger Edge"
    Case msgActiveAugFace
        tempLabel = "Ready to Start Auger Face"
  
  
    Case msgStarted
        If displayStr = "" Then GoTo strInError
        tempLabel = "Started - " & displayStr & " Remaining"
    Case msgStrike
        tempLabel = "Move to strike location, then press Release"
    Case msgToggle
        tempLabel = "Please flip switch to proceed"
    Case msgRunning
        tempLabel = "Running Pass. Flip Switch to Pause"
    Case msgPaused
        tempLabel = "Pass Paused. Flip Switch to Resume or Press Release to Finish"
    Case msgCompleted
        tempLabel = "Pass Finished. Press Release to finish, or flip switch to return to pass start"
    Case msgReturning
        tempLabel = "Returning to pass start - Please wait"
    Case msgNextPass
        tempLabel = "Flip Switch to begin next pass"
 
 
    Case msgNotFinished
        tempLabel = "N/F Pressed - Press Start to Resume"
    Case msgFinished
        If displayStr = "" Then GoTo strInError
        tempLabel = "Finished " & displayStr & " Parts." & Chr(13) & "Press start for next set of parts," & Chr(13) & "or press clear to enter new WO"
    Case msgTimeout
        tempLabel = "System Timed Out - Press Start to Resume"

    Case Else
        MsgBox "Error in Status Message Display - Case: " & newStatus
End Select

frmLine6.Var_Label_System_Status.Caption = tempLabel
frmLine6.Var_Label_System_Status.Visible = True
frmLine6.Var_Label_System_Status.Refresh

currentMsg = newStatus
currentStr = displayStr

Exit Sub

strInError:

MsgBox "Display String Error - String must not be null"

End Sub
