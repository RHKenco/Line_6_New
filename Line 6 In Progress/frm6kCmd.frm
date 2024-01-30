VERSION 5.00
Begin VB.Form frm6kCmd 
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   12435
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Terminal_Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11520
      Top             =   2880
   End
   Begin VB.TextBox Terminal_Textbox 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3315
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   12180
   End
End
Attribute VB_Name = "frm6kCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_GotFocus()
    If connected Then Terminal_Timer.Enabled = True
End Sub

Private Sub Form_LostFocus()
    Terminal_Timer.Enabled = False
End Sub

Private Sub Terminal_Textbox_Change()
    'the text box has a finite buffer so
    'make sure it doesn't overflow
    
    If Len(Terminal_Textbox.Text) > 16000 Then
        Terminal_Textbox.Text = Right$(Terminal_Textbox.Text, 500)    'buffer just the last 500 characters
    End If
    
End Sub

Private Sub Terminal_Textbox_DblClick()
    Terminal_Textbox.Text = ""     'clear the terminal display
End Sub

Private Sub Terminal_Textbox_KeyPress(KeyAscii As Integer)
'this routine processes the terminal's key presses
On Error GoTo Terminal_Textboxkeypress_error
    
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
                Terminal_Textbox.Text = ""
                KeyAscii = 0
            Else
                buffer = buffer & Chr$(13)      'append the CR
                Terminal_Timer.Enabled = False          'disable response polling to avoid simultaneous read/write
                temp = c6k.Write(buffer)        'send commands to 6k
                Terminal_Timer.Enabled = True           'enable response polling
            End If
            buffer = ""                         'empty the command local buffer
        
        
        'anything else just add to the buffer
        Case Else
            buffer = buffer & Chr$(KeyAscii)    'append char to the local command buffer
            
    End Select
    Exit Sub
    
Terminal_Textboxkeypress_error:
    Unload Me
End Sub

Private Sub Terminal_Timer_Timer()
On Error GoTo Terminal_Timer_Err

    'this timer routine polls for response from the controller
    Dim temp$
    temp = c6k.Read()                           'get response
    If Len(temp) Then Terminal_Textbox.SelText = temp      'if not empty then display in the text box
    Exit Sub
    
Terminal_Timer_Err:
    Disconnect
End Sub
