VERSION 5.00
Begin VB.Form frm6kCmd 
   Caption         =   "Terminal"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   12435
   StartUpPosition =   3  'Windows Default
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
      Height          =   675
      Index           =   1
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3840
      Width           =   12180
   End
   Begin VB.Timer Terminal_Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11880
      Top             =   3360
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
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   12195
   End
   Begin VB.Label Label1 
      Caption         =   "Command Line:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   2655
   End
End
Attribute VB_Name = "frm6kCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prevCmd(9) As String
Dim lastCmd As Integer
Dim cursorPos As Integer

Private Sub Form_Load()
    Terminal_Timer.Enabled = True
End Sub

Private Sub Form_GotFocus()
    Terminal_Timer.Enabled = True
End Sub

Private Sub Form_LostFocus()
    'Terminal_Timer.Enabled = False
End Sub



Private Sub Terminal_Textbox_Change(Index As Integer)
    'the text box has a finite buffer so
    'make sure it doesn't overflow
    
    If Len(Terminal_Textbox(Index).Text) > 16000 Then
        Terminal_Textbox(Index).Text = Right$(Terminal_Textbox(Index).Text, 500)    'buffer just the last 500 characters
    End If
    
End Sub

Private Sub Terminal_Textbox_DblClick(Index As Integer)
    Terminal_Textbox(Index).Text = ""     'clear the terminal display
End Sub

Private Sub Terminal_Textbox_KeyPress(Index As Integer, KeyAscii As Integer)
'this routine processes the terminal's key presses
On Error GoTo Terminal_Textboxkeypress_error

    
Dim temp%
Static buffer$      'local command buffer

'Reset the cursor position
'Terminal_Textbox(1).SelStart = Len(Text1.Text)


'perform action based on value of key being pressed
Select Case KeyAscii
    'backspace
    Case 8
        If Len(buffer) > 0 Then buffer = Left$(buffer, Len(buffer) - 1) 'erase one char from buffer
    
    'CR or colon - 6000 command delimeter
    Case 13, Asc(":")
        
        prevCmd(lastCmd) = buffer
        Terminal_Textbox(0).SelText = " >    " & buffer & Chr(13) & Chr(13)
        
        buffer = buffer & Chr$(13)      'append the CR
        Terminal_Timer.Enabled = False          'disable response polling to avoid simultaneous read/write
        temp = c6k.Write(buffer)        'send commands to 6k
        Terminal_Timer.Enabled = True           'enable response polling
            
        buffer = ""                         'empty the command local buffer
    
    
    'Any Normal Input, load into buffer
    Case Else
        If KeyAscii > 31 And KeyAscii < 127 Then
    
            If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
            buffer = buffer & Chr$(KeyAscii)    'append char to the local command buffer
        End If
End Select

'Reset terminal to match the buffer
Terminal_Textbox(1).Text = " >    " & buffer

'Reset the cursor position
'Terminal_Textbox(1).SelStart = Len(Text1.Text)

Exit Sub

    
Terminal_Textboxkeypress_error:
    'Unload Me
End Sub

Private Sub Terminal_Timer_Timer()
On Error GoTo Terminal_Timer_Err

    'this timer routine polls for response from the controller
    Dim temp$
    temp = c6k.Read()                           'get response
    If Len(temp) Then Terminal_Textbox(0).SelText = temp      'if not empty then display in the text box
    Exit Sub
    
Terminal_Timer_Err:
    'Unload Me
End Sub
