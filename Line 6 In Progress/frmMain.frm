VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Line 6"
   ClientHeight    =   2175
   ClientLeft      =   1260
   ClientTop       =   2130
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton FastStatus_View_Cmd 
      Caption         =   "FastStatus"
      Height          =   420
      Left            =   5520
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Launch_Cmd 
      Caption         =   "Launch Line 6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4200
      TabIndex        =   7
      Top             =   240
      Width           =   2295
   End
   Begin VB.OptionButton Eth_Rs_2 
      Caption         =   "RS232"
      Height          =   195
      Left            =   225
      TabIndex        =   6
      Top             =   180
      Width           =   1080
   End
   Begin VB.OptionButton Eth_Rs_1 
      Caption         =   "Ethernet"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   465
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Upload_Program_Cmd 
      Caption         =   "Upload File"
      Height          =   420
      Left            =   2160
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.CommandButton Download_OS_Cmd 
      Caption         =   "Download OS"
      Height          =   420
      Left            =   4200
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Terminal_Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   225
      Top             =   975
   End
   Begin VB.CommandButton Connect_Cmd 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1395
      TabIndex        =   2
      Top             =   240
      Width           =   2670
   End
   Begin VB.CommandButton Download_Program_Cmd 
      Caption         =   "Download File"
      Height          =   420
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1830
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
      Height          =   675
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   780
      Width           =   10500
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit





Private Sub Download_Program_Cmd_Click()
On Error GoTo cmd1err

    If (Not connected) Then Exit Sub    'exit if not connected
    
    Terminal_Timer.Enabled = False          'disable response polling to avoid simultaneous read/write
    If (c6k.SendFile("") > 0) Then  'download program files - empty string means to prompt for filename
        c6k.Write ("TDIR" & Chr$(13))         'send TDIR command
    End If
    Terminal_Timer.Enabled = True           'enable response polling
    Terminal_Textbox.SetFocus
    Exit Sub
    
cmd1err:
    Disconnect
End Sub


Private Sub Connect_Cmd_Click()
On Error GoTo cmd2err
    Dim fh As Long      'file handle
Dim i



Shell ("c:\6K.BAT")
For i = 1 To 100000
Next i

        fh = FreeFile   'get first avaiable handle
    
    'disconnect if already connected
    If connected Then
        
        Terminal_Timer.Enabled = False      'disable response polling
         Set c6k = Nothing           'disconnect and free up the c6k object
        connected = False           'set connection flag to false
    Else
    
    If Eth_Rs_1.value Then   'ethernet
        ' use this code for Ethernet
        Dim ipaddr$
            ipaddr = "192.168.2.2"
    
        Set c6k = CreateObject("COM6SRVR.NET")
        If c6k.Connect(ipaddr) > 0 Then
            c6k.Write "TREV" & vbCr    'send TREV command
 
            Terminal_Timer.Enabled = True       'enable response polling
            connected = True            'set connected flag to true
        Else
            Terminal_Timer.Enabled = False      'disable response polling (default)
            connected = False           'set connected flag to false
            MsgBox "Connection attempt failed...", 0, "Status"
        End If
Shell ("f:\apps\exe\flow sensor\fs6.exe")
    '***************2nd 6k *****************
'   Dim ipaddr1$
'            ipaddr1 = "192.168.2.3"
'
'        Set c6k1 = CreateObject("COM6SRVR.NET")
'        If c6k1.Connect(ipaddr1) > 0 Then
'            c6k1.Write "TREV" & vbCr    'send TREV command
'
'            Terminal_Timer.Enabled = True       'enable response polling
'            connected1 = True            'set connected flag to true
'        Else
'            Terminal_Timer.Enabled = False      'disable response polling (default)
'            connected = False           'set connected flag to false
'            MsgBox "Connection attempt failed...", 0, "Status"
'        End If
'    Else    'RS232
'        ' use this code for RS232
'        Dim commport$
'            commport = "2"
'
'        'attempt to open file where ip address is stored if file exists
'        If Len(Dir$("commport.dat")) Then
'            Open "commport.dat" For Input As #fh
'                Line Input #fh, commport
'            Close #fh
'        End If
'
'        'prompt for com port number using default
'        commport = InputBox("Enter PC COMPORT number.", "Port Setting", commport)
'        If Len(commport) = 0 Then Exit Sub
'
'        'save user specified ipaddr
'        Open "commport.dat" For Output As #fh
'            Print #fh, commport
'        Close #fh
'
'
'        Set c6k = CreateObject("COM6SRVR.RS232")
'        If c6k.Connect(CInt(commport)) > 0 Then
'            c6k.Write "TREV" & vbCr     'send TREV command
'            Terminal_Timer.Enabled = True       'enable response polling
'            connected = True            'set connected flag to true
'        Else
'            Terminal_Timer.Enabled = False      'disable response polling (default)
'            connected = False           'set connected flag to false
'            MsgBox "Connection attempt failed...", 0, "Status"
'        End If
    End If
    
    End If
    
    If connected Then
        Connect_Cmd.Caption = "Disconnect"
        Eth_Rs_1.Enabled = False
        Eth_Rs_2.Enabled = False
       
        
    Else
        Connect_Cmd.Caption = "Connect"
        Eth_Rs_1.Enabled = True
        Eth_Rs_2.Enabled = True
        Set c6k = Nothing           'release the comm server
        Launch_Cmd.Enabled = False
    End If
    'Terminal_Textbox.SetFocus
     Launch_Cmd.Enabled = True
    Exit Sub
    
cmd2err:
    Disconnect
End Sub

Private Sub Download_OS_Cmd_Click()
On Error GoTo cmd3err

    If (connected And Eth_Rs_2.value) Then
        Terminal_Timer.Enabled = False      'disable response polling
        c6k.SendOS ("")             'download the Operating System - prompt for OS file
        Terminal_Timer.Enabled = True       'enable response polling
        c6k.Write ("TREV" & Chr$(13))
        Terminal_Textbox.SetFocus
    Else
        MsgBox "Operating System download is only supported via RS232.", 0, "OS Download Unavailable"
    End If
    
    Exit Sub
    
cmd3err:
    Disconnect
End Sub

Private Sub Upload_Program_Cmd_Click()
On Error GoTo cmd4err

    If (Not connected) Then Exit Sub    'exit if not connected
    
    Terminal_Timer.Enabled = False      'disable response polling to avoid simultaneous read/write
    c6k.GetFile ("")            'upload program files - empty string means to prompt for filename
    Terminal_Timer.Enabled = True       'enable response polling
    Terminal_Textbox.SetFocus
    Exit Sub
    
cmd4err:
    Disconnect
End Sub


Private Sub FastStatus_View_Cmd_Click()
If (Terminal_Timer.Enabled And Eth_Rs_1.value) Then
        Me.Hide
       frmFastStatus.Show
    Else
        MsgBox "An ethernet connection is needed for Fast Staus display.", 0, "Display Unavailable"
    End If

End Sub

Private Sub Launch_Cmd_Click()
    If (Terminal_Timer.Enabled And Eth_Rs_1.value) Then
        Terminal_Timer.Enabled = False
        Me.Hide
        frmLine6.Show
       
       ' frmFastStatus.Show
    Else
        MsgBox "An ethernet connection is needed for Fast Staus display.", 0, "Display Unavailable"
    End If
End Sub

Private Sub Form_GotFocus()
    If connected Then Terminal_Timer.Enabled = True
End Sub

Private Sub Form_Load()
    connected = False   'connection disabled by default
    Launch_Cmd.Enabled = False
    
End Sub


Private Sub Form_LostFocus()
    Terminal_Timer.Enabled = False
End Sub


Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        Terminal_Textbox.Width = Me.Width - 345
        Terminal_Textbox.Height = Me.Height - 1275
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
'make sure to disconnect on unload
    If connected Then
        Set c6k = Nothing
    End If
    'c6kOps = Nothing
    'fsmMain = Nothing
    'fsmRun = Nothing
    'woMgr = Nothing
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
    Disconnect
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
