VERSION 5.00
Begin VB.Form BAR4 
   BackColor       =   &H00C00000&
   Caption         =   "RESET"
   ClientHeight    =   1575
   ClientLeft      =   1560
   ClientTop       =   1770
   ClientWidth     =   6690
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form4"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1575
   ScaleWidth      =   6690
   Begin VB.CommandButton Command3 
      Caption         =   "CLEAR ENTRY"
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
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CHANGE W/O"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CANCEL"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Change Work Order Number On The Product You are Working On"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "BAR4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
Label1.Caption = "CANCEL"
BAR4.Hide
End Sub


Private Sub Command1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Const ENTER = 0
Const LEAVE = 1
Text1.Text = State
Select Case State
Case ENTER
Label1.Caption = "RETURN TO MAIN MENU   MAKE NO CHANGES"

Case LEAVE
Label1.Caption = ""
End Select


End Sub

Private Sub Command1_GotFocus()
Label1.Caption = "RETURN TO MAIN MENU   MAKE NO CHANGES"
End Sub


Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X > 500 And X < 1500 And Y > 100 And Y < 400 Then
Label1.Caption = "RETURN TO MAIN MENU   MAKE NO CHANGES"
Else
'Case LEAVE
Label1.Caption = ""
End If

End Sub

Private Sub Command2_Click()
Label1.Caption = "Change W/O"
'StartCount = 0
BAR4.Hide

End Sub


Private Sub Command2_GotFocus()
Label1.Caption = "YOU ARE ABOUT TO CHAGE WORK ORDER NUMBER  " & InString$
End Sub


Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X > 200 And X < 1600 And Y > 100 And Y < 400 Then
Label1.Caption = "YOU ARE ABOUT TO CHAGE WORK ORDER NUMBER  " & InString$
Else
'Case LEAVE
Label1.Caption = ""
End If


End Sub

Private Sub Command3_Click()
Label1.Caption = "CLEAR"
BAR4.Hide

End Sub

Private Sub Command3_GotFocus()
Label1.Caption = "CLEAR LAST ENTRY"
End Sub


Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X > 200 And X < 1600 And Y > 100 And Y < 400 Then
Label1.Caption = "CLEAR LAST ENTRY"
'Label1.Caption = "YOU ARE ABOUT TO CHAGE WORK ORDER NUMBER  " & InString$
Else
'Case LEAVE
Label1.Caption = ""
End If


End Sub

Private Sub Form_Activate()
 If Dir("F:\BARCODE\temp5.tmp") = "" Then
            'FileExists = False
            Label1.Caption = "Work Order Is Not Active"
            Exit Sub
        Else
            Open "F:\BARCODE\temp5.tmp" For Input As #1
            Input #1, InString$, Date1$
            Close #1
            Label1.Caption = "YOU ARE ABOUT TO CHAGE WORK ORDER NUMBER  " & InString$
            Exit Sub
        End If
End Sub

