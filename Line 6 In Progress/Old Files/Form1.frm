VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Lift & Tilt Table"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Use arrow keys to lift or tilt table"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text1.Text = "STOP"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 40
If Text1.Text = "STOP" Then
'If Text6.Text <> "DOWN" Then
Text1.Text = "DOWN"
c6k.Write ("1OUT.11-1" & Chr$(13))
End If

Case 38
If Text1.Text = "STOP" Then
'If Text6.Text <> "UP" Then
Text1.Text = "UP"
c6k.Write ("1OUT.28-1:" & Chr$(13))
End If

Case 37
If Text1.Text = "STOP" Then
    'If Text6.Text <> "LEFT" Then
    Text1.Text = "TILT DOWN"
    c6k.Write ("1OUT.29-1:" & Chr$(13))
    
End If

Case 39
If Text1.Text = "STOP" Then
    'If Text6.Text <> "RIGHT" Then
    Text1.Text = "TILT UP"
    c6k.Write ("1OUT.26-1:" & Chr$(13))
End If
Case Else
c6k.Write ("!1OUTALL26,29,0:!1OUT.11-0" & Chr$(13))
Text1.Text = "STOP"

End Select
End Sub



Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
Text1.Text = "STOP"
c6k.Write ("!1OUTALL26,29,0:!1OUT.11-0" & Chr$(13))
Text1.Text = "STOP"
End Sub
