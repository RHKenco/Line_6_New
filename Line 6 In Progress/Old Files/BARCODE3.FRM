VERSION 5.00
Begin VB.Form BAR3 
   BackColor       =   &H00C00000&
   Caption         =   "ERROR"
   ClientHeight    =   1740
   ClientLeft      =   4635
   ClientTop       =   2175
   ClientWidth     =   3345
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1740
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "FINISH"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "N/F"
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
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "BAR3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
If RunCondition5 = "TimeOut" Then
    Let RunCondition5 = ""
    BAR3.Hide
    cutlist.Label21.Caption = "Press Start  or  Ganged  or  Line 5/4"
    cutlist.Command6.Enabled = True
    cutlist.Command8.Enabled = True
    'cutlist.Command10.Enabled = True
    cutlist.Label20.Visible = False
    'cutlist.Text5.Visible = False
'    cutlist.Command13.Value = True
Kill "F:\BARCODE\temp6.tmp"
End If
End Sub

Private Sub Command2_Click()
If RunCondition5 = "TimeOut" Then
'Let RunCondition5 = ""
cutlist.Command3.Value = True
BAR3.Hide
End If
End Sub


Private Sub Form_Activate()

Label1.Caption = "LINE #6 TIMED-OUT"

End Sub


