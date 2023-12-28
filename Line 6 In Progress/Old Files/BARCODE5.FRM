VERSION 5.00
Begin VB.Form BAR5 
   BackColor       =   &H00FF0000&
   Caption         =   "Mat'l Input"
   ClientHeight    =   2400
   ClientLeft      =   4500
   ClientTop       =   2115
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2400
   ScaleWidth      =   4470
   Begin VB.CommandButton Command2 
      Caption         =   "   CUST. SUPPLIED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "Mat'l"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1560
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
      Begin VB.OptionButton Option2 
         Caption         =   "  T-1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   " A-36"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TC Inches"
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
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Length"
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
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
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
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Thickness"
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
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Information for Part# ________________"
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
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "BAR5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Val(Text1.Text) <= 0 Then
  MsgBox ("Enter Thickness  "), , ("ERROR")
  Text1.SetFocus
  Exit Sub
End If
If Val(Text2.Text) <= 0 Then
  MsgBox ("Enter Width  "), , ("ERROR")
  Text2.SetFocus
  Exit Sub
End If
If Val(Text3.Text) <= 0 Then
  MsgBox ("Enter Length  "), , ("ERROR")
  Text3.SetFocus
  Exit Sub
End If
If Val(Text4.Text) <= 0 Then
  MsgBox ("Enter TC Inches "), , ("ERROR")
  Text4.SetFocus
  Exit Sub
End If

Let THICK2 = Val(Text1.Text)
Let WIDTH2 = Val(Text2.Text)
Let LENGTH2 = Val(Text3.Text)
Let PN2 = PartNumber
Let DAT2 = Date
Let Tc = Val(Text4.Text)
If Option1.Value = True Then
  Let MATL2 = "A-36"
 Else
  Let MATL2 = "T-1"
End If
Let COM5$ = "OPERATOR"
Open "F:\MFG\DWGINFO.ASC" For Append As #3
     Write #3, PN2, DWG$, THICK2, WIDTH2, LENGTH2, MATL2, DAT2, Tc
 Close #3
Unload BAR5
cutlist.Show
End Sub

Private Sub Command2_Click()
If Val(Text4.Text) <= 0 Then
  MsgBox ("Enter TC Inches "), , ("ERROR")
  Text4.SetFocus
  Exit Sub
End If

Let THICK2 = 0
Let WIDTH2 = 0
Let LENGTH2 = 0
Let PN2 = PartNumber
Let DAT2 = Date
Let Tc = Val(Text4.Text)
  Let MATL2 = "CS"
Let COM5$ = "CUST SUPPLIED"
Open "F:\MFG\DWGINFO.ASC" For Append As #3
     Write #3, PN2, DWG$, THICK2, WIDTH2, LENGTH2, MATL2, DAT2, Tc
 Close #3
Unload Form5
Form1.Show
End Sub


Private Sub Form_Activate()
Label1.Caption = "Enter Info For Part #   " & PartNumber
Text1.SetFocus
Text1.Text = DWG2$
End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3.SetFocus

End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text4.SetFocus

End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Option1.SetFocus

End Sub


