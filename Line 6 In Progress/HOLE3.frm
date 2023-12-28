VERSION 5.00
Begin VB.Form Save_File 
   BackColor       =   &H00C00000&
   Caption         =   "SAVE FILE"
   ClientHeight    =   4170
   ClientLeft      =   4665
   ClientTop       =   1515
   ClientWidth     =   4515
   LinkTopic       =   "Form13"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4170
   ScaleWidth      =   4515
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   3795
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Save_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim YA(5)
checkfile = Dir("G:\ACAD\ABLASCII\" & Text1.Text & ".dta")
checkfile1 = Dir("G:\ACAD\ABLASCII\" & Text1.Text & ".DTA")
checkfile2 = Dir("G:\ACAD\ABLASCII\" & Text1.Text)
If checkfile <> "" Or checkfile1 <> "" Or checkfile2 <> "" Then
    Msg = "Can Not Overwrite An Exsisting File !"   ' Define message.
    Style = vbOKOnly + vbCritical + vbDefaultButton1 ' Define buttons.
    Title = "ERROR"  ' Define title.
    help = "DEMO.HLP"   ' Define Help file.
    Ctxt = 1000 ' Define topic
    response = MsgBox(Msg, Style, Title, help, Ctxt)
    Exit Sub
End If

Form1.Caption = "Hole Punch       DWG#  " & Text1.Text

'THICK = Val(Form1.Text3.Text)
'BWIDTH = Val(Form1.Text4.Text)
'Length1 = Val(Form1.Text5.Text)
'BSIZE = Form1.Combo2
'H1 = Val(Form1.Text1(0).Text)
'Let HoleTotal = H1
'If Val(Form1.Text1(1).Text) > 0.01 Then H2 = Val(Form1.Text1(1).Text) - HoleTotal Else H2 = 0
'Let HoleTotal = HoleTotal + H2
'If Val(Form1.Text1(2).Text) > 0.01 Then H3 = Val(Form1.Text1(2).Text) - HoleTotal Else H3 = 0
'Let HoleTotal = HoleTotal + H3
'If Val(Form1.Text1(3).Text) > 0.01 Then H4 = Val(Form1.Text1(3).Text) - HoleTotal Else H4 = 0
'Let HoleTotal = HoleTotal + H4
'If Val(Form1.Text1(4).Text) > 0.01 Then H5 = Val(Form1.Text1(4).Text) - HoleTotal Else H5 = 0
'Let HoleTotal = HoleTotal + H5
'If Val(Form1.Text1(5).Text) > 0.01 Then H6 = Val(Form1.Text1(5).Text) - HoleTotal Else H6 = 0
'Let HoleTotal = HoleTotal + H6
'If Val(Form1.Text1(6).Text) > 0.01 Then H7 = Val(Form1.Text1(6).Text) - HoleTotal Else H7 = 0
'Let HoleTotal = HoleTotal + H7
'If Val(Form1.Text1(7).Text) > 0.01 Then H8 = Val(Form1.Text1(7).Text) - HoleTotal Else H8 = 0
'Let HoleTotal = HoleTotal + H8
'If Val(Form1.Text1(8).Text) > 0.01 Then H9 = Val(Form1.Text1(8).Text) - HoleTotal Else H9 = 0
'Let HoleTotal = HoleTotal + H9
'If Val(Form1.Text1(9).Text) > 0.01 Then H10 = Val(Form1.Text1(9).Text) - HoleTotal Else H10 = 0
'Let HoleTotal = HoleTotal + H10
'If Val(Form1.Text1(10).Text) > 0.01 Then H11 = Val(Form1.Text1(10).Text) - HoleTotal Else H11 = 0
'Let HoleTotal = HoleTotal + H11
'If Val(Form1.Text1(11).Text) > 0.01 Then H12 = Val(Form1.Text1(11).Text) - HoleTotal Else H12 = 0
'Let HoleTotal = HoleTotal + H12
'If Val(Form1.Text1(12).Text) > 0.01 Then H13 = Val(Form1.Text1(12).Text) - HoleTotal Else H13 = 0
'Let HoleTotal = HoleTotal + H13
'If Val(Form1.Text1(13).Text) > 0.01 Then H14 = Val(Form1.Text1(13).Text) - HoleTotal Else H14 = 0
'Let HoleTotal = HoleTotal + H14
'If Val(Form1.Text1(14).Text) > 0.01 Then H15 = Val(Form1.Text1(14).Text) - HoleTotal Else H15 = 0
'Let HoleTotal = HoleTotal + H15
'If Val(Form1.Text1(15).Text) > 0.01 Then H16 = Val(Form1.Text1(15).Text) - HoleTotal Else H16 = 0
'Let HoleTotal = HoleTotal + H16
'If Val(Form1.Text1(16).Text) > 0.01 Then H17 = Val(Form1.Text1(16).Text) - HoleTotal Else H17 = 0
'Let HoleTotal = HoleTotal + H17
'If Val(Form1.Text1(17).Text) > 0.01 Then H18 = Val(Form1.Text1(17).Text) - HoleTotal Else H18 = 0
'Let HoleTotal = HoleTotal + H18
'If Val(Form1.Text1(18).Text) > 0.01 Then H19 = Val(Form1.Text1(18).Text) - HoleTotal Else H19 = 0
'Let HoleTotal = HoleTotal + H19
'If Val(Form1.Text1(19).Text) > 0.01 Then H20 = Val(Form1.Text1(19).Text) - HoleTotal Else H20 = 0
'Let HoleTotal = HoleTotal + H20
'If Val(Form1.Text1(20).Text) > 0.01 Then H21 = Val(Form1.Text1(20).Text) - HoleTotal Else H21 = 0
'Let HoleTotal = HoleTotal + H21
'If Val(Form1.Text1(21).Text) > 0.01 Then H22 = Val(Form1.Text1(21).Text) - HoleTotal Else H22 = 0
'Let HoleTotal = HoleTotal + H22
'If Val(Form1.Text1(22).Text) > 0.01 Then H23 = Val(Form1.Text1(22).Text) - HoleTotal Else H23 = 0
'Let HoleTotal = HoleTotal + H23
'If Val(Form1.Text1(23).Text) > 0.01 Then H24 = Val(Form1.Text1(23).Text) - HoleTotal Else H24 = 0
'Let HoleTotal = HoleTotal + H24
'If Val(Form1.Text1(24).Text) > 0.01 Then H25 = Val(Form1.Text1(24).Text) - HoleTotal Else H25 = 0

'If Val(Form1.Text11.Text) <> Val(Form1.Text4.Text) Then
'For i = 0 To 3
'         If Form1.Text2(i).Text <> "" Then
'            Let YA(i) = Form1.Text2(i).Text - ((Val(Form1.Text11.Text) - Val(Form1.Text4.Text)) / 2)
'            End If
'    Next i
'End If
'If Val(YA(0)) > 0.01 Then Y1A = YA(0) Else Y1A = 0
'If Val(YA(1)) > 0.01 Then Y2A = YA(1) Else Y2A = 0
'If Val(YA(2)) > 0.01 Then Y3A = YA(2) Else Y3A = 0
'If Val(YA(3)) > 0.01 Then Y4A = YA(3) Else Y4A = 0

 'Open File1.Path & "\" & Text1.Text & ".DTA" For Output As #1
'Write #1, "HolePunch", Date, Text1.Text, "", "", "", Length1, BWIDTH, THICK, 0, 0, BSIZE, 0, Y1A, Y2A, Y3A, Y4A, 0, H1, H2, H3, H4, H5, H6, H7, H8, H9, H10, H11, H12, H13, H14, H15, H16, H17, H18, H19, H20, H21, H22, H23, H24, H25, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
'Close #1
'Unload Save_File
End Sub

Private Sub Command2_Click()
Unload Save_File
End Sub


Private Sub File1_Click()
Dim Index As Integer
For Index = 0 To File1.ListCount - 1
If File1.Selected(Index) Then
Text1.Text = File1.List(Index)
End If
Next Index
End Sub


Private Sub Form_Activate()

File1.Path = "G:\ACAD\ABLASCII\"
Text1.SetFocus
Let Index = 0
Text1.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.Value = True
End If
End Sub




