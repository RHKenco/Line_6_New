VERSION 5.00
Begin VB.Form Open_File 
   BackColor       =   &H00C00000&
   Caption         =   "OPEN FILE"
   ClientHeight    =   3450
   ClientLeft      =   4680
   ClientTop       =   1515
   ClientWidth     =   4215
   LinkTopic       =   "Form12"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3450
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   2055
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
      Left            =   3240
      TabIndex        =   2
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
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
      Left            =   2160
      TabIndex        =   1
      Top             =   2880
      Width           =   855
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Open_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
checkfile = Dir("G:\ACAD\ABLASCII\" & Text1.Text & ".dta")
checkfile1 = Dir("G:\ACAD\ABLASCII\" & Text1.Text & ".DTA")
checkfile2 = Dir("G:\ACAD\ABLASCII\" & Text1.Text)
If checkfile <> "" Or checkfile1 <> "" Or checkfile2 <> "" Then
Else
    Msg = "Can Not Find An Exsisting Dwg !"   ' Define message.
    Style = vbOKOnly + vbCritical + vbDefaultButton1 ' Define buttons.
    Title = "ERROR"  ' Define title.
    help = "DEMO.HLP"   ' Define Help file.
    Ctxt = 1000 ' Define topic
    response = MsgBox(Msg, Style, Title, help, Ctxt)
   Exit Sub
End If

If Right(Text1.Text, 3) = "dta" Or Right(Text1.Text, 3) = "DTA" Then
    Open File1.Path & "\" & Text1.Text For Input As #1
Else
    Open File1.Path & "\" & Text1.Text & ".DTA" For Input As #1
End If

Input #1, init, DAT, DWG, PART, cust, disp, Length1, BWIDTH, THICK, Offset, offset2, BSIZE, offsetdim, Y1A, Y2A, Y3A, Y4A, eofblade, H1, H2, H3, H4, H5, H6, H7, H8, H9, H10, H11, H12, H13, H14, H15, H16, H17, H18, H19, H20, H21, H22, H23, H24, H25, boltqty, fstart, fyaxis, fstop, fpassw, bstart, byaxis, bstop, bpassw, ex1start, ex1yaxis, ex1stop, ex1passW, ex2start, ex2yaxis, ex2stop, ex2passW, pass5start, pass5yaxis, pass5stop, pass5passW, pass6start, pass6yaxis, pass6stop, pass6passW, fstart, fpassW1, leftpass, bwidthW1, lengthB1, fpassB1, Lengthr1, bwidthB2, toltc, FRONTBEVEL, BACKBEVEL, matl
Close #1


cutlist.Caption = "Impregnator       DWG#  " & DWG & "        Part#  " & PART
'cutlist.List1.AddItem (Chr(9) + "X_axis Start" + Chr(9) + "Y_axis" + Chr(9) + "X_axis Stop" + Chr(9) + "Pass Width"), 0
'cutlist.List1.AddItem ("Front Pass" + Chr(9) + Str$(fstart) + Chr(9) + Str$(fyaxis) + Chr(9) + Chr(9) + Str$(fstop) + Chr(9) + Chr(9) + Str$(fpassw)), 1
'cutlist.List1.AddItem ("Back Pass " + Chr(9) + Str$(bstart) + Chr(9) + Str$(byaxis) + Chr(9) + Chr(9) + Str$(bstop) + Chr(9) + Chr(9) + Str$(bpassw)), 2
'Let A = Len(STEEL_ASS(i))
'Let A = 14 - A
'Let B = Len(PART_NUMB(i))
'Let B = 16 - B
'Let C = Len(Length(i))
'Let C = 8 - C
'List1.AddItem (FILE_DATE(i) & "  " & WORK_ORDER(i) & "  " & PART_NUMB(i) & Space(B) & STEEL_ASS(i) & Space(A) & Length(i))
'Let i = i + 1
'Loop 'Until EOF(1)
'Close #1
'Form1.Text3.Text = THICK
'Form1.Text4.Text = BWIDTH
'Form1.Text5.Text = Length1
'Form1.Combo2 = BSIZE

'Form1.Text4.BackColor = QBColor(15)

Unload Open_File
'Form1.Text6.SetFocus
'Form1.Text6.BackColor = QBColor(14)

End Sub

Private Sub Command2_Click()
Unload Open_File
End Sub

Private Sub File1_Click()
Dim Index As Integer
For Index = 0 To File1.ListCount - 1
If File1.Selected(Index) Then
Text1.Text = File1.List(Index)
End If
Next Index

End Sub

Private Sub File1_DblClick()
Dim Index As Integer
For Index = 0 To File1.ListCount - 1
If File1.Selected(Index) Then
Text1.Text = File1.List(Index)
Command1.Value = True
Exit Sub
End If
Next Index

End Sub


Private Sub Form_Activate()
Open_File.Caption = "Open ABLADE File"
File1.Path = "G:\ACAD\ABLASCII\"


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.Value = True
End If

End Sub


