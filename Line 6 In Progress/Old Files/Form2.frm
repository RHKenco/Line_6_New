VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C00000&
   Caption         =   "Speed Control Set-Up"
   ClientHeight    =   1590
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   8055
   LinkTopic       =   "Form2"
   ScaleHeight     =   1590
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   3
      Left            =   6000
      TabIndex        =   11
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   6000
      TabIndex        =   10
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   8
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   3360
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3360
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   250
      Index           =   2
      Left            =   960
      TabIndex        =   14
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Oss Speed"
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   13
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "X Axis Speed"
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   12
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PASS > 2.5"
      Height          =   255
      Index           =   3
      Left            =   6000
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PASS 2.0 - 2.5"
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PASS 1.25 - 2.0"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PASS <1.25"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Menu save 
      Caption         =   "&Save"
   End
   Begin VB.Menu exit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub exit_Click()

End Sub

Private Sub Form_Load()
Text1(0).Text = Format(Val(XSpeed1) * 37.8, "###0.00")
Text1(1).Text = Format(Val(XSpeed2) * 37.8, "###0.00")
Text1(2).Text = Format(Val(XSpeed3) * 37.8, "###0.00")
Text1(3).Text = Format(Val(XSpeed4) * 37.8, "###0.00")

Text2(0).Text = OssSpeed1
Text2(1).Text = OssSpeed2
Text2(2).Text = OssSpeed3
Text2(3).Text = OssSpeed4

End Sub

Private Sub save_Click()
Open "f:\apps\exe\line 6\files\SpeedSetup.asc" For Output As #1
Write #1, Format(Text1(0).Text / 37.8, "0.000"), Format(Text1(1).Text / 37.8, "0.000"), Format(Text1(2).Text / 37.8, "0.000"), Format(Text1(3).Text / 37.8, "0.000"), Text2(0).Text, Text2(1).Text, Text2(2).Text, Text2(3).Text
Close #1
XSpeed1 = Val(Text1(0).Text) / 37.8
XSpeed2 = Val(Text1(1).Text) / 37.8
XSpeed3 = Val(Text1(2).Text) / 37.8
XSpeed4 = Val(Text1(3).Text) / 37.8

 OssSpeed1 = Text2(0).Text
 OssSpeed2 = Text2(1).Text
 OssSpeed3 = Text2(2).Text
 OssSpeed4 = Text2(3).Text
End Sub

