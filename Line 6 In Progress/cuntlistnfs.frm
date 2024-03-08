VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form cutlist 
   BackColor       =   &H00C00000&
   Caption         =   "Impregnator     REV.17"
   ClientHeight    =   9435
   ClientLeft      =   7710
   ClientTop       =   4200
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   12495
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   11760
      TabIndex        =   121
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Disk Rotation CW"
      Height          =   375
      Left            =   10680
      TabIndex        =   119
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text24 
      Height          =   375
      Left            =   10680
      TabIndex        =   116
      Text            =   "Text24"
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox Text22 
      Height          =   375
      Left            =   10680
      TabIndex        =   115
      Text            =   "Text22"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Disk Rotation CCW"
      Height          =   375
      Left            =   10680
      TabIndex        =   114
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text23 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   113
      Text            =   "Text23"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text21 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   10560
      TabIndex        =   112
      Text            =   "0"
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   8040
      TabIndex        =   110
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text19 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   106
      Text            =   "Text19"
      Top             =   8280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2400
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1455
      Left            =   2160
      TabIndex        =   104
      Top             =   6840
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   2566
      _Version        =   393216
      Rows            =   20
      Cols            =   8
      ScrollBars      =   2
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "START"
      Height          =   375
      Left            =   9000
      TabIndex        =   93
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "FINISH"
      Height          =   375
      Left            =   9000
      TabIndex        =   92
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "GANGED"
      Height          =   375
      Left            =   10200
      TabIndex        =   91
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "RESET"
      Height          =   375
      Left            =   10200
      TabIndex        =   90
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "N/F"
      Height          =   375
      Left            =   7800
      TabIndex        =   89
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   5280
      TabIndex        =   87
      Text            =   "0"
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9480
      TabIndex        =   85
      Text            =   "0"
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   6720
      TabIndex        =   80
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   3840
      TabIndex        =   79
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   2400
      TabIndex        =   78
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   960
      TabIndex        =   77
      Top             =   4800
      Width           =   975
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1920
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   11400
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10440
      Top             =   0
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   76
      Text            =   "0"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C00000&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   9
      Left            =   9480
      TabIndex        =   75
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C00000&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   8
      Left            =   9480
      TabIndex        =   74
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   3000
      TabIndex        =   73
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   3000
      TabIndex        =   72
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   4560
      TabIndex        =   71
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   4560
      TabIndex        =   70
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   5760
      TabIndex        =   69
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   5760
      TabIndex        =   68
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   8160
      TabIndex        =   67
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   6960
      TabIndex        =   66
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8160
      TabIndex        =   65
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   6960
      TabIndex        =   64
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C00000&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   7
      Left            =   9480
      TabIndex        =   63
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C00000&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   6
      Left            =   9480
      TabIndex        =   62
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   6960
      TabIndex        =   61
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   8160
      TabIndex        =   60
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6960
      TabIndex        =   59
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   8160
      TabIndex        =   58
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   5760
      TabIndex        =   57
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   5760
      TabIndex        =   56
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   4560
      TabIndex        =   55
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4560
      TabIndex        =   54
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   3000
      TabIndex        =   53
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   3000
      TabIndex        =   52
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C00000&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   5
      Left            =   9480
      TabIndex        =   47
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C00000&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   4
      Left            =   9480
      TabIndex        =   46
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C00000&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   3
      Left            =   9480
      TabIndex        =   45
      Top             =   1440
      Width           =   1700
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C00000&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   2
      Left            =   9480
      TabIndex        =   44
      Top             =   1080
      Width           =   1700
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C00000&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   9480
      TabIndex        =   43
      Top             =   720
      Width           =   1700
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   8160
      TabIndex        =   42
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   8160
      TabIndex        =   41
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   8160
      TabIndex        =   40
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   8160
      TabIndex        =   39
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   38
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8160
      TabIndex        =   37
      Text            =   "1.5"
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
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
      Index           =   0
      Left            =   8160
      TabIndex        =   36
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   6960
      TabIndex        =   35
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6960
      TabIndex        =   34
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6960
      TabIndex        =   33
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6960
      TabIndex        =   32
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6960
      TabIndex        =   31
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6960
      TabIndex        =   30
      Text            =   "1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
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
      Index           =   0
      Left            =   6960
      TabIndex        =   29
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5760
      TabIndex        =   28
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   5760
      TabIndex        =   27
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   5760
      TabIndex        =   26
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5760
      TabIndex        =   25
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5760
      TabIndex        =   24
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5760
      TabIndex        =   23
      Text            =   "100"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
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
      Index           =   0
      Left            =   5760
      TabIndex        =   22
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4560
      TabIndex        =   21
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4560
      TabIndex        =   20
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4560
      TabIndex        =   19
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4560
      TabIndex        =   18
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4560
      TabIndex        =   17
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   3000
      TabIndex        =   16
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   3000
      TabIndex        =   15
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3000
      TabIndex        =   14
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3000
      TabIndex        =   13
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   12
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   11
      Text            =   "0"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   10
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
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
      Index           =   0
      Left            =   4560
      TabIndex        =   9
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000004&
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   8
      Top             =   0
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C00000&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   9480
      TabIndex        =   7
      Top             =   360
      Value           =   1  'Checked
      Width           =   1700
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New Blade"
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
      Left            =   7800
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
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
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO!"
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
      TabIndex        =   0
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C00000&
      Caption         =   "Rotation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   10680
      TabIndex        =   120
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C00000&
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   10680
      TabIndex        =   118
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C00000&
      Caption         =   "Disk Dia."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   10680
      TabIndex        =   117
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C00000&
      Caption         =   "Velocity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   10680
      TabIndex        =   111
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C00000&
      Caption         =   "#5 Rotation SPEED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   8040
      TabIndex        =   109
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   3000
      TabIndex        =   108
      Top             =   3960
      Width           =   6495
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   1095
      Left            =   3120
      TabIndex        =   107
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "SHUT-OFF TIMER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   103
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   102
      Top             =   6360
      Width           =   9015
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   101
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C0C0&
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
      Left            =   4560
      TabIndex        =   100
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C0C0&
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
      Left            =   6240
      TabIndex        =   99
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Order #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   98
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   97
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Stop Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   96
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Order #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   95
      Top             =   5400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   8
      X1              =   0
      X2              =   11415
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C00000&
      Caption         =   "Width Offset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5280
      TabIndex        =   88
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C00000&
      Caption         =   "Velocity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   9480
      TabIndex        =   86
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C00000&
      Caption         =   "#4 Length"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6720
      TabIndex        =   84
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C00000&
      Caption         =   "Z Postion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3840
      TabIndex        =   83
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C00000&
      Caption         =   "Y Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2400
      TabIndex        =   82
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C00000&
      Caption         =   "X Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   960
      TabIndex        =   81
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
      Caption         =   "Due Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C00000&
      Caption         =   "Material"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "Dwg Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "Part Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Menu PopDelete 
      Caption         =   "PopDelete"
      Visible         =   0   'False
      Begin VB.Menu Delete 
         Caption         =   "Delete"
      End
      Begin VB.Menu ClearSettings 
         Caption         =   "Clear Settings"
      End
   End
   Begin VB.Menu FILE 
      Caption         =   "&FILE"
      Begin VB.Menu OPEN 
         Caption         =   "&OPEN"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu joystick 
      Caption         =   "&Joystick"
   End
   Begin VB.Menu maintenance 
      Caption         =   "&Maintenance"
   End
   Begin VB.Menu calibration 
      Caption         =   "&Calibration"
   End
   Begin VB.Menu set0 
      Caption         =   "&Set0,0,0,0"
   End
   Begin VB.Menu water 
      Caption         =   "&Water/Exhaust"
   End
   Begin VB.Menu speed 
      Caption         =   "S&peed Control"
   End
   Begin VB.Menu RESET 
      Caption         =   "RESET"
   End
   Begin VB.Menu LiftTable 
      Caption         =   "Lift Table"
   End
End
Attribute VB_Name = "cutlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub calibration_Click()
Cailb.Show
End Sub

Private Sub Check1_Click(Index As Integer)

'If Check1(Index).Caption = "OFF" Then
'    Check1(Index).Caption = "ON"
'    Check1(Index).Value = 1
'    Check1(Index).Refresh
'ElseIf Check1(Index).Caption = "ON" Then
'    Check1(Index).Caption = "OFF"
'    Check1(Index).Value = 0
'    Check1(Index).Refresh
'End If
    

End Sub



Private Sub Command1_Click()
Dim EStop
Dim temp() As Byte
Dim RunPass As Boolean
c6k.FSEnabled = True                'enable fast status
temp = c6k.FastStatus                  'get fast status information
Call CopyMemory(fsinfo, temp(0), 280)
XSpeed1 = ".125"
XSpeed2 = ".119"
XSpeed3 = ".100"
XSpeed4 = ".075"
OssSpeed1 = "3.5"
OssSpeed2 = "3.4"
OssSpeed3 = "3.0"
OssSpeed4 = "2.8"
Label7.Caption = ""
Label7.Refresh
RunPass = False

For i = 1 To 6
    If Check1(i - 1).value = 1 Then
        RunPass = True
    End If
Next i
If RunPass = False Then
    MsgBox "Run Pass is Not Checked (On/Off)"
    Exit Sub
End If


' -------------- comment this block out for testing ---------
If RunCondition5 <> "YES" Then
    If RunCondition5 = "TimeOut" Then
       BAR3.Show
       Exit Sub
    End If
    If RunCondition5 = "NO" Then
        MsgBox "WORK ORDER NOT FOUND!"
        Exit Sub
    End If
    MsgBox "RunMode is Not ON!"
    Exit Sub
End If
' ------------ stop comments -------------------------------


Call Error:
StartCount5 = 0
c6k.Write ("ERASE" & Chr$(13))
c6k.Write ("MC000000:MA110000:COMEXC0:1INFNC20-D:INENXXXX1X" & Chr$(13)) ' 20= E-Stop
c6k.Write ("!JOG000000:OUT.9-1:OUT.16-0:1OUT.13-1:T1:1OUT.15-1:" & Chr$(13)) '9= Lens Fan / 16= Welder Contact / 13= Exhust Fan / 15= Water Pump
c6k.Write ("COMEXS0:COMEXL0:SCALE1:LH0,0,0,0,0,0:SCLD26550,257143,257143,12500,62500,62500,62500:" & Chr$(13))
      'HOME OSS
c6k.Write ("1INFNC18-4T:" & Chr$(13)) '18 = Oss Home Limit
c6k.Write ("@A10:@AD10:@V4:D,,,-.3,,:GO000100:" & Chr$(13))
c6k.Write ("HOMA,,,1,,:HOMAD,,,50,,:@HOMZ0:HOMV,,,1,,:HOMVF,,,1,,:" & Chr$(13))
c6k.Write ("HOMBAC111011:HOMEDG111111:HOMDF000100:HOM,,,0,,:" & Chr$(13))
c6k.Write ("WAIT(6AS=XXX1X):T.1:D,,,-2.375,,:GO0001000:" & Chr$(13))
    
For i = 1 To 6
    If Check1(i - 1).value = 1 Then
        Yoffset1 = Format((Val(Text8(i).Text) + Val(Text9(i).Text * 0.5)), "####0.000") 'Y START POS
        Xoffset1 = Format(Val(Text7(i).Text - 2), "####0.000") 'X STOP -2"
        Yoffset2 = Format(Val(CalTig_Y) + (Val(Text8(i).Text) + Val(Text9(i).Text * 0.5)), "####0.000")
        Xoffset2 = Format((Val(CalTig_X) - 2), "####0.000")
        Exit For
    End If
Next i
   
    'JOYSTICK ON
Label7.Caption = "Set 0,0    Press JoyStick Release"
Label7.Refresh
tom = 0
JogInput1 = 0
JogInput2 = 0
Do
    'E-STOP
Last_Pcut_State = 0 'attempt to recognize if e-stop is engaged during startup
temp = c6k.FastStatus
Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
     If (Last_Pcut_State = 0) Then
            c6k.Write "JOG000000:1OUTALL9,16,0:1OUTALL25,32,0:T2:1OUT.9-1:1OUT.13-1" & Chr$(13)
            Last_Pcut_State = 1
            cutlist.Label7.Caption = "E-STOP!!!"
            cutlist.Label7.Refresh
            Command1.Enabled = True
            
            Exit Sub
     End If
Else
     If (Last_Pcut_State = 1) Then
        Last_Pcut_State = 0
     End If
End If



    Call CopyMemory(fsinfo, temp(0), 280)
    temp = c6k.FastStatus
' -------------------------------------------------------------------------- Joystick stuff -----------------------------------------
    'If JogOn = True Then
    ' axis select on, limit switch off
        If CStr((fsinfo.ProgIn(1) And Input22) / Input22) = 0 And ((fsinfo.ProgIn(1) And Input8) / Input8) = 0 Then
            If JogInput1 = 0 And ((fsinfo.ProgIn(1) And Input24) / Input24) > 0 Then
                c6k.Write ("!JOG000000:" & Chr$(13))
                c6k.Write ("JOG000000:1INFNC1-5J:1INFNC2-5K:1INFNC3-2K:1INFNC4-2J:JOGA2,5,5,1,5,5:JOGAD50,99,99,99,99,15:JOGVH5,12,10,2,5,3:JOGVL5,12,10,5,5,5:JOG010010" & Chr$(13))
                JogInput1 = 1
                JogInput2 = 0
                JogInput3 = 0
                JogInput4 = 0
                JogInput5 = 0
                Label7.Caption = "    Joystick 1"
                Label7.Refresh
            End If
        End If
        ' axis select off, limit switch off
        If CStr((fsinfo.ProgIn(1) And Input22) / Input22) And ((fsinfo.ProgIn(1) And Input24) / Input24) > 0 Then
            If JogInput2 = 0 Then
                c6k.Write ("!JOG0000000:" & Chr$(13))
                c6k.Write ("JOG0000000:1INFNC2-1J:1INFNC1-1K:1INFNC3-6K:1INFNC4-6J:JOGA2,5,5,10,5,5:JOGAD50,99,99,99,99,15:JOGVH5,8,10,2,5,3:JOGVL5,15,10,5,5,5:JOG1000010" & Chr$(13))
                JogInput2 = 1
                JogInput1 = 0
                JogInput3 = 0
                JogInput4 = 0
                JogInput5 = 0
                Label7.Caption = "    Joystick 2"
                Label7.Refresh
            End If
        End If
        ' axis select on, limit switch on
        If CStr((fsinfo.ProgIn(1) And Input22) / Input22) = 0 And ((fsinfo.ProgIn(1) And Input8) / Input8) > 0 Then
            If JogInput3 = 0 Then
                c6k.Write ("!JOG000000:!PSET,,0" & Chr$(13))
                c6k.Write ("JOG000000:1INFNC1-5J:1INFNC2-5K:1INFNC3-3K:1INFNC4-3J:JOGA2,5,5,10,5,5:JOGAD50,99,99,99,99,15:JOGVH6,8,10,2,5,3:JOGVL6,15,10,5,5,5:JOG001010" & Chr$(13))
                JogInput1 = 0
                JogInput2 = 0
                JogInput3 = 1
                JogInput4 = 0
                JogInput5 = 0
                Label7.Caption = "    Joystick 3"
                Label7.Refresh
                End If
            End If
        ' axis select on, limit switch on, joystick down, text16 motor position
         If CStr((fsinfo.ProgIn(1) And Input22) / Input22) = 0 And ((fsinfo.ProgIn(1) And Input8) / Input8) > 0 And ((fsinfo.ProgIn(1) And Input3) / Input3) > 0 And Val(Text16.Text) < 1 Then
            If JogInput4 = 0 Then
                c6k.Write ("!JOG000000:" & Chr$(13))
                c6k.Write ("JOG000000:1INFNC1-5K:1INFNC2-5J:1INFNC3-2K:1INFNC4-2J:JOGA2,5,5,10,5,5:JOGAD50,99,99,99,99,15:JOGVH6,12,10,2,5,3:JOGVL5,12,10,5,5,5:JOG010010" & Chr$(13))
                JogInput1 = 0
                JogInput2 = 0
                JogInput3 = 1
                JogInput4 = 1
                JogInput5 = 0
                Label7.Caption = "   Joystick 4"
                Label7.Refresh
            End If
          End If
          ' Toggle trigger for axis 3 select
          If CStr((fsinfo.ProgIn(1) And Input24) / Input24) = 0 Then
            If JogInput5 = 0 Then
                c6k.Write ("!JOG000000:" & Chr$(13))
                c6k.Write ("JOG000000:1INFNC1-5K:1INFNC2-5J:1INFNC3-3K:1INFNC4-3J:JOGA2,5,5,10,5,5:JOGAD50,99,99,99,99,15:JOGVH5,5,8,5,5,3:JOGVL5,5,8,5,5,5:JOG001010" & Chr$(13))
                JogInput1 = 0
                JogInput2 = 0
                JogInput3 = 0
                JogInput4 = 0
                JogInput5 = 1
                Label7.Caption = "   Jogging Lead Screw"
                Label7.Refresh
            End If
          End If
        If CStr((fsinfo.ProgIn(1) And Input23) / Input23) = 0 Then
            c6k.Write ("JOG000000:" & Chr$(13))
            JogInput2 = 0
            JogInput1 = 0
            JogInput3 = 0
            JogInput4 = 0
            JogInput5 = 0
            Label7.Caption = ""
            Label7.Refresh
            jogOn = False
            Exit Do
        End If
        
        
    If JogInput1 = 1 Or JogInput2 = 1 Or JogInput3 = 1 Or JogInput4 = 1 Then
        cutlist.Text1.BackColor = &HFFFF&
        cutlist.Text1.ForeColor = QBColor(1)
        cutlist.Text1.Text = " JOYSTICK ON"
        cutlist.Text1.Refresh
        
       Maintenance1.Text8.BackColor = &HFFFF&
       Maintenance1.Text8.ForeColor = QBColor(1)
       Maintenance1.Text8.Text = " JOYSTICK ON"
       Maintenance1.Text8.Refresh
    End If
Loop ' UNTIL

Do
    temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
    'E-Stop
    If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
        If (Last_Pcut_State = 0) Then
            c6k.Write "1OUTALL10,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1" & Chr$(13) '9= Lens Fan / 13= Exhust Fan
            c6k.Write "!COMEXS0:" & Chr$(13)
            Last_Pcut_State = 1
            Label7.Caption = "E-STOP!!!!!"
            Label7.Refresh
            Timer2.Enabled = True
            EStopPos = ""
            Exit Sub
        End If
    Else
        If (Last_Pcut_State = 1) Then
            Last_Pcut_State = 0
        End If
    End If
 
    temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)
     '23= Jog Release
    If CStr((fsinfo.ProgIn(1) And Input23) / Input23) > 0 Then
        c6k.Write "JOG0000:PSET0,0,0,0" & Chr$(13)
        Label7.Caption = "JOYSTICK DONE"
        Label7.Refresh
        Exit Do 'get out of loop
     End If
Loop ' UNTIL
'c6k.Write "OUT.7-0" & Chr$(13)
Timer1.Enabled = False
Timer2.Enabled = False
Let Text12.Text = 0
c6k.Write "COMEXS0:DRFLVL111111:" & Chr$(13)
c6k.Write "MC000000:MA110000:COMEXC0:" & Chr$(13)
c6k.Write "JOG000000:COMEXC1" & Chr$(13)
Text19.Text = Yoffset2
Text19.Refresh
c6k.Write "JOG000000" & Chr$(13)

'CHECK WATER FLOW
Label7.Caption = "Checking Water Pump"
Label7.Refresh
TESTWATERPUMP = 0
Do
    'E-Stop
    temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)
    If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
        If (Last_Pcut_State = 0) Then
            c6k.Write "1OUTALL10,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1" & Chr$(13) '9= Lens Fan / 13= Exhust Fan
            c6k.Write "!COMEXS0:" & Chr$(13)
            Last_Pcut_State = 1
            Label7.Caption = "E-STOP!!!!!"
            Label7.Refresh
            Timer2.Enabled = True
            EStopPos = ""
            Exit Sub
        End If
    Else
        If (Last_Pcut_State = 1) Then
            Last_Pcut_State = 0
            Label7.Caption = ""
            Label7.Refresh
        End If
    End If
    'Water Pump
    Call CopyMemory(fsinfo, temp(0), 280)
    If CStr((fsinfo.ProgIn(1) And Input19) / Input19) > 0 Then
        Label7.Caption = ""
        Label7.Refresh
        Exit Do
    End If
  
    If TESTWATERPUMP > 1000 Then
        Label7.Caption = " Water Pump HAS A PROBLEM!"
        Label7.Refresh
        Exit Sub
    End If
Loop

Label7.Caption = "Lower Head Until Contacting Blade"
Label7.Refresh

'Check "danger zone" before striking arc
If ((fsinfo.ProgIn(1) And Input8) / Input8) > 0 And (Text16.Text) < 1.5 Then
    Form4.Show 1
End If

'STRIKE ARC
If ((fsinfo.ProgIn(1) And Input8) / Input8) = 0 Then
    c6k.Write "PSET0,0:" & Chr$(13)
    c6k.Write "1OUT.12-1:1OUT.9-1:1OUT.30-1:A,10:AD,3:V,13:D,0.9:GO01:1OUT.16-1:WAIT(MOV=b000000):JOG100001" & Chr$(13)
    '9= Lens Fan / 12= Argon / 30= open test / 16= Welder Contact
    cutlist.Refresh
End If



A = 0
'MOVE ONTO TAB

If Val(Text9(i).Text) < 1.25 Then XSpeed = XSpeed1: OssSpeed = OssSpeed1
If Val(Text9(i).Text) > 1.249 Then XSpeed = XSpeed2: OssSpeed = OssSpeed2
If Val(Text9(i).Text) > 1.99 Then XSpeed = XSpeed3: OssSpeed = OssSpeed3
If Val(Text9(i).Text) > 2.49 Then XSpeed = XSpeed4: OssSpeed = OssSpeed4
c6k.Write "VAR10=" + Str(XSpeed) + ":VAR11=" + Str(OssSpeed) + ":" & Chr$(13)
c6k.Write "MC000000:MA110000:COMEXC1:" & Chr$(13)
If Check1(i - 1).value = 1 Then
    If Val(Text9(i).Text) < 1 Then
        passWidth = 0
    Else
        passWidth = ((Val(Text9(i).Text) - 0.75))
    End If
    c6k.Write "VAR1=1" & Chr$(13)
    
    'START OF CHECK PROGRAM Sub loop during Oss
    c6k.Write "DEL CHECK:DEF CHECK" & Chr$(13)
        'Exit Check Program with release button
        'c6k.Write "IF(1IN.23=b0):1OUT.31-1:VAR4=0:VAR5=0:NIF" & Chr$(13) '23
        c6k.Write "IF(1IN.23=b0):1OUT.31-1:NIF" & Chr$(13) '23
        'Normal XY joystick
        c6k.Write "IF(1IN.22=b1 AND VAR4=0):10UT.30-1:1OUT.32-0:VAR4=1:VAR5=0:JOG000X00:1INFNC1-1K:1INFNC2-1J:1INFNC3-6K:1INFNC4-6J:JOGA2,4,15,1,5,5:JOGAD50,99,99,99,99,10:JOGVH5,8,10,2,5,3:JOGVL5,15,40,5,5,3:JOG100X01:NIF" & Chr$(13)
        'Axis select for Z stage and rotate
        c6k.Write "IF(1IN.22=b0 AND VAR5=0):10UT.30-1:1OUT.32-0:VAR4=0:VAR5=1:JOG000X00:1INFNC1-5K:1INFNC2-5J:1INFNC3-2K:1INFNC4-2J:JOGA4,4,15,1,5,5:JOGAD50,99,99,99,99,10:JOGVH8,12,10,2,5,5:JOGVL8,12,40,5,5,5:JOG010X10:NIF" & Chr$(13)
        'Axis select AND limit switch engaged and rotate
        c6k.Write "IF(1IN.22=b0 AND 1IN.8=b1 AND VAR5=0):1OUT.30.1:1OUT.32-0:VAR4=0:VAR5=1:JOG000X00:1INFNC1-5K:1INFNC2-5J:1INFNC3-3K:1INFNC4-3J:JOGA4,4,15,1,5,5:JOGAD50,99,99,99,99,15:JOGVH8,8,10,2,5,5:JOGVL8,15,40,5,5,5:JOG001010:NIF" & Chr$(13)
    c6k.Write "END" & Chr$(13)
    'END OF CHECK PROGRAM
    
    'START OF CHECK1 PROGRAM Sub loop during Oss
    c6k.Write "DEL CHECK1:DEF CHECK1" & Chr$(13)
        
        'Control of X and Y - axis select off
        c6k.Write "IF(1IN.22=b1 AND VAR4=0):1OUT.30-1:1OUT.32-0:VAR4=1:VAR5=0:JOGX00X00:1INFNC2-1J:1INFNC1-1K:1INFNC3-6K:1INFNC4-6J:JOGA2,1,15,1,5,3:JOGAD50,99,99,99,99,3:JOGVH5,.5,10,2,5,5:JOGVL5,.8,40,5,5,5:JOG100X01:NIF" & Chr$(13) '22
        'Control of rotate and Z1 - axis select on
        c6k.Write "IF(1IN.22=b0 AND VAR5=0):1OUT.32-1:1OUT.30-0:VAR5=1:VAR4=0:JOGX00X00:1INFNC1-5K:1INFNC2-5J:1INFNC3-2K:1INFNC4-2J:JOGAX,4,15,10,5,5:JOGADX,99,99,99,99,3:JOGVHX,8,10,7,5,5:JOGVLX,8,12,1,5,5:JOGX10X10:NIF" & Chr$(13) '22
        'Press release button to exit check 1
        c6k.Write "IF(1IN.23=b0):1OUT.31-1:VAR4=0:VAR5=0:NIF" & Chr$(13) '23
        
        'PAUSE
        c6k.Write "IF(1IN.24=b0 AND VAR1=1):VAR2=1VEL:MC100000:V0:D0:GO1:NIF " & Chr$(13)
        c6k.Write "IF(1IN.24=b0 AND VAR1=1):MC000000:JOG100X01:1INFNC2-1J:1INFNC1-1K:1INFNC3-6K:1INFNC4-6J:JOGA2,1,15,1,5,2:JOGAD50,99,99,99,99,5:JOGVH5,.5,10,2,5,3:JOGVL5,.8,40,5,5,3:JOG100X01:VAR1=2:NIF:" & Chr$(13)
        c6k.Write "IF(1IN.24=b1 AND VAR1=2):VAR10=VAR2:JOG010X11:MC100000:D-1:V(VAR10):GO1:VAR1=1:NIF " & Chr$(13)

        'X AXIS SPEED CONTROL
        c6k.Write "IF(1IN.22=b1 AND 1IN.24=b1 AND VAR1=1 AND 1IN=b01):VAR10=VAR10-.001:V(VAR10):GO1:NIF " & Chr$(13)
        c6k.Write "IF(1IN.22=b1 AND 1IN.24=b1 AND VAR1=1 AND 1IN=b10):VAR10=VAR10+.001:V(VAR10):GO1:NIF " & Chr$(13)
    
    c6k.Write "END" & Chr$(13)
    'END OF CHECK1 PROGRAM
    
    'Jog on / Oss on / Move to Tab
    c6k.Write "MC110000:VAR4=0:VAR5=0:VAR1=0:A,,,5,,:V,,,(VAR11),,:D,,,-" + Str$(passWidth / 2) + ":GO000100:WAIT(MOV=bXXX0XX):REPEAT:D,,," + Str$(passWidth) + ":GO000100:REPEAT:GOSUB CHECK:UNTIL(MOV=bXXX0XX):D,,,-" + Str$(passWidth) + ":GO000100:REPEAT:GOSUB CHECK:UNTIL(MOV=bXXX0XX):UNTIL(1OUT.31=b1):WAIT(1IN.23=b1):1OUT.31-0" & Chr$(13)
         
    'START TC AND X AXIS Move
    c6k.Write "JOG0:" & Chr$(13)
    c6k.Write "JOG0:MC100000:VAR4=0:VAR5=0:VAR3=1:1OUT.14-1:A20:V(VAR10):D-1:GO1:WAIT(1OUT.14=b1):PSET0,0,0,X,0,0:VAR1=1:INEN.2-1" & Chr$(13)
    c6k.Write "VAR1=1:V,,,(VAR11),,:REPEAT:D,,," + Str$(passWidth) + ":GO000100:REPEAT:GOSUB CHECK1:UNTIL(MOV=bXXX0XX):D,,,-" + Str$(passWidth) + ":GO000100:REPEAT:GOSUB CHECK1:UNTIL(MOV=bXXX0XX):UNTIL(1OUT.31=b1):WAIT(1IN.23=b1):1OUT.31-0:" & Chr$(13)
     
    'STOP X AXIS & FILL END / Jog
    c6k.Write "MC100000:V0:GO1:1OUT.25-1" & Chr$(13)
    c6k.Write "VAR4=0:VAR5=0:VAR1=0:V,,,(VAR11),,:REPEAT:D,,," + Str$(passWidth) + ":GO000100:REPEAT:GOSUB CHECK:UNTIL(MOV=bXXX0XX):D,,,-" + Str$(passWidth) + ":GO000100:REPEAT:GOSUB CHECK:UNTIL(MOV=bXXX0XX):UNTIL(1OUT.31=b1):WAIT(1IN.23=b1):1OUT.31-0:INENXXXX0XX:" & Chr$(13)
       
    ' FAST STATUS UPDATE ESTOP/WATER/MOTOR POS
    TESTWATERPUMP = 0
    
    VoltCount = 0
    JogInput1 = 0
    Do
        temp = c6k.FastStatus
        Call CopyMemory(fsinfo, temp(0), 280)
        'E-Stop
        If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
            If (Last_Pcut_State = 0) Then
                c6k.Write "!1OUTALL10,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1:!S" & Chr$(13)
                c6k.Write "!COMEXS0:" & Chr$(13)
                Last_Pcut_State = 1
                Label7.Caption = "E-STOP!!!!!"
                Label7.Refresh
                Timer2.Enabled = True
                Exit Sub
                EStopPos = ""
            End If
        Else
            If (Last_Pcut_State = 1) Then
                Last_Pcut_State = 0
                Label7.Caption = ""
                Label7.Refresh
            End If
        End If
        TESTWATERPUMP = TESTWATERPUMP + 1
        Do
            temp = c6k.FastStatus
            Call CopyMemory(fsinfo, temp(0), 280)
            'Water Pump
            If CStr((fsinfo.ProgIn(1) And Input19) / Input19) > 0 Then
                Label7.Caption = ""
                Label7.Refresh
                Exit Do
            End If
  
            If TESTWATERPUMP > 1000 Then
                c6k.Write "!S:!1OUTALL10,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1" & Chr$(13)
                Label7.Caption = " Water Pump HAS A PROBLEM!"
                Label7.Refresh
                Exit Sub
            End If
        Loop
        '' ******* Update Motor Position ****************
        temp = c6k.FastStatus                  'get fast status information
        Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure

' ------------------------------------------- Motor Position --------------------------
Let X = InDex3
    'SCLD26550,257143,257143,12500,62500,62500
        XMOTOR = (fsinfo.MotorPos(1))
            Let Text14.Text = Val(XMOTOR)
            Let Text14.Text = Format((Text14.Text / 26550), "0.000")
        
        YMOTOR = (fsinfo.MotorPos(6))
            Let Text15.Text = Val(YMOTOR)
            Let Text15.Text = Format((Text15.Text / 62500), "0.000")
        
        ZMOTOR = (fsinfo.MotorPos(3))
            Let Text16.Text = Val(ZMOTOR)
            Let Text16.Text = Format((Text16.Text / 257153), "0.000")
        
        XSpeedSHOW = (fsinfo.MotorVel(1))
            Let Text18.Text = Val(XSpeedSHOW)
            Let Text18.Text = Format((Text18.Text / 660), "0.000")
        
        ROTATION = (fsinfo.MotorPos(5))
            Let Text20.Text = Val(ROTATION)
            Let Text20.Text = Format(((Text20.Text / 4432) / 1), "0.000")

        Call CopyMemory(fsinfo, temp(0), 280)
        If CStr((fsinfo.ProgIn(0) And Input5) / Input5) = 0 Then
            For P = 1 To 100000
            Next P
            Exit Do
        End If
        cutlist.Text20.Refresh
        cutlist.Text14.Refresh
        cutlist.Text15.Refresh
        cutlist.Text16.Refresh
        cutlist.Text17.Refresh
        cutlist.Text18.Refresh
        cutlist.Text12.Refresh
        For P = 1 To 100000
        Next P
        
    Loop
    c6k.Write "WAIT(MOV=bXXXXXX):MC000000:MA000000:1OUT.16-0:1OUT.12-0:1OUT.9-1:1OUT.14-0:OUT.5-1:1OUT.13-1:" & Chr$(13)
    tom = Val(Text14.Text * -1)
    c6k.Write "D,2:V,5:GO01:WAIT(MOV=b000000):A5:V5:D" + Str(Val(Text14.Text * -1)) + ":GO100000:WAIT(MOV=b000000)" & Chr$(13)
   
End If

Do
    temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)
    Let Text14.Text = (fsinfo.MotorPos(1) / 26550)
    Text14.Refresh
    tom = (Text14.Text * -1)
    If ((Text14.Text * -1) < 0.06) Then
Joystick_Click
    Timer2.Enabled = True
    Exit Do
    End If
Loop

    
End Sub




Private Sub Command13_Click()
If RunCondition5 = "TimeOut" Then
       BAR3.Show
        Exit Sub
End If
RunCondition5 = ""
StartCount5 = 0
 If Dir("F:\BARCODE\temp6.tmp") = "" Then
           MsgBox ("No Active Work Order "), , ("ERROR")
           Exit Sub
        Else
            
             Open "F:\BARCODE\temp6.tmp" For Input As #2
            Input #2, temp6, DAT
            Close #2
            
            Open "F:\BARCODE\" + temp6 + "6.tmp" For Input As #2
            Input #2, InString1, STIME!, ETIME!, TTime!, Com2$
            Close #2
            
            Let ETIME! = Timer: Let TTime! = ((ETIME! - STIME!) / 60) + TTime!
            Open "F:\BARCODE\" + temp6 + "6.tmp" For Output As #2
            Write #2, InString1, STIME!, ETIME!, TTime!, Com2$
            Close #2
            
            Let STIME! = STIME! / 3600: Let ETIME! = ETIME! / 3600:
            Label21.Caption = "RunMode - OFF"
            Label17.Caption = Format(Now, "SHORT TIME")
            
            
            Kill "F:\BARCODE\temp6.tmp"
If Com2$ = "TOPS" Or InString2 = "BEVEL" Then
Sta$ = "Line 3&5"
ElseIf Com2$ = "" Then
Sta$ = "6"
Else
Sta$ = "6"
Typ$ = "ganged"
End If

Open "F:\MFG\Bar_Data.dta" For Append As #1
Write #1, InString1, Sta$, "N\F", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss")
    If Typ$ = "ganged" Then
    Write #1, Com2$, Sta$, "N\F", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss")
    End If
Close #1
Let InString1 = ""
            
             Timer2.Enabled = False
             'Text4.Visible = False
             'Label19.Visible = False
            Label21.Caption = "Press Start  or  Ganged  "
        End If
Command6.Enabled = True
Command8.Enabled = True
'Command10.Enabled = True
End Sub

Private Sub Command15_Click()
'Disk Rotation CW
Dim EStop
Dim temp() As Byte
Dim RunPass As Boolean
c6k.FSEnabled = True                'enable fast status
temp = c6k.FastStatus                  'get fast status information
Call CopyMemory(fsinfo, temp(0), 280)
'CMDDIR changes the commanded polarity of the motor but doesn't lock it into one direction
c6k.Write ("DRIVE0,X,X,0,0,0,0:CMDDIR0000001:DRIVE1,X,X,1,1,1,1,0" & Chr$(13))
RotSpeed = ".05"
XSpeed1 = ".125"
XSpeed2 = ".119"
XSpeed3 = ".100"
XSpeed4 = ".075"
OssSpeed1 = "3.5"
OssSpeed2 = "3.4"
OssSpeed3 = "3.0"
OssSpeed4 = "2.8"
Label7.Caption = ""
Label7.Refresh
RunPass = False

For i = 1 To 6
    If Check1(i - 1).value = 1 Then
        RunPass = True
    End If
Next i
If RunPass = False Then
    MsgBox "Run Pass is Not Checked (On/Off)"
    Exit Sub
End If


' -------------- comment this block out for testing ---------
If RunCondition5 <> "YES" Then
    If RunCondition5 = "TimeOut" Then
       BAR3.Show
       Exit Sub
    End If
    If RunCondition5 = "NO" Then
        MsgBox "WORK ORDER NOT FOUND!"
        Exit Sub
    End If
    MsgBox "RunMode is Not ON!"
    Exit Sub
End If
' ------------ stop comments -------------------------------


Call Error:
StartCount5 = 0
c6k.Write ("ERASE" & Chr$(13))
c6k.Write ("MC0000000:MA1100000:COMEXC0:1INFNC20-D:INENXXXX1XX" & Chr$(13)) ' 20= E-Stop
c6k.Write ("!JOG0000000:OUT.9-1:OUT.16-0:1OUT.13-1:T1:1OUT.15-1:" & Chr$(13)) '9= Lens Fan / 16= Welder Contact / 13= Exhust Fan / 15= Water Pump
c6k.Write ("COMEXS0:COMEXL0:SCALE1:LH0,0,0,0,0,0,0:SCLD26550,257143,257143,12500,62500,62500,25000:" & Chr$(13))
      'HOME OSS
c6k.Write ("1INFNC18-4T:" & Chr$(13)) '18 = Oss Home Limit
c6k.Write ("@A10:@AD10:@V4:D,,,-.3,,:GO0001000:" & Chr$(13))
c6k.Write ("HOMA,,,1,,,:HOMAD,,,50,,,:@HOMZ0:HOMV,,,1,,,:HOMVF,,,1,,,:" & Chr$(13))
c6k.Write ("HOMBAC1110111:HOMEDG1111111:HOMDF0001000:HOM,,,0,,,:" & Chr$(13))
c6k.Write ("WAIT(6AS=XXX1XXX):T.1:D,,,-2.375,,,:GO0001000:" & Chr$(13))
    
For i = 1 To 6
    If Check1(i - 1).value = 1 Then
        Yoffset1 = Format((Val(Text8(i).Text) + Val(Text9(i).Text * 0.5)), "####0.000") 'Y START POS
        Xoffset1 = Format(Val(Text7(i).Text - 2), "####0.000") 'X STOP -2"
        Yoffset2 = Format(Val(CalTig_Y) + (Val(Text8(i).Text) + Val(Text9(i).Text * 0.5)), "####0.000")
        Xoffset2 = Format((Val(CalTig_X) - 2), "####0.000")
        Exit For
    End If
Next i
   
    'JOYSTICK ON
Label7.Caption = "Set 0,0    Press JoyStick Release"
Label7.Refresh
tom = 0
JogInput1 = 0
JogInput2 = 0
Do
    'E-STOP
Last_Pcut_State = 0 'attempt to recognize if e-stop is engaged during startup
temp = c6k.FastStatus
Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
     If (Last_Pcut_State = 0) Then
            c6k.Write "JOG0000000:1OUTALL9,16,0:1OUTALL25,32,0:T2:1OUT.9-1:1OUT.13-1" & Chr$(13)
            Last_Pcut_State = 1
            cutlist.Label7.Caption = "E-STOP!!!"
            cutlist.Label7.Refresh
            Command1.Enabled = True
            
            Exit Sub
     End If
Else
     If (Last_Pcut_State = 1) Then
        Last_Pcut_State = 0
     End If
End If



    Call CopyMemory(fsinfo, temp(0), 280)
    temp = c6k.FastStatus
' -------------------------------------------------------------------------- Joystick stuff -----------------------------------------
    'If JogOn = True Then
    ' axis select on, limit switch off
        If CStr((fsinfo.ProgIn(1) And Input22) / Input22) = 0 And ((fsinfo.ProgIn(1) And Input8) / Input8) = 0 Then
            If JogInput1 = 0 And ((fsinfo.ProgIn(1) And Input24) / Input24) > 0 Then
                c6k.Write ("!JOG0000000:" & Chr$(13))
                c6k.Write ("JOG0000000:1INFNC1-7J:1INFNC2-7K:1INFNC3-2K:1INFNC4-2J:JOGA4,5,5,1,5,5,5:JOGAD50,99,99,99,99,15,99:JOGVH8,8,10,2,5,3,5:JOGVL8,15,10,5,5,5,5:JOG0100001" & Chr$(13))
                JogInput1 = 1
                JogInput2 = 0
                JogInput3 = 0
                JogInput4 = 0
                JogInput5 = 0
                Label7.Caption = "    Joystick 1"
                Label7.Refresh
            End If
        End If
        ' axis select off, limit switch off
        If CStr((fsinfo.ProgIn(1) And Input22) / Input22) And ((fsinfo.ProgIn(1) And Input24) / Input24) > 0 Then
            If JogInput2 = 0 Then
                c6k.Write ("!JOG0000000:" & Chr$(13))
                c6k.Write ("JOG0000000:1INFNC2-1J:1INFNC1-1K:1INFNC3-6K:1INFNC4-6J:JOGA2,5,5,10,5,5,5:JOGAD50,99,99,99,99,15,99:JOGVH5,8,10,2,5,3,5:JOGVL5,15,10,5,5,5,5:JOG1000010" & Chr$(13))
                JogInput2 = 1
                JogInput1 = 0
                JogInput3 = 0
                JogInput4 = 0
                JogInput5 = 0
                Label7.Caption = "    Joystick 2"
                Label7.Refresh
            End If
        End If
        ' axis select on, limit switch on
        If CStr((fsinfo.ProgIn(1) And Input22) / Input22) = 0 And ((fsinfo.ProgIn(1) And Input8) / Input8) > 0 Then
            If JogInput3 = 0 Then
                c6k.Write ("!JOG0000000:!PSET,,0" & Chr$(13))
                c6k.Write ("JOG0000000:1INFNC1-7J:1INFNC2-7K:1INFNC3-3K:1INFNC4-3J:JOGA3,5,5,10,5,5,5:JOGAD50,99,99,99,99,15,99:JOGVH8,8,10,2,5,3,5:JOGVL8,15,10,5,5,5,5:JOG0010001" & Chr$(13))
                JogInput1 = 0
                JogInput2 = 0
                JogInput3 = 1
                JogInput4 = 0
                JogInput5 = 0
                Label7.Caption = "    Joystick 3"
                Label7.Refresh
                End If
            End If
        ' axis select on, limit switch on, joystick down, text16 motor position
         If CStr((fsinfo.ProgIn(1) And Input22) / Input22) = 0 And ((fsinfo.ProgIn(1) And Input8) / Input8) > 0 And ((fsinfo.ProgIn(1) And Input3) / Input3) > 0 And Val(Text16.Text) < 1 Then
            If JogInput4 = 0 Then
                c6k.Write ("!JOG0000000:" & Chr$(13))
                c6k.Write ("JOG0000000:1INFNC1-7K:1INFNC2-7J:1INFNC3-2K:1INFNC4-2J:JOGA3,5,5,10,5,5,5:JOGAD50,99,99,99,99,15,99:JOGVH8,8,10,2,5,3,5:JOGVL8,15,10,5,5,5,5:JOG0100001" & Chr$(13))
                JogInput1 = 0
                JogInput2 = 0
                JogInput3 = 1
                JogInput4 = 1
                JogInput5 = 0
                Label7.Caption = "   Joystick 4"
                Label7.Refresh
            End If
          End If
          ' Toggle trigger for axis 3 select
          If CStr((fsinfo.ProgIn(1) And Input24) / Input24) = 0 Then
            If JogInput5 = 0 Then
                c6k.Write ("!JOG0000000:" & Chr$(13))
                c6k.Write ("JOG0000000:1INFNC1-7K:1INFNC2-7J:1INFNC3-3K:1INFNC4-3J:JOGA3,5,5,10,5,5,5:JOGAD50,99,99,99,99,15,99:JOGVH8,5,8,5,5,3,5:JOGVL8,5,8,5,5,5,5:JOG0010001" & Chr$(13))
                JogInput1 = 0
                JogInput2 = 0
                JogInput3 = 0
                JogInput4 = 0
                JogInput5 = 1
                Label7.Caption = "   Jogging Lead Screw"
                Label7.Refresh
            End If
          End If
        If CStr((fsinfo.ProgIn(1) And Input23) / Input23) = 0 Then
            c6k.Write ("JOG0000000:" & Chr$(13))
            JogInput2 = 0
            JogInput1 = 0
            JogInput3 = 0
            JogInput4 = 0
            JogInput5 = 0
            Label7.Caption = ""
            Label7.Refresh
            jogOn = False
            Exit Do
        End If
        
        
    If JogInput1 = 1 Or JogInput2 = 1 Or JogInput3 = 1 Or JogInput4 = 1 Then
        cutlist.Text1.BackColor = &HFFFF&
        cutlist.Text1.ForeColor = QBColor(1)
        cutlist.Text1.Text = " JOYSTICK ON"
        cutlist.Text1.Refresh
        
       Maintenance1.Text8.BackColor = &HFFFF&
       Maintenance1.Text8.ForeColor = QBColor(1)
       Maintenance1.Text8.Text = " JOYSTICK ON"
       Maintenance1.Text8.Refresh
    End If
Loop ' UNTIL

Do
    temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
    'E-Stop
    If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
        If (Last_Pcut_State = 0) Then
            c6k.Write "1OUTALL10,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1" & Chr$(13) '9= Lens Fan / 13= Exhust Fan
            c6k.Write "!COMEXS0:" & Chr$(13)
            Last_Pcut_State = 1
            Label7.Caption = "E-STOP!!!!!"
            Label7.Refresh
            Timer2.Enabled = True
            EStopPos = ""
            Exit Sub
        End If
    Else
        If (Last_Pcut_State = 1) Then
            Last_Pcut_State = 0
        End If
    End If
 
    temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)
     '23= Jog Release
    If CStr((fsinfo.ProgIn(1) And Input23) / Input23) > 0 Then
        c6k.Write "JOG0000:PSET0,0,0,0" & Chr$(13)
        Label7.Caption = "JOYSTICK DONE"
        Label7.Refresh
        Exit Do 'get out of loop
     End If
Loop ' UNTIL
'c6k.Write "OUT.7-0" & Chr$(13)
Timer1.Enabled = False
Timer2.Enabled = False
Let Text12.Text = 0
c6k.Write "COMEXS0:DRFLVL1111111:" & Chr$(13)
c6k.Write "MC0000000:MA1100000:COMEXC0:" & Chr$(13)
c6k.Write "JOG0000000:COMEXC1" & Chr$(13)
Text19.Text = Yoffset2
Text19.Refresh
c6k.Write "JOG0000000" & Chr$(13)

'CHECK WATER FLOW
Label7.Caption = "Checking Water Pump"
Label7.Refresh
TESTWATERPUMP = 0
Do
    'E-Stop
    temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)
    If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
        If (Last_Pcut_State = 0) Then
            c6k.Write "1OUTALL10,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1" & Chr$(13) '9= Lens Fan / 13= Exhust Fan
            c6k.Write "!COMEXS0:" & Chr$(13)
            Last_Pcut_State = 1
            Label7.Caption = "E-STOP!!!!!"
            Label7.Refresh
            Timer2.Enabled = True
            EStopPos = ""
            Exit Sub
        End If
    Else
        If (Last_Pcut_State = 1) Then
            Last_Pcut_State = 0
            Label7.Caption = ""
            Label7.Refresh
        End If
    End If
    'Water Pump
    Call CopyMemory(fsinfo, temp(0), 280)
    If CStr((fsinfo.ProgIn(1) And Input19) / Input19) > 0 Then
        Label7.Caption = ""
        Label7.Refresh
        Exit Do
    End If
  
    If TESTWATERPUMP > 1000 Then
        Label7.Caption = " Water Pump HAS A PROBLEM!"
        Label7.Refresh
        Exit Sub
    End If
Loop

Label7.Caption = "Lower Head Until Contacting Blade"
Label7.Refresh

'Check "danger zone" before striking arc
If ((fsinfo.ProgIn(1) And Input8) / Input8) > 0 And (Text16.Text) < 1.5 Then
    Form4.Show 1
End If

'STRIKE ARC
If ((fsinfo.ProgIn(1) And Input8) / Input8) = 0 Then
    c6k.Write "PSET0,0:" & Chr$(13)
    c6k.Write "1OUT.12-1:1OUT.9-1:1OUT.30-1:A,10:AD,3:V,13:D,0.9:GO01:1OUT.16-1:WAIT(MOV=b0000000):JOG1000010" & Chr$(13)
    '9= Lens Fan / 12= Argon / 30= open test / 16= Welder Contact
    cutlist.Refresh
End If



A = 0
'MOVE ONTO TAB

If Val(Text9(i).Text) < 1.25 Then XSpeed = XSpeed1: OssSpeed = OssSpeed1
If Val(Text9(i).Text) > 1.249 Then XSpeed = XSpeed2: OssSpeed = OssSpeed2
If Val(Text9(i).Text) > 1.99 Then XSpeed = XSpeed3: OssSpeed = OssSpeed3
If Val(Text9(i).Text) > 2.49 Then XSpeed = XSpeed4: OssSpeed = OssSpeed4
c6k.Write "VAR10=" + Str(RotSpeed) + ":VAR11=" + Str(OssSpeed) + ":" & Chr$(13)
c6k.Write "MC0000000:MA1100000:COMEXC1:" & Chr$(13)
If Check1(i - 1).value = 1 Then
    If Val(Text9(i).Text) < 1 Then
        passWidth = 0
    Else
        passWidth = ((Val(Text9(i).Text) - 0.75))
    End If
    c6k.Write "VAR1=1" & Chr$(13)
    
    'START OF CHECK PROGRAM Sub loop during Oss
    c6k.Write "DEL CHECK:DEF CHECK" & Chr$(13)
        'Exit Check Program with release button
        'c6k.Write "IF(1IN.23=b0):1OUT.31-1:VAR4=0:VAR5=0:NIF" & Chr$(13) '23
        c6k.Write "IF(1IN.23=b0):1OUT.31-1:NIF" & Chr$(13) '23
        'Normal Y and X joystick
        c6k.Write "IF(1IN.22=b1 AND VAR4=0):10UT.30-1:1OUT.32-0:VAR4=1:VAR5=0:JOG000X00X:1INFNC1-1K:1INFNC2-1J:1INFNC3-6K:1INFNC4-6J:JOGA2,4,15,1,5,5,5:JOGAD50,99,99,99,99,10,99:JOGVH5,8,10,2,5,3,5:JOGVL5,15,40,5,5,3,5:JOG100X01X:NIF" & Chr$(13)
        'Axis select for Z stage and Rotate
        c6k.Write "IF(1IN.22=b0 AND VAR5=0):10UT.30-1:1OUT.32-0:VAR4=0:VAR5=1:JOG000X00X:1INFNC1-7K:1INFNC2-7J:1INFNC3-2K:1INFNC4-2J:JOGA4,4,15,1,5,5,5:JOGAD50,99,99,99,99,10,99:JOGVH8,8,10,2,5,5,5:JOGVL8,15,40,5,5,5,5:JOG010X001:NIF" & Chr$(13)
        'Axis select AND limit switch engaged and rotate
        c6k.Write "IF(1IN.22=b0 AND 1IN.8=b1 AND VAR5=0):1OUT.30.1:1OUT.32-0:VAR4=0:VAR5=1:JOG000X00X:1INFNC1-5K:1INFNC2-5J:1INFNC3-3K:1INFNC4-3J:JOGA4,4,15,1,5,5,5:JOGAD50,99,99,99,99,15,99:JOGVH8,8,10,2,5,5,5:JOGVL8,15,40,5,5,5,5:JOG001010X:NIF" & Chr$(13)
    c6k.Write "END" & Chr$(13)
    'END OF CHECK PROGRAM
    
    'START OF CHECK1 PROGRAM Sub loop during Oss
    c6k.Write "DEL CHECK1:DEF CHECK1" & Chr$(13)
        
        'Control of X and Y - axis select off
        c6k.Write "IF(1IN.22=b1 AND VAR4=0):1OUT.30-1:1OUT.32-0:VAR4=1:VAR5=0:JOG000X00X:1INFNC2-1J:1INFNC1-1K:1INFNC3-6K:1INFNC4-6J:JOGA2,1,15,1,5,3,5:JOGAD50,99,99,99,99,3,99:JOGVH5,.5,10,2,5,5,5:JOGVL5,.8,40,5,5,5,5:JOG100X01X:NIF" & Chr$(13) '22
        'Control of rotate and Z1 - axis select on
        c6k.Write "IF(1IN.22=b0 AND VAR5=0):1OUT.32-1:1OUT.30-0:VAR5=1:VAR4=0:JOG000X00X:1INFNC1-7K:1INFNC2-7J:1INFNC3-2K:1INFNC4-2J:JOGAX,1,15,10,5,5,5:JOGADX,99,99,99,99,3,99:JOGVHX,.2,10,7,5,5,5:JOGVLX,15,12,1,5,5,5:JOGX10X101:NIF" & Chr$(13) '22
        'Press release button to exit check 1
        c6k.Write "IF(1IN.23=b0):1OUT.31-1:VAR4=0:VAR5=0:NIF" & Chr$(13) '23
        
        'PAUSE
        c6k.Write "IF(1IN.24=b0 AND VAR1=1):VAR2=7VEL:MC1000000:V,,,,,,0:D,,,,,,0:GO,,,,,,1:NIF " & Chr$(13)
        c6k.Write "IF(1IN.24=b0 AND VAR1=1):MC0000000:JOG100X010:1INFNC2-1J:1INFNC1-1K:1INFNC3-6K:1INFNC4-6J:JOGA2,1,15,1,5,2:JOGAD50,99,99,99,99,5:JOGVH5,.5,10,2,5,3:JOGVL5,.8,40,5,5,3:JOG100X010:VAR1=2:NIF:" & Chr$(13)
        c6k.Write "IF(1IN.24=b1 AND VAR1=2):VAR10=VAR2:JOG010X11X:MC0000001:D,,,,,,-1:V,,,,,,(VAR10):GO,,,,,,1:VAR1=1:NIF " & Chr$(13)

        'X AXIS SPEED CONTROL  --- will be rotation speed control?
        c6k.Write "IF(1IN.22=b0 AND 1IN.24=b1 AND VAR1=1 AND 1IN=b01):VAR10=VAR10-.001:V,,,,,,(VAR10):GO,,,,,,1:NIF " & Chr$(13)
        c6k.Write "IF(1IN.22=b0 AND 1IN.24=b1 AND VAR1=1 AND 1IN=b10):VAR10=VAR10+.001:V,,,,,,(VAR10):GO,,,,,,1:NIF " & Chr$(13)
    
    c6k.Write "END" & Chr$(13)
    'END OF CHECK1 PROGRAM
    
    'Jog on / Oss on / Move to Tab
    c6k.Write "MC0100001:VAR4=0:VAR5=0:VAR1=0:A,,,5,,:V,,,(VAR11),,:D,,,-" + Str$(passWidth / 2) + ":GO0001000:WAIT(MOV=bXXX0XXX):REPEAT:D,,," + Str$(passWidth) + ":GO0001000:REPEAT:GOSUB CHECK:UNTIL(MOV=bXXX0XXX):D,,,-" + Str$(passWidth) + ":GO0001000:REPEAT:GOSUB CHECK:UNTIL(MOV=bXXX0XXX):UNTIL(1OUT.31=b1):WAIT(1IN.23=b1):1OUT.31-0" & Chr$(13)
         
    'START TC AND X AXIS Move
    c6k.Write "JOG0:" & Chr$(13)
    c6k.Write "JOG0:MC0000001:VAR4=0:VAR5=0:VAR3=1:1OUT.14-1:A,,,,,,10:V,,,,,,(VAR10):D,,,,,,-1:GO,,,,,,1:WAIT(1OUT.14=b1):PSET0,0,0,,0,0,0:VAR1=1:INEN.2-1" & Chr$(13)
    c6k.Write "VAR1=1:V,,,(VAR11),,:REPEAT:D,,," + Str$(passWidth) + ":GO0001000:REPEAT:GOSUB CHECK1:UNTIL(MOV=bXXX0XXX):D,,,-" + Str$(passWidth) + ":GO0001000:REPEAT:GOSUB CHECK1:UNTIL(MOV=bXXX0XXX):UNTIL(1OUT.31=b1):WAIT(1IN.23=b1):1OUT.31-0:" & Chr$(13)
     
    'STOP X AXIS & FILL END / Jog
    c6k.Write "MC0000001:V,,,,,,0:GO,,,,,,1:1OUT.25-1" & Chr$(13)
    c6k.Write "VAR4=0:VAR5=0:VAR1=0:V,,,(VAR11),,:REPEAT:D,,," + Str$(passWidth) + ":GO000100X:REPEAT:GOSUB CHECK:UNTIL(MOV=bXXX0XXX):D,,,-" + Str$(passWidth) + ":GO000100X:REPEAT:GOSUB CHECK:UNTIL(MOV=bXXX0XXX):UNTIL(1OUT.31=b1):WAIT(1IN.23=b1):1OUT.31-0:INENXXXX0XX:" & Chr$(13)
       
    ' FAST STATUS UPDATE ESTOP/WATER/MOTOR POS
    TESTWATERPUMP = 0
    
    VoltCount = 0
    JogInput1 = 0
    Do
        temp = c6k.FastStatus
        Call CopyMemory(fsinfo, temp(0), 280)
        'E-Stop
        If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
            If (Last_Pcut_State = 0) Then
                c6k.Write "!1OUTALL10,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1:!S" & Chr$(13)
                c6k.Write "!COMEXS0:" & Chr$(13)
                Last_Pcut_State = 1
                Label7.Caption = "E-STOP!!!!!"
                Label7.Refresh
                Timer2.Enabled = True
                Exit Sub
                EStopPos = ""
            End If
        Else
            If (Last_Pcut_State = 1) Then
                Last_Pcut_State = 0
                Label7.Caption = ""
                Label7.Refresh
            End If
        End If
        TESTWATERPUMP = TESTWATERPUMP + 1
        Do
            temp = c6k.FastStatus
            Call CopyMemory(fsinfo, temp(0), 280)
            'Water Pump
            If CStr((fsinfo.ProgIn(1) And Input19) / Input19) > 0 Then
                Label7.Caption = ""
                Label7.Refresh
                Exit Do
            End If
  
            If TESTWATERPUMP > 1000 Then
                c6k.Write "!S:!1OUTALL10,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1" & Chr$(13)
                Label7.Caption = " Water Pump HAS A PROBLEM!"
                Label7.Refresh
                Exit Sub
            End If
        Loop
        '' ******* Update Motor Position ****************
        temp = c6k.FastStatus                  'get fast status information
        Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure

' ------------------------------------------- Motor Position --------------------------
Let X = InDex3
    'SCLD26550,257143,257143,12500,62500,62500
        XMOTOR = (fsinfo.MotorPos(1))
            Let Text14.Text = Val(XMOTOR)
            Let Text14.Text = Format((Text14.Text / 26550), "0.000")
        
        YMOTOR = (fsinfo.MotorPos(6))
            Let Text15.Text = Val(YMOTOR)
            Let Text15.Text = Format((Text15.Text / 62500), "0.000")
        
        ZMOTOR = (fsinfo.MotorPos(3))
            Let Text16.Text = Val(ZMOTOR)
            Let Text16.Text = Format((Text16.Text / 257153), "0.000")
        
        XSpeedSHOW = (fsinfo.MotorVel(1))
            Let Text18.Text = Val(XSpeedSHOW)
            Let Text18.Text = Format((Text18.Text / 660), "0.000")
        
        ROTATION = (fsinfo.MotorVel(7))
            Let Text20.Text = Val(ROTATION)
            Let Text20.Text = Format(((Text20.Text / 4432) / 1), "0.000")

        Call CopyMemory(fsinfo, temp(0), 280)
        If CStr((fsinfo.ProgIn(0) And Input5) / Input5) = 0 Then
            For P = 1 To 100000
            Next P
            Exit Do
        End If
        cutlist.Text20.Refresh
        cutlist.Text14.Refresh
        cutlist.Text15.Refresh
        cutlist.Text16.Refresh
        cutlist.Text17.Refresh
        cutlist.Text18.Refresh
        cutlist.Text12.Refresh
        For P = 1 To 100000
        Next P
        
    Loop
    
    c6k.Write "WAIT(MOV=bXX0XXX):MC000000:MA000000:1OUT.16-0:1OUT.12-0:1OUT.9-1:1OUT.14-0:OUT.5-1:1OUT.13-1:" & Chr$(13)
    tom = Val(Text17.Text * -1)
    'c6k.Write "A10,5,10:V10,10,10:D" + Str(Val(Text17.Text * -1)) + ",-3:GOXX1:WAIT(MOV=b000000):GO10:JOG000000:" & Chr$(13)
End If


      
Timer2.Enabled = True
'joystick
End Sub
Private Sub Command14_Click()
'Disk Rotation CCW
Dim EStop
Dim temp() As Byte
Dim RunPass As Boolean
c6k.FSEnabled = True                'enable fast status
temp = c6k.FastStatus                  'get fast status information
Call CopyMemory(fsinfo, temp(0), 280)
c6k.Write ("DRIVE0,X,X,0,0,0,0:CMDDIR0000000:DRIVE1,X,X,1,1,1,1,0" & Chr$(13))
RotSpeed = ".05"
XSpeed1 = ".125"
XSpeed2 = ".119"
XSpeed3 = ".100"
XSpeed4 = ".075"
OssSpeed1 = "3.5"
OssSpeed2 = "3.4"
OssSpeed3 = "3.0"
OssSpeed4 = "2.8"
Label7.Caption = ""
Label7.Refresh
RunPass = False

For i = 1 To 6
    If Check1(i - 1).value = 1 Then
        RunPass = True
    End If
Next i
If RunPass = False Then
    MsgBox "Run Pass is Not Checked (On/Off)"
    Exit Sub
End If


' -------------- comment this block out for testing ---------
If RunCondition5 <> "YES" Then
    If RunCondition5 = "TimeOut" Then
       BAR3.Show
       Exit Sub
    End If
    If RunCondition5 = "NO" Then
        MsgBox "WORK ORDER NOT FOUND!"
        Exit Sub
    End If
    MsgBox "RunMode is Not ON!"
    Exit Sub
End If
' ------------ stop comments -------------------------------


Call Error:
StartCount5 = 0
c6k.Write ("ERASE" & Chr$(13))
c6k.Write ("MC0000000:MA1100000:COMEXC0:1INFNC20-D:INENXXXX1XX" & Chr$(13)) ' 20= E-Stop
c6k.Write ("!JOG0000000:OUT.9-1:OUT.16-0:1OUT.13-1:T1:1OUT.15-1:" & Chr$(13)) '9= Lens Fan / 16= Welder Contact / 13= Exhust Fan / 15= Water Pump
c6k.Write ("COMEXS0:COMEXL0:SCALE1:LH0,0,0,0,0,0,0:SCLD26550,257143,257143,12500,62500,62500,25000:" & Chr$(13))
      'HOME OSS
c6k.Write ("1INFNC18-4T:" & Chr$(13)) '18 = Oss Home Limit
c6k.Write ("@A10:@AD10:@V4:D,,,-.3,,:GO0001000:" & Chr$(13))
c6k.Write ("HOMA,,,1,,,:HOMAD,,,50,,,:@HOMZ0:HOMV,,,1,,,:HOMVF,,,1,,,:" & Chr$(13))
c6k.Write ("HOMBAC1110111:HOMEDG1111111:HOMDF0001000:HOM,,,0,,,:" & Chr$(13))
c6k.Write ("WAIT(6AS=XXX1XXX):T.1:D,,,-2.375,,,:GO0001000:" & Chr$(13))
    
For i = 1 To 6
    If Check1(i - 1).value = 1 Then
        Yoffset1 = Format((Val(Text8(i).Text) + Val(Text9(i).Text * 0.5)), "####0.000") 'Y START POS
        Xoffset1 = Format(Val(Text7(i).Text - 2), "####0.000") 'X STOP -2"
        Yoffset2 = Format(Val(CalTig_Y) + (Val(Text8(i).Text) + Val(Text9(i).Text * 0.5)), "####0.000")
        Xoffset2 = Format((Val(CalTig_X) - 2), "####0.000")
        Exit For
    End If
Next i
   
    'JOYSTICK ON
Label7.Caption = "Set 0,0    Press JoyStick Release"
Label7.Refresh
tom = 0
JogInput1 = 0
JogInput2 = 0
Do
    'E-STOP
Last_Pcut_State = 0 'attempt to recognize if e-stop is engaged during startup
temp = c6k.FastStatus
Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
     If (Last_Pcut_State = 0) Then
            c6k.Write "JOG0000000:1OUTALL9,16,0:1OUTALL25,32,0:T2:1OUT.9-1:1OUT.13-1" & Chr$(13)
            Last_Pcut_State = 1
            cutlist.Label7.Caption = "E-STOP!!!"
            cutlist.Label7.Refresh
            Command1.Enabled = True
            
            Exit Sub
     End If
Else
     If (Last_Pcut_State = 1) Then
        Last_Pcut_State = 0
     End If
End If



    Call CopyMemory(fsinfo, temp(0), 280)
    temp = c6k.FastStatus
' -------------------------------------------------------------------------- Joystick stuff -----------------------------------------
    'If JogOn = True Then
    ' axis select on, limit switch off
        If CStr((fsinfo.ProgIn(1) And Input22) / Input22) = 0 And ((fsinfo.ProgIn(1) And Input8) / Input8) = 0 Then
            If JogInput1 = 0 And ((fsinfo.ProgIn(1) And Input24) / Input24) > 0 Then
                c6k.Write ("!JOG0000000:" & Chr$(13))
                c6k.Write ("JOG0000000:1INFNC1-7J:1INFNC2-7K:1INFNC3-2K:1INFNC4-2J:JOGA2,5,5,1,5,5,5:JOGAD50,99,99,99,99,15,99:JOGVH5,8,10,2,5,3,5:JOGVL5,15,10,5,5,5,5:JOG0100100" & Chr$(13))
                JogInput1 = 1
                JogInput2 = 0
                JogInput3 = 0
                JogInput4 = 0
                JogInput5 = 0
                Label7.Caption = "    Joystick 1"
                Label7.Refresh
            End If
        End If
        ' axis select off, limit switch off
        If CStr((fsinfo.ProgIn(1) And Input22) / Input22) And ((fsinfo.ProgIn(1) And Input24) / Input24) > 0 Then
            If JogInput2 = 0 Then
                c6k.Write ("!JOG0000000:" & Chr$(13))
                c6k.Write ("JOG0000000:1INFNC2-7J:1INFNC1-7K:1INFNC3-6K:1INFNC4-6J:JOGA3,5,5,10,5,5,5:JOGAD50,99,99,99,99,15,99:JOGVH8,8,10,2,5,3,5:JOGVL8,15,10,5,5,5,5:JOG0000011" & Chr$(13))
                JogInput2 = 1
                JogInput1 = 0
                JogInput3 = 0
                JogInput4 = 0
                JogInput5 = 0
                Label7.Caption = "    Joystick 2"
                Label7.Refresh
            End If
        End If
        ' axis select on, limit switch on
        If CStr((fsinfo.ProgIn(1) And Input22) / Input22) = 0 And ((fsinfo.ProgIn(1) And Input8) / Input8) > 0 Then
            If JogInput3 = 0 Then
                c6k.Write ("!JOG0000000:!PSET,,0" & Chr$(13))
                c6k.Write ("JOG0000000:1INFNC1-7J:1INFNC2-7K:1INFNC3-3K:1INFNC4-3J:JOGA3,5,5,10,5,5,5:JOGAD50,99,99,99,99,15,99:JOGVH8,8,10,2,5,3,5:JOGVL8,15,10,5,5,5,5:JOG0010001" & Chr$(13))
                JogInput1 = 0
                JogInput2 = 0
                JogInput3 = 1
                JogInput4 = 0
                JogInput5 = 0
                Label7.Caption = "    Joystick 3"
                Label7.Refresh
                End If
            End If
        ' axis select on, limit switch on, joystick down, text16 motor position
         If CStr((fsinfo.ProgIn(1) And Input22) / Input22) = 0 And ((fsinfo.ProgIn(1) And Input8) / Input8) > 0 And ((fsinfo.ProgIn(1) And Input3) / Input3) > 0 And Val(Text16.Text) < 1 Then
            If JogInput4 = 0 Then
                c6k.Write ("!JOG0000000:" & Chr$(13))
                c6k.Write ("JOG0000000:1INFNC1-7K:1INFNC2-7J:1INFNC3-2K:1INFNC4-2J:JOGA3,5,5,10,5,5,5:JOGAD50,99,99,99,99,15,99:JOGVH8,8,10,2,5,3,5:JOGVL8,15,10,5,5,5,5:JOG0100001" & Chr$(13))
                JogInput1 = 0
                JogInput2 = 0
                JogInput3 = 1
                JogInput4 = 1
                JogInput5 = 0
                Label7.Caption = "   Joystick 4"
                Label7.Refresh
            End If
          End If
          ' Toggle trigger for axis 3 select
          If CStr((fsinfo.ProgIn(1) And Input24) / Input24) = 0 Then
            If JogInput5 = 0 Then
                c6k.Write ("!JOG0000000:" & Chr$(13))
                c6k.Write ("JOG0000000:1INFNC1-7K:1INFNC2-7J:1INFNC3-3K:1INFNC4-3J:JOGA3,5,5,10,5,5,5:JOGAD50,99,99,99,99,15,99:JOGVH8,5,8,5,5,3,5:JOGVL8,5,8,5,5,5,5:JOG0010001" & Chr$(13))
                JogInput1 = 0
                JogInput2 = 0
                JogInput3 = 0
                JogInput4 = 0
                JogInput5 = 1
                Label7.Caption = "   Jogging Lead Screw"
                Label7.Refresh
            End If
          End If
        If CStr((fsinfo.ProgIn(1) And Input23) / Input23) = 0 Then
            c6k.Write ("JOG0000000:" & Chr$(13))
            JogInput2 = 0
            JogInput1 = 0
            JogInput3 = 0
            JogInput4 = 0
            JogInput5 = 0
            Label7.Caption = ""
            Label7.Refresh
            jogOn = False
            Exit Do
        End If
        
        
    If JogInput1 = 1 Or JogInput2 = 1 Or JogInput3 = 1 Or JogInput4 = 1 Then
        cutlist.Text1.BackColor = &HFFFF&
        cutlist.Text1.ForeColor = QBColor(1)
        cutlist.Text1.Text = " JOYSTICK ON"
        cutlist.Text1.Refresh
        
       Maintenance1.Text8.BackColor = &HFFFF&
       Maintenance1.Text8.ForeColor = QBColor(1)
       Maintenance1.Text8.Text = " JOYSTICK ON"
       Maintenance1.Text8.Refresh
    End If
Loop ' UNTIL

Do
    temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
    'E-Stop
    If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
        If (Last_Pcut_State = 0) Then
            c6k.Write "1OUTALL10,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1" & Chr$(13) '9= Lens Fan / 13= Exhust Fan
            c6k.Write "!COMEXS0:" & Chr$(13)
            Last_Pcut_State = 1
            Label7.Caption = "E-STOP!!!!!"
            Label7.Refresh
            Timer2.Enabled = True
            EStopPos = ""
            Exit Sub
        End If
    Else
        If (Last_Pcut_State = 1) Then
            Last_Pcut_State = 0
        End If
    End If
 
    temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)
     '23= Jog Release
    If CStr((fsinfo.ProgIn(1) And Input23) / Input23) > 0 Then
        c6k.Write "JOG0000:PSET0,0,0,0" & Chr$(13)
        Label7.Caption = "JOYSTICK DONE"
        Label7.Refresh
        Exit Do 'get out of loop
     End If
Loop ' UNTIL
'c6k.Write "OUT.7-0" & Chr$(13)
Timer1.Enabled = False
Timer2.Enabled = False
Let Text12.Text = 0
c6k.Write "COMEXS0:DRFLVL1111111:" & Chr$(13)
c6k.Write "MC0000000:MA1100000:COMEXC0:" & Chr$(13)
c6k.Write "JOG0000000:COMEXC1" & Chr$(13)
Text19.Text = Yoffset2
Text19.Refresh
c6k.Write "JOG0000000" & Chr$(13)

'CHECK WATER FLOW
Label7.Caption = "Checking Water Pump"
Label7.Refresh
TESTWATERPUMP = 0
Do
    'E-Stop
    temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)
    If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
        If (Last_Pcut_State = 0) Then
            c6k.Write "1OUTALL10,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1" & Chr$(13) '9= Lens Fan / 13= Exhust Fan
            c6k.Write "!COMEXS0:" & Chr$(13)
            Last_Pcut_State = 1
            Label7.Caption = "E-STOP!!!!!"
            Label7.Refresh
            Timer2.Enabled = True
            EStopPos = ""
            Exit Sub
        End If
    Else
        If (Last_Pcut_State = 1) Then
            Last_Pcut_State = 0
            Label7.Caption = ""
            Label7.Refresh
        End If
    End If
    'Water Pump
    Call CopyMemory(fsinfo, temp(0), 280)
    If CStr((fsinfo.ProgIn(1) And Input19) / Input19) > 0 Then
        Label7.Caption = ""
        Label7.Refresh
        Exit Do
    End If
  
    If TESTWATERPUMP > 1000 Then
        Label7.Caption = " Water Pump HAS A PROBLEM!"
        Label7.Refresh
        Exit Sub
    End If
Loop

Label7.Caption = "Lower Head Until Contacting Blade"
Label7.Refresh

'Check "danger zone" before striking arc
If ((fsinfo.ProgIn(1) And Input8) / Input8) > 0 And (Text16.Text) < 1.5 Then
    Form4.Show 1
End If

'STRIKE ARC
If ((fsinfo.ProgIn(1) And Input8) / Input8) = 0 Then
    c6k.Write "PSET0,0:" & Chr$(13)
    c6k.Write "1OUT.12-1:1OUT.9-1:1OUT.30-1:A,10:AD,3:V,13:D,0.9:GO01:1OUT.16-1:WAIT(MOV=b0000000):JOG1000010" & Chr$(13)
    '9= Lens Fan / 12= Argon / 30= open test / 16= Welder Contact
    cutlist.Refresh
End If



A = 0
'MOVE ONTO TAB

If Val(Text9(i).Text) < 1.25 Then XSpeed = XSpeed1: OssSpeed = OssSpeed1
If Val(Text9(i).Text) > 1.249 Then XSpeed = XSpeed2: OssSpeed = OssSpeed2
If Val(Text9(i).Text) > 1.99 Then XSpeed = XSpeed3: OssSpeed = OssSpeed3
If Val(Text9(i).Text) > 2.49 Then XSpeed = XSpeed4: OssSpeed = OssSpeed4
c6k.Write "VAR10=" + Str(RotSpeed) + ":VAR11=" + Str(OssSpeed) + ":" & Chr$(13)
c6k.Write "MC0000000:MA1100000:COMEXC1:" & Chr$(13)
If Check1(i - 1).value = 1 Then
    If Val(Text9(i).Text) < 1 Then
        passWidth = 0
    Else
        passWidth = ((Val(Text9(i).Text) - 0.75))
    End If
    c6k.Write "VAR1=1" & Chr$(13)
    
    'START OF CHECK PROGRAM Sub loop during Oss
    c6k.Write "DEL CHECK:DEF CHECK" & Chr$(13)
        'Exit Check Program with release button
        'c6k.Write "IF(1IN.23=b0):1OUT.31-1:VAR4=0:VAR5=0:NIF" & Chr$(13) '23
        c6k.Write "IF(1IN.23=b0):1OUT.31-1:NIF" & Chr$(13) '23
        'Normal Y and Clockwise Auger joystick
        c6k.Write "IF(1IN.22=b1 AND VAR4=0):10UT.30-1:1OUT.32-0:VAR4=1:VAR5=0:JOG000X00X:1INFNC1-1K:1INFNC2-1J:1INFNC3-6K:1INFNC4-6J:JOGA4,4,15,1,5,5,5:JOGAD50,99,99,99,99,10,99:JOGVH8,8,10,2,5,3,5:JOGVL8,15,40,5,5,3,5:JOG100X01X:NIF" & Chr$(13)
        'Axis select for Z stage and Rotate
        c6k.Write "IF(1IN.22=b0 AND VAR5=0):10UT.30-1:1OUT.32-0:VAR4=0:VAR5=1:JOG000X00X:1INFNC1-7K:1INFNC2-7J:1INFNC3-2K:1INFNC4-2J:JOGA4,4,15,1,5,5,5:JOGAD50,99,99,99,99,10,99:JOGVH8,8,10,2,5,5,5:JOGVL8,15,40,5,5,5,5:JOG010X001:NIF" & Chr$(13)
        'Axis select AND limit switch engaged and rotate
        c6k.Write "IF(1IN.22=b0 AND 1IN.8=b1 AND VAR5=0):1OUT.30.1:1OUT.32-0:VAR4=0:VAR5=1:JOG000X00X:1INFNC1-5K:1INFNC2-5J:1INFNC3-3K:1INFNC4-3J:JOGA4,4,15,1,5,5,5:JOGAD50,99,99,99,99,15,99:JOGVH8,8,10,2,5,5,5:JOGVL8,15,40,5,5,5,5:JOG001010X:NIF" & Chr$(13)
    c6k.Write "END" & Chr$(13)
    'END OF CHECK PROGRAM
    
    'START OF CHECK1 PROGRAM Sub loop during Oss
    c6k.Write "DEL CHECK1:DEF CHECK1" & Chr$(13)
        
        'Control of X and Y - axis select off
        c6k.Write "IF(1IN.22=b1 AND VAR4=0):1OUT.30-1:1OUT.32-0:VAR4=1:VAR5=0:JOG000X00X:1INFNC2-7J:1INFNC1-7K:1INFNC3-6K:1INFNC4-6J:JOGA4,1,15,1,5,3,5:JOGAD50,99,99,99,99,3,99:JOGVH8,.5,10,2,5,5,5:JOGVL8,.8,40,5,5,5,5:JOG100X01X:NIF" & Chr$(13) '22
        'Control of rotate and Z1 - axis select on
        c6k.Write "IF(1IN.22=b0 AND VAR5=0):1OUT.32-1:1OUT.30-0:VAR5=1:VAR4=0:JOG000X00X:1INFNC1-5K:1INFNC2-5J:1INFNC3-2K:1INFNC4-2J:JOGAX,1,15,10,5,5,5:JOGADX,99,99,99,99,3,99:JOGVHX,.2,10,7,5,5,5:JOGVLX,15,12,1,5,5,5:JOGX10X1X:NIF" & Chr$(13) '22
        'Press release button to exit check 1
        c6k.Write "IF(1IN.23=b0):1OUT.31-1:VAR4=0:VAR5=0:NIF" & Chr$(13) '23
        
        'PAUSE
        c6k.Write "IF(1IN.24=b0 AND VAR1=1):VAR2=7VEL:MC1000000:V,,,,,,0:D,,,,,,0:GO,,,,,,1:NIF " & Chr$(13)
        c6k.Write "IF(1IN.24=b0 AND VAR1=1):MC0000000:JOG100X010:1INFNC2-1J:1INFNC1-1K:1INFNC3-6K:1INFNC4-6J:JOGA4,1,15,1,5,2:JOGAD50,99,99,99,99,5:JOGVH8,.5,10,2,5,3:JOGVL8,.8,40,5,5,3:JOG100X010:VAR1=2:NIF:" & Chr$(13)
        c6k.Write "IF(1IN.24=b1 AND VAR1=2):VAR10=VAR2:JOG010X110:MC0000001:D,,,,,,-1:V,,,,,,(VAR10):GO,,,,,,1:VAR1=1:NIF " & Chr$(13)

        'X AXIS SPEED CONTROL  --- will be rotation speed control?
        c6k.Write "IF(1IN.22=b1 AND 1IN.24=b1 AND VAR1=1 AND 1IN=b01):VAR10=VAR10-.001:V,,,,,,(VAR10):GO,,,,,,1:NIF " & Chr$(13)
        c6k.Write "IF(1IN.22=b1 AND 1IN.24=b1 AND VAR1=1 AND 1IN=b10):VAR10=VAR10+.001:V,,,,,,(VAR10):GO,,,,,,1:NIF " & Chr$(13)
    
    c6k.Write "END" & Chr$(13)
    'END OF CHECK1 PROGRAM
    
    'Jog on / Oss on / Move to Tab
    c6k.Write "MC0100001:VAR4=0:VAR5=0:VAR1=0:A,,,5,,:V,,,(VAR11),,:D,,,-" + Str$(passWidth / 2) + ":GO0001000:WAIT(MOV=bXXX0XXX):REPEAT:D,,," + Str$(passWidth) + ":GO0001000:REPEAT:GOSUB CHECK:UNTIL(MOV=bXXX0XXX):D,,,-" + Str$(passWidth) + ":GO0001000:REPEAT:GOSUB CHECK:UNTIL(MOV=bXXX0XXX):UNTIL(1OUT.31=b1):WAIT(1IN.23=b1):1OUT.31-0" & Chr$(13)
         
    'START TC AND X AXIS Move
    c6k.Write "JOG0:" & Chr$(13)
    c6k.Write "JOG0:MC0000001:VAR4=0:VAR5=0:VAR3=1:1OUT.14-1:A,,,,,,10:V,,,,,,(VAR10):D,,,,,,-1:GO,,,,,,1:WAIT(1OUT.14=b1):PSET0,0,0,,0,0,0:VAR1=1:INEN.2-1" & Chr$(13)
    c6k.Write "VAR1=1:V,,,(VAR11),,:REPEAT:D,,," + Str$(passWidth) + ":GO0001000:REPEAT:GOSUB CHECK1:UNTIL(MOV=bXXX0XXX):D,,,-" + Str$(passWidth) + ":GO0001000:REPEAT:GOSUB CHECK1:UNTIL(MOV=bXXX0XXX):UNTIL(1OUT.31=b1):WAIT(1IN.23=b1):1OUT.31-0:" & Chr$(13)
     
    'STOP X AXIS & FILL END / Jog
    c6k.Write "MC1000001:V,,,,,,0:GO,,,,,,1:1OUT.25-1" & Chr$(13)
    c6k.Write "VAR4=0:VAR5=0:VAR1=0:V,,,(VAR11),,:REPEAT:D,,," + Str$(passWidth) + ":GO0001000:REPEAT:GOSUB CHECK:UNTIL(MOV=bXXX0XXX):D,,,-" + Str$(passWidth) + ":GO0001000:REPEAT:GOSUB CHECK:UNTIL(MOV=bXXX0XXX):UNTIL(1OUT.31=b1):WAIT(1IN.23=b1):1OUT.31-0:INENXXXX0XX:" & Chr$(13)
       
    ' FAST STATUS UPDATE ESTOP/WATER/MOTOR POS
    TESTWATERPUMP = 0
    
    VoltCount = 0
    JogInput1 = 0
    Do
        temp = c6k.FastStatus
        Call CopyMemory(fsinfo, temp(0), 280)
        'E-Stop
        If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
            If (Last_Pcut_State = 0) Then
                c6k.Write "!1OUTALL10,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1:!S" & Chr$(13)
                c6k.Write "!COMEXS0:" & Chr$(13)
                Last_Pcut_State = 1
                Label7.Caption = "E-STOP!!!!!"
                Label7.Refresh
                Timer2.Enabled = True
                Exit Sub
                EStopPos = ""
            End If
        Else
            If (Last_Pcut_State = 1) Then
                Last_Pcut_State = 0
                Label7.Caption = ""
                Label7.Refresh
            End If
        End If
        TESTWATERPUMP = TESTWATERPUMP + 1
        Do
            temp = c6k.FastStatus
            Call CopyMemory(fsinfo, temp(0), 280)
            'Water Pump
            If CStr((fsinfo.ProgIn(1) And Input19) / Input19) > 0 Then
                Label7.Caption = ""
                Label7.Refresh
                Exit Do
            End If
  
            If TESTWATERPUMP > 1000 Then
                c6k.Write "!S:!1OUTALL10,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1" & Chr$(13)
                Label7.Caption = " Water Pump HAS A PROBLEM!"
                Label7.Refresh
                Exit Sub
            End If
        Loop
        '' ******* Update Motor Position ****************
        temp = c6k.FastStatus                  'get fast status information
        Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure

' ------------------------------------------- Motor Position --------------------------
Let X = InDex3
    'SCLD26550,257143,257143,12500,62500,62500
        XMOTOR = (fsinfo.MotorPos(1))
            Let Text14.Text = Val(XMOTOR)
            Let Text14.Text = Format((Text14.Text / 26550), "0.000")
        
        YMOTOR = (fsinfo.MotorPos(6))
            Let Text15.Text = Val(YMOTOR)
            Let Text15.Text = Format((Text15.Text / 62500), "0.000")
        
        ZMOTOR = (fsinfo.MotorPos(3))
            Let Text16.Text = Val(ZMOTOR)
            Let Text16.Text = Format((Text16.Text / 257153), "0.000")
        
        XSpeedSHOW = (fsinfo.MotorVel(1))
            Let Text18.Text = Val(XSpeedSHOW)
            Let Text18.Text = Format((Text18.Text / 660), "0.000")
        
        ROTATION = (fsinfo.MotorVel(7))
            Let Text20.Text = Val(ROTATION)
            Let Text20.Text = Format(((Text20.Text / 4432) / 1), "0.000")

        Call CopyMemory(fsinfo, temp(0), 280)
        If CStr((fsinfo.ProgIn(0) And Input5) / Input5) = 0 Then
            For P = 1 To 100000
            Next P
            Exit Do
        End If
        cutlist.Text20.Refresh
        cutlist.Text14.Refresh
        cutlist.Text15.Refresh
        cutlist.Text16.Refresh
        cutlist.Text17.Refresh
        cutlist.Text18.Refresh
        cutlist.Text12.Refresh
        For P = 1 To 100000
        Next P
        
    Loop
    
    c6k.Write "WAIT(MOV=bXX0XXX):MC000000:MA000000:1OUT.16-0:1OUT.12-0:1OUT.9-1:1OUT.14-0:OUT.5-1:1OUT.13-1:" & Chr$(13)
    tom = Val(Text17.Text * -1)
    'c6k.Write "A10,5,10:V10,10,10:D" + Str(Val(Text17.Text * -1)) + ",-3:GOXX1:WAIT(MOV=b000000):GO10:JOG000000:" & Chr$(13)
End If


      
Timer2.Enabled = True
'joystick
End Sub


Private Sub Command2_Click()
'Timer1.Enabled = True
StartCount5 = 0
StartTimer = Timer
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text4.BackColor = QBColor(15)
Text10.Text = ""
Label7.Caption = "Wand BarCode"
For i = 1 To 10
Text5(i).Text = ""
Text6(i).Text = ""
Text7(i).Text = ""
Text8(i).Text = ""
Text9(i).Text = ""

Check1(i - 1).value = 0
Check1(i - 1).Caption = "OFF"
Next i
BARCODE_READER
WORK_ORDER$ = InString$
'WORK_ORDER$ = InputBox("ENTER WORK ORDER #" + Chr(13) + "OR WAND BARCODE", "WORK ORDER#")
'FIND DWG NUMBER IN FORM.ASC
Open "F:\MFG\FORM.ASC" For Input As #1
While WORK_ORDER$ <> WOK$
     Input #1, WOK$, PartNumber, TQTY!, PRICE!, Tc!, COMM$, DAT$, DWG$, CSCOST!, IDontKnow!
        If WORK_ORDER$ = WOK$ Then
            Close #1
            RunCondition5 = "YES"
            Let DWG4 = DWG$
            Text1.Text = DWG4
            For i = 1 To Len(DWG4)
                If Left(DWG4, 1) = "M" Then DWG4 = Mid(DWG4, 2)
                If Mid(DWG4, i, 1) = Chr(45) Then DWG4 = Left(DWG4, i - 1) & Mid(DWG4, i + 1)
                DICK = Mid(DWG4, i, 1)
                tom = InStr(1, DWG4, Chr(45))
                Text1.Text = PartNumber
                Text2.Text = DWG$
                If Date >= CDate(DAT$) Then Text4.BackColor = QBColor(12)
                Text4.Text = DAT$
                Text10.Text = WOK$
            Next i
'CHECK FOR ACAD DWG
            'checkDWG = Dir("G:\ACAD\ABLDWG\" & DWG4 & ".DWG")
            'checkDWG1 = Dir("G:\DWG\" & DWG4 & ".DWG")
                'If checkDWG <> "" Or checkDWG1 <> "" Then
'VIEW DWG
                   ' Shell "C:\VOLOVIEW.EXE G:\ACAD\ABLDWG\" + DWG4 + ".DWG", vbMaximizedFocus
                'Else
                '    Msg = "Can Not Find An Exsisting Dwg !"   ' Define message.
                '    Style = vbOKOnly + vbCritical + vbDefaultButton1 ' Define buttons.
                '    Title = "ERROR"  ' Define title.
                '    help = "DEMO.HLP"   ' Define Help file.
                '    Ctxt = 1000 ' Define topic
                '    response = MsgBox(Msg, Style, Title, help, Ctxt)
                '    Exit Sub
                'End If
'CHECK FOR DTA. FILE
            
            checkfile = Dir("G:\ACAD\ABLASCII\" & DWG4 & ".dta")
            checkfile1 = Dir("G:\ACAD\ABLASCII\" & DWG4 & ".DTA")
            checkfile2 = Dir("G:\ACAD\ABLASCII\" & DWG4)
                If checkfile <> "" Or checkfile1 <> "" Or checkfile2 <> "" Then
                Else
                    msg = "Can Not Find An Exsisting DATA FILE !"   ' Define message.
                    Style = vbOKOnly + vbCritical + vbDefaultButton1 ' Define buttons.
                    Title = "ERROR"  ' Define title.
                    help = "DEMO.HLP"   ' Define Help file.
                    Ctxt = 1000 ' Define topic
                    response = MsgBox(msg, Style, Title, help, Ctxt)
                    Exit Sub
                End If

                If Right(DWG4, 3) = "dta" Or Right(DWG4, 3) = "DTA" Then
                    Open "G:\ACAD\ABLASCII\" & DWG4 For Input As #2
                Else
                    Open "G:\ACAD\ABLASCII\" & DWG4 & ".DTA" For Input As #2
                End If
'INPUT DATA FROM DTA FILE
                Input #2, init, DAT, DWG, PART, cust, disp, Length1, BWIDTH, THICK, Offset, offset2, BSIZE, offsetdim, Y1A, Y2A, Y3A, Y4A, eofblade, H1, H2, H3, H4, H5, H6, H7, H8, H9, H10, H11, H12, H13, H14, H15, H16, H17, H18, H19, H20, H21, H22, H23, H24, H25, boltqty, fstart, fyaxis, fstop, fpassw, bstart, byaxis, bstop, bpassw, ex1start, ex1yaxis, ex1stop, ex1passW, ex2start, ex2yaxis, ex2stop, ex2passW, pass5start, pass5yaxis, pass5stop, pass5passW, pass6start, pass6yaxis, pass6stop, pass6passW, fstart, fpassW1, leftpass, bwidthW1, lengthB1, fpassB1, Lengthr1, bwidthB2, toltc, FRONTBEVEL, BACKBEVEL, matl
                Close #2
                Text3.Text = Str(THICK) + " x" + Str(BWIDTH) + " x" + Str(Length1) + "    " + matl
                
               Let fy1 = fyaxis
               Let by1 = byaxis
               Let ex1 = ex1yaxis
               Let ex2 = ex2yaxis
               Let pass5y = pass5yaxis
               Let pass6y = pass6yaxis
               
               Let mintext = 500
               Let G = 1
               For C = 1 To 6
               If fy1 < mintext Then currmin = 1: mintext = fy1
               If by1 < mintext Then currmin = 2: mintext = by1
               If ex1 < mintext Then currmin = 3: mintext = ex1
               If ex2 < mintext Then currmin = 4: mintext = ex2
               If pass5y < mintext Then currmin = 5: mintext = pass5y
               If pass6y < mintext Then currmin = 6: mintext = pass6y
               
               If currmin = 1 Then
                    If fpassw > 0 Then
                        Text5(G).Text = "Front Pass": Text6(G).Text = fstart: Text7(G).Text = fstop: Text8(G).Text = fyaxis: Text9(G).Text = (fpassw - fyaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    fy1 = 500
               ElseIf currmin = 2 Then
                    If bpassw > 0 Then
                        Text5(G).Text = "Back Pass": Text6(G).Text = bstart: Text7(G).Text = bstop: Text8(G).Text = byaxis: Text9(G).Text = (bpassw - byaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    by1 = 500
               ElseIf currmin = 3 Then
                    If ex1passW > 0 Then
                        Text5(G).Text = "Front Extra Pass": Text6(G).Text = ex1start: Text7(G).Text = ex1stop: Text8(G).Text = ex1yaxis: Text9(G).Text = (ex1passW - ex1yaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    ex1 = 500
               ElseIf currmin = 4 Then
                    If ex2passW > 0 Then
                        Text5(G).Text = "Back Extra Pass": Text6(G).Text = ex2start: Text7(G).Text = ex2stop: Text8(G).Text = ex2yaxis: Text9(G).Text = (ex2passW - ex2yaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    ex2 = 500
              ElseIf currmin = 5 Then
                    If pass5passW > 0 Then
                        Text5(G).Text = "Pass 5": Text6(G).Text = pass5start: Text7(G).Text = pass5stop: Text8(G).Text = pass5yaxis: Text9(G).Text = (pass5passW - pass5yaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    pass5y = 500
              ElseIf currmin = 6 Then
                    If pass6passW > 0 Then
                        Text5(G).Text = "Pass 6": Text6(G).Text = pass6start: Text7(G).Text = pass6stop: Text8(G).Text = pass6yaxis: Text9(G).Text = (pass6passW - pass6yaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    pass6y = 500
              End If
              Let mintext = 500
              Next C
'Form1.Text4.BackColor = QBColor(15)
              If Val(fstart) <> Val(leftpass) Then
                     Text5(7).Text = "Left Pass": Text6(7).Text = fstart: Text7(7).Text = leftpass: Text8(7).Text = fpassW1: Text9(7).Text = (bwidthW1 - fpassW1)
                        Check1(6).value = 1: Check1(6).Caption = "ON"
              End If
              If Val(lengthB1) <> Val(Lengthr1) Then
                     Text5(8).Text = "Right Pass": Text6(8).Text = lengthB1: Text7(8).Text = Lengthr1: Text8(8).Text = fpassB1: Text9(8).Text = (bwidthB2 - fpassB1)
                        Check1(7).value = 1: Check1(7).Caption = "ON"
              End If
              If Val(FRONTBEVEL) > 0 Then
                     Text5(9).Text = "Front Bevel": Text9(9).Text = Format(FRONTBEVEL, "#.###")
                        Check1(8).value = 1: Check1(8).Caption = "ON"
              End If
              If Val(BACKBEVEL) > 0 Then
                     Text5(10).Text = "Back Bevel": Text9(10).Text = Format(BACKBEVEL, "#.###")
                        Check1(9).value = 1: Check1(9).Caption = "ON"
              End If
           If RunCondition5 = "YES" Then
             Timer2.Enabled = True
             Label6.Caption = "Run Mode On"
           End If
           Exit Sub
        End If
       If EOF(1) Then
            Close #1
            RunCondition5 = "NO"
            MsgBox "WORK ORDER NOT FOUND!"
            'Form2.Label1.Caption "WORK ORDER# " & InString$ & " Not Found in Form.asc"
            
            Exit Sub
        End If
Wend
Close #1

'For i = 0 To 5
'If Check1(i).Value = 1 Then
'If i = 0 Then Check1(0).Caption = "ON LEFT"
'If i <> 0 And Check1(i - 1).Caption = "ON LEFT" Then
'Check1(i).Caption = "ON RIGHT"
'Else
'Check1(i).Caption = "ON LEFT"
'End If
'Next i
End Sub





Private Sub Command3_Click()
If Dir("F:\BARCODE\temp6.tmp") = "" Then
           MsgBox ("No Active Work Order "), , ("ERROR")
           Exit Sub
        End If
             
RunCondition5 = ""
 
 Open "F:\BARCODE\temp6.tmp" For Input As #2
            Input #2, temp6, DAT
            Close #2
            
            Open "F:\BARCODE\" + temp6 + "6.tmp" For Input As #2
            Input #2, InString$, STIME!, ETIME!, TTime!, InString2
            Close #2
If InString2 = "TOPS" Or InString2 = "BEVEL" Then
'Finish_344
Exit Sub
ElseIf InString2 = "" Then
Finish_One4
Else
Finish_Ganged4
End If

End Sub





Private Sub Command4_Click()
'SCLD26550,257143,257143,12500,62500,62500

    Timer1.Enabled = False
    Timer2.Enabled = False
    c6k.Write ("MA0:MC0:SCLD26550:V5:A5:D10:GO1" & Chr$(13))

End Sub

Private Sub Command6_Click()
Text10.Visible = False
Label20.Visible = False
'Command13.Enabled = False
Command8.Enabled = False
'Command3.Enabled = False


    If RunCondition5 = "TimeOut" Then
       BAR3.Show
        Exit Sub
    End If
'Timer1.Enabled = True
StartCount5 = 0
StartTimer = Timer
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text4.BackColor = QBColor(15)
'Text10.Text = ""
'Label7.Caption = "Wand BarCode"
For i = 1 To 10
Text5(i).Text = ""
Text6(i).Text = ""
Text7(i).Text = ""
Text8(i).Text = ""
Text9(i).Text = ""

Check1(i - 1).value = 0
Check1(i - 1).Caption = "OFF"
Next i
 If Dir("F:\BARCODE\temp6.tmp") = "" Then
            'FileExists = False
            Label21.Caption = "Wand Work Order Bar Code"
        Else
            Open "F:\BARCODE\temp6.tmp" For Input As #1
            Let Date1$ = Date$
            Input #1, InString$, Date1$
            Close #1
            Label21.Caption = "Work Order " & InString$ & " Was Not Finished"
            Exit Sub
        End If
InString$ = Text13.Text

WORK_ORDER$ = InString$
'WORK_ORDER$ = InputBox("ENTER WORK ORDER #" + Chr(13) + "OR WAND BARCODE", "WORK ORDER#")
'FIND DWG NUMBER IN FORM.ASC
Open "f:\mfg\FORM.ASC" For Input As #1
While WORK_ORDER$ <> WOK$
     Input #1, WOK$, PartNumber, TQTY!, PRICE!, Tc!, COMM$, DAT$, DWG$, CSCOST!, IDontKnow!
        If WOK$ = "100000" Then
        tom = dich
        End If
        
        If WORK_ORDER$ = WOK$ Then
            Close #1
            RunCondition5 = "YES"
            Let DWG4 = DWG$
            Text1.Text = DWG4
            For i = 1 To Len(DWG4)
            If Left(DWG4, 1) = "M" Then DWG4 = Mid(DWG4, 2)
            If Mid(DWG4, i, 1) = Chr(45) Then DWG4 = Left(DWG4, i - 1) & Mid(DWG4, i + 1)
            DICK = Mid(DWG4, i, 1)
            tom = InStr(1, DWG4, Chr(45))
            Text1.Text = PartNumber
            Text2.Text = DWG$
            If Date >= CDate(DAT$) Then Text4.BackColor = QBColor(12)
            Text4.Text = DAT$
           'new Text10.Text = WOK$
            Next i
'CHECK FOR ACAD DWG
            'checkDWG = Dir("G:\ACAD\ABLDWG\" & DWG4 & ".DWG")
            'checkDWG1 = Dir("G:\DWG\" & DWG4 & ".DWG")
                'If checkDWG <> "" Or checkDWG1 <> "" Then
'VIEW DWG
                   ' Shell "C:\VOLOVIEW.EXE G:\ACAD\ABLDWG\" + DWG4 + ".DWG", vbMaximizedFocus
                'Else
                '    Msg = "Can Not Find An Exsisting Dwg !"   ' Define message.
                '    Style = vbOKOnly + vbCritical + vbDefaultButton1 ' Define buttons.
                '    Title = "ERROR"  ' Define title.
                '    help = "DEMO.HLP"   ' Define Help file.
                '    Ctxt = 1000 ' Define topic
                '    response = MsgBox(Msg, Style, Title, help, Ctxt)
                '    Exit Sub
                'End If
'CHECK FOR DTA. FILE

            checkfile = Dir("G:\ACAD\ABLASCII\" & DWG4 & ".dta")
            checkfile1 = Dir("G:\ACAD\ABLASCII\" & DWG4 & ".DTA")
            checkfile2 = Dir("G:\ACAD\ABLASCII\" & DWG4)
                If checkfile <> "" Or checkfile1 <> "" Or checkfile2 <> "" Then
                Else
                    Text5(1).Text = "Front Pass"
                    Text5(4).Text = "Back Pass"
                    Text5(2).Text = "Front Extra Pass"
                    Text5(3).Text = "Back Extra Pass"
                    Text5(5).Text = "Pass 5"
                    Text5(6).Text = "Pass 6"
                    Text5(7).Text = "Left Pass"
                    Text5(8).Text = "Right Pass"
                    Text5(9).Text = "Front Bevel"
                    Text5(10).Text = "Back Bevel"
                    Text6(1).SetFocus
                    msg = "Can Not Find An Exsisting DATA FILE !"   ' Define message.
                    Style = vbOKOnly + vbCritical + vbDefaultButton1 ' Define buttons.
                    Title = "ERROR"  ' Define title.
                    help = "DEMO.HLP"   ' Define Help file.
                    Ctxt = 1000 ' Define topic
                    response = MsgBox(msg, Style, Title, help, Ctxt)
                    RunCondition5 = "YES"
                    GoTo 10:
                End If

                If Right(DWG4, 3) = "dta" Or Right(DWG4, 3) = "DTA" Then
                    Open "G:\ACAD\ABLASCII\" & DWG4 For Input As #2
                Else
                   Open "G:\ACAD\ABLASCII\" & DWG4 & ".DTA" For Input As #2
                End If
'INPUT DATA FROM DTA FILE
                Input #2, init, DAT, DWG, PART, cust, disp, Length1, BWIDTH, THICK, Offset, offset2, BSIZE, offsetdim, Y1A, Y2A, Y3A, Y4A, eofblade, H1, H2, H3, H4, H5, H6, H7, H8, H9, H10, H11, H12, H13, H14, H15, H16, H17, H18, H19, H20, H21, H22, H23, H24, H25, boltqty, fstart, fyaxis, fstop, fpassw, bstart, byaxis, bstop, bpassw, ex1start, ex1yaxis, ex1stop, ex1passW, ex2start, ex2yaxis, ex2stop, ex2passW, pass5start, pass5yaxis, pass5stop, pass5passW, pass6start, pass6yaxis, pass6stop, pass6passW, fstart, fpassW1, leftpass, bwidthW1, lengthB1, fpassB1, Lengthr1, bwidthB2, toltc, FRONTBEVEL, BACKBEVEL, matl
                Close #2
                Text3.Text = Str(THICK) + " x" + Str(BWIDTH) + " x" + Str(Length1)
'CHECK FOR TC DATA
                If Val((fpassw - fyaxis)) + Val((bpassw - byaxis)) + Val((ex1passW - ex1yaxis)) + Val((ex2passW - ex2yaxis)) + Val((pass5passW - pass5yaxis)) + Val((pass6passW - pass6yaxis)) + Val((bwidthW1 - fpassW1)) + Val((bwidthB2 - fpassB1)) + Val(FRONTBEVEL) + Val(BACKBEVEL) < 1 Then
                     Text5(1).Text = "Front Pass"
                    Text5(4).Text = "Back Pass"
                    Text5(2).Text = "Front Extra Pass"
                    Text5(3).Text = "Back Extra Pass"
                    Text5(5).Text = "Pass 5"
                    Text5(6).Text = "Pass 6"
                    Text5(7).Text = "Left Pass"
                    Text5(8).Text = "Right Pass"
                    Text5(9).Text = "Front Bevel"
                    Text5(10).Text = "Back Bevel"
                    Text6(1).SetFocus
                    msg = "Can Not Find T.C. DATA !"   ' Define message.
                    Style = vbOKOnly + vbCritical + vbDefaultButton1 ' Define buttons.
                    Title = "ERROR"  ' Define title.
                    help = "DEMO.HLP"   ' Define Help file.
                    Ctxt = 1000 ' Define topic
                    response = MsgBox(msg, Style, Title, help, Ctxt)
                    RunCondition5 = "YES"
                    GoTo 10:
                End If
               
               Let fy1 = fyaxis
               Let by1 = byaxis
               Let ex1 = ex1yaxis
               Let ex2 = ex2yaxis
               Let pass5y = pass5yaxis
               Let pass6y = pass6yaxis
               
               Let mintext = 500
               Let G = 1
               For C = 1 To 6
               If fy1 < mintext Then currmin = 1: mintext = fy1
               If by1 < mintext Then currmin = 2: mintext = by1
               If ex1 < mintext Then currmin = 3: mintext = ex1
               If ex2 < mintext Then currmin = 4: mintext = ex2
               If pass5y < mintext Then currmin = 5: mintext = pass5y
               If pass6y < mintext Then currmin = 6: mintext = pass6y
               
              
               If currmin = 1 Then
                    If fpassw > 0 Then
                        Text5(G).Text = "Front Pass": Text6(G).Text = fstart: Text7(G).Text = fstop: Text8(G).Text = fyaxis: Text9(G).Text = (fpassw - fyaxis)
                        If Text9(G).Text > 2.625 Then Text9(G).Text = 2.625
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    fy1 = 500
               ElseIf currmin = 2 Then
                    If bpassw > 0 Then
                        Text5(G).Text = "Back Pass": Text6(G).Text = bstart: Text7(G).Text = bstop: Text8(G).Text = byaxis: Text9(G).Text = (bpassw - byaxis)
                        If Text9(G).Text > 2.625 Then Text9(G).Text = 2.625
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    by1 = 500
               ElseIf currmin = 3 Then
                    If ex1passW > 0 Then
                        Text5(G).Text = "Front Extra Pass": Text6(G).Text = ex1start: Text7(G).Text = ex1stop: Text8(G).Text = ex1yaxis: Text9(G).Text = (ex1passW - ex1yaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    ex1 = 500
               ElseIf currmin = 4 Then
                    If ex2passW > 0 Then
                        Text5(G).Text = "Back Extra Pass": Text6(G).Text = ex2start: Text7(G).Text = ex2stop: Text8(G).Text = ex2yaxis: Text9(G).Text = (ex2passW - ex2yaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    ex2 = 500
              ElseIf currmin = 5 Then
                    If pass5passW > 0 Then
                        Text5(G).Text = "Pass 5": Text6(G).Text = pass5start: Text7(G).Text = pass5stop: Text8(G).Text = pass5yaxis: Text9(G).Text = (pass5passW - pass5yaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    pass5y = 500
              ElseIf currmin = 6 Then
                    If pass6passW > 0 Then
                        Text5(G).Text = "Pass 6": Text6(G).Text = pass6start: Text7(G).Text = pass6stop: Text8(G).Text = pass6yaxis: Text9(G).Text = (pass6passW - pass6yaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    pass6y = 500
              End If
              Let mintext = 500
              Next C
'Form1.Text4.BackColor = QBColor(15)

              If Val(fstart) <> Val(leftpass) Then
                     Text5(7).Text = "Left Pass": Text6(7).Text = fstart: Text7(7).Text = leftpass: Text8(7).Text = fpassW1: Text9(7).Text = (bwidthW1 - fpassW1)
                        Check1(6).value = 1: Check1(6).Caption = "ON"
              End If
              If Val(lengthB1) <> Val(Lengthr1) Then
                     Text5(8).Text = "Right Pass": Text6(8).Text = lengthB1: Text7(8).Text = Lengthr1: Text8(8).Text = fpassB1: Text9(8).Text = (bwidthB2 - fpassB1)
                        Check1(7).value = 1: Check1(7).Caption = "ON"
              End If
              If Val(FRONTBEVEL) > 0 Then
                     Text5(9).Text = "Front Bevel": Text9(9).Text = Format(FRONTBEVEL, "#.###")
                        Check1(8).value = 1: Check1(8).Caption = "ON"
              End If
              If Val(BACKBEVEL) > 0 Then
                     Text5(10).Text = "Back Bevel": Text9(10).Text = Format(BACKBEVEL, "#.###")
                        Check1(9).value = 1: Check1(9).Caption = "ON"
              End If
10:
           If RunCondition5 = "YES" Then
              Timer2.Enabled = True
              Label21.Caption = "Run Mode On"
              Text13.Text = InString$
              Let Impregnator = 5
              Open "f:\mfg\atimp.asc" For Append As #1
              Let Date1$ = Date$
              Write #1, InString$, Date1$
              Close #1
             '***********
             'Shell ("f:\apps\exe\atimp.exe")

                If Dir("F:\BARCODE\" + InString$ + "6.tmp") = "" Then
                Else
                    Open "F:\BARCODE\" + InString$ + "6.tmp" For Input As #1
                    Input #1, InString$, STIME!, ETIME!, TTime!, Com2$
                    Close #1
                    Let ETIME! = 0
                    Let TTime! = TTime!
                    Let Com2$ = Com2$
                End If

            Let STIME! = Timer
            Open "F:\BARCODE\" + InString$ + "6.tmp" For Output As #2
            Write #2, InString$, STIME!, ETIME!, TTime!, Com2$
            Close #2

            Open "F:\BARCODE\temp6.tmp" For Output As #1
            Let Date1$ = Date$
            Write #1, InString$, Date1$
            Close #1
            Label21.Caption = "RunMode - ON"
            Label18.Caption = Format(Now, "SHORT TIME")
            'Label6.Caption = ""
            Command6.Enabled = False
            Open "f:\mfg\Bar_data.dta" For Append As #1
            Write #1, InString$, "6", "Start", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss")
            Close #1

           End If
'        For i = 0 To 5
'            If Check1(i).Value = 1 Then
'                If i = 0 Then Check1(0).Caption = "ON"
'                If i > 0 Then
'                    If Check1(i - 1).Caption = "ON LEFT" Then
'                        Check1(i).Caption = "ON RIGHT"
'                    Else
'                        Check1(i).Caption = "ON LEFT"
'                    End If
'                End If
'            End If
'        Next i
          Exit Sub
        End If
       If EOF(1) Then
            Close #1
            RunCondition5 = "NO"
            MsgBox "WORK ORDER NOT FOUND!"
            'Form2.Label1.Caption "WORK ORDER# " & InString$ & " Not Found in Form.asc"
            
            Exit Sub
        End If
Wend
Close #1

'For i = 0 To 5
'If Check1(i).Value = 1 Then
'If i = 0 Then Check1(0).Caption = "ON LEFT"
'If i <> 0 And Check1(i - 1).Caption = "ON LEFT" Then
'Check1(i).Caption = "ON RIGHT"
'Else
'Check1(i).Caption = "ON LEFT"
'End If
'Next i
'************* Start Old ***********



   
 
''Text5.Visible = False
'Label20.Visible = False
'Let Text2.Text = ""
'Let Label4.Caption = ""
'Let Label5.Caption = ""
'Label22.Caption = ""




    If RunCondition5 <> "YES" Then
        InString$ = ""
        Text13.Text = ""
        Let Impregnator = 0
        Exit Sub
    End If

Exit Sub
errorhand:
If Err.Number = 70 Then
Form7.Show 1
Resume 0
End If



End Sub



Private Sub Command8_Click()
Text10.Visible = True
Label20.Visible = True
'Command13.Enabled = False
Command8.Enabled = False
'Command3.Enabled = False


    If RunCondition5 = "TimeOut" Then
       BAR3.Show
        Exit Sub
    End If
'Timer1.Enabled = True
StartCount5 = 0
StartTimer = Timer
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text4.BackColor = QBColor(15)
'Text10.Text = ""
'Label7.Caption = "Wand BarCode"
For i = 1 To 10
Text5(i).Text = ""
Text6(i).Text = ""
Text7(i).Text = ""
Text8(i).Text = ""
Text9(i).Text = ""

Check1(i - 1).value = 0
Check1(i - 1).Caption = "OFF"
Next i
 If Dir("F:\BARCODE\temp6.tmp") = "" Then
            'FileExists = False
            Label21.Caption = "Wand First Work Order Bar Code (Lowest Number First)"
        Else
            Open "F:\BARCODE\temp6.tmp" For Input As #1
            Let Date1$ = Date$
            Input #1, InString$, Date1$
            Close #1
            Label21.Caption = "Work Order " & InString$ & " Was Not Finished"
            Exit Sub
        End If
'******************** First barcode **************
BARCODE_READER
WORK_ORDER$ = InString$
'WORK_ORDER$ = InputBox("ENTER WORK ORDER #" + Chr(13) + "OR WAND BARCODE", "WORK ORDER#")
'FIND DWG NUMBER IN FORM.ASC
Open "f:\mfg\FORM.ASC" For Input As #1
While WORK_ORDER$ <> WOK$
     Input #1, WOK$, PartNumber, TQTY!, PRICE!, Tc!, COMM$, DAT$, DWG$, CSCOST!, IDontKnow!
        If WORK_ORDER$ = WOK$ Then
            Close #1
            RunCondition5 = "YES"
            Let DWG4 = DWG$
            Text1.Text = DWG4
            For i = 1 To Len(DWG4)
            If Left(DWG4, 1) = "M" Then DWG4 = Mid(DWG4, 2)
            If Mid(DWG4, i, 1) = Chr(45) Then DWG4 = Left(DWG4, i - 1) & Mid(DWG4, i + 1)
            DICK = Mid(DWG4, i, 1)
            tom = InStr(1, DWG4, Chr(45))
            Text1.Text = PartNumber
            Text2.Text = DWG$
            If Date >= CDate(DAT$) Then Text4.BackColor = QBColor(12)
            Text4.Text = DAT$
           'new Text10.Text = WOK$
            Next i
'CHECK FOR ACAD DWG
            'checkDWG = Dir("G:\ACAD\ABLDWG\" & DWG4 & ".DWG")
            'checkDWG1 = Dir("G:\DWG\" & DWG4 & ".DWG")
                'If checkDWG <> "" Or checkDWG1 <> "" Then
'VIEW DWG
                   ' Shell "C:\VOLOVIEW.EXE G:\ACAD\ABLDWG\" + DWG4 + ".DWG", vbMaximizedFocus
                'Else
                '    Msg = "Can Not Find An Exsisting Dwg !"   ' Define message.
                '    Style = vbOKOnly + vbCritical + vbDefaultButton1 ' Define buttons.
                '    Title = "ERROR"  ' Define title.
                '    help = "DEMO.HLP"   ' Define Help file.
                '    Ctxt = 1000 ' Define topic
                '    response = MsgBox(Msg, Style, Title, help, Ctxt)
                '    Exit Sub
                'End If
'CHECK FOR DTA. FILE

            checkfile = Dir("G:\ACAD\ABLASCII\" & DWG4 & ".dta")
            checkfile1 = Dir("G:\ACAD\ABLASCII\" & DWG4 & ".DTA")
            checkfile2 = Dir("G:\ACAD\ABLASCII\" & DWG4)
                If checkfile <> "" Or checkfile1 <> "" Or checkfile2 <> "" Then
                Else
                    Text5(1).Text = "Front Pass"
                    Text5(4).Text = "Back Pass"
                    Text5(2).Text = "Front Extra Pass"
                    Text5(3).Text = "Back Extra Pass"
                    Text5(5).Text = "Pass 5"
                    Text5(6).Text = "Pass 6"
                    Text5(7).Text = "Left Pass"
                    Text5(8).Text = "Right Pass"
                    Text5(9).Text = "Front Bevel"
                    Text5(10).Text = "Back Bevel"
                    Text6(1).SetFocus
                    msg = "Can Not Find An Exsisting DATA FILE !"   ' Define message.
                    Style = vbOKOnly + vbCritical + vbDefaultButton1 ' Define buttons.
                    Title = "ERROR"  ' Define title.
                    help = "DEMO.HLP"   ' Define Help file.
                    Ctxt = 1000 ' Define topic
                    response = MsgBox(msg, Style, Title, help, Ctxt)
                    GoTo 10:
                End If

                If Right(DWG4, 3) = "dta" Or Right(DWG4, 3) = "DTA" Then
                    Open "G:\ACAD\ABLASCII\" & DWG4 For Input As #2
                Else
                    Open "G:\ACAD\ABLASCII\" & DWG4 & ".DTA" For Input As #2
                End If
'INPUT DATA FROM DTA FILE
                Input #2, init, DAT, DWG, PART, cust, disp, Length1, BWIDTH, THICK, Offset, offset2, BSIZE, offsetdim, Y1A, Y2A, Y3A, Y4A, eofblade, H1, H2, H3, H4, H5, H6, H7, H8, H9, H10, H11, H12, H13, H14, H15, H16, H17, H18, H19, H20, H21, H22, H23, H24, H25, boltqty, fstart, fyaxis, fstop, fpassw, bstart, byaxis, bstop, bpassw, ex1start, ex1yaxis, ex1stop, ex1passW, ex2start, ex2yaxis, ex2stop, ex2passW, pass5start, pass5yaxis, pass5stop, pass5passW, pass6start, pass6yaxis, pass6stop, pass6passW, fstart, fpassW1, leftpass, bwidthW1, lengthB1, fpassB1, Lengthr1, bwidthB2, toltc, FRONTBEVEL, BACKBEVEL, matl
                Close #2
                Text3.Text = Str(THICK) + " x" + Str(BWIDTH) + " x" + Str(Length1)
               If Val((fpassw - fyaxis)) + Val((bpassw - byaxis)) + Val((ex1passW - ex1yaxis)) + Val((ex2passW - ex2yaxis)) + Val((pass5passW - pass5yaxis)) + Val((pass6passW - pass6yaxis)) + Val((bwidthW1 - fpassW1)) + Val((bwidthB2 - fpassB1)) + Val(FRONTBEVEL) + Val(BACKBEVEL) < 1 Then
                     Text5(1).Text = "Front Pass"
                    Text5(4).Text = "Back Pass"
                    Text5(2).Text = "Front Extra Pass"
                    Text5(3).Text = "Back Extra Pass"
                    Text5(5).Text = "Pass 5"
                    Text5(6).Text = "Pass 6"
                    Text5(7).Text = "Left Pass"
                    Text5(8).Text = "Right Pass"
                    Text5(9).Text = "Front Bevel"
                    Text5(10).Text = "Back Bevel"
                    Text6(1).SetFocus
                    msg = "Can Not Find T.C. DATA !"   ' Define message.
                    Style = vbOKOnly + vbCritical + vbDefaultButton1 ' Define buttons.
                    Title = "ERROR"  ' Define title.
                    help = "DEMO.HLP"   ' Define Help file.
                    Ctxt = 1000 ' Define topic
                    response = MsgBox(msg, Style, Title, help, Ctxt)
                    RunCondition5 = "YES"
                    GoTo 10:
                End If
               Let fy1 = fyaxis
               Let by1 = byaxis
               Let ex1 = ex1yaxis
               Let ex2 = ex2yaxis
               Let pass5y = pass5yaxis
               Let pass6y = pass6yaxis
               
               Let mintext = 500
               Let G = 1
               For C = 1 To 6
               If fy1 < mintext Then currmin = 1: mintext = fy1
               If by1 < mintext Then currmin = 2: mintext = by1
               If ex1 < mintext Then currmin = 3: mintext = ex1
               If ex2 < mintext Then currmin = 4: mintext = ex2
               If pass5y < mintext Then currmin = 5: mintext = pass5y
               If pass6y < mintext Then currmin = 6: mintext = pass6y
               
               If currmin = 1 Then
                    If fpassw > 0 Then
                        Text5(G).Text = "Front Pass": Text6(G).Text = fstart: Text7(G).Text = fstop: Text8(G).Text = fyaxis: Text9(G).Text = (fpassw - fyaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    fy1 = 500
               ElseIf currmin = 2 Then
                    If bpassw > 0 Then
                        Text5(G).Text = "Back Pass": Text6(G).Text = bstart: Text7(G).Text = bstop: Text8(G).Text = byaxis: Text9(G).Text = (bpassw - byaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    by1 = 500
               ElseIf currmin = 3 Then
                    If ex1passW > 0 Then
                        Text5(G).Text = "Front Extra Pass": Text6(G).Text = ex1start: Text7(G).Text = ex1stop: Text8(G).Text = ex1yaxis: Text9(G).Text = (ex1passW - ex1yaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    ex1 = 500
               ElseIf currmin = 4 Then
                    If ex2passW > 0 Then
                        Text5(G).Text = "Back Extra Pass": Text6(G).Text = ex2start: Text7(G).Text = ex2stop: Text8(G).Text = ex2yaxis: Text9(G).Text = (ex2passW - ex2yaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    ex2 = 500
              ElseIf currmin = 5 Then
                    If pass5passW > 0 Then
                        Text5(G).Text = "Pass 5": Text6(G).Text = pass5start: Text7(G).Text = pass5stop: Text8(G).Text = pass5yaxis: Text9(G).Text = (pass5passW - pass5yaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    pass5y = 500
              ElseIf currmin = 6 Then
                    If pass6passW > 0 Then
                        Text5(G).Text = "Pass 6": Text6(G).Text = pass6start: Text7(G).Text = pass6stop: Text8(G).Text = pass6yaxis: Text9(G).Text = (pass6passW - pass6yaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    pass6y = 500
              End If
              Let mintext = 500
              Next C
'Form1.Text4.BackColor = QBColor(15)
              If Val(fstart) <> Val(leftpass) Then
                     Text5(7).Text = "Left Pass": Text6(7).Text = fstart: Text7(7).Text = leftpass: Text8(7).Text = fpassW1: Text9(7).Text = (bwidthW1 - fpassW1)
                        Check1(6).value = 1: Check1(6).Caption = "ON"
              End If
              If Val(lengthB1) <> Val(Lengthr1) Then
                     Text5(8).Text = "Right Pass": Text6(8).Text = lengthB1: Text7(8).Text = Lengthr1: Text8(8).Text = fpassB1: Text9(8).Text = (bwidthB2 - fpassB1)
                        Check1(7).value = 1: Check1(7).Caption = "ON"
              End If
              If Val(FRONTBEVEL) > 0 Then
                     Text5(9).Text = "Front Bevel": Text9(9).Text = Format(FRONTBEVEL, "#.###")
                        Check1(8).value = 1: Check1(8).Caption = "ON"
              End If
              If Val(BACKBEVEL) > 0 Then
                     Text5(10).Text = "Back Bevel": Text9(10).Text = Format(BACKBEVEL, "#.###")
                        Check1(9).value = 1: Check1(9).Caption = "ON"
              End If
       GoTo 10
        End If
If EOF(1) Then
            Close #1
            RunCondition5 = "NO"
            MsgBox "WORK ORDER NOT FOUND!"
            'Form2.Label1.Caption "WORK ORDER# " & InString$ & " Not Found in Form.asc"
            
            Exit Sub
        End If
Wend

Close #1
'******************** Second Barcode ************************
10:
For A = 1 To 1000000
Next A

            Label21.Caption = "Wand Second Work Order Bar Code"
       
BARCODE_READER
WORK_ORDER2$ = InString$
'WORK_ORDER$ = InputBox("ENTER WORK ORDER #" + Chr(13) + "OR WAND BARCODE", "WORK ORDER#")
'FIND DWG NUMBER IN FORM.ASC
Open "f:\mfg\FORM.ASC" For Input As #1
While WORK_ORDER2$ <> WOK$
     Input #1, WOK$, PartNumber, TQTY!, PRICE!, Tc!, COMM$, DAT$, DWG$, CSCOST!, IDontKnow!
        If WORK_ORDER2$ = WOK$ Then
            Close #1
            RunCondition5 = "YES"
            
'CHECK FOR DTA. FILE
            
            checkfile = Dir("G:\ACAD\ABLASCII\" & DWG4 & ".dta")
            checkfile1 = Dir("G:\ACAD\ABLASCII\" & DWG4 & ".DTA")
            checkfile2 = Dir("G:\ACAD\ABLASCII\" & DWG4)
                If checkfile <> "" Or checkfile1 <> "" Or checkfile2 <> "" Then
                Else
                    msg = "Can Not Find An Exsisting DATA FILE !"   ' Define message.
                    Style = vbOKOnly + vbCritical + vbDefaultButton1 ' Define buttons.
                    Title = "ERROR"  ' Define title.
                    help = "DEMO.HLP"   ' Define Help file.
                    Ctxt = 1000 ' Define topic
                    response = MsgBox(msg, Style, Title, help, Ctxt)
                    Exit Sub
                End If

              
       GoTo 20
        End If
If EOF(1) Then
            Close #1
            RunCondition5 = "NO"
            MsgBox "WORK ORDER NOT FOUND!"
            'Form2.Label1.Caption "WORK ORDER# " & InString$ & " Not Found in Form.asc"
            
            Exit Sub
        End If
Wend

Close #1
'******************** End Second Barcode ************************
20:
           If WORK_ORDER2$ < WORK_ORDER$ Then
           InString$ = ""
           Text13.Text = ""
           Text10.Text = ""
           WORK_ORDER$ = ""
           WORK_ORDER2$ = ""
           Impregnator = 0
           Label21.Caption = "Press Start  or  Ganged  "
           Command6.Enabled = True
           Command8.Enabled = True
           Command3.Enabled = True
           Command9.Enabled = True
           Command13.Enabled = True
           MsgBox ("First BarCode Must Be Less Than Second"), , ("ERROR")
           Exit Sub
           End If
           If RunCondition5 = "YES" Then
              Timer2.Enabled = True
              Label21.Caption = "Run Mode On"
              Text13.Text = WORK_ORDER$
              Text10.Text = WORK_ORDER2$
              Let Impregnator = 4
              Open "f:\mfg\atimp.asc" For Append As #1
              Let Date1$ = Date$
              Write #1, WORK_ORDER$, Date1$
              Write #1, WORK_ORDER2$, Date1$
              Close #1
             '***********
             'Shell ("f:\apps\exe\atimp.exe")

              If Dir("F:\BARCODE\" + WORK_ORDER$ + Right(WORK_ORDER2$, 2) + "6.tmp") = "" Then
              Else
                Open "F:\BARCODE\" + WORK_ORDER$ + Right(WORK_ORDER2$, 2) + "6.tmp" For Input As #1
                Input #1, WORK_ORDER2$, STIME!, ETIME!, TTime!, WORK_ORDER$
                Close #1
              End If

            Let STIME! = Timer
                Open "F:\BARCODE\" + WORK_ORDER$ + Right(WORK_ORDER2$, 2) + "6.tmp" For Output As #1
               Write #1, WORK_ORDER2$, STIME!, ETIME!, TTime!, WORK_ORDER$
                Close #1
            
            Open "F:\BARCODE\temp6.tmp" For Output As #1
            Let Date1$ = Date$
            Write #1, WORK_ORDER$ + Right(WORK_ORDER2$, 2), Date1$
            Close #1
            Label21.Caption = "RunMode - ON"
            Label18.Caption = Format(Now, "SHORT TIME")
            'Label6.Caption = ""
            Command6.Enabled = False
            Open "f:\mfg\Bar_data.dta" For Append As #1
            Write #1, WORK_ORDER$, "6", "Ganged", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss")
            Write #1, WORK_ORDER2$, "6", "Ganged", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss")
            Close #1

           End If
        ' GoTo 10
         ' Exit Sub
        

'************* Start Old ***********

    If RunCondition5 <> "YES" Then
        InString$ = ""
       ' Text13.Text = ""
        Let Impregnator = 0
        Exit Sub
        Else
'        For i = 0 To 5
'            If Check1(i).Value = 1 Then
'                If i = 0 Then Check1(0).Caption = "ON LEFT"
'                If i > 0 Then
'                    If Check1(i - 1).Caption = "ON LEFT" Then
'                        Check1(i).Caption = "ON RIGHT"
'                    Else
'                        Check1(i).Caption = "ON LEFT"
'                    End If
'                End If
'            End If
'        Next i
        
    End If

Exit Sub
errorhand:
If Err.Number = 70 Then
Form7.Show 1
Resume 0
End If

End Sub

Private Sub Command9_Click()
If RunCondition5 = "TimeOut" Then
        BAR3.Show
        Exit Sub
    End If
RunCondition5 = ""
If Dir("F:\BARCODE\temp6.tmp") = "" Then
       MsgBox ("No Active Work Order "), , ("ERROR")
   Label21.Caption = "Press Start  or  Ganged  "
   Command6.Enabled = True
   Command8.Enabled = True
   Command3.Enabled = True
   Label20.Visible = False
   Text10.Visible = False
   'Text5.Visible = False
   Label21.Caption = "Press Start  or  Ganged  "
Exit Sub
End If
    Timer2.Enabled = False
    StartCount5 = 0
    BAR4.Show 1
    Let Label21.Caption = BAR4!Label1.Caption

        If BAR4!Label1.Caption = "CANCEL" Then
        Let Label21.Caption = ""
           Exit Sub
        End If

    If BAR4!Label1.Caption = "CLEAR" Then
         msg = "THIS WILL CLEAR ALL INFO.!!   CONTINUE?"
            Style = vbYesNo + vbCritical + vbDefaultButton2
            Title = "CLEAR ALL"
            response = MsgBox(msg, Style, Title, help, Ctxt)
            If response <> vbYes Then Exit Sub
Open "f:\mfg\Bar_Data.dta" For Append As #1
Write #1, InString$, "6", "Clear W/O", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss")
Close #1
        
        RunCondition5 = "NO"
        Let InString$ = ""
        Let InString1 = ""
        Let InString2 = ""
        Label21.Caption = "RunMode - OFF"
        Label5.Caption = ""
        'Text5.Visible = False
        Label20.Visible = False
        Let Text2.Text = ""
        Let Label4.Caption = ""
        Kill "F:\BARCODE\temp6.tmp"
        'ULStat% = cbDBitOut%(0, 10, 2, 0)
        'If ULStat% <> 0 Then Stop
        Command6.Enabled = True
        Command8.Enabled = True
        Command13.Enabled = True

        Exit Sub
     End If
        
    If BAR4!Label1.Caption = "Change W/O" Then
        Label21.Caption = "WAND NEW WORK ORDER NUMBER"
        Let Impregnator = 4
        Let InStringOLD$ = InString$
        Call BARCODE_READER
        Label21.Caption = ""

Open "f:\mfg\Bar_Data.dta" For Append As #1
Write #1, InString$, "6", "Change W/O", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss")
Close #1

Call CheckForm1
        If RunCondition5 <> "YES" Then
            InString$ = InStringOLD$
            Text13.Text = InString$
            Let Impregnator = 0
            Exit Sub
        End If
    
        Open "f:\mfg\atimp.asc" For Append As #1
        Let Date1$ = Date$
        Write #1, InString$, Date1$
        Close #1
'Shell ("f:\apps\exe\atimp.exe")
 
        Open "F:\BARCODE\" + InStringOLD$ + "6.tmp" For Input As #2
        Input #2, InStringOLD$, STIME!, ETIME!, TTime!, Com2$
        Close #2

        Open "F:\BARCODE\" + InString$ + "6.tmp" For Output As #2
        Write #2, InString$, STIME!, ETIME!, TTime!, Com2$
        Close #2

        Kill "F:\BARCODE\" + InStringOLD$ + "6.tmp"
        Kill "F:\BARCODE\temp6.tmp"

        Open "F:\BARCODE\temp6.tmp" For Output As #1
        Let Date1$ = Date$
        Write #1, InString$, Date1$
        Close #1
        Text13.Text = InString$
        Label21.Caption = "RunMode - ON"
        Label18.Caption = Format(Now, "SHORT TIME")


'******* relay on ********
      '  ULStat% = cbDBitOut%(0, 10, 2, 1)
      '  If ULStat% <> 0 Then Stop
     Timer2.Enabled = True



        If RunCondition5 = "TimeOut" Then
            BAR3.Show
            Exit Sub
        End If
 End If
End Sub



Private Sub Form_Activate()
Call Error:
c6k.Write (":COMEXS0" & Chr$(13))
c6k.Write ("1INEN.5-0:" & Chr$(13))
c6k.Write ("DRIVE0,0,0,0,0,0,0,0:CMDDIR000000:DRFEN00000000" & Chr$(13))
c6k.Write "AXSDEF00000000:@DRES250000:SCLD26550,39683,510204,,62500,62500,62500:@SCLA25000:@SCLV25000:SCALE1:LH0,0,0,0,0,0,0,0:LHAD100,100,100,100,100,100,100,100:" + Chr$(13)
c6k.Write ("ENCPOL11011111:ERES4000,4000,4000,4000,4000,4000,4000,4000:ENCCNT11111111:" & Chr$(13))
c6k.Write ("COMEXC0" & Chr$(13))
c6k.Write ("DRIVE1,1,1,1,1,1,1,0" & Chr$(13))
c6k.Write "INEN.20-E:!JOG0000000:1OUT.9-1:1OUT.13-1" & Chr$(13)
Timer2.Enabled = False
Timer1.Enabled = False
Text5(0).Text = ""
Text6(0).Text = "X_Start"
Text7(0).Text = "X_Stop"
Text8(0).Text = "Y_Start"
Text9(0).Text = "Pass Width"
MSFlexGrid1.ColWidth(0) = 800
MSFlexGrid1.ColWidth(1) = 1800
MSFlexGrid1.ColWidth(2) = 400
MSFlexGrid1.ColWidth(3) = 500
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.ColWidth(5) = 1100
MSFlexGrid1.ColWidth(6) = 1200
MSFlexGrid1.ColWidth(7) = 2000

MSFlexGrid1.col = 0
MSFlexGrid1.row = 0
MSFlexGrid1.Text = "W.O. #"
MSFlexGrid1.col = 1
MSFlexGrid1.Text = "Part Number"
MSFlexGrid1.col = 2
MSFlexGrid1.Text = "Inp"
MSFlexGrid1.col = 3
MSFlexGrid1.Text = "QTY"
MSFlexGrid1.col = 4
MSFlexGrid1.Text = "Avg. Time"
MSFlexGrid1.col = 5
MSFlexGrid1.Text = "Run Time"
MSFlexGrid1.col = 6
MSFlexGrid1.Text = "Difference (min)"
MSFlexGrid1.col = 7
MSFlexGrid1.Text = "Comments"
Timer1.Enabled = True

End Sub

Private Sub Form_Load()
0 c6k.Write (":COMEXS0" & Chr$(13))
 c6k.Write ("1INEN.5-0:" & Chr$(13))
c6k.Write ("DRIVE0,0,0,0,0,0,0,0:DRFEN00000000" & Chr$(13))
c6k.Write "AXSDEF00000000:@DRES250000:SCLD26550,39683,510204,,62500,62500,62500:@SCLA25000:SCALE1:LH0,0,0,0,0,0,0,0:LHAD100,100,100,100,100,100,100,100:" + Chr$(13)
c6k.Write ("ENCPOL11011111:ERES4000,4000,4000,4000,4000,4000,4000,4000:ENCCNT11111111:" & Chr$(13))
c6k.Write ("COMEXC0" & Chr$(13))
c6k.Write ("DRIVE1,1,1,1,1,1,1,0" & Chr$(13))
c6k.Write "!JOG00000000:1OUT.9-1:1OUT.13-1" & Chr$(13)

'----------------------------------------------------------------------------
'c6k.FSEnabled = True
'temp = c6k.FastStatus
'Call CopyMemory(fsinfo, temp(0), 280)

'Do While Counter < 2000
'    Counter = Counter + 1
'Loop


'If CStr(fsinfo.ProgIn(1) And Input20 / Input20) > 0 Then
'    Form5.Show 1
'    End If
'---------------------------------------------------------------------------

'Open_File.Show
'Form1.Show
'Text10.Visible = False
'cmd$ = "OUT.7-1" + Chr$(ENTER)
'       tmp% = SendAT6400Block(Device_Address%, cmd$, 0)
'        If Dir("F:\BARCODE\temp6.tmp") = "" Then
'            Label21.Caption = "Press Start  or  Ganged  "
'        Else
'            Open "F:\BARCODE\temp6.tmp" For Input As #1
'            Let Date1$ = Date$
'            Input #1, InString$, Date1$
'            Close #1
'            Label21.Caption = "Work Order " & InString$ & " Was Not Finished"
'            Exit Sub
'        End If
MSFlexGrid1.ColWidth(0) = 800
MSFlexGrid1.ColWidth(1) = 1800
MSFlexGrid1.ColWidth(2) = 500
MSFlexGrid1.ColWidth(3) = 500
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.ColWidth(5) = 1100
MSFlexGrid1.ColWidth(6) = 1100
MSFlexGrid1.ColWidth(7) = 2000
'msflexgrid1.AddItem "W.O. #" & Chr(9) & "Part Number" & Chr(9) & "QTY" & Chr(9) & "Avg. Time" & Chr(9) & "Steel Cost" & Chr(9) & "Gross Profit", 0
MSFlexGrid1.col = 0
MSFlexGrid1.row = 0
MSFlexGrid1.Text = "W.O. #"
MSFlexGrid1.col = 1
MSFlexGrid1.Text = "Part Number"
MSFlexGrid1.col = 2
MSFlexGrid1.Text = "Inp"
MSFlexGrid1.col = 3
MSFlexGrid1.Text = "QTY"
MSFlexGrid1.col = 4
MSFlexGrid1.Text = "Avg. Time"
MSFlexGrid1.col = 5
MSFlexGrid1.Text = "Steel Cost"
MSFlexGrid1.col = 6
MSFlexGrid1.Text = "Gross Profit"
MSFlexGrid1.col = 7
MSFlexGrid1.Text = "Comments"
'Form1.Hide

End Sub

Private Sub List1_Click()
'Dim Index As Integer
''Form1.Clearall

'List2.Clear

'For Index = 0 To List1.ListCount - 1
'If List1.Selected(Index) Then
'Text1.Text = WORK_ORDER(Index) & "  " & PART_NUMB(Index) 'List1.List(Index)
'Text2.Text = DWG_NUMB(Index)
'Text4.Text = WANT_DATE(Index)
'If Scrap(Index) = "Y" Then Text3.Text = "Load Steel Rack  SCRAP " & Inv(Index) Else Text3.Text = "Inventory Length = " & Inv(Index)
'List2.AddItem STEEL_ASS(Index)
''If Mid(STEEL_ASS(Index), 3, 1) = "A" Then
''    List2.AddItem (Left(STEEL_ASS(Index), 2) & "B" & Mid(STEEL_ASS(Index), 4))
'' End If
''If Mid(STEEL_ASS(Index), 3, 1) = "B" Then
''    List2.AddItem (Left(STEEL_ASS(Index), 2) & "A" & Mid(STEEL_ASS(Index), 4))
'' End If

''List2.AddItem "NON-STOCK"
'InDex1 = Index
'End If

'Next Index

'List2.Selected(0) = True

End Sub
Public Sub Check_Dwg()
'checkfile = Dir("g:\acad\ablascii\" & Text2.Text & ".dta")
'checkfile1 = Dir("g:\acad\ablascii\" & Text2.Text & ".DTA")
'checkfile2 = Dir("g:\acad\ablascii\OpAscii\" & Text2.Text)
'If checkfile <> "" Or checkfile1 <> "" Or checkfile2 <> "" Then
'cutlist.Caption = "Dwg Found"
'Else
'    Msg = "Can Not Find An Exsisting Dwg !"   ' Define message.
'    Style = vbOKOnly + vbCritical + vbDefaultButton1 ' Define buttons.
'    Title = "ERROR"  ' Define title.
'    help = "DEMO.HLP"   ' Define Help file.
'    Ctxt = 1000 ' Define topic
'    response = MsgBox(Msg, Style, Title, help, Ctxt)
'    Exit Sub
'End If

'If checkfile > "" Or checkfile1 > "" Then
'    Open "g:\acad\ablascii\" & Text2.Text & ".dta" For Input As #1
'Else
'    Open "g:\acad\ablascii\OpAscii\" & Text2.Text For Input As #1
'End If

'Input #1, init, dat, dwg, PART, cust, disp, Length1, BWIDTH, THICK, Offset, offset2, BSIZE, offsetdim, Y1A, Y2A, Y3A, Y4A, eofblade, H1, H2, H3, H4, H5, H6, H7, H8, H9, H10, H11, H12, H13, H14, H15, H16, H17, H18, H19, H20, H21, H22, H23, H24, H25, boltqty, fstart, fyaxis, fstop, FPASSW, bstart, byaxis, bstop, bpassW, ex1start, ex1yaxis, ex1stop, ex1passW, ex2start, ex2yaxis, ex2stop, ex2passW, pass5start, pass5yaxis, pass5stop, pass5passW, pass6start, pass6yaxis, pass6stop, pass6passW, fstart, fpassW1, leftpass, bwidthW1, lengthB1, fpassB1, Lengthr1, bwidthB2, toltc, FRONTBEVEL, BACKBEVEL, matl
'Close #1
'Form1.Caption = "Hole Punch   " & WORK_ORDER(INDEX1) & "   DWG#  " & DWG & "        Part#  " & PART
'Form1.Text3.Text = THICK
'Form1.Text4.Text = BWIDTH
'Form1.Text5.Text = Str(Length1)
'Form1.Combo2 = BSIZE
'Form1.Text6.SetFocus
'Form1.Text6.BackColor = QBColor(14)


    'Place Code Body Here!


End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 2 Then
'PopupMenu PopDelete
'End If
End Sub

Private Sub List2_Click()
'Dim Index As Integer
'For Index = 0 To List2.ListCount - 1
'If List2.Selected(Index) Then
'Let STEEL_USED(Index) = List2.List(Index)
'Form1.Combo3.Clear

'Form1.Combo3.Text = STEEL_USED(Index)
' End If



'Next Index
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unload Form1
End Sub

Public Sub Joystick_Click()
'c6k.Write "1INFNC22-N:1INFNC23-M:JOYAXL1-2,1-1,1-0,1-0:JOYAXH1-0,1-0,1-1,,1-2:JOYVH7,10,14,2,4:JOYVL7,10,14,5,4:JOYA5,10,10,10,2:JOYAD30,100,100,10,10:" & Chr$(13)
'c6k.Write "1JOYZ.1=1:1JOYZ.2=1:1JOYEDB.1=1.18:1JOYEDB.2=1.18:1JOYCDB.2=.5:1JOYCDB.1=.5:" & Chr$(13)
'c6k.Write "JOY11101" & Chr$(13)
'Jog
Command1.Enabled = False
Command14.Enabled = False
Command15.Enabled = False
JogInput1 = 0
JogInput2 = 0
c6k.FSEnabled = True
Dim temp() As Byte
Do
    'E-STOP
temp = c6k.FastStatus
Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
     If (Last_Pcut_State = 0) Then
            c6k.Write "JOG000000:1OUTALL9,16,0:1OUTALL25,32,0:T2:1OUT.9-1:1OUT.13-1" & Chr$(13)
            Last_Pcut_State = 1
            cutlist.Label7.Caption = "E-STOP!!!"
            cutlist.Label7.Refresh
            Command1.Enabled = True
            
            Exit Sub
     End If
Else
     If (Last_Pcut_State = 1) Then
        Last_Pcut_State = 0
     End If
End If


' -------------------------------------------------------------------------- Joystick stuff -----------------------------------------
    Call CopyMemory(fsinfo, temp(0), 280)
    temp = c6k.FastStatus
    'If JogOn = True Then
    ' axis select on, limit switch off
        If CStr((fsinfo.ProgIn(1) And Input22) / Input22) = 0 And ((fsinfo.ProgIn(1) And Input8) / Input8) = 0 Then
            If JogInput1 = 0 And ((fsinfo.ProgIn(1) And Input24) / Input24) > 0 Then
                c6k.Write ("!JOG000000:" & Chr$(13))
                c6k.Write ("JOG000000:1INFNC1-5J:1INFNC2-5K:1INFNC3-2K:1INFNC4-2J:JOGA4,5,5,1,5,5:JOGAD50,99,99,99,99,15:JOGVH8,8,10,2,5,3:JOGVL8,15,10,5,5,5:JOG01010" & Chr$(13))
                JogInput1 = 1
                JogInput2 = 0
                JogInput3 = 0
                JogInput4 = 0
                JogInput5 = 0
                Label7.Caption = "    Joystick 1"
                Label7.Refresh
            End If
        End If
        ' axis select off, limit switch off (Normal x/y operation)
        If CStr((fsinfo.ProgIn(1) And Input22) / Input22) And ((fsinfo.ProgIn(1) And Input24) / Input24) > 0 Then
            If JogInput2 = 0 Then
                c6k.Write ("!JOG000000:" & Chr$(13))
                c6k.Write ("JOG000000:1INFNC2-1J:1INFNC1-1K:1INFNC3-6K:1INFNC4-6J:JOGA2,5,5,10,5,5:JOGAD50,99,99,99,99,15:JOGVH10,8,10,2,5,6:JOGVL10,15,10,5,5,10:JOG100001" & Chr$(13))
                JogInput2 = 1
                JogInput1 = 0
                JogInput3 = 0
                JogInput4 = 0
                JogInput5 = 0
                Label7.Caption = "    Joystick 2"
                Label7.Refresh
            End If
        End If
        ' axis select on, limit switch on
        If CStr((fsinfo.ProgIn(1) And Input22) / Input22) = 0 And ((fsinfo.ProgIn(1) And Input8) / Input8) > 0 Then
            If JogInput3 = 0 Then
                c6k.Write ("!JOG000000:!PSET,,0" & Chr$(13))
                c6k.Write ("JOG000000:1INFNC1-5J:1INFNC2-5K:1INFNC3-3K:1INFNC4-3J:JOGA3,5,5,10,5,5:JOGAD50,99,99,99,99,15:JOGVH8,8,10,2,5,3:JOGVL8,15,10,5,5,5:JOG001010" & Chr$(13))
                JogInput1 = 0
                JogInput2 = 0
                JogInput3 = 1
                JogInput4 = 0
                JogInput5 = 0
                Label7.Caption = "    Joystick 3"
                Label7.Refresh
                End If
            End If
        ' axis select on, limit switch on, joystick down, text16 motor position
         If CStr((fsinfo.ProgIn(1) And Input22) / Input22) = 0 And ((fsinfo.ProgIn(1) And Input8) / Input8) > 0 And ((fsinfo.ProgIn(1) And Input3) / Input3) > 0 And Val(Text16.Text) < 1 Then
            If JogInput4 = 0 Then
                c6k.Write ("!JOG000000:" & Chr$(13))
                c6k.Write ("JOG000000:1INFNC1-5K:1INFNC2-5J:1INFNC3-2K:1INFNC4-2J:JOGA3,5,5,10,5,5:JOGAD50,99,99,99,99,15:JOGVH8,8,10,2,5,3:JOGVL8,15,10,5,5,5:JOG010010" & Chr$(13))
                JogInput1 = 0
                JogInput2 = 0
                JogInput3 = 1
                JogInput4 = 1
                JogInput5 = 0
                Label7.Caption = "   Joystick 4"
                Label7.Refresh
            End If
          End If
          ' Toggle trigger for axis 3 select
          If CStr((fsinfo.ProgIn(1) And Input24) / Input24) = 0 Then
            If JogInput5 = 0 Then
                c6k.Write ("!JOG000000:" & Chr$(13))
                c6k.Write ("JOG000000:1INFNC1-5K:1INFNC2-5J:1INFNC3-3K:1INFNC4-3J:JOGA3,5,5,10,5,5:JOGAD50,99,99,99,99,15:JOGVH8,5,8,2,5,3:JOGVL8,5,8,5,5,5:JOG001010" & Chr$(13))
                JogInput1 = 0
                JogInput2 = 0
                JogInput3 = 0
                JogInput4 = 0
                JogInput5 = 1
                Label7.Caption = "   Jogging Lead Screw"
                Label7.Refresh
            End If
          End If
        If CStr((fsinfo.ProgIn(1) And Input23) / Input23) = 0 Then
            c6k.Write ("JOG000000:" & Chr$(13))
            JogInput2 = 0
            JogInput1 = 0
            JogInput3 = 0
            JogInput4 = 0
            JogInput5 = 0
            Label7.Caption = ""
            Label7.Refresh
            jogOn = False
            Exit Do
        End If
' ------------------------------------------- Motor Position --------------------------
Let X = InDex3
    XMOTOR = (fsinfo.MotorPos(1))
        Let Text14.Text = Val(XMOTOR)
        Let Text14.Text = Format((Text14.Text / 25000), "0.000")
    YMOTOR = (fsinfo.MotorPos(6))
        Let Text15.Text = Val(YMOTOR)
        Let Text15.Text = Format((Text15.Text / 25000), "0.000")
    ZMOTOR = (fsinfo.MotorPos(3))
        Let Text16.Text = Val(ZMOTOR)
        Let Text16.Text = Format((Text16.Text / 257153), "0.000")
        
        
    If JogInput1 = 1 Or JogInput2 = 1 Or JogInput3 = 1 Or JogInput4 = 1 Then
        cutlist.Text1.BackColor = &HFFFF&
        cutlist.Text1.ForeColor = QBColor(1)
        cutlist.Text1.Text = " JOYSTICK ON"
        cutlist.Text1.Refresh
        
       Maintenance1.Text8.BackColor = &HFFFF&
       Maintenance1.Text8.ForeColor = QBColor(1)
       Maintenance1.Text8.Text = " JOYSTICK ON"
       Maintenance1.Text8.Refresh
    End If
Loop

cutlist.Text1.BackColor = &H800000
cutlist.Text1.ForeColor = QBColor(1)
cutlist.Text1.Text = ""
cutlist.Text1.Refresh
Command1.Enabled = True
Command14.Enabled = True
Command15.Enabled = True
End Sub

Private Sub LiftTable_Click()
Timer1.Enabled = False
'c6k.Write "!JOG000000:" & Chr$(13)
Form1.Show
End Sub

Private Sub maintenance_Click()
Timer1.Enabled = False

Maintenance1.Show

End Sub



Private Sub OPEN_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text4.BackColor = QBColor(15)
'Text10.Text = ""
For i = 1 To 6
Text5(i).Text = ""
Text6(i).Text = ""
Text7(i).Text = ""
Text8(i).Text = ""
Text9(i).Text = ""

Check1(i - 1).value = 0
Check1(i - 1).Caption = "OFF"
Next i
WORK_ORDER$ = InputBox("ENTER WORK ORDER #" + Chr(13) + "OR WAND BARCODE", "WORK ORDER#")
'FIND DWG NUMBER IN FORM.ASC
Open "f:\mfg\FORM.ASC" For Input As #1
While WORK_ORDER$ <> WOK$
     Input #1, WOK$, PartNumber, TQTY!, PRICE!, Tc!, COMM$, DAT$, DWG$, CSCOST!, IDontKnow!
 
        If WORK_ORDER$ = WOK$ Then
            Close #1
            RunCondition5 = "YES"
            Let DWG4 = DWG$
            Text1.Text = DWG4
            For i = 1 To Len(DWG4)
            If Left(DWG4, 1) = "M" Then DWG4 = Mid(DWG4, 2)
            If Mid(DWG4, i, 1) = Chr(45) Then DWG4 = Left(DWG4, i - 1) & Mid(DWG4, i + 1)
            DICK = Mid(DWG4, i, 1)
            tom = InStr(1, DWG4, Chr(45))
            Text1.Text = PartNumber
            Text2.Text = DWG$
            If Date >= CDate(DAT$) Then Text4.BackColor = QBColor(12)
            Text4.Text = DAT$
            'Text10.Text = WOK$
            Next i
'CHECK FOR ACAD DWG
            'checkDWG = Dir("G:\ACAD\ABLDWG\" & DWG4 & ".DWG")
            'checkDWG1 = Dir("G:\DWG\" & DWG4 & ".DWG")
                'If checkDWG <> "" Or checkDWG1 <> "" Then
'VIEW DWG
                   ' Shell "C:\VOLOVIEW.EXE G:\ACAD\ABLDWG\" + DWG4 + ".DWG", vbMaximizedFocus
                'Else
                '    Msg = "Can Not Find An Exsisting Dwg !"   ' Define message.
                '    Style = vbOKOnly + vbCritical + vbDefaultButton1 ' Define buttons.
                '    Title = "ERROR"  ' Define title.
                '    help = "DEMO.HLP"   ' Define Help file.
                '    Ctxt = 1000 ' Define topic
                '    response = MsgBox(Msg, Style, Title, help, Ctxt)
                '    Exit Sub
                'End If
'CHECK FOR DTA. FILE

            checkfile = Dir("G:\ACAD\ABLASCII\" & DWG4 & ".dta")
            checkfile1 = Dir("G:\ACAD\ABLASCII\" & DWG4 & ".DTA")
            checkfile2 = Dir("G:\ACAD\ABLASCII\" & DWG4)
                If checkfile <> "" Or checkfile1 <> "" Or checkfile2 <> "" Then
                Else
                    msg = "Can Not Find An Exsisting DATA FILE !"   ' Define message.
                    Style = vbOKOnly + vbCritical + vbDefaultButton1 ' Define buttons.
                    Title = "ERROR"  ' Define title.
                    help = "DEMO.HLP"   ' Define Help file.
                    Ctxt = 1000 ' Define topic
                    response = MsgBox(msg, Style, Title, help, Ctxt)
                    Exit Sub
                End If

                If Right(DWG4, 3) = "dta" Or Right(DWG4, 3) = "DTA" Then
                    Open "G:\ACAD\ABLASCII\" & DWG4 For Input As #2
                Else
                    Open "G:\ACAD\ABLASCII\" & DWG4 & ".DTA" For Input As #2
                End If
'INPUT DATA FROM DTA FILE
                Input #2, init, DAT, DWG, PART, cust, disp, Length1, BWIDTH, THICK, Offset, offset2, BSIZE, offsetdim, Y1A, Y2A, Y3A, Y4A, eofblade, H1, H2, H3, H4, H5, H6, H7, H8, H9, H10, H11, H12, H13, H14, H15, H16, H17, H18, H19, H20, H21, H22, H23, H24, H25, boltqty, fstart, fyaxis, fstop, fpassw, bstart, byaxis, bstop, bpassw, ex1start, ex1yaxis, ex1stop, ex1passW, ex2start, ex2yaxis, ex2stop, ex2passW, pass5start, pass5yaxis, pass5stop, pass5passW, pass6start, pass6yaxis, pass6stop, pass6passW, fstart, fpassW1, leftpass, bwidthW1, lengthB1, fpassB1, Lengthr1, bwidthB2, toltc, FRONTBEVEL, BACKBEVEL, matl
                Close #2
                Text3.Text = Str(THICK) + " x" + Str(BWIDTH) + " x" + Str(Length1) + "    " + matl
                
               Let fy1 = fyaxis
               Let by1 = byaxis
               Let ex1 = ex1yaxis
               Let ex2 = ex2yaxis
               Let pass5y = pass5yaxis
               Let pass6y = pass6yaxis
               
               Let mintext = 500
               Let G = 1
               For C = 1 To 6
               If fy1 < mintext Then currmin = 1: mintext = fy1
               If by1 < mintext Then currmin = 2: mintext = by1
               If ex1 < mintext Then currmin = 3: mintext = ex1
               If ex2 < mintext Then currmin = 4: mintext = ex2
               If pass5y < mintext Then currmin = 5: mintext = pass5y
               If pass6y < mintext Then currmin = 6: mintext = pass6y
               
               If currmin = 1 Then
                    If fpassw > 0 Then
                        Text5(G).Text = "Front Pass": Text6(G).Text = fstart: Text7(G).Text = fstop: Text8(G).Text = fyaxis: Text9(G).Text = (fpassw - fyaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    fy1 = 500
               ElseIf currmin = 2 Then
                    If bpassw > 0 Then
                        Text5(G).Text = "Back Pass": Text6(G).Text = bstart: Text7(G).Text = bstop: Text8(G).Text = byaxis: Text9(G).Text = (bpassw - byaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    by1 = 500
               ElseIf currmin = 3 Then
                    If ex1passW > 0 Then
                        Text5(G).Text = "Front Extra Pass": Text6(G).Text = ex1start: Text7(G).Text = ex1stop: Text8(G).Text = ex1yaxis: Text9(G).Text = (ex1passW - ex1yaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    ex1 = 500
               ElseIf currmin = 4 Then
                    If ex2passW > 0 Then
                        Text5(G).Text = "Back Extra Pass": Text6(G).Text = ex2start: Text7(G).Text = ex2stop: Text8(G).Text = ex2yaxis: Text9(G).Text = (ex2passW - ex2yaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    ex2 = 500
              ElseIf currmin = 5 Then
                    If pass5passW > 0 Then
                        Text5(G).Text = "Pass 5": Text6(G).Text = pass5start: Text7(G).Text = pass5stop: Text8(G).Text = pass5yaxis: Text9(G).Text = (pass5passW - pass5yaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    pass5y = 500
              ElseIf currmin = 6 Then
                    If pass6passW > 0 Then
                        Text5(G).Text = "Pass 6": Text6(G).Text = pass6start: Text7(G).Text = pass6stop: Text8(G).Text = pass6yaxis: Text9(G).Text = (pass6passW - pass6yaxis)
                        Check1(G - 1).value = 1: Check1(G - 1).Caption = "ON"
                        G = G + 1
                    End If
                    pass6y = 500
              End If
              Let mintext = 500
              Next C
'Form1.Text4.BackColor = QBColor(15)
              If Val(fstart) <> Val(leftpass) Then
                     Text5(7).Text = "Left Pass": Text6(7).Text = fstart: Text7(7).Text = leftpass: Text8(7).Text = fpassW1: Text9(7).Text = (bwidthW1 - fpassW1)
                        Check1(6).value = 1: Check1(6).Caption = "ON"
              End If
              If Val(lengthB1) <> Val(Lengthr1) Then
                     Text5(8).Text = "Right Pass": Text6(8).Text = lengthB1: Text7(8).Text = Lengthr1: Text8(8).Text = fpassB1: Text9(8).Text = (bwidthB2 - fpassB1)
                        Check1(7).value = 1: Check1(7).Caption = "ON"
              End If
              If Val(FRONTBEVEL) > 0 Then
                     Text5(9).Text = "Front Bevel": Text9(9).Text = Format(FRONTBEVEL, "#.###")
                        Check1(8).value = 1: Check1(8).Caption = "ON"
              End If
              If Val(BACKBEVEL) > 0 Then
                     Text5(10).Text = "Back Bevel": Text9(10).Text = Format(BACKBEVEL, "#.###")
                        Check1(9).value = 1: Check1(9).Caption = "ON"
              End If
           If RunCondition5 = "YES" Then
             Timer2.Enabled = True
             'Label6.Caption = "Run Mode On"
'For i = 0 To 5
'    If Check1(i).Value = 1 Then
'        If i = 0 Then Check1(0).Caption = "ON LEFT"
'            If i > 0 Then
'                If Check1(i - 1).Caption = "ON LEFT" Then
'                Check1(i).Caption = "ON RIGHT"
'                Else
'                Check1(i).Caption = "ON LEFT"
'                End If
'            End If
'    End If
'Next i
           End If
           Exit Sub
        End If
       If EOF(1) Then
            Close #1
            RunCondition5 = "NO"
            MsgBox "WORK ORDER NOT FOUND!"
            'Form2.Label1.Caption "WORK ORDER# " & InString$ & " Not Found in Form.asc"
            
            Exit Sub
        End If
Wend
Close #1
End Sub

Private Sub Reset_Click()
c6k.Write ("!RESET" & Chr$(13))
End Sub

Private Sub set0_Click()


c6k.Write "JOG000000:PSET0,0,0,0,0,0" & Chr$(13)

End Sub





Private Sub speed_Click()
Form2.Show
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = (13) Then
Command6.value = True
End If

End Sub

Private Sub Text5_Click(Index As Integer)
For i = 1 To 10
If i <> Index Then
    Check1(i - 1).value = 0: Check1(i - 1).Caption = "OFF"
    Text5(i).BackColor = QBColor(15)
Else
    If Text5(i).Text > "" Then
        Check1(i - 1).value = 1: Check1(i - 1).Caption = "ON"
        Text5(i).BackColor = QBColor(14)
    End If
End If
Next i
'For i = 0 To 5
'    If Check1(i).Value = 1 Then
'        If i = 0 Then Check1(0).Caption = "ON LEFT"
'            If i > 0 Then
'                If Check1(i - 1).Caption = "ON LEFT" Then
'                Check1(i).Caption = "ON RIGHT"
'                Else
'                Check1(i).Caption = "ON LEFT"
'                End If
'            End If
'    End If
'Next i
End Sub

Private Sub Text6_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = (13) Then
Text7(Index).SetFocus
End If
End Sub


Private Sub Text7_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = (13) Then
Text8(Index).SetFocus
End If

End Sub


Private Sub Text8_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = (13) Then
Text9(Index).SetFocus
End If

End Sub


Private Sub Text9_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = (13) Then
    If Index < 10 Then
        Text6(Index + 1).SetFocus
    Else
        Text6(1).SetFocus
    End If
    
    
    
    'If Check1(Index).Value = 0 Then
    '    Check1(Index).Caption = "OFF": Text5(Index + 1).BackColor = QBColor(15)
    'Else
        If Text5(Index).Text > "" And Val(Text9(Index).Text) > 0 Then
            Check1(Index - 1).Caption = "ON": Check1(Index - 1).value = 1: Text5(Index).BackColor = QBColor(14)
        'Else
        '    Check1(Index).Value = 0
        End If
    'End If
    'For i = 0 To 5
    '    If Check1(i).Value = 1 Then
    '    If i = 0 Then Check1(0).Caption = "ON LEFT"
    '            If i > 0 Then
    '                If Check1(i - 1).Caption = "ON LEFT" Then
    '                    Check1(i).Caption = "ON RIGHT"
    '                Else
    '                    Check1(i).Caption = "ON LEFT"
    '                End If
    '            End If
    '    End If
    'Next i
    
End If
End Sub


Private Sub Timer1_Timer()
c6k.FSEnabled = True
Dim temp() As Byte
    
    'E-STOP
    temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
     If (Last_Pcut_State = 0) Then
            c6k.Write "JOG000000:1OUTALL9,16,0:1OUTALL25,32,0:T1.5:1OUT.9-1:1OUT.13-1" & Chr$(13)
            Last_Pcut_State = 1
            Label7.Caption = "E-STOP!!!"
            Label7.Refresh
            Timer2.Enabled = True
            EStopPos = ""
     End If
Else
     If (Last_Pcut_State = 1) Then
        Last_Pcut_State = 0
       ' Label7.Caption = ""
       ' Label7.Refresh
     End If
End If
 

temp = c6k.FastStatus
Call CopyMemory(fsinfo, temp(0), 280)
If CStr((fsinfo.AxisStatus(1) And joystick) / joystick) > 0 Then
             Label7.Caption = "joystick on"
            Label7.Refresh
            'Exit Do
        Else
            If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
            Label7.Caption = "JOYSTICK OFF"
            Label7.Refresh
            End If
End If
  '' ******* Update Motor Position ****************
        temp = c6k.FastStatus                  'get fast status information
        Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure

        YMOTOR = (fsinfo.MotorPos(6))
            Let Text16.Text = Val(YMOTOR)
            Let Text16.Text = Format((Text16.Text / 26550), "0.000")
                        
        XMOTOR = (fsinfo.MotorPos(2))
            Let Text17.Text = Val(XMOTOR)
            Let Text17.Text = Format((Text17.Text / 26550), "0.000")
            

        UDMOTOR = (fsinfo.MotorPos(3))
            Let Text14.Text = Val(UDMOTOR)
            Let Text14.Text = Format((Text14.Text / 500000), "0.000")
            
        OSSMOTOR = (fsinfo.MotorPos(6))
            Let Text15.Text = Val(OSSMOTOR)
            Let Text15.Text = Format((Text15.Text / 25000), "0.000")
                    
        ROTATION = (fsinfo.MotorPos(5))
            Let Text20.Text = Val(ROTATION)
            Let Text20.Text = Format(((Text20.Text / 4432)), "0.000")
        
 
        XSpeedSHOW = (fsinfo.MotorVel(2))
            Let Text18.Text = Val(XSpeedSHOW)
            Let Text18.Text = Format((Text18.Text / 660), "0.000")
 

       
        cutlist.Text14.Refresh
        cutlist.Text15.Refresh
        cutlist.Text16.Refresh
        cutlist.Text17.Refresh
        cutlist.Text18.Refresh
        cutlist.Text12.Refresh
 
End Sub


Private Sub Timer2_Timer()
Dim temp() As Byte
 temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
    If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
        If (Last_Pcut_State = 0) Then
            c6k.Write "!COMEXS0:" & Chr$(13)
            c6k.Write "1OUTALL9,16,0:1OUTALL25,32,0:1OUT.9-1:1OUT.13-1:JOG0000:" & Chr$(13)
            Last_Pcut_State = 1
            Label7.Caption = "E-STOP!!!!!"
            Label7.Refresh
            Timer2.Enabled = True
            EStopPos = ""
            Exit Sub
        End If
    Else
        If (Last_Pcut_State = 1) Then
            Last_Pcut_State = 0
            'Label7.Caption = ""
            Label7.Refresh
        End If
    End If
         
         
         
         
         
         
         
         
         
         If StartCount5 = 0 Then
            Let StartCount5 = 1
            Let StartTimer4 = Timer
        End If
    Let endtime = Timer
    Let elptime = (500 - (endtime - StartTimer4))
    Let Text11.Text = Format(elptime, "###")
         If elptime < 0 Then
            If Dir("F:\BARCODE\temp6.tmp") = "" Then
                'FileExists = False
                Label21.Caption = "Work Order is Not Open"
                RunCondition5 = "TimeOut"
                Timer2.Enabled = False
            Exit Sub
        Else
            Open "F:\BARCODE\temp6.tmp" For Input As #2
            Input #2, temp6, DAT
            Close #2
            
            Open "F:\BARCODE\" + temp6 + "6.tmp" For Input As #2
            Input #2, InString1, STIME!, ETIME!, TTime!, Com2$
            Close #2
            
            Let ETIME! = Timer: Let TTime! = ((ETIME! - STIME!) / 60) + TTime!
            Open "F:\BARCODE\" + temp6 + "6.tmp" For Output As #2
            Write #2, InString1, STIME!, ETIME!, TTime!, Com2$
            Close #2
            
            Let STIME! = STIME! / 3600: Let ETIME! = ETIME! / 3600:
            Label21.Caption = "Timed-Out RunMode - OFF"
            Label17.Caption = Format(Now, "SHORT TIME")
'*************
If Com2$ = "TOPS" Or InString2 = "BEVEL" Then
Sta$ = "Line 5&4"
ElseIf Com2$ = "" Then
Sta$ = "6"
Else
Sta$ = "6"
Typ$ = "ganged"
End If

Open "f:\MFG\Bar_Data.dta" For Append As #1
Write #1, InString1, Sta$, "Timed-Out", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss")
    If Typ$ = "ganged" Then
    Write #1, Com2$, Sta$, "Timed-Out", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss")
    End If
Close #1
'*********
  
Let InString1 = ""
            
           ' Kill "F:\BARCODE\temp3.tmp"
            '********** relay off **********
             'ULStat% = cbDBitOut%(0, 10, 2, 0)
             '  If ULStat% <> 0 Then Stop
             StartCount5 = 0
             Timer2.Enabled = False
            Let RunCondition5 = "TimeOut"
            BAR3.Show
       '  End If
       End If
'Else
StartCount5 = 0
End If


End Sub

Public Sub BARCODE_READER()
 
If MSComm1.PortOpen = False Then
   ' Use COM1.
    MSComm1.CommPort = 1
    ' 9600 baud, no parity, 8 data, and 1 stop bit.
    MSComm1.Settings = "9600,N,8,1"
    
    ' Open the port.
    MSComm1.PortOpen = True
    ' Send the attention command to the modem.
    End If
    Do
        
        
        MSComm1.output = "a" & Chr$(13)
        ' Wait for data to come back to the serial port.
        Dummy = DoEvents()
            For j = 1 To 1000
Beep
            Next j
        ' Read the "OK" response data in the serial port.
'        If Form4!Label1.Caption = "CLEAR" Then
'        If RunCondition3 <> "OK" Then
'        MSComm1.PortOpen = False
'        Exit Sub
'        End If
If MSComm1.PortOpen = False Then
Exit Sub
End If
        InString$ = MSComm1.Input
        ' Close the serial port.
    
    Let SEL = INKEY$
    If SEL <> "" Then End
    Loop Until Len(InString$) >= 3 Or KeyAscii = Chr$(27)
    MSComm1.PortOpen = False
'   Form4!Label1.Caption = ""
   If Len(InString$) >= 3 Then
      
   If Len(InString$) = 7 Then
        Let InString$ = Mid(InString$, 1, 5) ', 5)
   End If
   If Len(InString$) = 8 Then
        Let InString$ = Mid(InString$, 1, 6) ', 5)
   End If
  'Label7.Caption = InString$
   Else: InString$ = ""
   End If
'Printer.Print InString$ & "  " & Format(Now) & "ABCDEFGHIJKLMNOPQRSTUVWXYZ 12345678910"
'Printer.EndDoc
End Sub

Public Sub CheckForm1()
On Error GoTo errorhand
Open "f:\mfg\FORM.ASC" For Input As #5
While InString$ <> WOK$
    Input #5, WOK$, PartNumber, TQTY!, PRICE!, Tc!, COMM$, DAT$, DWG$, CSCOST!, IDontKnow!
        If InString$ = WOK$ Then
            Close #5
            RunCondition5 = "YES"
            Let DWG4 = DWG$
            Exit Sub
        End If
        If EOF(5) Then
            Close #5
            RunCondition5 = "NO"
            Form2.Show 1
            'Form2.Label1.Caption "WORK ORDER# " & InString$ & " Not Found in Form.asc"
            
            Exit Sub
        End If
Wend
Close #5

Exit Sub
errorhand:
If Err.Number = 70 Then
Form7.Show 1
Resume 0
End If
End Sub

Public Sub Finish_One4()
Dim DAT As String
On Error GoTo errorhand
'If RunCondition5 = "TimeOut" Then
'        Form3.Show
'        Exit Sub
'End If
StartCount5 = 0
 If Dir("F:\BARCODE\temp6.tmp") = "" Then
            'FileExists = False
             MsgBox ("No Active Work Order "), , ("ERROR")
             Exit Sub
 End If
            Let Qty! = Val(InputBox("ENTER QTY FINISHED", "Qty", 0, 4000, 1500))
            If Qty! = 0 Then Exit Sub
                
            Open "f:\mfg\FORM.ASC" For Input As #5
            Do Until EOF(5)
            Input #5, WOK$, PartNumber, TQTY!, PRICE!, Tc!, COMM$, DAT$, DWG$, CSCOST!, SBT_AvgTime!
               If InString$ = WOK$ Then
                  If TQTY! < Qty! Then
                    MsgBox ("You Entered More Than The Work Order QTY"), , ("ERROR")
                  Close #5
                  Exit Sub
                 End If
                Exit Do
                Else
                Let WOK$ = "": PartNumber = "": TQTY! = 0: PRICE! = 0.00001: Tc! = 0: COMM$ = "": DAT$ = "": DWG$ = "no_dwg": CSCOST! = 0
               End If
             Loop
             Close #5
            If PartNumber = "" Then Let PartNumber = InString$ & " Override"
            Open "F:\BARCODE\temp6.tmp" For Input As #2
            Input #2, temp6, DAT
            Close #2
            
            Open "F:\BARCODE\" + temp6 + "6.tmp" For Input As #2
            Input #2, InString1, STIME!, ETIME!, TTime!, Com2$
            Close #2
            
If RunCondition5 = "TimeOut" Then
    'RunCondition5 = ""
Else
    Let ETIME! = Timer: Let TTime! = ((ETIME! - STIME!) / 60) + TTime!
End If
            Kill "F:\BARCODE\" + temp6 + "6.tmp"
            Let STIME! = STIME! / 3600: Let ETIME! = ETIME! / 3600:
            Label21.Caption = "RunMode - OFF"
            Label17.Caption = Format(Now, "SHORT TIME")
            
            
            Kill "F:\BARCODE\temp6.tmp"
            '********** relay off **********
            ' ULStat% = cbDBitOut%(0, 10, 2, 0)
            '   If ULStat% <> 0 Then Stop
             Timer2.Enabled = False
            'Text4.Visible = False
            'Label19.Visible = False
             Label21.Caption = "Press Start  or  Ganged  "

Command6.Enabled = True
Command8.Enabled = True
Command13.Enabled = True




Let STYPE$ = "0"


'*********************CALC
'T.C COST/LB        4.81
'INCHES/LB         28.00
'T.C. COST/SQ"     $0.172
'MACH.P/DAY     $1344.00
'LABOR P/DAY    $1888.00
'TOTAL P/DAY    $3232.00
'CAP. UTIL%        80.00
'CAP.DAILY COST $3232.00
'PLANT COST/HR    404.00
'*************************************************************************

Let TCCOST! = 0.171821
Let HOCOST! = 404!
Let TTime! = TTime! / 60
Let AvgTime! = Format$((TTime! / Qty!), "CURRENCY")
Let TCCOST! = Format$((TCCOST! * Tc!), "CURRENCY")
Let HORATE! = 0.5


Let CONVCOST! = Format$((HOCOST! * AvgTime! * HORATE!), "CURRENCY")
Let UNITCOST! = Format$((TCCOST! + CONVCOST! + STLCOST!), "CURRENCY")
Let UNITPROFIT! = Format$((PRICE! - UNITCOST!), "CURRENCY")
'Let GROSSPROFIT! = Format$(((UNITPROFIT! / PRICE!) * 100), "CURRENCY")
Let TSALES! = Qty! * PRICE!
Let TTC! = Tc! * Qty!
Let TTCCOST! = TCCOST! * Qty!
Let TSTLCOST! = STLCOST! * Qty!
Let TCONVCOST! = CONVCOST! * Qty!
Let TOTALCOST! = UNITCOST! * Qty!
Let TOTALPROFIT! = UNITPROFIT! * Qty!
Let COM$ = COM1$ + Com2$ + Com3$ + COM4$ + COM5$ + COM6$ + COM7$

Let diff = SBT_AvgTime! - AvgTime!
If AvgTime! < SBT_AvgTime! Then Let diffcom = " Faster  "
If AvgTime! > SBT_AvgTime! Then Let diffcom = " Slower  "
MSFlexGrid1.AddItem InString$ & Chr(9) & PartNumber & Chr(9) & "4" & Chr(9) & Qty! & Chr(9) & SBT_AvgTime! & Chr(9) & AvgTime! & Chr(9) & Format((diff * 60), "####0.00") & diffcom & Chr(9) & COM$, 1
Open "F:\MFG\Bar_Data.dta" For Append As #1
If RunCondition5 = "TimeOut" Then
    RunCondition5 = ""
    Write #1, InString$, "6", "Finish", Format(Now, "yy-mm-dd"), "Timed Out", PartNumber, Qty!, AvgTime!, STLCOST!, GROSSPROFIT! & "%", COM$
Else
    Write #1, InString$, "6", "Finish", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss"), PartNumber, Qty!, AvgTime!, STLCOST!, GROSSPROFIT! & "%", COM$
End If
Close #1

Let Sta$ = "6"

19495 Open "F:\MFG\GROSSPIT.GPA" For Append As #1
Write #1, WOK$, PartNumber, TQTY!, Qty!, Sta$, Date$, PRICE!, Tc!, STLCOST!, AvgTime!, TCCOST!, CONVCOST!, UNITCOST!, UNITPROFIT!, GROSSPROFIT!, COM$, TTime!, TSALES!, TTC!, TTCCOST!, TSTLCOST!, TCONVCOST!, TOTALCOST!, TOTALPROFIT!, DAT$, DWG$, CSCOST!
Close #1

Open "F:\mfg\GARY.GPA" For Append As #1
Write #1, WOK$, PartNumber, TQTY!, Qty!, Sta$, Date$, PRICE!, Tc!, STLCOST!, AvgTime!
Close #1
Open "F:\mfg\GOAL.GPA" For Append As #1
Write #1, WOK$, PartNumber, TQTY!, Qty!, Sta$, Date$, PRICE!, Tc!, STLCOST!, AvgTime!
Close #1

Let TTime! = 0: Let ATSALES! = 0: Let ATTC! = 0: Let BTTIME! = 0: Let BTSALES! = 0: Let BTTC! = 0: Let CTTIME! = 0: Let CTSALES! = 0: Let CTTC! = 0: Let TUNITPROFIT! = 0
Let TTC! = 0
Let COM1$ = "": Let Com2$ = "": Let Com3$ = "": Let COM4$ = "": Let COM5$ = "": Let COM6$ = "": Let COM7$ = ""
Let STIME! = 0: Let ETIME! = 0: Let Qty! = 0: Let TTime! = 0: Let C$ = "": Let B$ = "": Let A$ = ""
Let WOK$ = ""
Let InString1 = ""

Exit Sub
errorhand:
If Err.Number = 70 Then
Form7.Show 1
Resume 0
Else
Resume Next
End If
End Sub

Public Sub Finish_Ganged4()
Dim DAT As String
Dim AvgTime2 As Single
Dim Time_Per_Inch As Single
Dim Tc2 As Single
Dim AvgTime1 As Single
Dim TC1 As Single
Dim Toltal_TC12 As Single
On Error GoTo errorhand
'If RunCondition5 = "TimeOut" Then
'        Form3.Show
'        Exit Sub
'End If
StartCount5 = 0
 If Dir("F:\BARCODE\temp6.tmp") = "" Then
            'FileExists = False
             MsgBox ("No Active Work Order "), , ("ERROR")
             Exit Sub
 End If
 Open "F:\BARCODE\temp6.tmp" For Input As #2
            Input #2, temp6, DAT
            Close #2
 Open "F:\BARCODE\" + temp6 + "6.tmp" For Input As #2
            Input #2, InString1, STIME!, ETIME!, TTime!, InString2
            Close #2
    
 If RunCondition5 = "TimeOut" Then
   'RunCondition5 = ""
Else
    Let ETIME! = Timer: Let TTime! = ((ETIME! - STIME!) / 60) + TTime!
End If
            Kill "F:\BARCODE\" + temp6 + "6.tmp"
                        
            Let Qty1! = Val(InputBox("ENTER QTY FINISHED (W.O.# " & InString1 & ")", "Qty", 0, 4000, 1500))
            If Qty1! = 0 Then Exit Sub
                
            Let Qty2! = Val(InputBox("ENTER QTY FINISHED (W.O.# " & InString2 & ")", "Qty", 0, 4000, 1500))
            If Qty2! = 0 Then Exit Sub
                
            Open "F:\MFG\FORM.ASC" For Input As #5
            Do Until EOF(5)
            Input #5, WOK$, PartNumber, TQTY!, PRICE!, Tc!, COMM$, DAT$, DWG$, CSCOST!, SBT_AvgTime!
               If InString1 = WOK$ Then
                  If TQTY! < Qty1! Then
                    MsgBox ("You Entered More Than The Work Order QTY"), , ("ERROR")
                  Close #5
                  Exit Sub
                 End If
                WOK1$ = WOK$: PartNumber1 = PartNumber: TQTY1! = TQTY!: PRICE1! = PRICE!
                TC1 = Tc!: COMM1$ = COMM$: DAT1$ = DAT$: DWG1$ = DWG$: CSCOST1! = CSCOST!: SBT_AvgTime1! = SBT_AvgTime!
                Let WOK$ = "": PartNumber = "": TQTY! = 0: PRICE! = 0.00001: Tc! = 0: COMM$ = "": DAT$ = "": DWG$ = "no_dwg": CSCOST! = 0
               Exit Do
                Else
                Let WOK$ = "": PartNumber = "": TQTY! = 0: PRICE1! = 0.00001: Tc! = 0: COMM$ = "": DAT$ = "": DWG1$ = "no_dwg": CSCOST! = 0
               End If
             Loop
             Close #5
      
            Open "F:\MFG\FORM.ASC" For Input As #6
            Do Until EOF(6)
            Input #6, WOK$, PartNumber, TQTY!, PRICE!, Tc!, COMM$, DAT$, DWG$, CSCOST!, SBT_AvgTime!
               If InString2 = WOK$ Then
                  If TQTY! < Qty2! Then
                    MsgBox ("You Entered More Than The Work Order QTY"), , ("ERROR")
                  Close #6
                  Exit Sub
                 End If
                WOK2$ = WOK$: PartNumber2 = PartNumber: TQTY2! = TQTY!: PRICE2! = PRICE!
                Tc2 = Tc!: COMM2$ = COMM$: DAT2$ = DAT$: DWG2$ = DWG$: CSCOST2! = CSCOST!: SBT_AvgTime2! = SBT_AvgTime!
                Let WOK$ = "": PartNumber = "": TQTY! = 0: PRICE! = 0.00001: Tc! = 0: COMM$ = "": DAT$ = "": DWG$ = "no_dwg": CSCOST! = 0
               Exit Do
                Else
                Let WOK$ = "": PartNumber = "": TQTY! = 0: PRICE2! = 0.00001: Tc! = 0: COMM$ = "": DAT$ = "": DWG2$ = "no_dwg": CSCOST! = 0
               End If
             Loop
             Close #6
            
            
            If PartNumber1 = "" Then Let PartNumber1 = InString1 & " Override"
            If PartNumber2 = "" Then Let PartNumber2 = InString2 & " Override"
            
            'Let ETIME! = Timer: Let TTime! = ((ETIME! - STIME!) / 60) + TTime!
            
            Let STIME! = STIME! / 3600: Let ETIME! = ETIME! / 3600:
            Label21.Caption = "RunMode - OFF"
            Label17.Caption = Format(Now, "SHORT TIME")
            
            
            Kill "F:\BARCODE\temp6.tmp"
            '********** relay off **********
            ' ULStat% = cbDBitOut%(0, 10, 2, 0)
            '   If ULStat% <> 0 Then Stop
             Timer2.Enabled = False
            'Text4.Visible = False
            'Label19.Visible = False
             Label21.Caption = "Press Start  or  Ganged  "

Command6.Enabled = True
Command8.Enabled = True
Command13.Enabled = True


'et Label6.Caption = DWG1$ & " " & DWG2$

Let STYPE1$ = "0"
Let STYPE2$ = "0"

'*********************CALC
'T.C COST/LB        4.81
'INCHES/LB         28.00
'T.C. COST/SQ"     $0.172
'MACH.P/DAY     $1344.00
'LABOR P/DAY    $1888.00
'TOTAL P/DAY    $3232.00
'CAP. UTIL%        80.00
'CAP.DAILY COST $3232.00
'PLANT COST/HR    404.00
'*************************************************************************

Let TCCOST! = 0.171821
Let HOCOST! = 404!
Let TTime! = TTime! / 60
'******** NEW CALC **************
Let Toltal_TC12 = (TC1 * Qty1!) + (Tc2 * Qty2!)
Let Time_Per_Inch = TTime! / Toltal_TC12
'*************
Let AvgTime1 = Format((Time_Per_Inch * TC1), "CURRENCY")
'Let AvgTime1 = (Time_Per_Inch * TC1)
'Let AvgTime2 = (Time_Per_Inch * TC2)
Let AvgTime2 = Format((Time_Per_Inch * Tc2), "CURRENCY")

Let TCCOST1! = Format$((TCCOST! * TC1), "CURRENCY")
Let TCCOST2! = Format$((TCCOST! * Tc2), "CURRENCY")

Let HORATE! = 0.5


Let CONVCOST1! = Format$((HOCOST! * AvgTime1 * HORATE!), "CURRENCY")
Let CONVCOST2! = Format$((HOCOST! * AvgTime2 * HORATE!), "CURRENCY")
Let UNITCOST1! = Format$((TCCOST1! + CONVCOST1! + STLCOST1!), "CURRENCY")
Let UNITCOST2! = Format$((TCCOST2! + CONVCOST2! + STLCOST2!), "CURRENCY")
Let UNITPROFIT1! = Format$((PRICE1! - UNITCOST1!), "CURRENCY")
Let UNITPROFIT2! = Format$((PRICE2! - UNITCOST2!), "CURRENCY")
Let GROSSPROFIT1! = Format$(((UNITPROFIT1! / PRICE1!) * 100), "CURRENCY")
Let GROSSPROFIT2! = Format$(((UNITPROFIT2! / PRICE2!) * 100), "CURRENCY")
Let TSALES1! = Qty1! * PRICE1!
Let TSALES2! = Qty2! * PRICE2!
Let TTC1 = TC1 * Qty1!
Let TTC2 = Tc2 * Qty2!
Let TTCCOST1! = TCCOST1! * Qty1!
Let TTCCOST2! = TCCOST2! * Qty2!
Let TSTLCOST1! = STLCOST1! * Qty1!
Let TSTLCOST2! = STLCOST2! * Qty2!
Let TCONVCOST1! = CONVCOST1! * Qty1!
Let TCONVCOST2! = CONVCOST2! * Qty2!
Let TOTALCOST1! = UNITCOST1! * Qty1!
Let TOTALCOST2! = UNITCOST2! * Qty2!
Let TOTALPROFIT1! = UNITPROFIT1! * Qty1!
Let TOTALPROFIT2! = UNITPROFIT2! * Qty2!
Let COM11$ = COM1$ + Com2$ + Com3$ + COM4$ + COM51$ + COM6$ + COM7$
Let COM12$ = COM1$ + Com2$ + Com3$ + COM4$ + COM52$ + COM6$ + COM7$

Let diff1 = SBT_AvgTime1! - AvgTime1!
If AvgTime1! < SBT_AvgTime1! Then Let diffcom1 = " Faster  "
If AvgTime1! > SBT_AvgTime1! Then Let diffcom1 = " Slower  "

Let diff2 = SBT_AvgTime2! - AvgTime2!
If AvgTime2! < SBT_AvgTime2! Then Let diffcom2 = " Faster  "
If AvgTime2! > SBT_AvgTime2! Then Let diffcom2 = " Slower  "

MSFlexGrid1.AddItem InString1 & Chr(9) & PartNumber1 & Chr(9) & "4" & Chr(9) & Qty1! & Chr(9) & SBT_AvgTime1 & Chr(9) & AvgTime1! & Chr(9) & (diff1 * 60) & diffcom1 & Chr(9) & COM11$, 1
MSFlexGrid1.AddItem InString2 & Chr(9) & PartNumber2 & Chr(9) & "4" & Chr(9) & Qty2! & Chr(9) & SBT_AvgTime2 & Chr(9) & AvgTime2! & Chr(9) & (diff2 * 60) & diffcom2 & Chr(9) & COM12$, 1
Open "F:\MFG\Bar_Data.dta" For Append As #1
If RunCondition5 = "TimeOut" Then
    RunCondition5 = ""
    Write #1, InString1, "5", "Finish", Format(Now, "yy-mm-dd"), "Timed Out", PartNumber1, Qty1!, AvgTime1, STLCOST1!, GROSSPROFIT1! & "%", COM11$
    Write #1, InString2, "5", "Finish", Format(Now, "yy-mm-dd"), "Timed Out", PartNumber2, Qty2!, AvgTime2, STLCOST2!, GROSSPROFIT2! & "%", COM12$
Else
    Write #1, InString1, "5", "Finish", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss"), PartNumber1, Qty1!, AvgTime1, STLCOST1!, GROSSPROFIT1! & "%", COM11$
    Write #1, InString2, "5", "Finish", Format(Now, "yy-mm-dd"), Format$(Time, "hh:mm:ss"), PartNumber2, Qty2!, AvgTime2, STLCOST2!, GROSSPROFIT2! & "%", COM12$
End If
Close #1

Let Sta$ = "5"

 
 Open "F:\MFG\GROSSPIT.GPA" For Append As #1
Write #1, WOK1$, PartNumber1, TQTY1!, Qty1!, Sta$, Date$, PRICE1!, TC1, STLCOST1!, AvgTime1, TCCOST1!, CONVCOST1!, UNITCOST1!, UNITPROFIT1!, GROSSPROFIT1!, COM11$, TTime1!, TSALES1!, TTC1, TTCCOST1!, TSTLCOST1!, TCONVCOST1!, TOTALCOST1!, TOTALPROFIT1!, DAT1$, DWG1$, CSCOST1!
Close #1
 Open "F:\MFG\GROSSPIT.GPA" For Append As #1
Write #1, WOK2$, PartNumber2, TQTY2!, Qty2!, Sta$, Date$, PRICE2!, Tc2, STLCOST2!, AvgTime2, TCCOST2!, CONVCOST2!, UNITCOST2!, UNITPROFIT2!, GROSSPROFIT2!, COM12$, TTIME2!, TSALES2!, TTC2, TTCCOST2!, TSTLCOST2!, TCONVCOST2!, TOTALCOST2!, TOTALPROFIT2!, DAT2$, DWG2$, CSCOST2!
Close #1

Open "F:\mfg\GARY.GPA" For Append As #1
Write #1, WOK1$, PartNumber1, TQTY1!, Qty1!, Sta$, Date$, PRICE1!, TC1!, STLCOST1!, AvgTime1!
Close #1

Open "F:\mfg\GARY.GPA" For Append As #1
Write #1, WOK2$, PartNumber2, TQTY2!, Qty2!, Sta$, Date$, PRICE2!, Tc2!, STLCOST2!, AvgTime2!
Close #1

Open "F:\mfg\GOAL.GPA" For Append As #1
Write #1, WOK1$, PartNumber1, TQTY1!, Qty1!, Sta$, Date$, PRICE1!, TC1!, STLCOST1!, AvgTime1!
Close #1

Open "F:\mfg\GOAL.GPA" For Append As #1
Write #1, WOK2$, PartNumber2, TQTY2!, Qty2!, Sta$, Date$, PRICE2!, Tc2!, STLCOST2!, AvgTime2!
Close #1

Let ATTIME! = 0: Let ATSALES! = 0: Let ATTC! = 0: Let BTTIME! = 0: Let BTSALES! = 0: Let BTTC! = 0: Let CTTIME! = 0: Let CTSALES! = 0: Let CTTC! = 0: Let TUNITPROFIT! = 0
Let TTC! = 0
Let COM1$ = "": Let Com2$ = "": Let Com3$ = "": Let COM4$ = "": Let COM5$ = "": Let COM6$ = "": Let COM7$ = ""
Let STIME! = 0: Let ETIME! = 0: Let Qty! = 0: Let TTime! = 0: Let C$ = "": Let B$ = "": Let A$ = ""
Let WOK$ = ""
Let InString1 = ""

Exit Sub
errorhand:
If Err.Number = 70 Then
Form7.Show 1
Resume 0
Else
Resume Next
End If

End Sub

Private Sub water_Click()

c6k.Write "1OUT.13-1:1OUT.15-1" & Chr$(13)
End Sub
