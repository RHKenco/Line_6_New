VERSION 5.00
Begin VB.Form Maintenance1 
   BackColor       =   &H00FF0000&
   Caption         =   "Maintenance"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10200
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
   ScaleHeight     =   8985
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   8520
      TabIndex        =   56
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   7200
      TabIndex        =   55
      Text            =   "Text11"
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   7200
      TabIndex        =   54
      Text            =   "Text10"
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   7200
      TabIndex        =   53
      Text            =   "Text9"
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   3960
      TabIndex        =   42
      Text            =   "Text8"
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CUT LIST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   40
      Top             =   3960
      Width           =   972
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HOME OSS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   39
      Top             =   4680
      Width           =   972
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HOME ROTATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   38
      Top             =   4080
      Width           =   972
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF0000&
      Caption         =   "JoyStick"
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
      Height          =   3132
      Left            =   3840
      TabIndex        =   28
      Top             =   1800
      Width           =   2652
      Begin VB.CommandButton Command12 
         Caption         =   "JOYSTICK "
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   1452
      End
      Begin VB.TextBox Text5 
         Height          =   372
         Left            =   960
         TabIndex        =   30
         Text            =   "Text5"
         Top             =   1920
         Width           =   1332
      End
      Begin VB.TextBox Text7 
         Height          =   372
         Left            =   960
         TabIndex        =   29
         Text            =   "Text7"
         Top             =   2640
         Width           =   1332
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Up-Down"
         ForeColor       =   &H8000000E&
         Height          =   252
         Index           =   1
         Left            =   960
         TabIndex        =   32
         Top             =   2400
         Width           =   1092
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Left-Right"
         ForeColor       =   &H8000000E&
         Height          =   252
         Index           =   0
         Left            =   960
         TabIndex        =   31
         Top             =   1680
         Width           =   1092
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF0000&
      Caption         =   "Encoder"
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
      Height          =   852
      Left            =   3840
      TabIndex        =   26
      Top             =   600
      Width           =   2652
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   480
         TabIndex        =   27
         Text            =   "Text2"
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF0000&
      Caption         =   "Jog Motors"
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
      Height          =   3012
      Left            =   6720
      TabIndex        =   21
      Top             =   600
      Width           =   3132
      Begin VB.CommandButton Command7 
         Caption         =   "Stop"
         Height          =   255
         Left            =   2160
         TabIndex        =   35
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Start"
         Height          =   255
         Left            =   2160
         TabIndex        =   34
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "PSET"
         Height          =   255
         Left            =   2160
         TabIndex        =   33
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   372
         Left            =   1560
         TabIndex        =   25
         Text            =   "Text6"
         Top             =   720
         Width           =   1212
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Text            =   "Text4"
         Top             =   1680
         Width           =   1212
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Text            =   "Text3"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Up-Down"
         ForeColor       =   &H8000000E&
         Height          =   252
         Index           =   3
         Left            =   240
         TabIndex        =   37
         Top             =   2160
         Width           =   1092
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Left-Right"
         ForeColor       =   &H8000000E&
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   36
         Top             =   1320
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Use Arrow Keys to Jog Motors"
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
         Height          =   732
         Index           =   16
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   2412
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF0000&
      Caption         =   "InPut"
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
      Height          =   3495
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   3495
      Begin VB.TextBox Text12 
         Height          =   495
         Left            =   2640
         TabIndex        =   57
         Text            =   "Text12"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF0000&
         Caption         =   "4 JOY BACK"
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
         Index           =   3
         Left            =   240
         TabIndex        =   61
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF0000&
         Caption         =   "3 JOY FRONT"
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
         Index           =   2
         Left            =   240
         TabIndex        =   60
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF0000&
         Caption         =   "2 JOY RIGHT"
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
         Index           =   1
         Left            =   240
         TabIndex        =   59
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF0000&
         Caption         =   "1 JOY LEFT"
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
         Index           =   0
         Left            =   240
         TabIndex        =   58
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "24 = JOYSTICK PAUSE"
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
         Index           =   15
         Left            =   240
         TabIndex        =   17
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "23 = JOYSTICK RELEASE"
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
         Index           =   14
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "22 = JOYSTICK AXIS SELECT"
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
         Index           =   13
         Left            =   240
         TabIndex        =   15
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "21 = ROTATE PROX SW"
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
         Index           =   12
         Left            =   240
         TabIndex        =   14
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "20 = E-STOP"
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
         Index           =   11
         Left            =   240
         TabIndex        =   13
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "19 = WATER FLOW SW"
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
         Index           =   10
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "18 = OSS LIMIT SW"
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
         Index           =   9
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "17 = TC HEAD LIMIT SW"
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
         Index           =   8
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "OutPut"
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
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   3492
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
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "32 = OPEN"
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
         Height          =   252
         Index           =   15
         Left            =   120
         TabIndex        =   50
         Top             =   4080
         Width           =   3252
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "31 = OPEN"
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
         Height          =   252
         Index           =   14
         Left            =   120
         TabIndex        =   49
         Top             =   3840
         Width           =   3252
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "30 = OPEN"
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
         Height          =   252
         Index           =   13
         Left            =   120
         TabIndex        =   48
         Top             =   3600
         Width           =   3252
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "29 = LIFT TILT DOWN"
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
         Height          =   252
         Index           =   12
         Left            =   120
         TabIndex        =   47
         Top             =   3360
         Width           =   3252
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "28 = LIFT UP"
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
         Height          =   252
         Index           =   11
         Left            =   120
         TabIndex        =   46
         Top             =   3120
         Width           =   3252
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "27 = BAD"
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
         Height          =   252
         Index           =   10
         Left            =   120
         TabIndex        =   45
         Top             =   2880
         Width           =   3252
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "26 = LIFT TILT UP"
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
         Height          =   252
         Index           =   9
         Left            =   120
         TabIndex        =   44
         Top             =   2640
         Width           =   3252
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "25 =OPEN"
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
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   43
         Top             =   2400
         Width           =   3252
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "16 = WELDER CONTACT"
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
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "15 = WATER PUMP"
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
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "14 = TC FEEDER"
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
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "13 = EXHAUST FAN"
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
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "12 = ARGON"
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
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " 0 = Off"
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "11 = LIFT DOWN"
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
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "10 = Y DRIVE SOL"
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
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "  9 = LENS COOLING"
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
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3375
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   140
      Left            =   2040
      Top             =   0
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   615
      Left            =   4320
      TabIndex        =   52
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   735
      Left            =   4320
      TabIndex        =   51
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "E-STOP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Maintenance1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub EStop()
Dim temp() As Byte
    temp = c6k.FastStatus                  'get fast status information





Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure


'************* E STOP
If CStr((fsinfo.ProgIn(1) And Input5) / Input5) > 0 Then
        If (Last_Pcut_State = 0) Then
            Last_Pcut_State = 1
            c6k.Write ("1OUTALL17,32,0:" & Chr$(13))
            c6k.Write ("OUT.1-0:" & Chr$(13))
            Text9.BackColor = &HFFFF&
            Text9.ForeColor = QBColor(1)
            Text9.Text = "E-Stop"
            Text9.Refresh
            Timer2.Enabled = True
            EStopOn = True
            
            Exit Sub
        End If
    Else
        If (Last_Pcut_State = 1) Then
            Last_Pcut_State = 0
            Text9.BackColor = &H400000
            Text9.ForeColor = QBColor(1)
            Text9.Text = ""
            Text9.Refresh
            EStopOn = False
        End If
    End If
End Sub







Private Sub Command1_Click()
c6k.Write ("IF(1ANI.6>5):1OUT.13-0:T2:1OUT.13-1:T2:1OUT.13-0:NIF" & Chr$(13))
End Sub

Private Sub Command12_Click()
 Jog
'c6k.Write "1INFNC22-N:1INFNC23-M:JOYAXL1-1,1-2,1-0,1-0:JOYAXH1-0,1-0,1-1,,1-2:JOYVH15,12,14,2,4:JOYVl15,12,14,2,4:JOYA10,10,10,10,2:JOYAD30,100,100,10,10:" & Chr$(13)
'c6k.Write "1JOYZ.3=1:1JOYZ.2=1:1JOYEDB.3=1.18:1JOYEDB.2=1.18:1JOYCDB.2=.5:1JOYCDB.3=.5:" & Chr$(13)
'c6k.Write "JOY11101" & Chr$(13)
End Sub





Private Sub Command2_Click()
c6k.Write ("1INFNC21-5T:" & Chr$(13))
c6k.Write ("comexc0:MA00000000:@A10:@AD10:@V2:D,,,,-4:GO00001:" & Chr$(13))
c6k.Write ("HOMA,,,,2:HOMAD,,,,50:@HOMZ0:HOMV,,,,3:HOMVF,,,,1:" & Chr$(13))
c6k.Write ("HOMBAC11111:HOMEDG11110:HOMDF000001:HOM,,,,0:" & Chr$(13))
c6k.Write ("WAIT(5AS=XXXX1):T.1:D,,,,.37:GO00001:" & Chr$(13))
End Sub

Private Sub Command3_Click()
c6k.Write ("1INFNC18-6T:" & Chr$(13))
c6k.Write ("comexc0:MA00000000:@A10:@AD10:@V4:D,,,,,.3:GO000001:" & Chr$(13))
c6k.Write ("HOMA,,,,,2:HOMAD,,,,,50:@HOMZ0:HOMV,,,,,2:HOMVF,,,,,.25:" & Chr$(13))
c6k.Write ("HOMBAC111111:HOMEDG111110:HOMDF000001:HOM,,,,,1:" & Chr$(13))
c6k.Write ("WAIT(6AS=XXXX1):T1:D,,,,,1.375:GO000001:" & Chr$(13))
End Sub

Private Sub Command4_Click()
Timer1.Enabled = False
Maintenance1.Hide
cutlist.Show

End Sub

Private Sub Command5_Click()
c6k.Write ("PSET0,0,0,0,0,0,0,0:PESET0,0,0,0,0,0,0,0:" & Chr$(13))
'c6k1.Write ("PSET0,0,0,0,0,0,0,0:PESET0,0,0,0,0,0,0,0:" & Chr$(13))
End Sub

Private Sub Command6_Click()
Text6.SetFocus
End Sub

Private Sub Command7_Click()
Text1.SetFocus
End Sub

Private Sub Form_Load()
    c6k.FSEnabled = True                'enable fast status
c6k.Write (":COMEXS0" & Chr$(13))
 c6k.Write ("1INEN.5-0:" & Chr$(13))
c6k.Write ("DRIVE0,0,0,0,0,0,0,0:DRFEN00000000" & Chr$(13))
c6k.Write "AXSDEF00000000:@DRES250000:SCLD39683,39683,510204,,62500,62500:@SCLA25000:@SCLV25000:SCALE1:LH0,0,0,0,0,0,0,0:LHAD100,100,100,100,100,100,100,100:" + Chr$(13)
c6k.Write ("ENCPOL11011111:ERES4000,4000,4000,4000,4000,4000,4000,4000:ENCCNT11111111:" & Chr$(13))
c6k.Write ("COMEXS0" & Chr$(13))
c6k.Write ("DRIVE1,1,1,1,1,1,0,0" & Chr$(13))
c6k.Write "1ANIEN.1=E,E,E,E,E,E,E,E:" & Chr$(13)
c6k.Write "1ANIRNG.1=3:1ANIRNG.3=3:1ANIRNG.2=3:1ANIRNG.4=3:1ANIRNG.5=3:1ANIRNG.6=3:" & Chr$(13)
c6k.Write ("INFNC1-2J:INFNC2-2K:INFNC3-2J:INFNC4-2K:INFNC6-1J:INFNC7-1K:INFNC5-L:JOGA10,10,110:JOGAD999,999,999:JOGVH3,10,5:JOGVL3,10,12:" & Chr$(13))
c6k.Write ("1INFNC10-1J:JOGA6,10:JOGAD10,999:JOGVH8,1:JOGVL8,10:" & Chr$(13))


'    c6k1.FSEnabled = True                'enable fast status
'c6k1.Write ("DRIVE0,0,0,0,0,0,0,0:DRFEN00000000" & Chr$(13))
'c6k1.Write "AXSDEF00000000:@DRES25000:@SCLD25000:@SCLV25000:SCALE1:LH0,0,0,0,0,0,0,0:" + Chr$(13)

'c6k1.Write ("COMEXS0" & Chr$(13))
'c6k1.Write ("DRIVE1,1,1,1,0,1,0,1" & Chr$(13))

'c6k.Write ("comexc0:MA00000000:@A10:@AD100:@V13:@D-10:GO001:" & Chr$(13))

End Sub


Private Sub Form_LostFocus()
Timer1.Enabled = False

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = (13) And Text1.Text = "" Then
    c6k.Write "1OUTALL9,16,0:1OUTALL25,32,0" & Chr$(13)
    Label2.ForeColor = QBColor(12)
     For i = 1 To 16
       Label1(i - 1).ForeColor = QBColor(15)
     Next i
End If

If KeyAscii = (13) And Text1.Text <> "" Then
    c6k.Write "OUTFEN1:OUTFNC1-A:1OUT." + Str(Text1.Text) + "-1" & Chr$(13)
    
    'c6k.Write "1OUTALL1," + Str((Text1.Text) - 1) + ",0" & Chr$(13)
    
    'c6k.Write "1OUTALL" + Str(Text1.Text + 1) + ",32,0" & Chr$(13)
    
    If Val(Text1.Text) = 0 Then
     Label2.ForeColor = QBColor(12)
     c6k.Write "1OUTALL9,32,0" & Chr$(13)
     Else
     Label2.ForeColor = QBColor(15)
     End If
    For i = 9 To 16
        If Val(Text1.Text) = i Then
            Label1(i - 9).ForeColor = QBColor(12)
        Else
            ' Label1(i - 9).ForeColor = QBColor(15)
        End If
    Next i
    For i = 25 To 32
        If Val(Text1.Text) = i Then
            Label1(i - 17).ForeColor = QBColor(12)
        Else
            ' Label1(i - 9).ForeColor = QBColor(15)
        End If
     Next i
    Text1.Text = ""
End If
If KeyAscii = (13) And Text1.Text = "" Then
'c6k.write "1OUTALL17,24,0" & Chr$(13)
End If

End Sub



Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)



Select Case KeyCode
Case 40
If Text6.Text = "STOP" Then
'If Text6.Text <> "DOWN" Then
Text6.Text = "DOWN"
c6k.Write ("INEN.1-1:" & Chr$(13))
End If

Case 38
If Text6.Text = "STOP" Then
'If Text6.Text <> "UP" Then
Text6.Text = "UP"
c6k.Write ("INEN.2-1:" & Chr$(13))
End If

Case 37
If Text6.Text = "STOP" Then
    'If Text6.Text <> "LEFT" Then
    Text6.Text = "LEFT"
    If CtrlKeyOn = True Then
        c6k.Write ("INEN.3-1:" & Chr$(13))
    Else
        c6k.Write ("INEN.6-1:" & Chr$(13))
    End If
End If

Case 39
If Text6.Text = "STOP" Then
    'If Text6.Text <> "RIGHT" Then
    Text6.Text = "RIGHT"
    If CtrlKeyOn = True Then
        c6k.Write ("INEN.4-1:" & Chr$(13))
    Else
        c6k.Write ("INEN.7-1:" & Chr$(13))
    End If
End If
Case Else
c6k.Write ("!INEN.1-0:!INEN.2-0:" & Chr$(13))
c6k.Write ("!INEN.3-0:!INEN.4-0:" & Chr$(13))
c6k.Write ("!INEN.6-0:!INEN.7-0:" & Chr$(13))
Text6.Text = "STOP"

End Select
End Sub


Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
Text6.Text = "STOP"
c6k.Write ("!INEN.1-0:!INEN.2-0:" & Chr$(13))
c6k.Write ("!INEN.3-0:!INEN.4-0:" & Chr$(13))
c6k.Write ("!INEN.6-0:!INEN.7-0:" & Chr$(13))
'c6k.write ("!STOP" & Chr$(13))
End Sub


Private Sub Timer1_Timer()
Dim tom As Double
Dim temp() As Byte                      'create temporary byte array
temp = c6k.FastStatus                  'get fast status information
Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
'************************ INPUTS ************
  temp = c6k.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)
    ENCODER750 = (fsinfo.EncoderPos(1))
    Let Text2.Text = Val(ENCODER750) / 8160
    Text2.Text = Format(Text2.Text, "0.0000")
    Text2.BackColor = &HFFFF&
    Text2.Refresh
    If CStr((fsinfo.ProgIn(1) And Input1) / Input1) > 0 Then
        Label8(0).ForeColor = QBColor(12)
Else
        Label8(0).ForeColor = QBColor(15)
End If

    If CStr((fsinfo.ProgIn(1) And Input2) / Input2) > 0 Then
        Label8(1).ForeColor = QBColor(12)
Else
        Label8(1).ForeColor = QBColor(15)
End If

    If CStr((fsinfo.ProgIn(1) And Input3) / Input3) > 0 Then
        Label8(3).ForeColor = QBColor(12)
Else
        Label8(3).ForeColor = QBColor(15)
End If

    If CStr((fsinfo.ProgIn(1) And Input4) / Input4) > 0 Then
        Label8(2).ForeColor = QBColor(12)
Else
        Label8(2).ForeColor = QBColor(15)
End If

If CStr((fsinfo.ProgIn(1) And Input17) / Input17) > 0 Then
        Label4(8).ForeColor = QBColor(12)
Else
        Label4(8).ForeColor = QBColor(15)
End If
If CStr((fsinfo.ProgIn(1) And Input18) / Input18) > 0 Then
        Label4(9).ForeColor = QBColor(12)
Else
        Label4(9).ForeColor = QBColor(15)
End If
If CStr((fsinfo.ProgIn(1) And Input19) / Input19) > 0 Then
        Label4(10).ForeColor = QBColor(12)
        
        Else
        Label4(10).ForeColor = QBColor(15)
        
End If
If CStr((fsinfo.ProgIn(1) And Input20) / Input20) > 0 Then
        Label4(11).ForeColor = QBColor(12)
        Else
        Label4(11).ForeColor = QBColor(15)
End If
If CStr((fsinfo.ProgIn(1) And Input21) / Input21) > 0 Then
        Label4(12).ForeColor = QBColor(12)
        Else
        Label4(12).ForeColor = QBColor(15)
End If
If CStr((fsinfo.ProgIn(1) And Input22) / Input22) > 0 Then
        Label4(13).ForeColor = QBColor(12)
        Else
        Label4(13).ForeColor = QBColor(15)
End If
If CStr((fsinfo.ProgIn(1) And Input23) / Input23) > 0 Then
        Label4(14).ForeColor = QBColor(12)
        c6k.Write "OUT.5-1" & Chr$(13)
        Else
        Label4(14).ForeColor = QBColor(15)
        c6k.Write "OUT.5-0" & Chr$(13)
End If
If CStr((fsinfo.ProgIn(1) And Input24) / Input24) > 0 Then
        Label4(15).ForeColor = QBColor(12)
        Else
        Label4(15).ForeColor = QBColor(15)
End If
If CStr((fsinfo.ProgIn(1) And Input1) / Input1) > 0 Then
        Text12.Text = "1"
End If
If CStr((fsinfo.ProgIn(1) And Input2) / Input2) > 0 Then
        Text12.Text = "2"
End If
If CStr((fsinfo.ProgIn(1) And Input3) / Input3) > 0 Then
        Text12.Text = "3"
End If
If CStr((fsinfo.ProgIn(1) And Input4) / Input4) > 0 Then
        Text12.Text = "4"
End If


'angvolt = (fsinfo.Analog(2))
'            Text5.Text = Round((angvolt / 413), 2)
'            Text5.Refresh
'angvolt3 = (fsinfo.Analog(1))
'            Text7.Text = Round((angvolt3 / 413), 2)
'            Text7.Refresh
 
' angvolt4 = (fsinfo.Analog(4))
'            Text9.Text = Round((angvolt4 / 413), 2)
'            Text9.Refresh
' angvolt5 = (fsinfo.Analog(5))
'            Text10.Text = Round((angvolt5 / 413), 2)
'            Text10.Refresh
' angvolt6 = (fsinfo.Analog(6))
'            Text11.Text = Round((angvolt6 / 413), 2)
'            Text11.Refresh
 
 
 
 
 
 
 XMOTOR = (fsinfo.MotorPos(1))
            Let Text3.Text = Val(XMOTOR)
            Let Text3.Text = Text3.Text / 25000
            Text3.Text = Format(Text3.Text, "0.000")
           'Text18.Text = Yoffset2 - Text5.Text

'temp = c6k1.FastStatus
    Call CopyMemory(fsinfo, temp(0), 280)
 YMOTOR = (fsinfo.MotorPos(1))
          Let Text4.Text = Val(YMOTOR)
          Let Text4.Text = Text4.Text / 25000
          Text4.Text = Format(Text4.Text, "0.000")

    Call CopyMemory(fsinfo, temp(0), 280)
If CStr((fsinfo.AxisStatus(3) And joystick) / joystick) > 0 Then
          Let Text8.Text = "JoyStick On"
        Else
        Text8.Text = "JoyStick Off"
End If
    Call CopyMemory(fsinfo, temp(0), 280)   'copy byte array into structure
If CStr((fsinfo.ProgIn(1) And Input20) / Input20) = 0 Then
     If (Last_Pcut_State = 0) Then
            Last_Pcut_State = 1
            Label5.Caption = "E Stop!!!!!"
            c6k.Write "1OUTALL9,16,0:1OUTALL25,32,0:!S" & Chr$(13)
            Label5.Refresh
            'Timer2.Enabled = True
            EStopPos = ""
     End If
Else
     If (Last_Pcut_State = 1) Then
        Last_Pcut_State = 0
        Label5.Caption = ""
        Label5.Refresh
     End If
End If
TESTWATERPUMP = TESTWATERPUMP + 1
    Call CopyMemory(fsinfo, temp(0), 280)
    If CStr((fsinfo.ProgIn(1) And Input19) / Input19) > 0 Then
        Label5.Caption = ""
        Label5.Refresh
       ' Exit Do
    End If
  
    If TESTWATERPUMP > 10 Then
        Label5.Caption = " Water Pump HAS A PROBLEM!"
        Label5.Refresh
        Exit Sub
    End If


If CStr((fsinfo.ProgOut(0) And Input5) / Input5) > 0 Then
        Label5.Caption = "ON"
        
        Else
       Label5.Caption = "OOFF"
        
End If

'If CStr((fsinfo.ProgIn(1) And Input22) / Input22) > 0 And ((fsinfo.ProgIn(1) And Input24) / Input24) > 0 Then
'        Label4(13).ForeColor = QBColor(12)
'        c6k.Write ("1OUT.13-1:" & Chr$(13))
'        Else
'        Label4(13).ForeColor = QBColor(15)
'        c6k.Write ("1OUT.13-0:" & Chr$(13))
'End If
End Sub


