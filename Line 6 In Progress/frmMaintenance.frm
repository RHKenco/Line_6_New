VERSION 5.00
Begin VB.Form frmMaintenance 
   BackColor       =   &H00C00000&
   Caption         =   "Maintainance"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   10740
   LinkTopic       =   "Form6"
   ScaleHeight     =   5775
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer_c6kRead 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.Frame Frame_Motors 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   4935
      Left            =   5040
      TabIndex        =   54
      Top             =   600
      Width           =   4455
      Begin VB.CommandButton Button_Move_Axis 
         Caption         =   "GoY"
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   87
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Button_Move_Axis 
         Caption         =   "Go R"
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   86
         Top             =   3600
         Width           =   735
      End
      Begin VB.CommandButton Button_Move_Axis 
         Caption         =   "Go O"
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   85
         Top             =   3000
         Width           =   735
      End
      Begin VB.CommandButton Button_Move_Axis 
         Caption         =   "Go Za"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   84
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton Button_Move_Axis 
         Caption         =   "Go Z"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   83
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton Button_Move_Axis 
         Caption         =   "Go X"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   82
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text_In_Des 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   81
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text_In_Des 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   80
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox Text_In_Des 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   79
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text_In_Des 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   78
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text_In_Des 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   77
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text_In_Des 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   76
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text_Pop_DRO 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   600
         TabIndex        =   68
         Text            =   "0.000"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text_Pop_DRO 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   600
         TabIndex        =   67
         Text            =   "0.000"
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox Text_Pop_DRO 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   600
         TabIndex        =   66
         Text            =   "0.000"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Text_Pop_DRO 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   600
         TabIndex        =   65
         Text            =   "0.000"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text_Pop_DRO 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   600
         TabIndex        =   64
         Text            =   "0.000"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text_Pop_DRO 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   63
         Text            =   "0.000"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label_Drive_Units 
         BackColor       =   &H00C00000&
         Caption         =   "in"
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
         Index           =   5
         Left            =   1560
         TabIndex        =   75
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label_Drive_Units 
         BackColor       =   &H00C00000&
         Caption         =   "deg"
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
         Index           =   4
         Left            =   1560
         TabIndex        =   74
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label Label_Drive_Units 
         BackColor       =   &H00C00000&
         Caption         =   "in"
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
         Index           =   3
         Left            =   1560
         TabIndex        =   73
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label_Drive_Units 
         BackColor       =   &H00C00000&
         Caption         =   "in"
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
         Left            =   1560
         TabIndex        =   72
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label_Drive_Units 
         BackColor       =   &H00C00000&
         Caption         =   "in"
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
         Left            =   1560
         TabIndex        =   71
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label_Drive_Units 
         BackColor       =   &H00C00000&
         Caption         =   "in"
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
         Left            =   1560
         TabIndex        =   70
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Var_Label_Joystick_Status 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "Joystick Enabled:"
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
         Height          =   735
         Left            =   1440
         TabIndex        =   69
         Top             =   4080
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label_Drive_Num 
         BackColor       =   &H00C00000&
         Caption         =   "8 - Unused"
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
         Index           =   7
         Left            =   240
         TabIndex        =   62
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label_Drive_Num 
         BackColor       =   &H00C00000&
         Caption         =   "7 - Unused"
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
         Index           =   6
         Left            =   240
         TabIndex        =   61
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label_Drive_Num 
         BackColor       =   &H00C00000&
         Caption         =   "6 - Y-Axis"
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
         Index           =   5
         Left            =   240
         TabIndex        =   60
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label_Drive_Num 
         BackColor       =   &H00C00000&
         Caption         =   "5 - APT Rot"
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
         Index           =   4
         Left            =   240
         TabIndex        =   59
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label_Drive_Num 
         BackColor       =   &H00C00000&
         Caption         =   "4 - Osc"
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
         Index           =   3
         Left            =   240
         TabIndex        =   58
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label_Drive_Num 
         BackColor       =   &H00C00000&
         Caption         =   "3 - Z-Alt"
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
         Left            =   240
         TabIndex        =   57
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label_Drive_Num 
         BackColor       =   &H00C00000&
         Caption         =   "2 - Z-Axis"
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
         Left            =   240
         TabIndex        =   56
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label_Drive_Num 
         BackColor       =   &H00C00000&
         Caption         =   "1 - X-Axis"
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
         Left            =   240
         TabIndex        =   55
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame_Output 
      BackColor       =   &H00C00000&
      Caption         =   "Block4: Output"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2415
      Index           =   1
      Left            =   2520
      TabIndex        =   36
      Top             =   3120
      Width           =   2415
      Begin VB.CheckBox Check_Output 
         BackColor       =   &H00FF0000&
         Caption         =   "Check1"
         Height          =   255
         Index           =   15
         Left            =   2040
         TabIndex        =   44
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox Check_Output 
         BackColor       =   &H00FF0000&
         Caption         =   "Check1"
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   43
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox Check_Output 
         BackColor       =   &H00FF0000&
         Caption         =   "Check1"
         Height          =   255
         Index           =   13
         Left            =   2040
         TabIndex        =   42
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox Check_Output 
         BackColor       =   &H00FF0000&
         Caption         =   "Check1"
         Height          =   255
         Index           =   12
         Left            =   2040
         TabIndex        =   41
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Check_Output 
         BackColor       =   &H00FF0000&
         Caption         =   "Check1"
         Height          =   255
         Index           =   11
         Left            =   2040
         TabIndex        =   40
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox Check_Output 
         BackColor       =   &H00FF0000&
         Caption         =   "Check1"
         Height          =   255
         Index           =   10
         Left            =   2040
         TabIndex        =   39
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check_Output 
         BackColor       =   &H00FF0000&
         Caption         =   "Check1"
         Height          =   255
         Index           =   9
         Left            =   2040
         TabIndex        =   38
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox Check_Output 
         BackColor       =   &H00FF0000&
         Caption         =   "Check1"
         Height          =   255
         Index           =   8
         Left            =   2040
         TabIndex        =   37
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "32 - Unused"
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
         Index           =   31
         Left            =   120
         TabIndex        =   52
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "31 - Unused"
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
         Index           =   30
         Left            =   120
         TabIndex        =   51
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "30 - Unused"
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
         Index           =   29
         Left            =   120
         TabIndex        =   50
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "29 - Unused"
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
         Index           =   28
         Left            =   120
         TabIndex        =   49
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "28 - Unused"
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
         Index           =   27
         Left            =   120
         TabIndex        =   48
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "27 - Unused"
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
         Index           =   26
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "26 - Unused"
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
         Index           =   25
         Left            =   120
         TabIndex        =   46
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "25 - Unused"
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
         Index           =   24
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame_Input 
      BackColor       =   &H00C00000&
      Caption         =   "Block3: Input"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2415
      Index           =   1
      Left            =   240
      TabIndex        =   27
      Top             =   3120
      Width           =   2175
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "24 - Joy Tog"
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
         Index           =   23
         Left            =   120
         TabIndex        =   35
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "23 - Joy Rlse"
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
         Index           =   22
         Left            =   120
         TabIndex        =   34
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "22 - Joy Select"
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
         Index           =   21
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "21 - Rot Prox"
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
         Index           =   20
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "20 - E-Stop"
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
         Index           =   19
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "19 - H20 Flow"
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
         Index           =   18
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "18 - Osc. Lim"
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
         Index           =   17
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "17 - Unused"
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
         Index           =   16
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame_Output 
      BackColor       =   &H00C00000&
      Caption         =   "Block2: Output"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2415
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   2415
      Begin VB.CheckBox Check_Output 
         BackColor       =   &H00FF0000&
         Caption         =   "Check1"
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   25
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox Check_Output 
         BackColor       =   &H00FF0000&
         Caption         =   "Check1"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   24
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox Check_Output 
         BackColor       =   &H00FF0000&
         Caption         =   "Check1"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   23
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox Check_Output 
         BackColor       =   &H00FF0000&
         Caption         =   "Check1"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   22
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Check_Output 
         BackColor       =   &H00FF0000&
         Caption         =   "Check1"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   21
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox Check_Output 
         BackColor       =   &H00FF0000&
         Caption         =   "Check1"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   20
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check_Output 
         BackColor       =   &H00FF0000&
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   19
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox Check_Output 
         BackColor       =   &H00FF0000&
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   18
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "16 - Welder"
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
         Index           =   15
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "15 - H2O Pump"
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
         Index           =   14
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "14 - TC Feed"
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
         Index           =   13
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "13 - Exhaust"
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
         Index           =   12
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "12 - Argon"
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
         Index           =   11
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "11 - Unused"
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
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "10 - Unused"
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
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "  9 - Airblade"
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
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame_Input 
      BackColor       =   &H00C00000&
      Caption         =   "Block1: Input"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2415
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2175
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "8 - Unused"
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
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "7 - Unused"
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
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "6 - Unused"
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
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "5 - Unused"
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
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "4 - Joy Down"
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
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "3 - Joy Up"
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
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "2 - Joy Right"
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
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label_Block_Pin 
         BackColor       =   &H00C00000&
         Caption         =   "1 - Joy Left"
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
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label_Drives 
      BackColor       =   &H00C00000&
      Caption         =   "        Drives                                              "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5040
      TabIndex        =   53
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label_IO_Section 
      BackColor       =   &H00C00000&
      Caption         =   "        I/O                                                  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   240
      Width           =   4695
   End
   Begin VB.Menu Topbar_Joystick 
      Caption         =   "Joystick"
   End
   Begin VB.Menu Topbar_Stop_Output 
      Caption         =   "Stop Output"
   End
End
Attribute VB_Name = "frmMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Button_Move_Axis_Click(Index As Integer)

Const MoveVelocities = "1,1,1,1,1,1"

'Generate string of values for 6k with desired distance & GO command in correct position
Dim tempD As String
Dim tempGO As String

For i = 0 To 5

    If i = 1 Then
        If i = Index Then
            tempD = Format(CDbl(Text_In_Des(i).Text), "0.000")
            tempGO = "1"
        Else
            tempD = "0"
            tempGO = "0"
        End If
    Else
        If i = Index Then
            tempD = tempD & "," & Format(CDbl(Text_In_Des(i).Text), "0.000")
            tempGO = tempGO & "1"
        Else
            tempD = tempD & ",0"
            tempGO = "0"
        End If
    End If

Next i

'Instruct 6k
c6k.Write ("MC0:V" & MoveVelocities & ":D" & tempD & ":GO" & tempGO & Chr(13))

End Sub



Private Sub Form_Load()
    
    Timer_c6kRead.Enabled = True
    
    Call c6kOps.stopAllOut

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call FL6_Close_Maintainance
    
End Sub

Private Sub FL6_Close_Maintainance()
    Timer_c6kRead.Enabled = False
End Sub

Private Sub setInputText(currentInputState As Long)

Dim Index As Integer

For i = 0 To 15
    'Set up index for selecting proper label & Input binary
    If i > 7 Then Index = i + 8 Else Index = i
    
    'Set text to red if input is active
    If (currentInputState And (2 ^ Index)) Then Label_Block_Pin(Index).ForeColor = (&HCF&) Else Label_Block_Pin(Index).ForeColor = (&H8000000E)

Next i

End Sub

Private Sub setOutputs(currentOutputState As Long)

'Create local temp variables
Dim outputIndex As Integer
Dim outputOn As Boolean

'For all checkboxes
For Each i In Check_Output

    '-- Set up index for selecting proper label & Input binary
    If i < 8 Then outputIndex = i + 8 Else outputIndex = i + 16
    
    'Compare output long to single bit shifted to output location, then convert to bool
    outputOn = (currentOutputState And (2 ^ outputIndex))
    
    '-- If output state does not match checkbox, set output state accordingly
    ' If the current output is not enabled and checkbox is checked
    If Not outputOn And Check_Output(i).value = 1 Then
        'Acivate Output
        Call c6kOps.setOutputNum((outputIndex + 1), True)
        'Set text color to red
        Label_Block_Pin(i).ForeColor = (&HCF&)
        
    ' If the current output is enabled and the box not checked
    ElseIf outputOn And Check_Output(i).value = 0 Then
        'Disable Output
        Call c6kOps.setOutputNum((outputIndex + 1), False)
        'Set text color to white
        Label_Block_Pin(i).ForeColor = (&H8000000E)
        
    ' If the output is enabled and the checkbox is clicked
    ElseIf outputOn And Check_Output(i).value = 1 Then
        'Verify color is set correctly
        If Label_Block_Pin(i).ForeColor = (&H8000000E) Then Label_Block_Pin(i).ForeColor = (&HCF&)
        
    ' If the output is not enabled and the checkbox is not clicked
    ElseIf Not outputOn And Check_Output(i).value = 0 Then
        'Verify color is set correctly
        If Label_Block_Pin(i).ForeColor = (&HCF&) Then Label_Block_Pin(i).ForeColor = (&H8000000E)
    
    Else
        MsgBox "Error in Maintenance Output Control"
    End If
    
Next i

End Sub

Private Sub Timer_c6kRead_Timer()

'Call Fast Status
Call c6kOps.updFastStatus

'-- Check for E-Stop
If Not c6kOps.chkE_Stop Then

'-- Run Joystick if active
    If c6kOps.getJoyActive And Not c6kOps.chkE_Stop() Then
    
        'Run JoyRun function, and if it returns true,
        If c6kOps.runJoy("Free") Then
    
            'Set Joystick Status Message
            Var_Label_Joystick_Status.Caption = "Joystick" & Chr(13) & "Enabled:" & Chr(13) & Chr(13) & c6kOps.getJoyStr() & Chr(13) & "Mode"
            Var_Label_Joystick_Status.Visible = True
    
        Else
            'If the joystick becomes inactive hide label
            Var_Label_Joystick_Status.Visible = False
    
        End If
    
    '--Input State Debug - Uncomment these two lines to enter input debug mode
    'Var_Label_Joystick_Status.Caption = c6kOps.getInputState
    'Var_Label_Joystick_Status.Visible = True
        
    End If

    'Set input text to red if input is active
    Call setInputText(c6kOps.getInputState())
    Call setOutputs(c6kOps.getOutputState())
    
    Call c6kOps.updDro

'Else

'Add functionality for E-Stop label in frmMaintenance Here

End If

'Reset Fast Status Update Flag
Call c6kOps.resetFSupd

End Sub

Private Sub Topbar_Joystick_Click()

' Toggle joystick active boolean
If Not c6kOps.getJoyActive Then
    c6kOps.runJoy ("Enable")
Else
    c6kOps.runJoy ("Disable")
    Var_Label_Joystick_Status.Visible = False
End If

End Sub

Private Sub Topbar_Stop_Output_Click()

    Call c6kOps.stopAllOut

End Sub
