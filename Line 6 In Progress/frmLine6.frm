VERSION 5.00
Begin VB.Form frmLine6 
   BackColor       =   &H00C00000&
   Caption         =   "Line 6"
   ClientHeight    =   6660
   ClientLeft      =   4680
   ClientTop       =   1710
   ClientWidth     =   19410
   LinkTopic       =   "Form6"
   ScaleHeight     =   6660
   ScaleWidth      =   19410
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
      Height          =   4095
      Left            =   17040
      TabIndex        =   29
      Top             =   2400
      Visible         =   0   'False
      Width           =   2175
      Begin VB.TextBox Text_Pop_DRO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "0.000"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text_Pop_DRO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "0.000"
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox Text_Pop_DRO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "0.000"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Text_Pop_DRO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "0.000"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text_Pop_DRO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "0.000"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text_Pop_DRO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
         Top             =   600
         Width           =   375
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox Text_Focus_Trap 
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   6240
      Width           =   615
   End
   Begin VB.CommandButton Button_Set_Auger 
      Caption         =   "Auger Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5400
      TabIndex        =   26
      Top             =   5040
      Width           =   2175
   End
   Begin VB.OptionButton Option_Auger_Direction 
      BackColor       =   &H00C00000&
      Caption         =   "Run Edge"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   22
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Frame Frame_Auger_Switch 
      BackColor       =   &H00C00000&
      Caption         =   "Auger Direction"
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
      Height          =   1095
      Left            =   5400
      TabIndex        =   20
      Top             =   5040
      Width           =   2175
      Begin VB.CommandButton Button_Clear_Auger 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option_Auger_Direction 
         BackColor       =   &H00C00000&
         Caption         =   "Run Face"
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox Text_Enter_Pass_Width 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   18
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Button_Go 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   600
      TabIndex        =   17
      Top             =   4320
      Width           =   4575
   End
   Begin VB.CommandButton Button_Fin 
      Caption         =   "Finish"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   16
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Button_NF 
      Caption         =   "N / F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   15
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Button_Start 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3720
      TabIndex        =   14
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text_Pop_Total_Qty 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text_Pop_Part_Num 
      Enabled         =   0   'False
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
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Timer Timer_FSM 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Button_WO_Enter_Clear 
      Caption         =   "Enter"
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
      Left            =   4200
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text_Pop_Due_Date 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text_Pop_Dwg_Num 
      Enabled         =   0   'False
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
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox Text_Enter_WO 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Text            =   "Enter Work Order Number Here"
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label Var_Label_Pass_Speed 
      BackStyle       =   0  'Transparent
      Caption         =   "Pass Speed"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3375
      Left            =   5400
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label_Estop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E-Stop Enabled"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000CF&
      Height          =   495
      Left            =   6360
      TabIndex        =   24
      Top             =   5520
      Visible         =   0   'False
      Width           =   12615
   End
   Begin VB.Label Var_Label_Joystick_Status 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "Joystick Enabled:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3495
      Index           =   0
      Left            =   5400
      TabIndex        =   23
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label_Pass_Width 
      BackColor       =   &H00C00000&
      Caption         =   "Pass Width"
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
      Left            =   5400
      TabIndex        =   19
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Var_Label_System_Status 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ready to Run Blade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2055
      Left            =   6120
      TabIndex        =   13
      Top             =   840
      Width           =   12615
   End
   Begin VB.Label Var_Label_WO_Active 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No Active Work Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6120
      TabIndex        =   12
      Top             =   480
      Width           =   12615
   End
   Begin VB.Label Label_Total_Qty 
      BackColor       =   &H00C00000&
      Caption         =   "Quantity"
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
      Left            =   1920
      TabIndex        =   11
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label_Part_Num 
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
      Left            =   600
      TabIndex        =   9
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label_WO_Info 
      BackColor       =   &H00C00000&
      Caption         =   "Work Order Information:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label_Due_Date 
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
      Left            =   600
      TabIndex        =   4
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label_Dwg_Num 
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
      Left            =   600
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label_WO 
      BackStyle       =   0  'Transparent
      Caption         =   "Work_Order"
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
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Menu Topbar_Set_0 
      Caption         =   "Set 0"
   End
   Begin VB.Menu Topbar_Joystick 
      Caption         =   "Joystick"
   End
   Begin VB.Menu Topbar_Maintenance 
      Caption         =   "Maintenance"
      Begin VB.Menu Topbar_Maint_Open 
         Caption         =   "Open Maintenance Form"
      End
      Begin VB.Menu Topbar_Show_Terminal 
         Caption         =   "Show Terminal"
      End
      Begin VB.Menu Topbar_Show_DRO 
         Caption         =   "Show DRO"
      End
   End
   Begin VB.Menu Topbar_Test_Dropdown 
      Caption         =   "Test"
      Begin VB.Menu topbar_test_0 
         Caption         =   "c6k Check Input Functions"
      End
      Begin VB.Menu topbar_test_1 
         Caption         =   "Simple Move Motors"
      End
      Begin VB.Menu topbar_test_2 
         Caption         =   "Check Inputs (Generic)"
      End
      Begin VB.Menu topbar_test_3 
         Caption         =   "Check Inputs (Specific)"
      End
      Begin VB.Menu topbar_test_4 
         Caption         =   "Basic Joystick"
      End
   End
   Begin VB.Menu Topbar_Calibration 
      Caption         =   "Calibration"
   End
   Begin VB.Menu Topbar_Reset 
      Caption         =   "Reset"
   End
End
Attribute VB_Name = "frmLine6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button_Clear_Auger_Click()

' Clear Auger Parameters
Call c6kOps.clearAugerParam

' Reset Pass Type
Call c6kOps.setPassType

'Re-show set button
Button_Set_Auger.Visible = True
Button_Set_Auger.Refresh

Call statusMsg(msgActive)

End Sub

Private Sub Button_Fin_Click()
'finish Punch
Call woMgr.finishWO

End Sub

Private Sub Button_Go_Click()

'Set pass width
Call c6kOps.setPassWidth

'Set Button State
Call btnState(btnRunning)

fsmMain.State = 2

'Text_Focus_Trap.SetFocus

End Sub

Private Sub Button_NF_Click()

'NF Punch
Call woMgr.notfinishWO

End Sub

Private Sub Button_Set_Auger_Click()

'Show auger setup form
frmAugerSetup.Show

End Sub

Private Sub Button_Start_Click()

'Start Punch
Call woMgr.startWO
    
End Sub

Private Sub Button_WO_Enter_Clear_Click()

If Not woMgr.isWOactive() Then
    Call woMgr.loadWO
    
    '-- Start Airblade & Exhaust Fan
    Call c6kOps.setOutput(outAirblade, True)
    Call c6kOps.setOutput(outExhaust, True)
Else
    If Not woMgr.clearWO Then Exit Sub
    
    '-- Stop Airblade & Exhaust Fan
    Call c6kOps.setOutput(outAirblade, False)
    Call c6kOps.setOutput(outExhaust, False)
End If

Text_Enter_Pass_Width.SetFocus

End Sub

Private Sub Option_Auger_Direction_Click(Index As Integer)

If Index = 0 Then
    If Option_Auger_Direction(0).value = True Then
        Option_Auger_Direction(1).value = False
    Else
        Option_Auger_Direction(1).value = True
    End If
Else
    If Option_Auger_Direction(1).value = True Then
        Option_Auger_Direction(0).value = False
    Else
        Option_Auger_Direction(0).value = True
    End If
End If

Call statusMsg(msgActive)
Call c6kOps.setPassType

End Sub

Private Sub Text_Enter_Pass_Width_Change()

Static pwOldInput As Single
Dim pwNewInput As String

pwNewInput = Text_Enter_Pass_Width.Text
If IsNumeric(pwNewInput) Then
    pwOldInput = CSng(pwNewInput)
ElseIf pwNewInput = "" Then
    Exit Sub
Else
    Text_Enter_Pass_Width.Text = CStr(pwOldInput)
End If

End Sub


Private Sub Text_Enter_Pass_Width_KeyPress(KeyAscii As Integer)

If KeyAscii = (13) Then

Call Button_Go_Click

End If

End Sub

Private Sub Text_Enter_WO_KeyPress(KeyAscii As Integer)

If KeyAscii = (13) Then
    'If enter is pressed and the work order is not already active
    If Not woMgr.isWOactive Then Call Button_WO_Enter_Clear_Click
    
    Text_Enter_Pass_Width.SetFocus
End If

End Sub


Private Sub Form_Activate()

'When the form returns to focus, activate the FSM
Timer_FSM.Enabled = True

End Sub

Private Sub Form_Deactivate()

'When the form loses focus, deactivate the FSM and reset the FastStatus up-to-date flag
Timer_FSM.Enabled = False

End Sub

Private Sub Form_Load()
'When the form is loaded:

'-- Set Up Defaults for Text
Text_Enter_WO.Tag = "Enter Work Order Number Here"
Text_Enter_WO.Text = Text_Enter_WO.Tag

Var_Label_WO_Active.Tag = "No Active Work Order"
Var_Label_WO_Active.Caption = Var_Label_WO_Active.Tag
Var_Label_WO_Active.Visible = True

Call statusMsg(msgInactive)

Var_Label_Joystick_Status(0).Visible = False

Label_Estop.Visible = False

'Initialize Focus Trap Textbox
Text_Focus_Trap.Visible = False
Text_Focus_Trap.TabStop = False
'Text_Focus_Trap.SetFocus

'Initialize Form Buttons
Call btnState(btnInactive)


'--
Call c6kOps.Enable
Call Joy.createJoystick

'-- Run motor setup subroutine
Call c6kOps.bootDrives

'-- Check for previously active WO
Call woMgr.chkActiveWO

'-- Enable FSM
Timer_FSM.Enabled = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If User hits X
    If UnloadMode = 0 Then
        If MsgBox("Are you sure you want to close?", vbYesNo Or vbQuestion) = vbNo Then Cancel = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call FL6_End_Program
    
End Sub

Private Sub FL6_End_Program()

    'Shut Down c6k - Immediate jog off, disable outputs, disable drives, clear command buffer
    c6k.Write "!S:!JOG00000000:1OUTALL9,16,0:1OUTALL25,32,0:DRIVE0,0,0,0,0,0,0,0"
    
    'Reset c6k Controller to avoid retained settings & commands in next boot. Will disconnect ethernet
    'c6k.Write ("RESET")
    
    Var_Label_WO_Active.Caption = "Unloading all forms, Please Wait"
    Var_Label_WO_Active.Refresh
    
    'Clear All Variables
    
    
    
    'Unload all Forms except frmLine6 and frmMain
    For Each tmpForm In Forms
        If (tmpForm.Name <> "frmLine6") Or (tmpForm.Name <> "frmMain") Then
            Unload tmpForm
            Set tmpForm = Nothing
        End If
    Next
            
    'Show frmMain
    Set c6k = Nothing
    frmMain.Launch_Cmd.Enabled = False
    frmMain.Connect_Cmd.Caption = "Connect"
    frmMain.Connect_Cmd.Enabled = False         'Until variable clearing is set up and fucntional, this prevents problems with variables that got set but not cleared by forcing the user to close the program
    frmMain.Show
End Sub

Private Sub Text_Enter_WO_GotFocus()

If Text_Enter_WO.Text = Text_Enter_WO.Tag Then Text_Enter_WO = vbNullString

End Sub

Private Sub Text_Enter_WO_LostFocus()

If Trim(Text_Enter_WO) = vbNullString Then Text_Enter_WO = Text_Enter_WO.Tag
    
End Sub

Private Sub Text_Pop_Mat_Change()

End Sub

'-- Main FSM Timer - Operates the non-UI elements of the program ------------------------------------
Private Sub Timer_FSM_Timer()

'Run FSM
Call fsmMain.Run

'Update DRO if it is visible
If frmLine6.Frame_Motors.Visible Then Call c6kOps.updDro

End Sub


Private Sub Topbar_Calibration_Click()

frmCalibrate.Show

End Sub

Private Sub Topbar_Joystick_Click()

' Toggle joystick active boolean
If Not Joy.getJoyActive Then
    Call Joy.runJoy(joyFree)
Else
    Call Joy.runJoy(joyDisable)
    Var_Label_Joystick_Status(0).Visible = False
End If

End Sub

Private Sub Topbar_Maint_Open_Click()

'Show Maintenance Form
frmMaintenance.Show

End Sub

Private Sub Topbar_Reset_Click()

'Display Verification Prompt
temp = MsgBox("Reset Drives?", 1, "Reset")

If temp = 1 Then
    'Send Reset command to 6k - simulates hard reboot
    c6k.Write ("!RESET" & Chr$(13))
    
    'Unload the main form, as the reset command will disconnect ethernet
    Unload Me
End If
    
End Sub

Private Sub Topbar_Set_0_Click()

'Display Verification Prompt
temp = MsgBox("Set Machine Home?", 1, "Set 0")

If temp = 0 Then
    'Send command to 6k to set 0 position on all axes
    Call c6kOps.setMachineHome
End If

End Sub

Private Sub Topbar_Show_DRO_Click()
    If Frame_Motors.Visible Then
        Frame_Motors.Visible = False
    Else
        Frame_Motors.Visible = True
    End If
End Sub

Private Sub Topbar_Show_Terminal_Click()
    frm6kCmd.Show
End Sub

Private Sub topbar_test_0_Click()
    c6k.Write "1INFNC" & Chr(13)
End Sub

Private Sub topbar_test_1_Click()
    Call c6kOps.testDrives
End Sub

Private Sub topbar_test_2_Click()
    MsgBox ("Inputs Are: " & Chr$(13) & c6kOps.getInputStateStr)
End Sub

Private Sub topbar_test_4_Click()

c6k.Write ("!JOG000000:" & Chr$(13))
c6k.Write ("JOG000000:1INFNC1-5J:1INFNC2-5K:1INFNC3-2K:1INFNC4-2J:JOGA4,5,5,1,5,5:JOGAD50,99,99,99,99,15:JOGVH8,8,10,2,5,3:JOGVL8,15,10,5,5,5:JOG01010" & Chr$(13))

End Sub

