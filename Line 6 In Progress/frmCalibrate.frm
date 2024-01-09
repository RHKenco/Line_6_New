VERSION 5.00
Begin VB.Form frmCalibrate 
   BackColor       =   &H00C00000&
   Caption         =   "Calibration"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6240
   LinkTopic       =   "Form6"
   ScaleHeight     =   4050
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Button_Default 
      Caption         =   "Default"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   21
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton Button_Load 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   20
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Button_Save 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   19
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Button_Update 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   18
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text_Pass_Thrsh 
      Height          =   285
      Index           =   2
      Left            =   4080
      TabIndex        =   17
      Top             =   1560
      Width           =   1680
   End
   Begin VB.TextBox Text_Pass_Thrsh 
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   16
      Top             =   1200
      Width           =   1680
   End
   Begin VB.TextBox Text_Osc_Speed 
      Height          =   285
      Index           =   3
      Left            =   2640
      TabIndex        =   15
      Top             =   1680
      Width           =   1200
   End
   Begin VB.TextBox Text_Osc_Speed 
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   14
      Top             =   1320
      Width           =   1200
   End
   Begin VB.TextBox Text_Osc_Speed 
      Height          =   285
      Index           =   1
      Left            =   2640
      TabIndex        =   13
      Top             =   960
      Width           =   1200
   End
   Begin VB.TextBox Text_Pass_Thrsh 
      Height          =   285
      Index           =   0
      Left            =   4080
      TabIndex        =   12
      Top             =   840
      Width           =   1680
   End
   Begin VB.TextBox Text_Osc_Speed 
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   11
      Top             =   600
      Width           =   1200
   End
   Begin VB.TextBox Text_Pass_Speed 
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   10
      Top             =   1680
      Width           =   1200
   End
   Begin VB.TextBox Text_Pass_Speed 
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   9
      Top             =   1320
      Width           =   1200
   End
   Begin VB.TextBox Text_Pass_Speed 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   8
      Top             =   960
      Width           =   1200
   End
   Begin VB.TextBox Text_Pass_Speed 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   7
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label_Trns 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  Transition   Pass Width (in)"
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
      Height          =   615
      Left            =   4080
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label_Osc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Osc (in/s)"
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
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label_Pass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pass (in/s)"
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
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label_Speed 
      BackColor       =   &H00C00000&
      Caption         =   "Speed 4: "
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
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label_Speed 
      BackColor       =   &H00C00000&
      Caption         =   "Speed 3: "
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
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label_Speed 
      BackColor       =   &H00C00000&
      Caption         =   "Speed 2: "
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
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label_Speed 
      BackColor       =   &H00C00000&
      Caption         =   "Speed 1: "
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
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmCalibrate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button_Default_Click()

Call c6kOps.defaultPassSpeed

End Sub

Private Sub Button_Load_Click()

Call c6kOps.loadPassSpeed

Call c6kOps.dispPassSpeed

End Sub

Private Sub Button_Save_Click()

Call c6kOps.setPassSpeed(Text_Pass_Speed(0), Text_Pass_Speed(1), Text_Pass_Speed(2), Text_Pass_Speed(3), Text_Osc_Speed(0), Text_Osc_Speed(1), Text_Osc_Speed(2), Text_Osc_Speed(3), Text_Pass_Thrsh(0), Text_Pass_Thrsh(1), Text_Pass_Thrsh(2))

Call c6kOps.savePassSpeed

End Sub

Private Sub Button_Update_Click()

Call c6kOps.setPassSpeed(Text_Pass_Speed(0), Text_Pass_Speed(1), Text_Pass_Speed(2), Text_Pass_Speed(3), Text_Osc_Speed(0), Text_Osc_Speed(1), Text_Osc_Speed(2), Text_Osc_Speed(3), Text_Pass_Thrsh(0), Text_Pass_Thrsh(1), Text_Pass_Thrsh(2))

End Sub

Private Sub Form_Load()
    
    Call c6kOps.loadPassSpeed
    
    Call c6kOps.dispPassSpeed
    
End Sub

