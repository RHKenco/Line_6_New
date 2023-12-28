VERSION 5.00
Begin VB.Form frmAugerSetup 
   BackColor       =   &H00C00000&
   Caption         =   "Auger Setup"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3030
   LinkTopic       =   "Form6"
   ScaleHeight     =   6105
   ScaleWidth      =   3030
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text_Pop_Auger_Angle 
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
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton Button_Compute 
      Caption         =   "Compute"
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
      Left            =   480
      TabIndex        =   9
      Top             =   3360
      Width           =   2175
   End
   Begin VB.OptionButton Option_Auger_Twist 
      BackColor       =   &H00C00000&
      Caption         =   "Left-Handed"
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
      Left            =   600
      TabIndex        =   8
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Frame Frame_Auger_Twist 
      BackColor       =   &H00C00000&
      Caption         =   "Direction"
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
      Height          =   975
      Left            =   480
      TabIndex        =   6
      Top             =   2160
      Width           =   2175
      Begin VB.OptionButton Option_Auger_Twist 
         BackColor       =   &H00C00000&
         Caption         =   "Right-Handed"
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
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox Text_Enter_Auger_Dia 
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
      Left            =   480
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text_Enter_Auger_Pitch 
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
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label_Auger_Ready 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "Auger Program Ready"
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
      Left            =   480
      TabIndex        =   12
      Top             =   5280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label_Auger_Setup_Angle 
      BackColor       =   &H00C00000&
      Caption         =   "Set Auger to:"
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
      Left            =   480
      TabIndex        =   10
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label_Auger_Setup 
      BackColor       =   &H00C00000&
      Caption         =   "Auger Setup:"
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
      Left            =   360
      TabIndex        =   5
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label Label_Auger_Diameter 
      BackColor       =   &H00C00000&
      Caption         =   "Diameter"
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
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label_Auger_Pitch 
      BackColor       =   &H00C00000&
      Caption         =   "Pitch"
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
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label_Auger_Info 
      BackColor       =   &H00C00000&
      Caption         =   "Auger Information:"
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmAugerSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
