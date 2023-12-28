VERSION 5.00
Begin VB.Form Form7 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   Caption         =   "ERROR !"
   ClientHeight    =   1200
   ClientLeft      =   3435
   ClientTop       =   1965
   ClientWidth     =   2745
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1200
   ScaleWidth      =   2745
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2280
      Top             =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File Open      Please Wait !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Form7.Hide
End Sub


