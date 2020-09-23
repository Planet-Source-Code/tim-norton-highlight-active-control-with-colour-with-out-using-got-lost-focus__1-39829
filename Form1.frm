VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   240
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "5 October 2002"
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "nortont@rexel.com.au"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Code by: Tim Norton"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Call CurrentControl(Screen.ActiveControl, Screen.ActiveForm)
End Sub
