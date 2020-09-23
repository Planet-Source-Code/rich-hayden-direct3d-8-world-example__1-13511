VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Details"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Label lblFps 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label Label1 
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label2.Caption = "Controls: " & Chr(13) & "TURN LEFT=LEFT ARROW" & Chr(13) & "TURN RIGHT=RIGHT ARROW" & Chr(13) & "MOVE FORWARD=UP ARROW" & Chr(13) & "MOVE BACKWARD=DOWN ARROW" & Chr(13) & "JUMP=SPACE BAR" & Chr(13) & "RUN=HOLD DOWN SHIFT"
End Sub

