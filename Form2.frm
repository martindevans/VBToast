VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Game Over"
   ClientHeight    =   1185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3435
   LinkTopic       =   "Form2"
   ScaleHeight     =   1185
   ScaleWidth      =   3435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label EndGame_Score 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

EndGame_Score.Caption = "You got " & Form1.points & " points in the entire game."
Select Case Form1.points
    Case Is < 500
        EndGame_Score.Caption = EndGame_Score.Caption & " That's awful"
    Case Is > 499 And Form1.points < 1000
        EndGame_Score.Caption = EndGame_Score.Caption & " That's pretty bad"
    Case Is > 999 And Form1.points < 2500
        EndGame_Score.Caption = EndGame_Score.Caption & " That's quite good"
    Case Is > 2499
        EndGame_Score.Caption = EndGame_Score.Caption & " That's really good, well done"
    End Select
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
