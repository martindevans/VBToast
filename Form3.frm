VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "T - Main Menu"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   5295
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit Game"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Game"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Click in the grey box for more hints and tips..."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Stuff_Box 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Stuff_Var As Byte
Const NumStuffs = 5

Private Sub Command1_Click()
Load Form1
Form1.Show
Form3.Hide
Unload Form3
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Stuff_Var = 1
Timer1_Timer
End Sub

Private Sub Stuff_Box_Click()
Timer1_Timer
End Sub

Private Sub Timer1_Timer()

Debug.Print "menu printing" & Stuff_Var

'read the level file
filename = "GameData/MenuStuff/Stuff" & Stuff_Var & ".txt"
FileNumber = FreeFile           'Freefile returns the first unused file number
Open filename For Input As FileNumber   'Opens the file for input
length = LOF(FileNumber)        'LOF function returns length of file
If length < 32767 Then
    strreading = Input$(length, FileNumber)
Else
    MsgBox "This file is longer than 32Kb"
End If
Close FileNumber            'Close when done
Stuff_Box.Caption = strreading

Stuff_Var = (Stuff_Var + 1) Mod (NumStuffs + 1)
If Stuff_Var = 0 Then
    Stuff_Var = 1
End If

End Sub
