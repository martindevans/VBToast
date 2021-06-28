VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3105
   ClientLeft      =   12825
   ClientTop       =   7410
   ClientWidth     =   2040
   LinkTopic       =   "Form2"
   ScaleHeight     =   3105
   ScaleWidth      =   2040
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Text            =   "X"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "save walls"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   "0"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "select wall"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "enemy?"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "new melon"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "reset spawn"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "lay new wall"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Form1.Check1.Value = Form2.Check1.Value
Form1.Check1_Click
End Sub

Private Sub Command1_Click()
Form1.lay_new_wall_Click
End Sub

Private Sub Command2_Click()
Form1.reset_player_Click
End Sub

Private Sub Command3_Click()
Form1.Command1_Click
End Sub

Private Sub Command4_Click()
Form1.select_wall_Click
End Sub

Private Sub Command5_Click()
Form1.save_walls_Click
End Sub

Private Sub Text1_Change()
Form1.wall_select_box = Form2.Text1
End Sub

Private Sub Text2_Change()
Form1.lvlnum = Form2.Text2
End Sub
