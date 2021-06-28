VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "SAVE THE MELONS!"
   ClientHeight    =   10320
   ClientLeft      =   -2220
   ClientTop       =   -1845
   ClientWidth     =   13125
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer FPSTimer 
      Interval        =   1000
      Left            =   120
      Top             =   7680
   End
   Begin VB.Timer difficulty_timer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   120
      Tag             =   "makes the difficulty label disappear"
      Top             =   8160
   End
   Begin VB.Frame Help_credits 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Help / Credits"
      Height          =   5415
      Left            =   2160
      TabIndex        =   30
      Top             =   4440
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Help: -> Main controls are the arrow keys"
         Height          =   4695
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Thankyou for your enlightenment Master"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   5040
         Width           =   4695
      End
   End
   Begin VB.Timer labeltimer 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   120
      Tag             =   "makes points labels disappear"
      Top             =   8640
   End
   Begin VB.Frame Console 
      Caption         =   "Ze Cheat Console!"
      Height          =   855
      Left            =   1800
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton Command5 
         Caption         =   "Toggle FPS"
         Height          =   495
         Left            =   2160
         TabIndex        =   37
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "record"
         Height          =   255
         Left            =   1440
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox cheat_lvl_select 
         Height          =   285
         Left            =   720
         TabIndex        =   25
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Load Lvl"
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+1 life"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Label11"
         Height          =   255
         Left            =   1440
         TabIndex        =   35
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enter A Level Code:"
      Height          =   1335
      Left            =   10440
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Text            =   "ye"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Load Level"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Frame TheWay 
      BackColor       =   &H00C0C0C0&
      Caption         =   "The Tao of Toast"
      Height          =   3855
      Left            =   8160
      TabIndex        =   15
      Top             =   6360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   3135
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   17
         Text            =   "Form1.frx":0000
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Thankyou For Your Knoweledge Master"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3480
         Width           =   4575
      End
   End
   Begin VB.Frame Main_menu 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Menu"
      Height          =   2055
      Left            =   10560
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Help / Credits"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exit"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Return To Game"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "The Way Of The Toast"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enter A Level Code"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame progress_frame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "WELL DONE"
      Height          =   2415
      Left            =   7320
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox lvlCode_Box 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Text            =   "Enter a Level Code"
         Top             =   1560
         Width           =   4935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "LOAD LEVEL..."
         Height          =   375
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   6
         Top             =   1920
         Width           =   4935
      End
      Begin VB.Label lived_inform 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lived Bonus"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label Time_inform 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Time Bonus:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label Melons_inform 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Melons collected:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label lvlCode_label 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Well done my young toast, you have done well in completing level X, the code for the next level is:"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   4935
      End
   End
   Begin VB.PictureBox Offscreenarrowleft 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   120
      Picture         =   "Form1.frx":0235
      ScaleHeight     =   345
      ScaleWidth      =   180
      TabIndex        =   3
      Top             =   9840
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox Offscreenarrowright 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   360
      Picture         =   "Form1.frx":05B3
      ScaleHeight     =   345
      ScaleWidth      =   180
      TabIndex        =   2
      Top             =   9840
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox offscreenarrowup 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   120
      Picture         =   "Form1.frx":0931
      ScaleHeight     =   180
      ScaleWidth      =   345
      TabIndex        =   1
      Top             =   9600
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Timer Timer 
      Interval        =   1
      Left            =   120
      Tag             =   "the main game loop"
      Top             =   9120
   End
   Begin VB.Label FPSLabel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label12"
      Height          =   255
      Left            =   480
      TabIndex        =   36
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Difficulty_label 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Difficulty level:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   2400
      TabIndex        =   33
      Top             =   2280
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Image enemy 
      Height          =   450
      Index           =   0
      Left            =   1320
      Picture         =   "Form1.frx":0CD3
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Pause_indicator 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Paused"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   960
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lives_indicator 
      BackColor       =   &H0000FFFF&
      Caption         =   "Lives: 0"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Shape BT_indicator 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   3855
      Left            =   120
      Top             =   840
      Width           =   255
   End
   Begin VB.Label points_box 
      BackColor       =   &H0000FFFF&
      Caption         =   "Points: 0"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image player 
      Height          =   570
      Left            =   6360
      Picture         =   "Form1.frx":470D
      Stretch         =   -1  'True
      Top             =   360
      Width           =   570
   End
   Begin VB.Shape Box 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   480
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Melon 
      Height          =   210
      Index           =   0
      Left            =   2040
      Picture         =   "Form1.frx":8147
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   210
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NumLevels = 30

Const coefficientoffriction = 5
Const coefficientoffrictionforthesides = 3
Const enemyspeed = 9

Dim paused As Boolean
Dim replaying As Boolean
Dim replay_string As String

Dim recording As Boolean
Dim fso As Scripting.FileSystemObject
Dim txs As TextStream

Dim FPS As Integer

Dim time_left As Integer

Dim lostlife As Boolean

Dim adjust As Integer

Public points As Integer

Dim upforce As Integer
Dim rightforce As Integer

Dim level As Integer

Dim spawnX As Integer
Dim spawnY As Integer

Dim oldX As Integer
Dim oldY As Integer

Dim xpos As Integer
Dim ypos As Integer

Dim leftkey As Boolean
Dim rightkey As Boolean
Dim upkey As Boolean
Dim downkey As Boolean

Dim numenemies As Integer
Dim enemyarray() As Boolean

Dim lives As Integer

Dim onfloor As Boolean
Dim hitleftside As Boolean
Dim hitrightside As Boolean

Dim donemelons As Integer
Dim nummelons As Integer
Dim melons() As position

Dim numwalls As Integer
Dim walls() As position

Private Function make_code(levelX As Integer) As String

Dim code As String
Dim total As Integer
Dim desired As Integer
Dim number As Integer

code = ""
desired = levelX * 111

brunning = True
Do While brunning = True
    'a = 97
    'z = 122
    If desired - total < 240 Then
        Do While desired - total > 0
            number = desired - total
            If number > 122 Then
                number = 121
            End If
            total = total + number
            code = code & Chr(number)
            DoEvents
        Loop
        brunning = False
    Else
        number = random(97, 122)
        Do While (number > 122) Or (number < 97)
            number = random(97, 122)
        Loop
        total = total + number
        code = code & Chr(number)
    End If
Loop

make_code = code

End Function

Private Sub Command1_Click()

progress_frame.Visible = False

load_level

If level <> NumLevels + 1 Then
    xpos = spawnX
    ypos = spawnY
    
    Timer.Enabled = True
Else
    Timer.Enabled = False
    Unload Form1
End If

End Sub

Private Sub Command2_Click()
lives = lives + 1
lives_indicator = "lives:" & lives
End Sub

Private Sub Command3_Click()
level = cheat_lvl_select
load_level
End Sub

Private Sub Command4_Click()

If recording = True Then
    recording = False
Else
    recording = True
End If

Label11 = recording

End Sub

Private Sub Command5_Click()
FPSLabel.Visible = Not FPSLabel.Visible
End Sub

Private Sub difficulty_timer_Timer()

Difficulty_label.Visible = False
paused = False
difficulty_timer.Enabled = False
Timer.Enabled = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 223
        'console key
        Console.Visible = Not Console.Visible

    Case vbKeyP
        If paused = True Then
            paused = False
            Timer.Enabled = True
            Pause_indicator.Visible = False
            Main_menu.Visible = False
            Frame2.Visible = False
            TheWay.Visible = False
            Help_credits.Visible = False
        Else
            paused = True
            Pause_indicator.Visible = True
            Main_menu.Visible = True
        End If
    Case vbKeyRight
        If progress_frame.Visible = False Then
            rightkey = True
            player.Picture = LoadPicture("GameData/ninja2right.bmp")
        Else
        End If
    Case vbKeyLeft
        If progress_frame.Visible = False Then
            leftkey = True
            player.Picture = LoadPicture("GameData/ninja2left.bmp")
        Else
        End If
    Case vbKeyUp
        If progress_frame.Visible = False Then
            upkey = True
        Else
        End If
    Case 27
        If paused = True Then
            paused = False
            Timer.Enabled = True
            Pause_indicator.Visible = False
            Main_menu.Visible = False
            Frame2.Visible = False
            TheWay.Visible = False
            Help_credits.Visible = False
        Else
            paused = True
            Pause_indicator.Visible = True
            Main_menu.Visible = True
        End If
End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeyRight
        rightkey = False
    Case vbKeyLeft
        leftkey = False
    Case vbKeyUp
        upkey = False
End Select

End Sub

Private Sub load_level()

'decrypt level code
If lvlCode_Box.Text <> "Enter a Level Code" Then
    Do While Len(lvlCode_Box.Text) > 0
        strreading = Asc(Left(lvlCode_Box, 1))
        lvlCode_Box = Right(lvlCode_Box, Len(lvlCode_Box) - 1)
        total = total + strreading
        DoEvents
    Loop
    level = total / 111
End If

If level > 30 Then
    Load Form2
    Unload Form1
    Form2.Show
Else
    
    'difficulty notification
    Select Case level
        Case 1
            Difficulty_label.Caption = "Difficulty level: Easy :)"
            Difficulty_label.Left = (Form1.Width / 2) - (Difficulty_label.Width / 2)
            Difficulty_label.Top = (Form1.Height / 2) - (Difficulty_label.Height / 2)
            Difficulty_label.Visible = True
            paused = True
            difficulty_timer.Enabled = True
        Case 25
            Difficulty_label.Caption = "Difficulty level: Hard :o"
            Difficulty_label.Left = (Form1.Width / 2) - (Difficulty_label.Width / 2)
            Difficulty_label.Top = (Form1.Height / 2) - (Difficulty_label.Height / 2)
            Difficulty_label.Visible = True
            paused = True
            difficulty_timer.Enabled = True
        Case 50
            Difficulty_label.Caption = "Difficulty level: Insane! =>:("
            Difficulty_label.Left = (Form1.Width / 2) - (Difficulty_label.Width / 2)
            Difficulty_label.Top = (Form1.Height / 2) - (Difficulty_label.Height / 2)
            Difficulty_label.Visible = True
            paused = True
            difficulty_timer.Enabled = True
        Case 75
            Difficulty_label.Caption = "Difficulty level: Give up now :'("
            Difficulty_label.Left = (Form1.Width / 2) - (Difficulty_label.Width / 2)
            Difficulty_label.Top = (Form1.Height / 2) - (Difficulty_label.Height / 2)
            Difficulty_label.Visible = True
            paused = True
            difficulty_timer.Enabled = True
    End Select
    
    
    
'    'read the replay file
'    filename = "GameData/replays/replay" & level & ".txt"
'    FileNumber = FreeFile           'Freefile returns the first unused file number
'    Open filename For Input As FileNumber   'Opens the file for input
'    length = LOF(FileNumber)        'LOF function returns length of file
'    If length < 32767 Then
'        strreading = Input$(length, FileNumber)
'    Else
'        MsgBox "This file is longer than 32Kb"
'    End If
'    Close FileNumber            'Close when done
'    replay_string = strreading
'
'    'debugging
    replaying = False
    
    
    
    For j = 1 To Box.UBound
        Unload Box(j)
    Next j
    
    For j = 1 To Melon.UBound
        Unload Melon(j)
    Next j
    
    For j = 1 To enemy.UBound
        Unload enemy(j)
    Next j
    
    'read the level file
    filename = "GameData/levels/lvl" & level & ".txt"
    FileNumber = FreeFile           'Freefile returns the first unused file number
    Open filename For Input As FileNumber   'Opens the file for input
    length = LOF(FileNumber)        'LOF function returns length of file
    If length < 32767 Then
        strreading = Input$(length, FileNumber)
    Else
        MsgBox "This file is longer than 32Kb"
    End If
    Close FileNumber            'Close when done
    
    'spawn position
    spawnX = Left(strreading, InStr(strreading, ";") - 1)
    strreading = Right(strreading, Len(strreading) - InStr(strreading, ";"))
    spawnY = Left(strreading, InStr(strreading, ";") - 1)
    strreading = Right(strreading, Len(strreading) - InStr(strreading, ";"))
    
    'set numenemies, numenemies of course always = numwalls
    numenemies = Left(strreading, InStr(strreading, ";") - 1)
    strreading = Right(strreading, Len(strreading) - InStr(strreading, ";"))
    numwalls = numenemies
    
    For j = 0 To numwalls - 1
        If j <> 0 Then
            Load Box(j)
            Box(j).FillColor = &H0&
            Box(j).Visible = True
        End If
        Box(j).Left = Left(strreading, InStr(strreading, ";") - 1)
        strreading = Right(strreading, Len(strreading) - InStr(strreading, ";"))
        Box(j).Top = Left(strreading, InStr(strreading, ";") - 1)
        strreading = Right(strreading, Len(strreading) - InStr(strreading, ";"))
        Box(j).Width = Left(strreading, InStr(strreading, ";") - 1)
        strreading = Right(strreading, Len(strreading) - InStr(strreading, ";"))
        Box(j).Height = Left(strreading, InStr(strreading, ";") - 1)
        strreading = Right(strreading, Len(strreading) - InStr(strreading, ";"))
    Next j
    
    '/---------------------------------------------------\
    '|PUT IN ENEMY FILE READING, IT APPEARS TO BE MISSING|
    '\---------------------------------------------------/
    'ignore this, it's still applicable but black magic is...
    '...somehow pulling the data from the file anyway :/
    
    'no files here, just storing the form positions in an array
    ReDim walls(numwalls - 1)
    For j = 0 To numwalls - 1
        walls(j).xpos = Box(j).Left
        walls(j).ypos = Box(j).Top
        walls(j).xwidth = Box(j).Width
        walls(j).ywidth = Box(j).Height
    Next j
    
    nummelons = Left(strreading, InStr(strreading, ";") - 1)
    strreading = Right(strreading, Len(strreading) - InStr(strreading, ";"))
    
    ReDim melons(nummelons)
    
    Melon(0).Visible = False
    
    For j = 1 To nummelons
        If j <> 0 Then
            Load Melon(j)
            Melon(j).Visible = True
        End If
        Melon(j).ToolTipText = j
        Melon(j).Left = Left(strreading, InStr(strreading, ";") - 1)
        strreading = Right(strreading, Len(strreading) - InStr(strreading, ";"))
        Melon(j).Top = Left(strreading, InStr(strreading, ";") - 1)
        strreading = Right(strreading, Len(strreading) - InStr(strreading, ";"))
        Melon(j).Visible = True
        melons(j).xpos = Melon(j).Left
        melons(j).ypos = Melon(j).Top
        melons(j).xwidth = Melon(j).Width
        melons(j).ywidth = Melon(j).Height
    Next j
    
    ReDim enemyarray(numenemies - 1)
    For j = 0 To numenemies - 1
        If j <> 0 Then
            Load enemy(j)
            enemy(j).ZOrder (0)
        End If
        strreading = Right(strreading, Len(strreading) - InStr(strreading, ";"))
        If Left(strreading, InStr(strreading, ";") - 1) = "true" Then
            enemy(j).Visible = True
            If random(0, 1) = 1 Then
                enemyarray(j) = True
            Else
                enemyarray(j) = False
            End If
        Else
            enemy(j).Visible = False
        End If
        strreading = Right(strreading, Len(strreading) - InStr(strreading, ";"))
        
        enemy(j).Left = Box(j).Left + Box(j).Width / 2
        enemy(j).Top = Box(j).Top - enemy(j).Height
    Next j
    
    xpos = spawnX
    ypos = spawnY
    
    lostlife = False
    lives = lives + 1
    lives_indicator = "lives:" & lives
    time_left = 10000
End If

End Sub

Private Sub Form_Load()

If level <> NumLevels + 1 Then
    replaying = True
    
    time_left = 10000
    lives = 2
    level = 1
    
    If level = 1 Then
        Set fso = New Scripting.FileSystemObject
        Set txs = fso.CreateTextFile("log.txt", True)
    End If
    
    load_level
    
'    For j = 0 To nummelons
'        Melon(j).Picture = LoadPicture("melon.bmp")
'    Next j
    
    xpos = spawnX
    ypos = spawnY
    
    lives_indicator = "lives:" & lives
End If

End Sub

Private Sub Form_Resize()

'centre paused
Pause_indicator.Left = (Form1.Width / 2) - (Pause_indicator.Width / 2)
Pause_indicator.Top = (Form1.Height / 2) - (Pause_indicator.Height / 2) - (Main_menu.Height)

'centre menu
Main_menu.Left = (Form1.Width / 2) - (Main_menu.Width / 2)
Main_menu.Top = Pause_indicator.Top + Pause_indicator.Height

End Sub


Private Sub FPSTimer_Timer()
FPSLabel = FPS
FPS = 0
End Sub

Private Sub Label1_Click()

'set help stuff
Label10.Caption = "Help:" & vbNewLine & "-> Main controls are the arrow keys, left and right, up to jump, hold up to jump lots of times." & vbNewLine & vbNewLine & vbNewLine _
& "Tips:" & vbNewLine & "-> The aim of the game is to collect all the melons in a level and then get back to the red exit square thing" & vbNewLine & vbNewLine & _
"-> It is possible to NINJA WALL JUMP! ahem, Run towards a vertical wall, and press jump when you get there, hold jump and keep pressing towards the wall and you will achieve the mythical NINJA WALL JUMP!" _
& vbNewLine & vbNewLine & vbNewLine & "Credits:" & vbNewLine & "Ok, this game would not have been possible without:" & vbNewLine & "-Me! (Martin Evans), programming, level design, game testing" & vbNewLine & _
"-Dan Atkinson, level design, game art, game testing" & vbNewLine & "-Dave Copson, level design, game testing" & vbNewLine & "-Jon Rudd, game concept, level design, game art, game testing" & vbNewLine & _
"-Tom James, Level Design, game testing" & vbNewLine & "-Graham Evans, game testing (quite badly :D)" & vbNewLine & "-Philip Evans, game testing (officially worst T player EVER!"

Help_credits.Visible = True
Help_credits.Left = (Main_menu.Left) + (Main_menu.Width / 2) - (Help_credits.Width / 2)
Help_credits.Top = Main_menu.Top
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BorderStyle = 1
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BorderStyle = 0
End Sub

Private Sub Label2_Click()
Frame2.Left = (Main_menu.Left) + (Main_menu.Width / 2) - (Frame2.Width / 2)
Frame2.Top = Main_menu.Top
Frame2.Visible = True
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 1
End Sub
Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 0
End Sub
Private Sub Label3_Click()
TheWay.Left = (Form1.Width / 2) - (TheWay.Width / 2)
TheWay.Top = Main_menu.Top
TheWay.Visible = True
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BorderStyle = 1
End Sub
Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BorderStyle = 0
End Sub

Private Sub Label4_Click()
Help_credits.Visible = False
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.BorderStyle = 1
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.BorderStyle = 0
End Sub

Private Sub Label5_Click()
Pause_indicator.Visible = False
Main_menu.Visible = False
paused = False
Timer.Enabled = True
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BorderStyle = 1
End Sub
Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BorderStyle = 0
End Sub
Private Sub Label6_Click()
Unload Me
End Sub
Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.BorderStyle = 1
End Sub
Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.BorderStyle = 0
End Sub
Private Sub Label7_Click()
TheWay.Visible = False
End Sub
Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BorderStyle = 1
End Sub
Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BorderStyle = 0
End Sub
Private Sub Label8_Click()
lvlCode_Box.Text = Text2
Frame2.Visible = False
load_level
End Sub

Private Sub Label9_Click()
Frame2.Visible = False
End Sub

Private Sub labeltimer_Timer()

Dim endit As Boolean

If labeltimer.Interval = 3000 Then
    'begin shrinking the labels
    labeltimer.Interval = 10
Else
    'shrink the labels
    If points_box.Width > 16 Then
        points_box.Width = points_box.Width - 10
    Else
        endit = True
    End If
    
    If lives_indicator.Width > 16 Then
        lives_indicator.Width = lives_indicator.Width - 10
    Else
        If endit = True Then
            labeltimer.Interval = 3000
            labeltimer.Enabled = False
        End If
    End If
End If

End Sub

Private Sub lives_indicator_Change()

points_box.Width = 1695
lives_indicator.Width = 1695

labeltimer.Enabled = True

End Sub

Private Sub Timer_Timer()

FPS = FPS + 1

If paused = True Then
    Timer.Enabled = False
End If

If progress_frame.Visible = True Then
    upforce = 0
    rightforce = 0
    xpos = spawnX
    ypos = spawnY
Else
    If replaying = False Then
        'enemies
        For j = 0 To numenemies - 1
            
            If enemy(j).Visible = True Then
            
                'move enemies
                If enemyarray(j) = True Then
                    'moving right
                    enemy(j).Left = enemy(j).Left + random(enemyspeed - enemyspeed / 2, enemyspeed + enemyspeed / 2)
                    If (enemy(j).Left + enemy(j).Width) > Box(j).Left + Box(j).Width Then
                        enemyarray(j) = False
                        enemy(j).Picture = LoadPicture("GameData/heroictoastleft.BMP")
                    End If
                Else
                    'moving left
                    enemy(j).Left = enemy(j).Left + random(-(enemyspeed + enemyspeed / 2), -(enemyspeed - enemyspeed / 2))
                    If (enemy(j).Left) < Box(j).Left Then
                        enemyarray(j) = True
                        enemy(j).Picture = LoadPicture("GameData/heroictoastright.BMP")
                    End If
                End If
                
                'check collision with player
                If xpos < enemy(j).Left + enemy(j).Width * 0.85 Then
                        
                    If xpos + player.Width > enemy(j).Left + enemy(j).Width * 0.15 Then
                            
                        If ypos < enemy(j).Top + enemy(j).Height * 0.85 Then
                                
                            If ypos + player.Height > enemy(j).Top + enemy(j).Height * 0.15 Then
                                'collided with enemy
                                lose_life
                            End If
                        End If
                    End If
                End If
            End If
        Next j
        
        'level timer
        time_left = time_left - 1
        If time_left < 1 Then
            time_left = 10000
            lose_life
            BT_indicator.Height = Barheight(time_left, 10000, 3855)
        End If
        
        If adjust = 70 Then
            BT_indicator.Height = Int(Barheight(time_left, 10000, 3855))
            adjust = 0
        Else
            adjust = adjust + 1
        End If
        
        'gravity
        upforce = upforce - 1
        
        'carry out player controls
        If (rightkey = True) And (hitleftside = False) Then
            rightforce = rightforce + 3
        Else
        End If
        
        If (leftkey = True) And (hitrightside = False) Then
            rightforce = rightforce - 3
        Else
        End If
        
        If upkey = True Then
        
            If onfloor = True Then
                upforce = upforce + 60
            Else
            
                If hitleftside = True Then
                    upforce = upforce + 40
                    rightforce = rightforce - 20
                    hitleftside = False
                Else
                    If hitrightside = True Then
                        upforce = upforce + 40
                        rightforce = rightforce + 20
                        hitleftside = False
                    End If
                End If
            End If
        End If
        
        hitrightside = False
        hitleftside = False
        onfloor = False
        
        For j = 0 To numwalls - 1
            'collision detect
                            
            If initialdetect(xpos, ypos, player.Width, player.Height, walls(j).xpos, walls(j).ypos, walls(j).xwidth, walls(j).ywidth) = True Then
                'if you're here then you've collided with the current wall
                
                If j = 0 Then
                    'if you're here then you've landed on the exit portal
                    'check if all melons are gone
                    If donemelons = nummelons Then
                        'level end
                        'update points inform
                        Melons_inform = "Melons collected :" & nummelons
                        Time_inform = "Time Bonus: " & Int(time_left / 100)
                        If lost_life = False Then
                            lived_inform = "Lived Bonus: " & lives * 5
                            points = points + lives * 5
                        Else
                            lived_inform = "Lived Bonus: 0"
                        End If
                        
                        'update points
                        points = points + Int(time_left / 100)
                        Timer.Enabled = False
                        lvlCode_label.Caption = "Well done my young toast, you have done well in completing level " & level & ", the code for the next level is:"
                        lvlCode_Box.Text = make_code(level + 1)
                        progress_frame.Left = (Form1.Width / 2) - (progress_frame.Width / 2)
                        progress_frame.Top = (Form1.Height / 2) - (progress_frame.Height / 2)
                        progress_frame.Visible = True
                        xpos = spawnX
                        ypos = spawnY
                        upforce = 3
                        rightforce = 0
                        donemelons = 0
                        leftkey = False
                        rightkey = False
                        upkey = False
                        
                        'recording
                        recording = False
                        txs.WriteLine replay_string
                        
                        'goto almost the bottom of this subroutine,
                        'The place Just before player avatar is drawn at
                        'xpos and ypos
                        GoTo SkipToDrawAvatar
                    End If
                End If
                                
                'If (ypos + player.Height > walls(j).ypos - upforce) And (ypos < (walls(j).ypos + walls(j).ywidth - upforce)) Then
                'side check
                                 
                Select Case sidedetectshape(xpos, ypos, player.Width, player.Height, walls(j).xpos, walls(j).ypos, walls(j).xwidth, walls(j).ywidth)
                    Case "left"
                        'hitting the left side of the player
                        xpos = (walls(j).xpos + Box(j).Width) + 1
                                        
                        hitrightside = True
                        hitleftside = False
                            
                        If rightforce < 0 Then
                            rightforce = 3
                        Else
                            rightforce = rightforce + 3
                        End If
                        
                        dofrictionsides
                        
                    Case "right"
                        'hitting the right side
                        xpos = (walls(j).xpos - player.Width) - 1
                                        
                        hitleftside = True
                        hitrightside = False
                                        
                        If rightforce > 0 Then
                            rightforce = -3
                        Else
                            rightforce = rightforce - 3
                        End If
                        
                        dofrictionsides
                    Case "top"
                        'downforce of the roof
                        ypos = Box(j).Top + Box(j).Height + 1
                        upforce = 0
                        onfloor = False
                        
                        dofrictionwalls
                        
                    Case "bottom"
                        'upforce of the floor
                        ypos = Box(j).Top - player.Height - 1
                        upforce = -upforce * 0.3
                        onfloor = True
                            
                        dofrictionwalls
                            
                    End Select
                                    
                Else
                                        
            End If
                                    
                
        Next j
        
        'melon code
        For j = 1 To nummelons
            'collision detect
            
            If xpos < melons(j).xpos + melons(j).xwidth Then
                
                If xpos + player.Width > melons(j).xpos Then
                    
                    If ypos < melons(j).ypos + melons(j).ywidth Then
                        
                        If ypos + player.Height > melons(j).ypos Then
                        
                            'ooh! melon is collected :)
                            If Melon(j).Visible = True Then
                                donemelons = donemelons + 1
                                Melon(j).Visible = False
                                points = points + 1
                                points_box.Caption = "Points: " & points
                            Else
                            End If
                        End If
                    End If
                End If
            End If
        Next j
        
        'draw arrows if character is offscreen
        If player.Top + player.Height < 0 Then
            offscreenarrowup.Left = player.Left + (player.Width / 2) - (offscreenarrowup.Width / 2)
            offscreenarrowup.Top = 0
            If offscreenarrowup.Visible = False Then
                offscreenarrowup.Visible = True
            End If
        Else
            offscreenarrowup.Visible = False
        End If
        
        If player.Left + player.Width > Form1.Width Then
            Offscreenarrowright.Top = player.Top + (player.Height / 2) - (Offscreenarrowright.Width / 2)
            Offscreenarrowright.Left = Form1.Width - Offscreenarrowright.Width
            If Offscreenarrowright.Visible = False Then
                Offscreenarrowright.Visible = True
            End If
        Else
            Offscreenarrowright.Visible = False
        End If
        
        If player.Left + player.Width < 0 Then
            Offscreenarrowleft.Top = player.Top + (player.Height / 2) - (Offscreenarrowleft.Width / 2)
            Offscreenarrowleft.Left = 0
            If Offscreenarrowleft.Visible = False Then
                Offscreenarrowleft.Visible = True
            End If
        Else
            Offscreenarrowleft.Visible = False
        End If
        
        If ypos > 25000 Then
            lose_life
        Else
            If ypos < -25000 Then
                lose_life
            End If
        End If
        
        If xpos > 25000 Then
            lose_life
            lives_indicator = "lives:" & lives
        Else
            If xpos < -25000 Then
                lose_life
                lives_indicator = "lives:" & lives
            End If
        End If
        
        oldX = xpos
        oldY = ypos
        xpos = xpos + rightforce
        ypos = ypos - upforce
    Else
        If Len(replay_string) = 0 Then
            'end replay
            replaying = False
        Else
            'if you're here then you're replaying for the current level
            'if you're not here then go away ;)
            xpos = Left(replay_string, InStr(replay_string, ";") - 1)
            replay_string = Right(replay_string, (Len(replay_string)) - (Len((Left(replay_string, InStr(replay_string, ";"))))))
            ypos = Left(replay_string, InStr(replay_string, ";") - 1)
            replay_string = Right(replay_string, (Len(replay_string)) - (Len((Left(replay_string, InStr(replay_string, ";"))))))
        End If
        
    End If

End If

SkipToDrawAvatar:

'update positions
player.Left = xpos
player.Top = ypos

If recording = True Then
    replay_string = replay_string & xpos & ";" & ypos & ";"
End If

If lives < 0 Then
    Load Form2
    Unload Form1
    Form2.Show
End If

End Sub

Private Sub dofrictionwalls()

    If rightforce > 0 Then
                                            
        If rightforce > coefficientoffriction Then
            rightforce = rightforce - coefficientoffriction
        Else
            rightforce = 0
        End If
                                            
    Else
                                        
        If rightforce < 0 Then
                                                
            If -rightforce > coefficientoffriction Then
                rightforce = rightforce + coefficientoffriction
            Else
                rightforce = 0
            End If
                                                
        Else
            'rightforce = 0
        End If
    End If

End Sub

Private Sub dofrictionsides()
      
    If upforce > 0 Then
        If upforce > coefficientoffrictionforthesides Then
            upforce = upforce - coefficientoffrictionforthesides
        Else
            upforce = 0
        End If
                                        
    Else
                                        
        If upforce < 0 Then
                                                
            If -upforce > coefficientoffrictionforthesides Then
                upforce = upforce + coefficientoffrictionforthesides
            Else
                upforce = 0
            End If
                                                
        Else
            'rightforce = 0
        End If
    End If

End Sub

Private Sub lose_life()
lives = lives - 1
lostlife = True
lives_indicator = "lives:" & lives
xpos = spawnX
ypos = spawnY
rightforce = 0
upforce = 0
End Sub

Function Barheight(ByVal CurrentValue As Integer, ByVal MaxValue As Integer, ByVal MaxBarheight As Integer)

MaxValue = MaxValue + 1

    If CurrentValue >= MaxValue Then
        Barheight = MaxValue
    Else
        Barheight = MaxBarheight * (CurrentValue / MaxValue)
    End If

End Function
