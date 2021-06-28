VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T Map Maker"
   ClientHeight    =   9915
   ClientLeft      =   420
   ClientTop       =   510
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   13005
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   11040
      Top             =   7800
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      Begin VB.CommandButton save_walls 
         Caption         =   "save walls"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox lvlnum 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Text            =   "0"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "enemy?"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton select_wall 
         BackColor       =   &H00000000&
         Caption         =   "Select wall"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "New Melon"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton reset_player 
         Caption         =   "reset spawn"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox wall_select_box 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton lay_new_wall 
         Caption         =   "lay new wall"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Shape melon 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   2520
      Top             =   3480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape player 
      BorderWidth     =   2
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   3720
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape wall 
      BorderWidth     =   3
      Height          =   495
      Index           =   0
      Left            =   6720
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nummelons As Integer
Dim activemelon As Integer

Dim enemy() As Boolean

Dim numwalls As Integer
Dim selected As Integer
Dim fso As Scripting.FileSystemObject
Dim txs As TextStream
Dim set_start As Boolean

Public Sub Check1_Click()
If Check1.Value = 1 Then
    enemy(wall_select_box) = True
Else
    enemy(wall_select_box) = False
End If
End Sub

Public Sub Command1_Click()
nummelons = nummelons + 1
Load melon(nummelons)
melon(nummelons).Visible = True
activemelon = nummelons
selected = 0
End Sub

Private Sub Form_Load()
Load Form2
Form2.Show
offsetX = 0
offsetY = 0
numwalls = 1
selected = 0
wall_select_box.Text = selected
set_start = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If activemelon = 0 Then
    If set_start = False Then
        xstart = X
        ystart = Y
        player.Left = xstart
        player.Top = ystart
        If Button = 1 Then
            set_start = True
        Else
        End If
    Else
    
        If selected <> 0 Then
            If Button = 1 Then
                wall(selected).Top = Y
                wall(selected).Left = X
            Else
                If Button = 2 Then
                    If X < wall(selected).Left Then
                        wall(selected).Left = X
                    Else
                        If Y < wall(selected).Top Then
                            wall(selected).Top = Y
                        Else
                        End If
                        wall(selected).Width = ((X) - wall(selected).Left)
                        wall(selected).Height = ((Y) - wall(selected).Top)
                        wall(0).Width = ((X) - wall(selected).Left)
                        wall(0).Height = ((Y) - wall(selected).Top)
                    End If
                Else
                End If
            End If
            
        End If
    End If
Else
    melon(activemelon).Left = X
    melon(activemelon).Top = Y
    If Button = 1 Then
        activemelon = 0
    End If
End If

End Sub

Public Sub lay_new_wall_Click()
Check1.Value = 0
numwalls = numwalls + 1
selected = numwalls - 1
Load wall(numwalls - 1)
wall(numwalls - 1).Visible = True
wall_select_box.Text = selected
ReDim enemy(numwalls)
enemy(numwalls) = False
End Sub

Private Sub open_saved_Click()

FileName = "walls.txt"
FileNumber = FreeFile           'Freefile returns the first unused file number
Open FileName For Input As FileNumber   'Opens the file for input
length = LOF(FileNumber)        'LOF function returns length of file
If length < 32767 Then
    strreading = Input$(length, FileNumber)
Else
    MsgBox "This file is longer than 32Kb"
End If
Close FileNumber            'Close when done

'parse it
spacepos = InStr(strreading, ";")
length = Len(strreading)
data = Left$(strreading, spacepos - 1)
strreading = Right$(strreading, length - spacepos)

Number = data - 1
numwalls = data
For j = 1 To Number
    Load wall(j)
Next j

spacepos = InStr(strreading, ";")
length = Len(strreading)
data = Left$(strreading, spacepos - 1)
strreading = Right$(strreading, length - spacepos)

player.Left = data
set_start = True

spacepos = InStr(strreading, ";")
length = Len(strreading)
data = Left$(strreading, spacepos - 1)
strreading = Right$(strreading, length - spacepos)

player.Top = data

For j = 0 To Number
    spacepos = InStr(strreading, ";")
    length = Len(strreading)
    data = Left$(strreading, spacepos - 1)
    strreading = Right$(strreading, length - spacepos)
    wall(j).Left = data
    spacepos = InStr(strreading, ";")
    length = Len(strreading)
    data = Left$(strreading, spacepos - 1)
    strreading = Right$(strreading, length - spacepos)
    wall(j).Top = data
    spacepos = InStr(strreading, ";")
    length = Len(strreading)
    data = Left$(strreading, spacepos - 1)
    strreading = Right$(strreading, length - spacepos)
    wall(j).Width = data
    spacepos = InStr(strreading, ";")
    length = Len(strreading)
    data = Left$(strreading, spacepos - 1)
    strreading = Right$(strreading, length - spacepos)
    wall(j).Height = data
    wall(j).Visible = True
Next j

End Sub

Public Sub reset_player_Click()
set_start = False
End Sub

Public Sub save_walls_Click()

Set fso = New Scripting.FileSystemObject
Set txs = fso.CreateTextFile("lvl" & lvlnum & ".txt", True)
Dim data As String

data = player.Left & ";" & player.Top & ";" & numwalls - 1 & ";"

For j = 1 To numwalls - 1
    data = data & wall(j).Left & ";" & wall(j).Top & ";" & wall(j).Width & ";" & wall(j).Height & ";"
Next j

data = data & nummelons & ";"

For j = 1 To nummelons
    data = data & melon(j).Left & ";" & melon(j).Top & ";"
Next j

For j = 1 To numwalls - 1
    If enemy(j) = True Then
        data = data & j & ";" & "true" & ";"
    Else
        data = data & j & ";" & "false" & ";"
    End If
Next j

txs.WriteLine (data)

End Sub

Public Sub select_wall_Click()
selected = wall_select_box.Text
For j = 0 To numwalls - 1
    wall(j).BorderColor = &H0&
Next j
wall(selected).BorderColor = &HFF00&
End Sub
