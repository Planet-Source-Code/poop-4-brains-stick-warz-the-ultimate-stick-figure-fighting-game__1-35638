VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stick Warz"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   705
   ClientWidth     =   5520
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   251
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   368
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox BM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   7800
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox BS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   7560
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox Map 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3765
      Left            =   960
      ScaleHeight     =   251
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   368
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   5520
   End
   Begin VB.PictureBox S 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   6720
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   7800
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox MapS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   3765
      Left            =   6600
      ScaleHeight     =   251
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   368
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   5520
   End
   Begin VB.PictureBox Board 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Myriad Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   0
      ScaleHeight     =   251
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   368
      TabIndex        =   5
      Top             =   0
      Width           =   5520
      Begin VB.Frame frmWorld 
         Caption         =   "New Game Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   3855
         Begin VB.TextBox tNick 
            Height          =   285
            Index           =   1
            Left            =   960
            MaxLength       =   8
            TabIndex        =   22
            Text            =   "Player2"
            Top             =   2280
            Width           =   1095
         End
         Begin VB.TextBox tNick 
            Height          =   285
            Index           =   0
            Left            =   960
            MaxLength       =   8
            TabIndex        =   21
            Text            =   "Player1"
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CheckBox chkAI 
            Caption         =   "Human Control"
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   18
            Top             =   2280
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox chkAI 
            Caption         =   "Human Control"
            Height          =   255
            Index           =   0
            Left            =   2160
            TabIndex        =   16
            Top             =   1920
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1920
            TabIndex        =   14
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Start Game"
            Height          =   255
            Left            =   600
            TabIndex        =   13
            Top             =   2760
            Width           =   1335
         End
         Begin VB.PictureBox photo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1050
            Left            =   2160
            ScaleHeight     =   1050
            ScaleWidth      =   1050
            TabIndex        =   10
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   ">>>"
            Height          =   255
            Left            =   1920
            TabIndex        =   8
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CommandButton cmdBack 
            Caption         =   "<<<"
            Height          =   255
            Left            =   720
            TabIndex        =   9
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Navigation:"
            Height          =   255
            Left            =   600
            TabIndex        =   20
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            Height          =   255
            Left            =   600
            TabIndex        =   19
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Player2 "
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Player1"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label lblLLevel 
            BackStyle       =   0  'Transparent
            Caption         =   "Very Easy"
            Height          =   255
            Left            =   1080
            TabIndex        =   12
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblLName 
            BackStyle       =   0  'Transparent
            Caption         =   "Jungle"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   11
            Top             =   360
            Width           =   975
         End
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "Menu"
      Begin VB.Menu mnuStart 
         Caption         =   "Start Game"
      End
      Begin VB.Menu mnuP 
         Caption         =   "Pause\Unpause"
      End
      Begin VB.Menu mnuEX 
         Caption         =   "Exit Game"
      End
   End
   Begin VB.Menu mnuOp 
      Caption         =   "Options"
      Begin VB.Menu mnuMoves 
         Caption         =   "Moves"
      End
      Begin VB.Menu mnuKeys 
         Caption         =   "Keys"
      End
      Begin VB.Menu mnuSpeed 
         Caption         =   "Speed"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Paused
Dim C As Long

Private Sub cmdBack_Click()
WorldI = WorldI - 1: If WorldI < 1 Then WorldI = MaxWorlds
UpdateWorld
End Sub

Private Sub cmdNext_Click()
WorldI = WorldI + 1: If WorldI > MaxWorlds Then WorldI = 1
UpdateWorld
End Sub

Function UpdateWorld()
On Error Resume Next
lblLName.Caption = Split(GetLineFromFile(WorldI, App.path + "\Gfx\levels.dat"), ",")(0)
lblLLevel.Caption = Split(GetLineFromFile(WorldI, App.path + "\Gfx\levels.dat"), ",")(1)

LoadPic photo, "photo" & WorldI
LoadPic Board, "background" & WorldI
Main_Loop
LoadPic Map, "back_mask" & WorldI
LoadPic MapS, "back_sprite" & WorldI
End Function

Private Sub Command2_Click()
Dim I As Long

frmWorld.Visible = False

Map.Cls
MapS.Cls

For I = 0 To 255 Step 5
C = 0
Board.Line (0, 0)-(Board.ScaleWidth, Board.ScaleHeight), RGB(I, I, I), BF
DoEvents
Next I

ClearShots

For I = 1 To UBound(P())
InitPlayer P(I), IIf(I = 1, 80, 250)
P(I).ID = I
If chkAI(I - 1).Value = 1 Then
P(I).AI = False
Else
P(I).AI = True
End If
P(I).Nick = tNick(I - 1).Text
Next I

Paused = False
Fighting = True
Flashed = False
End Sub

Private Sub Command3_Click()
frmWorld.Visible = False
End Sub

Private Sub Form_Load()
On Error Resume Next
MaxWorlds = GetFileLines(App.path + "\GFX\levels.dat")

Dim K As String, HJ() As String
Open App.path + "\Gfx\newgame.dat" For Input As #1
Line Input #1, K
HJ() = Split(Mid(K, 3), ",")
WorldI = Val(HJ(0))
UpdateWorld
tNick(0).Text = HJ(1)
tNick(1).Text = HJ(2)
chkAI(0).Value = Val(HJ(3))
chkAI(1).Value = Val(HJ(4))
Speed = Val(HJ(5))
Close #1

Randomize

LoadPic M, "p_mask"
LoadPic S, "p_sprite"
LoadPic BS, "fire_s"
LoadPic BM, "fire_m"

GetAsyncKeyState (0)

Me.Show

Do
If GetTickCount > 300 Then
If C > (10000 - Speed) Then
C = 0
Main_Loop
Else
C = C + 1
End If
End If
DoEvents
Loop
End Sub

Function LoadPic(picbox As PictureBox, pic)
On Error Resume Next
picbox.Picture = LoadPicture(App.path + "\Gfx\" & pic & ".bmp")
End Function

Private Sub Form_Unload(Cancel As Integer)
Open App.path + "\Gfx\newgame.dat" For Binary As #1
Dim K
If Speed > 10000 Then Speed = 10000
If Speed < 0 Then Speed = 0
K = WorldI & "," & tNick(0).Text & "," & tNick(1).Text & "," & chkAI(0).Value & "," & chkAI(1).Value & "," & Speed
Put #1, , K
Close #1
End
End Sub



Private Sub mnuEX_Click()
Unload Me
End Sub

Private Sub mnuKeys_Click()
frmKeys.Show
End Sub

Private Sub mnuMoves_Click()
frmMoves.Show
End Sub

Private Sub mnuP_Click()
Select Case Paused
Case True: Paused = False
Case False: Paused = True
End Select
End Sub

Private Sub mnuSpeed_Click()
frmSpeed.Show
frmSpeed.Pauseit.Text = ""
If NumOfPlayers > 0 Then
If Paused = True Then frmSpeed.Pauseit.Text = "T"
Paused = True
End If
End Sub

Private Sub mnuStart_Click()
If NumOfPlayers <> 0 Then mnuP_Click
frmWorld.Visible = True
End Sub

Public Sub Main_Loop()
Dim I As Long, drawx, G As Long, HI As Long
Board.Cls

If Paused = True Then
Board.Line (0, 0)-(Board.ScaleWidth, Board.ScaleHeight), vbBlack, BF
Board.FontSize = 12
Board.CurrentX = Board.ScaleWidth \ 2 - Board.TextWidth("Paused") \ 2
Board.CurrentY = Board.ScaleHeight \ 2 - Board.TextHeight("|") \ 2
Board.ForeColor = vbWhite
Board.Print "Paused"
Exit Sub
ElseIf NumOfPlayers = 0 Then
If Fighting Then

If Flashed = False Then
For G = 255 To 0
Board.Line (0, 0)-(Board.ScaleWidth, Board.ScaleHeight), RGB(G, G, G), BF
For HI = 0 To 100
DoEvents
Next HI
Next G
Flashed = True
End If

Board.Line (0, 0)-(Board.ScaleWidth, Board.ScaleHeight), vbBlack, BF
Board.FontSize = 26
Board.CurrentX = Board.ScaleWidth \ 2 - Board.TextWidth("Its a Tie!") \ 2
Board.CurrentY = Board.ScaleHeight \ 2 - Board.TextHeight("|") \ 2
Board.ForeColor = vbWhite
Board.Print "Its a Tie!"
Board.FontSize = 12
Board.CurrentX = Board.ScaleWidth \ 2 - Board.TextWidth("Press Start Game for New game") \ 2
Board.CurrentY = Board.ScaleHeight \ 2 + (Board.TextHeight("|") * 1.5)
Board.ForeColor = vbWhite
Board.Print "Press Start Game for New game"
Else
Board.Line (0, 0)-(Board.ScaleWidth, Board.ScaleHeight), vbBlack, BF
Board.FontSize = 26
Board.CurrentX = Board.ScaleWidth \ 2 - Board.TextWidth("Stick Warz") \ 2
Board.CurrentY = Board.ScaleHeight \ 2 - Board.TextHeight("|") \ 2
Board.ForeColor = vbWhite
Board.Print "Stick Warz"
Board.FontSize = 12
Board.CurrentX = Board.ScaleWidth \ 2 - Board.TextWidth("Press Start Game for New game") \ 2
Board.CurrentY = Board.ScaleHeight \ 2 + (Board.TextHeight("|") * 1.5)
Board.ForeColor = vbWhite
Board.Print "Press Start Game for New game"
End If

ElseIf NumOfPlayers = 1 Then

If Flashed = False Then
For G = 255 To 0
Board.Line (0, 0)-(Board.ScaleWidth, Board.ScaleHeight), RGB(G, G, G), BF
For HI = 0 To 100
DoEvents
Next HI
Next G
Flashed = True
End If

Board.Line (0, 0)-(Board.ScaleWidth, Board.ScaleHeight), vbBlack, BF
Board.FontSize = 26
Board.CurrentY = Board.ScaleHeight \ 2 - Board.TextHeight("|") \ 2
Board.ForeColor = vbWhite

Fighting = False

Select Case P(1).Act
Case True
Board.CurrentX = Board.ScaleWidth \ 2 - Board.TextWidth(P(1).Nick & " has won!") \ 2
Board.Print P(1).Nick & " has won!"
Case False
Board.CurrentX = Board.ScaleWidth \ 2 - Board.TextWidth(P(2).Nick & " has won!") \ 2
Board.Print P(2).Nick & " has won!"
End Select

Board.FontSize = 12
Board.CurrentX = Board.ScaleWidth \ 2 - Board.TextWidth("Press Start Game for New game") \ 2
Board.CurrentY = Board.ScaleHeight \ 2 + (Board.TextHeight("|") * 1.5)
Board.ForeColor = vbWhite
Board.Print "Press Start Game for New game"

Else

BitBlt Board.hdc, 0, 0, Board.ScaleWidth, Board.ScaleHeight, Map.hdc, 0, 0, vbSrcAnd
BitBlt Board.hdc, 0, 0, Board.ScaleWidth, Board.ScaleHeight, MapS.hdc, 0, 0, vbSrcInvert

If P(1).AI = False Then
DoKeys P(1), vbKeyW, vbKeyA, vbKeyD, vbKeyS, vbKeyShift, vbKeyZ, vbKeyX, vbKeyC
Else
DoAI P(1)
End If

If P(2).AI = False Then
DoKeys P(2), vbKeyNumpad8, vbKeyNumpad4, vbKeyNumpad6, vbKeyNumpad5, vbKeyReturn, vbKeyNumpad1, vbKeyNumpad2, vbKeyNumpad3
Else
DoAI P(2)
End If

For I = 1 To UBound(P())
MovePlayer P(I)
DrawPlayer P(I)

Select Case I
Case 1
drawx = 4
Case 2
drawx = Board.ScaleWidth - 54
End Select

Board.Line (drawx, 4)-(drawx + 50, 14), vbRed, BF
Board.Line (drawx, 4)-(drawx + P(I).HP \ 4, 9), vbGreen, BF
Board.Line (drawx, 9)-(drawx + (P(I).MP \ 10), 14), vbBlue, BF
Next I

MoveShots
DrawShots
End If
End Sub
