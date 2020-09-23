Attribute VB_Name = "Module1"
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetTickCount Lib "kernel32" () As Long

Public MaxWorlds

'sprite's dimensions
Public Const PLAY_W = 30
Public Const PLAY_H = 30

'diff player pos's (as constants)
Public Const POS_STEP = 0
Public Const POS_FALL = PLAY_H * 2
Public Const POS_JUMP = PLAY_H * 3
Public Const POS_FIRE = PLAY_H * 4
Public Const POS_LEAN = PLAY_H * 5
Public Const POS_FALLEN = PLAY_H * 6
Public Const POS_GETUP = PLAY_H * 7
Public Const POS_KICK = PLAY_H * 8
Public Const POS_PUNCH = PLAY_H * 9
Public Const POS_BLAST = PLAY_H * 10

'sides player face
Public Const PLAY_LEFT = PLAY_W
Public Const PLAY_RIGHT = 0

Public Type Player
X As Double
y As Double
XS As Double
YS As Double

STEP As Long
StepC As Long

Ani As Long 'step, fall, jump, or fire
AniL As Long 'length that the ani lasts

Dirc As Long

Act As Boolean
ID As Long
AI As Boolean

HP As Long
MP As Long
Rel As Long

OnGround As Integer

Nick As String
End Type

Type Shot
X As Double
y As Double

XS As Double
YS As Double

Act As Long
End Type

Public P(1 To 2) As Player
Public S(0 To 20) As Shot
Public Fighting As Boolean
Public Flashed As Boolean
Public Speed As Long

Public WorldI As Long

Function ClearShots()
Dim I As Long
For I = 0 To 20
S(I).Act = False
Next I
End Function

Function GetLineFromFile(line, path) As String
On Error GoTo Ack
Dim I As Long
Open path For Input As #1
For I = 1 To line
Line Input #1, GetLineFromFile
Next I
Close #1
Ack:
End Function

Function GetFileLines(path) As Double
Dim K
Open path For Input As #1
Do Until EOF(1)
Line Input #1, K
GetFileLines = GetFileLines + 1
Loop
Close #1
End Function

Function mBlast(ID)
If P(ID).MP < 500 Then Exit Function
If P(ID).Ani = POS_BLAST Then Exit Function

P(ID).Ani = POS_BLAST
P(ID).AniL = 410

Dim I As Double, B, BX, BY
Form1.Main_Loop

For I = 0 To 10000
DoEvents
Next I

For I = 6 To 400 Step 3
C = 0

BX = Int(Rnd * Form1.Board.ScaleWidth)
BY = Int(Rnd * Form1.Board.ScaleHeight)

Form1.Map.DrawWidth = 10
Form1.MapS.DrawWidth = 10
Form1.Map.PSet (BX, BY), vbWhite
Form1.MapS.PSet (BX, BY), vbBlack
Form1.Map.DrawWidth = 1
Form1.MapS.DrawWidth = 1

Form1.Board.DrawWidth = 2
Form1.Board.Circle (P(ID).X + 15, P(ID).y + 15), I - 0.5, RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
Form1.Board.Circle (P(ID).X + 15, P(ID).y + 15), I, RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))

For B = 0 To 50
DoEvents
Next B

If P(ID).Ani <> POS_BLAST Then Exit Function

Next I

Form1.Board.DrawWidth = 1

P(ID).HP = P(ID).HP - 25
P(ID).MP = 0

For I = 1 To UBound(P())
If P(I).ID <> ID Then P(I).HP = P(I).HP - 150
Next I

P(ID).Ani = POS_STEP
P(ID).AniL = 0
End Function

Function mKick(ID)
Dim I As Long

If P(ID).Ani = POS_KICK Then Exit Function

P(ID).Ani = POS_KICK
P(ID).AniL = 3

For I = 1 To UBound(P())
If I <> ID And P(I).Act = True And PCollision(ID, I) = True Then
P(I).HP = P(I).HP - 5
P(I).Ani = POS_FALLEN
P(I).AniL = 4
End If
Next I
End Function

Function mPunch(ID)
Dim I As Long

If P(ID).Ani = POS_PUNCH Then Exit Function

P(ID).Ani = POS_PUNCH
P(ID).AniL = 3

For I = 1 To UBound(P())
If I <> ID And P(I).Act = True And PCollision(ID, I) = True Then
P(I).HP = P(I).HP - 2
P(I).Ani = POS_FALLEN
P(I).AniL = 2
End If
Next I
End Function

Function InitPlayer(P As Player, X)
P.Act = True
P.STEP = 0
P.Ani = 0
P.AniL = 0
P.Dirc = PLAY_LEFT
P.X = X
P.y = 5
P.YS = 1
P.XS = 0
P.HP = 200
P.MP = 250
P.Nick = ""
End Function

Function MovePlayer(P As Player)
Dim Side1, Side2, I, Blasting, BlastI

If P.HP <= 0 Then P.Act = False
If P.y > Form1.Board.ScaleHeight Then P.Act = False

If P.Ani = POS_BLAST Then Blasting = True: BlastI = P.AniL

P.Rel = P.Rel - 1: If P.Rel < 0 Then P.Rel = 0
P.MP = P.MP + 2: If P.MP > 500 Then P.MP = 500

If P.StepC <= 0 Then
P.StepC = 0
P.STEP = 0
Else
P.StepC = P.StepC - 1
End If

P.X = P.X + P.XS
P.y = P.y + P.YS

P.YS = P.YS + 0.2
P.XS = P.XS * 0.1

If P.X < -5 Then P.X = -5: P.XS = 0
If P.X > Form1.Board.ScaleWidth - 25 Then: P.X = Form1.Board.ScaleHeight - 25: P.XS = 0
 
If P.Ani = POS_FALL Or P.Ani = POS_JUMP Or P.Ani = POS_STEP Then If P.YS > 1 Then P.Ani = POS_FALL: P.AniL = 2: P.OnGround = False
If P.Ani = POS_FALL Or P.Ani = POS_JUMP Or P.Ani = POS_STEP Then If P.YS < -0.5 Then P.Ani = POS_JUMP: P.AniL = 2: P.OnGround = False

Side1 = Form1.Map.Point(P.X + 5, P.y + 15) = vbBlack
Side2 = Form1.Map.Point(P.X + 25, P.y + 15) = vbBlack

If Side1 = True Then P.X = P.X + 2
If Side2 = True Then P.X = P.X - 2

Side1 = Form1.Map.Point(P.X + 10, P.y + 27) = vbBlack
Side2 = Form1.Map.Point(P.X + 20, P.y + 27) = vbBlack

If Side1 = True And Side2 = True Then

P.OnGround = True
For I = P.y To 0 Step -1
P.y = I
Side1 = Form1.Map.Point(P.X + 20, P.y + 25) = vbBlack
Side2 = Form1.Map.Point(P.X + 10, P.y + 25) = vbBlack
If Side1 = False And Side2 = False Then Exit For
Next I

P.YS = -(P.YS * 0.2)
End If

If P.AniL <= 0 Then
P.AniL = 0

If P.Ani = POS_BLAST Then Blasting = False: BlastI = 0

If P.Ani = POS_FALLEN Then
P.Ani = POS_GETUP
P.AniL = 5
Else
P.Ani = POS_STEP
End If

Else
P.AniL = P.AniL - 1
End If

If Blasting = True Then P.Ani = POS_BLAST: P.AniL = BlastI
End Function

Function DrawShots()
Dim I As Long
For I = 0 To 20
If S(I).Act = True Then
BitBlt Form1.Board.hdc, S(I).X, S(I).y, 15, 15, Form1.BM.hdc, 0, 0, vbSrcAnd
BitBlt Form1.Board.hdc, S(I).X, S(I).y, 15, 15, Form1.BS.hdc, 0, 0, vbSrcInvert
End If
Next I
End Function

Function Shoot(P As Player)
Dim I As Long

If P.Rel > 0 Then Exit Function
If P.MP < 75 Then Exit Function

For I = 0 To 20
If S(I).Act = False Then

Select Case P.Dirc
Case PLAY_LEFT
S(I).XS = -5
S(I).X = P.X - 13
Case PLAY_RIGHT
S(I).XS = 5
S(I).X = P.X + 27
End Select

P.MP = P.MP - 75
P.Rel = 3

P.Ani = POS_FIRE
P.AniL = 10

Select Case P.Dirc
Case PLAY_LEFT
P.XS = 3
Case PLAY_RIGHT
P.XS = -3
End Select

S(I).y = P.y + 5
S(I).YS = 0

S(I).Act = True

Exit For
End If
Next I
End Function

Function MoveShots()
Dim I As Long, C
For I = 0 To 20
If S(I).Act = True Then
S(I).X = S(I).X + S(I).XS
S(I).y = S(I).y + S(I).YS

If CollisionDetect(0, 0, Form1.Board.Width, Form1.Board.Height, 0, 0, Form1.Map.hdc, S(I).X, S(I).y, 15, 15, 0, 0, Form1.BM.hdc) = True Then
S(I).Act = False
Form1.Map.DrawWidth = 10
Form1.MapS.DrawWidth = 10
Form1.MapS.PSet (S(I).X + 7.5, S(I).y + 7.5), vbBlack
Form1.Map.PSet (S(I).X + 7.5, S(I).y + 7.5), vbWhite
Form1.Map.DrawWidth = 1
Form1.MapS.DrawWidth = 1
End If

If S(I).y < -15 Or S(I).X < -15 Or S(I).X > Form1.Board.ScaleWidth Then S(I).Act = False

For C = 1 To UBound(P())
If CollisionDetect(P(C).X, P(C).y, 30, 30, P(C).Dirc, P(C).Ani, Form1.M.hdc, S(I).X, S(I).y, 15, 15, 0, 0, Form1.BM.hdc) = True Then
S(I).Act = False
P(C).HP = P(C).HP - 10
P(C).Ani = POS_FALLEN
P(C).AniL = 5
Exit For
End If
Next C

End If
Next I
End Function

Function PCollision(ID, ID2)
PCollision = CollisionDetect(P(ID).X, P(ID).y, 30, 30, P(ID).Dirc, P(ID).Ani, Form1.M.hdc, P(ID2).X, P(ID2).y, 30, 30, P(ID2).Dirc, P(ID2).Ani, Form1.M.hdc) = True
End Function

Function DoKeys(P As Player, Up, Left, Right, Down, Fire, Kick, Punch, Blast)

If GetAsyncKeyState(Left) Then
P.Dirc = PLAY_LEFT
P.XS = -2
P.STEP = P.STEP + 1: If P.STEP > 1 Then P.STEP = 0
P.StepC = 2
End If

If GetAsyncKeyState(Right) Then
P.Dirc = PLAY_RIGHT
P.XS = 2
P.STEP = P.STEP + 1: If P.STEP > 1 Then P.STEP = 0
P.StepC = 2
End If

If GetAsyncKeyState(Fire) Then
Shoot P
End If

If GetAsyncKeyState(Up) Then
If P.OnGround = True Then
P.YS = -3
End If
End If

If GetAsyncKeyState(Kick) Then mKick (P.ID)
If GetAsyncKeyState(Punch) Then mPunch (P.ID)
If GetAsyncKeyState(Blast) Then mBlast (P.ID)
End Function

Function DrawPlayer(P As Player)
If P.Act = False Then Exit Function
Select Case P.Ani
Case 0
BitBlt Form1.Board.hdc, P.X, P.y, PLAY_W, PLAY_H, Form1.M.hdc, P.Dirc, P.STEP * PLAY_H, vbSrcAnd
BitBlt Form1.Board.hdc, P.X, P.y, PLAY_W, PLAY_H, Form1.S.hdc, P.Dirc, P.STEP * PLAY_H, vbSrcInvert
Case Else
BitBlt Form1.Board.hdc, P.X, P.y, PLAY_W, PLAY_H, Form1.M.hdc, P.Dirc, P.Ani, vbSrcAnd
BitBlt Form1.Board.hdc, P.X, P.y, PLAY_W, PLAY_H, Form1.S.hdc, P.Dirc, P.Ani, vbSrcInvert
End Select
Form1.Board.FontSize = 5
Form1.Board.CurrentX = P.X + 15 - Form1.Board.TextWidth(P.Nick) \ 2
Form1.Board.CurrentY = P.y - Form1.Board.TextHeight("|")
Form1.Board.ForeColor = vbRed
Form1.Board.Print P.Nick
End Function

Function NumOfPlayers() As Integer
Dim I As Long
For I = 1 To UBound(P())
If P(I).Act = True Then NumOfPlayers = NumOfPlayers + 1
Next I
End Function

Function DoAI(D As Player)
Dim Target

If Int(Rnd * 10) = Int(Rnd * 5) Then

Select Case D.ID
Case 1: Target = 2
Case 2: Target = 1
End Select

If P(Target).y < D.y Then
If D.OnGround = True Then
D.YS = -3
End If
End If

If Int(Rnd * 3) = 1 Then
Select Case P(Target).X
Case Is < D.X + 15
D.XS = -5
D.Dirc = PLAY_LEFT
D.STEP = D.STEP + 1
D.StepC = 5
If D.STEP > 1 Then D.STEP = 0
Case Is > D.X + 15
D.XS = 5
D.Dirc = PLAY_RIGHT
D.STEP = D.STEP + 1
D.StepC = 5
If D.STEP > 1 Then D.STEP = 0
End Select
End If

If Int(Rnd * 20) = 12 And D.MP = 500 Then mBlast D.ID

If Int(Rnd * 7) = 4 Then

Dim tX, tY, pX, pY

pX = D.X + 15
pY = D.y + 15
tX = P(Target).X + 15
tY = P(Target).y + 15

If pX > tX - 40 And pX < tX + 55 And pY > tY - 45 And pY < tY + 70 Then

Select Case Int(Rnd * 3)
Case 0
mPunch D.ID
Case 1
mKick D.ID
Case 2
If D.MP >= 250 Then Shoot D
End Select
End If

End If

End If
End Function

