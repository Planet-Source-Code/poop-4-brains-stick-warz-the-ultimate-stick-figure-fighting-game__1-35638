VERSION 5.00
Begin VB.Form frmSpeed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Speed Control"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   Icon            =   "frmSpeed.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Pauseit 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.HScrollBar sp 
      Height          =   255
      LargeChange     =   1000
      Left            =   240
      Max             =   10000
      SmallChange     =   100
      TabIndex        =   0
      Top             =   480
      Value           =   1
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Faster"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Slower"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmSpeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Speed = sp.Value
If Speed > 10000 Then Speed = 10000
If Speed < 0 Then Speed = 0
Unload Me
End Sub

Private Sub Form_Load()
sp.Value = Speed
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Pauseit.Text = "T" Then Paused = False
End Sub

