VERSION 5.00
Begin VB.Form frmKeys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keys"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "frmKeys.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Player2"
      Height          =   1815
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   2535
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   1395
         ItemData        =   "frmKeys.frx":0A02
         Left            =   120
         List            =   "frmKeys.frx":0A1B
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Player1"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   1395
         ItemData        =   "frmKeys.frx":0AA3
         Left            =   120
         List            =   "frmKeys.frx":0ABC
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
