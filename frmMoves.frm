VERSION 5.00
Begin VB.Form frmMoves 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Moves"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "frmMoves.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lst 
      Appearance      =   0  'Flat
      Height          =   1395
      ItemData        =   "frmMoves.frx":0A02
      Left            =   0
      List            =   "frmMoves.frx":0A12
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmMoves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
