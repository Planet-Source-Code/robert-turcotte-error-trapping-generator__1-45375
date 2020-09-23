VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   Caption         =   "About"
   ClientHeight    =   705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3855
   LinkTopic       =   "Form2"
   ScaleHeight     =   705
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   645
      Left            =   2385
      TabIndex        =   1
      Top             =   45
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Celestor@hotmail.com"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      Top             =   315
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "Made By Himura Battousai"
      Height          =   240
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   2265
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Me.Hide
End Sub

