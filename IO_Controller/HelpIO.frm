VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7770
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "OK"
      Height          =   330
      Left            =   6405
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5145
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   $"HelpIO.frx":0000
      Height          =   3690
      Left            =   6405
      TabIndex        =   0
      Top             =   105
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   5580
      Left            =   0
      Picture         =   "HelpIO.frx":0116
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6315
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
