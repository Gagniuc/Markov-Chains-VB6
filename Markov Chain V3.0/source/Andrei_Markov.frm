VERSION 5.00
Begin VB.Form About 
   Caption         =   "Andrei A. Markov"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11685
   LinkTopic       =   "Form2"
   ScaleHeight     =   5880
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton OUT 
      Caption         =   "Ok"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00404040&
      Height          =   4935
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Andrei_Markov.frx":0000
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1856 - 1922"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5160
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   4890
      Left            =   240
      Picture         =   "Andrei_Markov.frx":0787
      Top             =   120
      Width           =   4230
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OUT_Click()
Unload Me
End Sub
