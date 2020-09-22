VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtReplace 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Replace String"
      Top             =   600
      Width           =   8175
   End
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Search String"
      Top             =   120
      Width           =   8175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "REPLACE"
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   5520
      Width           =   2535
   End
   Begin VB.TextBox txtString 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   1200
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    txtString.Text = Replace$(txtString, txtSearch.Text, txtReplace.Text)
End Sub
