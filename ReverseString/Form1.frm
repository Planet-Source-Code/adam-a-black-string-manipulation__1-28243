VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Enter text to reverse"
      Top             =   240
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&REVERSE"
      Default         =   -1  'True
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function StringReverse(Str As String) As String
    Dim i As Integer
    For i = Len(Str) To 1 Step -1
        StringReverse = StringReverse & Mid(Str, i, 1)
    Next
End Function

Private Sub Command1_Click()
    Text1 = StringReverse(Text1)
End Sub
