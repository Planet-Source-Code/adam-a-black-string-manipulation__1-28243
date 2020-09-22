VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Search String"
      Top             =   120
      Width           =   8175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CALCULATE"
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   5040
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
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   720
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error GoTo iNext
    
    Dim i As Long
    Dim Count As Long
    Dim ExitLoop As Boolean
    
    ExitLoop = False
    Count = 0
    i = 1
    
    Do
        i = InStr(i, LCase(txtString), LCase(txtSearch))
        If Not i = 0 Then
            Count = Count + 1
            i = i + 1
            ElseIf i = 0 Then
            GoTo iNext
        End If
    Loop
    
iNext:
If Count > 1 Or Count = 0 Then
    MsgBox "There were " & Count & " words that matched your query"
    Else
    MsgBox "There was " & Count & " word that matched your query"
End If
End Sub
