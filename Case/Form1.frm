VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim Name As String
    Name = InputBox("Enter your name")
    
    If LCase$(Trim$(Name)) = "john" Then
        MsgBox "Hello John, You are a valid user"
        ElseIf LCase$(Trim$(Name)) = "paul" Then
        MsgBox "Hey Paul, access granted"
        ElseIf LCase$(Trim$(Name)) = "adam" Then
        MsgBox "Hey Adam, access granted"
        ElseIf LCase$(Trim$(Name)) = "rob" Then
        MsgBox "Hey Rob, access granted"
        Else
        MsgBox "You are not a valid user, " & Name
    End If
End Sub

