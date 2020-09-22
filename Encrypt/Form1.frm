VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "ENCRYPT/DECRYPT"
      Height          =   975
      Left            =   2400
      TabIndex        =   1
      Top             =   4800
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   3855
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   480
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i As Double
    Dim Char As String
    Dim NewStr As String
    For i = 1 To Len(Text1)
        Char = Chr(Asc(Mid(Text1, i, 1)) Xor 17.7447446868468)
        NewStr = NewStr & Char
    Next
    Text1 = NewStr
End Sub

'This code is very basic. First we have the For/Next loop going through
'each character in the Text1 control. Then to narrow down the next part
'to make it simpler for you it would look like this

'Char = Mid(Text, i, 1)
'Char = Asc(Char)
'Char = Char Xor 17.7447446868468)
'Char = Chr(Char)

'The first line is getting the next character to encrypt.
'next it is converting that character to ascii
'then using the Xor operator to mix it around a bit
'now converting it back to a charcter.

'The Xor function is good because it doesn't need a seperate routine
'to decrpyt.

'You can mix it around a bit if you like, try changing the
'"Xor 17.7447446868468" to "+ 1" this will make it use the next
'letter in the alphabet for each character. using - 1 will
'use the previous letter for each character. If you use + or -
'instead of Xor you will need a decrypt function that does the opposite.

'Example

'To Encrypt use
'Char = Chr(Asc(Mid(Text1, i, 1)) + 1)

'To Decrypt use
'Char = Chr(Asc(Mid(Text1, i, 1)) - 1)


'Thanks for reading my tutorial, therefore visiting my site.

