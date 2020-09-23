VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text Effects"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optColored 
      Caption         =   "Colored"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
   End
   Begin VB.OptionButton optHacker 
      Caption         =   "HaCkEr"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.OptionButton optElite 
      Caption         =   "£|ïTë"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.OptionButton optRevSp 
      Caption         =   "Reverse/Space"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.OptionButton optStar 
      Caption         =   "Starred"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.OptionButton optSpace 
      Caption         =   "Space"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.OptionButton optReverse 
      Caption         =   "Reverse"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do It!"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "What To Effect..."
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.Text = "" 'it MUST always blank out text2
If optReverse.Value = True Then GoTo Reverse 'goes to specified effect...
If optSpace.Value = True Then GoTo Space 'see above
If optStar.Value = True Then GoTo Star 'see above
If optRevSp.Value = True Then GoTo RevSpace 'see above
If optElite.Value = True Then GoTo Elite 'see above
If optHacker.Value = True Then GoTo Hacker 'see above
If optColored.Value = True Then GoTo Colored: 'see above
Reverse:
NewString = StrReverse(Text1.Text) 'StrReverse is a VB keyword - it reverses a string!
Text2.Text = NewString 'puts reversed string in text2
Exit Sub
Space:
For i = 1 To Len(Text1) 'loop once for each letter
Char = Mid(Text1, i, 1) 'each loop, it stores the next character from the
                        'beginning of the string into "Char" and the next line
                        'of code adds it to text2 with an extra space. (" ")
Text2.Text = Text2.Text + (Char + " ")
Next i
Exit Sub
Star:
Text2.Text = "*"
For i = 1 To Len(Text1)
Char = Mid(Text1, i, 1)
Text2.Text = Text2.Text + (Char + "*") 'simply adds a star to Char
Next i
Exit Sub
RevSpace:
For i = 1 To Len(Text1)
XYZ = Mid(Text1, Len(Text1) + 1 - i, 1) 'gets last letter
Text2.Text = Text2.Text + (XYZ + " ") 'adds a space to it and puts it as first letter in text2
Next i
Exit Sub
Elite:
For i = 1 To Len(Text1)
strnew = Mid(Text1, i, 1)
 If strnew = "a" Then strnew = "ã" ' changes letters to cool letters
 If strnew = "A" Then strnew = "Ä" 'you can customize these
 If strnew = "b" Then strnew = "b"
 If strnew = "B" Then strnew = "ß"
 If strnew = "c" Then strnew = "ç"
 If strnew = "C" Then strnew = "Ç"
 If strnew = "d" Then strnew = "ð"
 If strnew = "D" Then strnew = "Ð"
 If strnew = "e" Then strnew = "ë"
 If strnew = "E" Then strnew = "£"
 If strnew = "f" Then strnew = "ƒ"
 If strnew = "F" Then strnew = "F"
 If strnew = "g" Then strnew = "g"
 If strnew = "G" Then strnew = "G"
 If strnew = "h" Then strnew = "h"
 If strnew = "H" Then strnew = "H"
 If strnew = "i" Then strnew = "ï"
 If strnew = "I" Then strnew = "Î"
 If strnew = "j" Then strnew = "J"
 If strnew = "J" Then strnew = "¿"
 If strnew = "k" Then strnew = "l‹"
 If strnew = "K" Then strnew = "\<"
 If strnew = "l" Then strnew = "|"
 If strnew = "L" Then strnew = "(_"
 If strnew = "m" Then strnew = "m"
 If strnew = "M" Then strnew = "/V\"
 If strnew = "n" Then strnew = "ñ"
 If strnew = "N" Then strnew = "Ñ"
 If strnew = "o" Then strnew = "ø"
 If strnew = "O" Then strnew = "Õ"
 If strnew = "p" Then strnew = "Þ"
 If strnew = "P" Then strnew = "þ"
 If strnew = "q" Then strnew = "q"
 If strnew = "Q" Then strnew = "Ø"
 If strnew = "r" Then strnew = "R"
 If strnew = "R" Then strnew = "r"
 If strnew = "s" Then strnew = "š"
 If strnew = "S" Then strnew = "Š"
 If strnew = "t" Then strnew = "†"
 If strnew = "T" Then strnew = "t"
 If strnew = "u" Then strnew = "ú"
 If strnew = "U" Then strnew = "Ü"
 If strnew = "v" Then strnew = "V"
 If strnew = "V" Then strnew = "\/"
 If strnew = "w" Then strnew = "vv"
 If strnew = "W" Then strnew = "VV "
 If strnew = "x" Then strnew = "X"
 If strnew = "X" Then strnew = "><"
 If strnew = "y" Then strnew = "ÿ"
 If strnew = "Y" Then strnew = "¥"
 If strnew = "z" Then strnew = "Z"
 If strnew = "Z" Then strnew = "z"
Text2.Text = Text2.Text + strnew
Next i
Exit Sub
Hacker:
'still working on this one...mail me if you know a good way!
Exit Sub
Colored: 'this effect just generates a bunch of random numbers and does a bunch of crazy crap that I can't really explain...
Ln = Len(Text1)
For i = 1 To Ln
ColorChar = Mid(Text1, i, 1)
Text2.Text = Text2.Text + ColorChar
Next i
For i = 1 To 300
rannum = Int(Rnd * 99)
Form1.BackColor = RGB(i, i * rannum, i + Int(Rnd * 500))
Next i
Exit Sub
End Sub

Private Sub Form_Load()
Randomize 'makes each set of rands different
End Sub
