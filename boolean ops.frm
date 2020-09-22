VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Bit Operations"
   ClientHeight    =   3945
   ClientLeft      =   5115
   ClientTop       =   5445
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5085
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   615
      Left            =   3000
      TabIndex        =   12
      Top             =   2160
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "boolean ops.frx":0000
      Left            =   3480
      List            =   "boolean ops.frx":000D
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "Output"
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   4815
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Output:"
         Height          =   195
         Left            =   600
         TabIndex        =   9
         Top             =   480
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1320
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "boolean ops.frx":0028
         Left            =   1800
         List            =   "boolean ops.frx":003E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Choose Operation:"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   960
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enter Second Number:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Enter First Number:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1350
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do it!"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y, temp As Double



Private Sub Combo1_Click()
'Same thing as below
If Combo1.ListIndex = 5 Then
    Text2.Visible = False
    Label2.Visible = False
    Frame1.Height = 1335
    Combo2.Visible = True
    Form1.Refresh
Else
    Text2.Visible = True
    Label2.Visible = True
    Frame1.Height = 1935
    Combo2.Visible = False
    Form1.Refresh
End If
End Sub

Private Sub Combo1_GotFocus()
'This will hide the things that arent needed for 1's complement, since only 1 number is needed for it

If Combo1.ListIndex = 5 Then
    Text2.Visible = False
    Label2.Visible = False
    Frame1.Height = 1335
    Combo2.Visible = True
    Form1.Refresh
Else
'If it isn't 1's complement, then return it to the normal state
    Text2.Visible = True
    Label2.Visible = True
    Frame1.Height = 1935
    Combo2.Visible = False
    Form1.Refresh
End If
End Sub

Private Sub Command1_Click()

'This is the select statement to see which function to call
'according to combo1
Select Case Combo1.ListIndex

Case 0
    Call booleanAND
Case 1
    Call booleanOR
Case 2
    Call booleanXOR
Case 3
    Call booleanSL
Case 4
    Call booleanSR
Case 5
    Call boolean1COMP
End Select
    



End Sub


Public Function booleanAND()
'When this is called, it will AND the two values and make text3 the return value
'When two bits are compared, if one and the other are 1, it will return a 1

x = Val(Text1.Text)
y = Val(Text2.Text)
Output = 0

For i = 23 To 0 Step -1 '24 bit numbers accepted
    temp = x \ (2 ^ i) 'gets the bit of text1.text
    x = x - temp * (2 ^ i) 'subtracts it from the number
    temp1 = y \ (2 ^ i) 'gets the bit of text2.text
    y = y - temp1 * (2 ^ i) 'subtracts it from the number
    If temp = 1 And temp1 = 1 Then 'If both are equal to 1 then return a 1
        Output = Output + (2 ^ i) 'This returns a decimal number
    End If
Next i

Text3.Text = Output


End Function
Public Function booleanOR()
'When this is called, it will or the two values and make text3 the return value
'When two bits are compared, if one or the other, or both are 1, it will return a 1

x = Val(Text1.Text)
y = Val(Text2.Text)
Output = 0

For i = 23 To 0 Step -1 '24 bit numbers accepted
    temp = x \ (2 ^ i) 'gets the bit of text1.text
    x = x - temp * (2 ^ i) 'subtracts it from the number
    temp1 = y \ (2 ^ i) 'gets the bit of text2.text
    y = y - temp1 * (2 ^ i) 'subtracts it from the number
    If temp = 1 Or temp1 = 1 Then ' If one or the other or both=1 then return a 1
        Output = Output + (2 ^ i) 'This returns a decimal number
    End If
Next i

Text3.Text = Output


End Function
Public Function booleanXOR()
'When this is called, it will xor the two values and make text3 the return value
'When two bits are compared, if one or the other, but not both are 1, it will return a 1
x = Val(Text1.Text)
y = Val(Text2.Text)
Output = 0

For i = 23 To 0 Step -1 '24 bit numbers accepted
    temp = x \ (2 ^ i) 'gets the bit of text1.text
    x = x - temp * (2 ^ i) 'subtracts it from the number
    temp1 = y \ (2 ^ i) 'gets the bit of text2.text
    y = y - temp1 * (2 ^ i) 'subtracts it from the number
    If temp = 1 Xor temp1 = 1 Then 'If one or the other but not both=1 then return a 1
        Output = Output + (2 ^ i) 'This returns a decimal number
    End If
Next i

Text3.Text = Output


End Function
Public Function booleanSL()
'When this function is called, it will shift left the bits
'This will move all of x's bits to the left y places

x = Val(Text1.Text)
y = Val(Text2.Text)
Text3.Text = x * 2 ^ y ' This shifts the bits
End Function


Public Function booleanSR()
'When this function is called, it will shift right the bits
'This will move all of x's bits to the right y places
x = Val(Text1.Text)
y = Val(Text2.Text)

Text3.Text = x \ 2 ^ y 'This shifts them

End Function

Public Function boolean1COMP()
'When this function is called, it will perform 1's complement on the given number
'1's complement is making all ones zeros and all zeros ones

x = Val(Text1.Text)
Output = 0


Select Case Combo2.ListIndex 'See how many bits the user selected
Case 0
    bits = 7 '8 bit
Case 1
    bits = 15 '16 bit
Case 2
    bits = 23 '24 bit
End Select

For i = bits To 0 Step -1
    temp = x \ (2 ^ i) ' This finds if the bit for the position i should be 1 or 0
    x = x - temp * (2 ^ i) 'This subtracts that from it
    If temp = 0 Then 'If bit=0, then bit=1
        Output = Output + (2 ^ i) 'Puts it in a decimal variable
    End If
Next i
Text3.Text = Output 'Displays it
End Function

Private Sub Command2_Click()
'This is my about section
msgtext = "Bit Operations by Derek Haas" + vbNewLine
msgtext = msgtext + "This will perform many operations" + vbNewLine
msgtext = msgtext + "on numbers you input." + vbNewLine + vbNewLine
msgtext = msgtext + "If you want to use any of the source in this program," + vbNewLine
msgtext = msgtext + "email me at kibblesnbits@snip.net and get my permission." + vbNewLine + vbNewLine
msgtext = msgtext + "Thank you"
q = MsgBox(msgtext, vbOKCancel, "About")
End Sub

