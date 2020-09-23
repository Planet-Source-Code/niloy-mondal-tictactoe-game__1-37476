VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TicTacToe"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "About me"
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4320
      TabIndex        =   11
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start a New Game"
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Line Lline 
      Index           =   7
      Visible         =   0   'False
      X1              =   960
      X2              =   4920
      Y1              =   4560
      Y2              =   960
   End
   Begin VB.Line Lline 
      Index           =   6
      Visible         =   0   'False
      X1              =   960
      X2              =   4920
      Y1              =   960
      Y2              =   4560
   End
   Begin VB.Line Lline 
      Index           =   5
      Visible         =   0   'False
      X1              =   4680
      X2              =   4680
      Y1              =   840
      Y2              =   4800
   End
   Begin VB.Line Lline 
      Index           =   4
      Visible         =   0   'False
      X1              =   3000
      X2              =   3000
      Y1              =   840
      Y2              =   4800
   End
   Begin VB.Line Lline 
      Index           =   3
      Visible         =   0   'False
      X1              =   1200
      X2              =   1200
      Y1              =   840
      Y2              =   4800
   End
   Begin VB.Line Lline 
      Index           =   2
      Visible         =   0   'False
      X1              =   840
      X2              =   5280
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Lline 
      Index           =   1
      Visible         =   0   'False
      X1              =   840
      X2              =   5280
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Lline 
      Index           =   0
      Visible         =   0   'False
      X1              =   840
      X2              =   5280
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "TICTACTOE The Unbeatable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   120
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      Height          =   4455
      Left            =   360
      Top             =   600
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   5
      Left            =   3960
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   4
      Left            =   2160
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Line Line4 
      X1              =   360
      X2              =   5640
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   5640
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      X1              =   3840
      X2              =   3840
      Y1              =   600
      Y2              =   5040
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   2040
      Y1              =   600
      Y2              =   5040
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   2
      Left            =   3960
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   6
      Left            =   480
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   7
      Left            =   2160
      TabIndex        =   6
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   8
      Left            =   3960
      TabIndex        =   7
      Top             =   3600
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public chance As Integer    'chance determines who plays first chance=even=user's first move
Public lineno As Integer      'lineno stores the line number used to draw when the pc wins
Dim gmove As Integer 'gmove is incremented every time computer makes a move

Private Function opening_move()
If Label1(4).Caption <> "X" Then
    Label1(4).Caption = "O"
Else
    Label1(8).Caption = "O"
End If
gmove = gmove + 1
End Function
Private Function perform_check(ch As String)
Dim i As Integer
For i = 0 To 6 Step 3           'checks horizontally 3 times for 3 row
If Label1(i).Caption = ch And Label1(i + 1).Caption = ch And Label1(i + 2).Caption = "" Then
    Label1(i + 2).Caption = "O"
    gmove = gmove + 1
    Exit Function
End If
If Label1(i).Caption = ch And Label1(i + 1).Caption = "" And Label1(i + 2).Caption = ch Then
    Label1(i + 1).Caption = "O"
    gmove = gmove + 1
    Exit Function
End If
If Label1(i).Caption = "" And Label1(i + 1).Caption = ch And Label1(i + 2).Caption = ch Then
    Label1(i + 0).Caption = "O"
    gmove = gmove + 1
    Exit Function
End If
Next i
For i = 0 To 2                  'checks vertically 3 times 3 coloumns
If Label1(i).Caption = ch And Label1(i + 3).Caption = ch And Label1(i + 6).Caption = "" Then
    Label1(i + 6).Caption = "O"
    gmove = gmove + 1
    Exit Function
End If
If Label1(i).Caption = ch And Label1(i + 3).Caption = "" And Label1(i + 6).Caption = ch Then
    Label1(i + 3).Caption = "O"
    gmove = gmove + 1
    Exit Function
End If
If Label1(i).Caption = "" And Label1(i + 3).Caption = ch And Label1(i + 6).Caption = ch Then
    Label1(i + 0).Caption = "O"
    gmove = gmove + 1
    Exit Function
End If
Next i
'check diagonal from upper right to lower left
If Label1(2).Caption = ch And Label1(4).Caption = ch And Label1(6).Caption = "" Then
    Label1(6).Caption = "O"
    gmove = gmove + 1
    Exit Function
End If
'check diagonal from upper right to lower left, 2nd condition
If Label1(2).Caption = "" And Label1(4).Caption = ch And Label1(6).Caption = ch Then
    Label1(2).Caption = "O"
    gmove = gmove + 1
    Exit Function
End If
'check diagonal from upper left to lower right
If Label1(0).Caption = "" And Label1(4).Caption = ch And Label1(8).Caption = ch Then
    Label1(0).Caption = "O"
    gmove = gmove + 1
    Exit Function
End If
'check diagonal from lower right to upper left
If Label1(8).Caption = "" And Label1(4).Caption = ch And Label1(0).Caption = ch Then
    Label1(8).Caption = "O"
    gmove = gmove + 1
    Exit Function
End If

End Function
Private Function perform_win_check(c As String)
Dim i As Integer
For i = 0 To 6 Step 3       'horizontal check
    If Label1(i).Caption = c And Label1(i + 1).Caption = c And Label1(i + 2).Caption = c Then
        Lline(i / 3).Visible = True
        lineno = i / 3
        If c = "O" Then
            MsgBox "Computer Wins, You lose"
        Else
            MsgBox "Very Good, You win, Computer loses, tell me how you did it"
        End If
    End If
Next i
For i = 0 To 2              'vertical check
    If Label1(i).Caption = c And Label1(i + 3).Caption = c And Label1(i + 6).Caption = c Then
        Lline(i + 3).Visible = True
        lineno = i + 3
        If c = "O" Then
            MsgBox "Computer Wins, You lose"
        Else
            MsgBox "Very Good, You win, Computer loses, tell me how you did it."
        End If
    End If
Next i
'two diagonal checks
If Label1(0).Caption = c And Label1(4).Caption = c And Label1(8).Caption = c Then
    Lline(6).Visible = True
    lineno = 6
    If c = "O" Then
        MsgBox "Computer Wins, You lose"
    Else
        MsgBox "Very Good, You win, Computer loses, tell me how you did it"
    End If
End If
If Label1(2).Caption = c And Label1(4).Caption = c And Label1(6).Caption = c Then
    Lline(7).Visible = True
    lineno = 7
    If c = "O" Then
        MsgBox "Computer Wins, You lose"
    Else
        MsgBox "Very Good, You win, Computer loses, tell me how you did it."
    End If
End If
End Function

Private Function second_move()
perform_check ("X")              'checks and destroys users chances of winning
If gmove = 1 Then      'If no move was made in perform_check then take care of some special trick moves
    If Label1(5).Caption = "X" And Label1(7).Caption = "X" Then
        Label1(8).Caption = "O"
        gmove = gmove + 1
    End If
    If Label1(0).Caption = "X" And Label1(8).Caption = "X" Then
        Label1(1).Caption = "O"
        gmove = gmove + 1
    End If
    If Label1(2).Caption = "X" And Label1(6).Caption = "X" Then
        Label1(1).Caption = "O"
        gmove = gmove + 1
    End If
    If Label1(0).Caption = "X" And Label1(4).Caption = "X" Then
        Label1(6).Caption = "O"
        gmove = gmove + 1
    End If
    If Label1(0).Caption = "X" And Label1(7).Caption = "X" Then
        Label1(6).Caption = "O"
        gmove = gmove + 1
    End If
    If Label1(2).Caption = "X" And Label1(7).Caption = "X" Then
        Label1(8).Caption = "O"
        gmove = gmove + 1
    End If
    If Label1(5).Caption = "X" And Label1(6).Caption = "X" Then
        Label1(8).Caption = "O"
        gmove = gmove + 1
    End If
    If Label1(3).Caption = "X" And Label1(5).Caption = "X" Then
        Label1(0).Caption = "O"
        gmove = gmove + 1
    End If
    If Label1(1).Caption = "X" And Label1(7).Caption = "X" Then
        Label1(0).Caption = "O"
        gmove = gmove + 1
    End If
End If
If gmove = 1 Then       'If still no move was made by computer then
    seek_empty_fill     'Find any empty box and fill it with "O"
End If
End Function
Private Function seek_empty_fill()
Dim i As Integer
For i = 0 To 8
    If Label1(i).Caption <> "X" And Label1(i).Caption <> "O" And flag = 0 Then
        Label1(i).Caption = "O"
        gmove = gmove + 1
        Exit For
    End If
Next i
End Function
Private Function third_move()
perform_check ("O")     'check and accomplish if computer has any chance of winning
If gmove = 2 Then       'if no move was made on the previous perform check
    perform_check ("X")     'check and destroy users chances of winning
End If
If Label1(0).Caption = "O" And Label1(4).Caption = "O" And Label1(3).Caption = "" And gmove = 2 Then
    Label1(3).Caption = "O"
    gmove = gmove + 1
End If
If Label1(0).Caption = "X" And Label1(7).Caption = "X" And Label1(2).Caption = "" And gmove = 2 Then
    Label1(6).Caption = "O"
    gmove = gmove + 1
End If
If gmove = 2 Then
    seek_empty_fill     'find any empty box and fill it with "O"
End If
End Function
Private Function forth_move()
perform_check ("O")     'check and accomplish if computer has any chance of winning
If gmove = 3 Then
    perform_check ("X") 'check and destroy users chances of winning
End If
If gmove = 3 Then       'If still no move was given then
    seek_empty_fill     'fill any empty box with "O"
End If
End Function

Private Sub Command1_Click() 'New Game
Dim i As Integer
chance = chance + 1
gmove = 0
Lline(lineno).Visible = False
For i = 0 To 8
    Label1(i).Caption = ""
Next i
If chance Mod 2 = 0 Then
    opening_move
End If
End Sub

Private Sub Command2_Click()    'Exit
End
End Sub

Private Sub Command3_Click() 'About me
MsgBox "The name is NILOY MONDAL. Email- niloygk@yahoo.com or niloy_mondal@indiatimes.com"
End Sub



Private Sub Label1_Click(index As Integer)
If Label1(index).Caption = "" Then
    Label1(index).Caption = "X" 'Mark X where user clicks
    If gmove = 4 Then
        MsgBox "Looks like we have a draw"  'Its a draw by this stage
    End If
    If gmove = 3 Then
        forth_move          'Perform the forth move
    End If
    If gmove = 2 Then
        third_move          'Perform the third move
    End If
    If gmove = 1 Then
        second_move         'Perform the second move, this is where some tricky things happen
    End If
    If gmove = 0 Then    'If this is the first move
        opening_move        'Perform the opening move
    End If
End If
perform_win_check ("X") 'check if user has won, I have tried my best this will not happen, but who knows
perform_win_check ("O") 'check if computer has won
End Sub
