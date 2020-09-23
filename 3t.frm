VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Smartest Tic-Tac-Toe (AI)"
   ClientHeight    =   5400
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3105
   Icon            =   "3t.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   3105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   585
      TabIndex        =   9
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   2160
      TabIndex        =   8
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   2160
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   2160
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   240
      X2              =   2880
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   240
      X2              =   2880
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   2040
      X2              =   2040
      Y1              =   360
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   1080
      X2              =   1080
      Y1              =   360
      Y2              =   3000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Human won:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   105
      TabIndex        =   11
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Computer won:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   105
      TabIndex        =   10
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Menu ab 
      Caption         =   "About"
      Begin VB.Menu abou 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'You can freely distribute the code only if you have mentioned my name
'You can use this code if you take my permission
Dim comuturn As Boolean
Dim compisfirst As Boolean
Dim aqq(0 To 8) As Boolean
Dim pos2ocu As Integer
Dim computer As Integer
Dim user As Integer
Dim brain As String
Dim noofmem As Integer
Dim arrofmem(10000) As String
Dim wonbyfbm As Boolean


Private Sub abou_Click()
MsgBox "Made by Paras Chopra. See readme.txt for all details about me and this program." & vbCrLf & "I know a bug whenever you win it tells you that it's your chance but takes it's own or places two crosses. In these cases just hit new game button."
End Sub

Private Sub Command1_Click(Index As Integer)
If Command1(Index).Caption = "" Then
Call moveto(Index, False)
brain = brain & Index
rer = winforcomp(True, 0)
If rer = 102 Then
Exit Sub
End If
Call moveto(CInt(rer), True)
brain = brain & rer

End If
End Sub

Function winforcomp(isitcomp As Boolean, level As Integer) As Integer
Dim p(0 To 8) As Integer
p(0) = 3
p(1) = 2
p(2) = 3
p(3) = 2
p(4) = 5
p(5) = 2
p(6) = 3
p(7) = 2
p(8) = 3
Dim valuee(1) As Integer
greatestval = -100
If isitcomp = True Then
For i = 0 To 8
If aqq(i) = False Then
aqq(i) = True
Command1(i).Caption = "X"
If cfw() = True Then
winforcomp = i
Exit Function
End If
aqq(i) = False
Command1(i).Caption = ""
End If
Next i
For i = 0 To 8
If i = 9 Then
    winforcomp = position
    Exit Function
End If
If aqq(CInt(i)) = False Then
efr = fbm
    If efr <> 10 Then
    wonbyfbm = True
    winforcomp = efr
    Exit Function
    Else
valuee(level) = p(i)
re = winforcomp(False, level + 1)
    If re = -100 Then
    winforcomp = pos2ocu
    Exit Function
    End If
valuee(level) = re + valuee(level)
    If valuee(level) >= greatestval Then
    greatestval = valuee(level)
    position = i
    End If
End If
End If
Next i
If (aqq(0) = True) And position = Empty Then
winforcomp = 102
Else
winforcomp = position
End If
Else
For j = 0 To 8
If aqq(j) = False Then
valuee(level) = -(p(j))
aqq(j) = True
Command1(j).Caption = "0"
If cfw() = True Then
valuee(level) = -100
pos2ocu = j
End If
aqq(j) = False
Command1(j).Caption = ""
If valuee(level) <= lowestval Then
lowestval = valuee(level)
End If
End If
Next j
winforcomp = lowestval
End If
End Function

Private Sub Command2_Click()
Form_Load
End Sub

Private Sub Form_Load()
wonbyfbm = False
Randomize
For i = 0 To noofmem
arrofmem(i) = ""
Next i
noofmem = 0
brain = ""
Label1 = "Computer won : " & computer
Label2 = "You won : " & user
inputbrain
For i = 0 To 8
aqq(i) = False
Command1(i).Caption = ""
Next i
de = Rnd
If Rnd < 0.5 Then
MsgBox "Computer's turn"
compisfirst = True
comuturn = True
r = Int(Rnd * 8) + 0
'r = winforcomp(True, 0)
Call moveto(CInt(r), True)
brain = r
comuturn = False
Else
MsgBox "Your turn"
compisfirst = False
comuturn = False
End If
End Sub

Function cfw() As Boolean

For i = 0 To 6 Step 3
        Owner = mboard(CInt(i))
        If Not Owner = 0 Then If Owner = mboard(i + 1) And Owner = mboard(i + 2) Then cfw = True
    Next i

        For i = 0 To 2
            Owner = mboard(CInt(i))
            If Not Owner = 0 Then If Owner = mboard(i + 3) And Owner = mboard(i + 6) Then cfw = True
        Next i

    
        Owner = mboard(0)
        If Not Owner = 0 Then If Owner = mboard(4) And Owner = mboard(8) Then cfw = True
        Owner = mboard(2)
        If Not Owner = 0 Then If Owner = mboard(4) And Owner = mboard(6) Then cfw = True
If cfw = True Then
Exit Function
Else
cfw = False
End If
End Function

Sub moveto(pose As Integer, comp As Boolean)
If comp = True Then
Command1(pose).Caption = "X"
aqq(pose) = True
comuturn = False
temp = cfw()
If temp = True Then
MsgBox "As usual: Computer(AI) wins!"
computer = computer + 1
Label1 = "Computer won : " & computer
If compisfirst = True Then
brain = brain & 1
Else
brain = brain & 2
End If
If wonbyfbm <> True Then
putintobrain
End If
Form_Load
End If
Else
Command1(pose).Caption = "0"
aqq(pose) = True
temp = cfw()
If temp = True Then
MsgBox "Rare message: Human wins!"
user = user + 1
Label2 = "You won : " & user
If compisfirst = True Then
brain = brain & 2
Else
brain = brain & 1
End If
putintobrain
Form_Load
End If
End If
End Sub

Function mboard(pose As Integer) As Integer
If Command1(pose).Caption = "X" Then
mboard = 1
ElseIf Command1(pose).Caption = "0" Then
mboard = 2
Else
mboard = 0
End If
End Function

Sub putintobrain()
Open App.Path & "\brain.txt" For Append As #1
Print #1, brain
Close #1
End Sub

Function fbm() As Integer
' it thinks of what is cuurent move & what is in the brain
For i = 0 To noofmem - 1
    j = 1
    r = Mid$(arrofmem(i), Len(arrofmem(i)), 1)
    If (compisfirst = True And r = 1) Or (compisfirst = False And r = 2) Then
    If Mid$(arrofmem(i), j, Len(brain)) = brain Then
    j = j + Len(brain)
    fbm = Mid$(arrofmem(i), j, 1)
    Exit Function
    End If
    End If
Next i
fbm = 10
End Function

Sub inputbrain()
Open App.Path & "\brain.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, r
arrofmem(noofmem) = r
noofmem = noofmem + 1
Loop
Close #1
End Sub
