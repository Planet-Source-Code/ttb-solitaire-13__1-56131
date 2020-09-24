VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "solitaire 13"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "New game"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   1560
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "score:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   1455
      Left            =   120
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Left            =   6840
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "pile"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "garbage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   51
      Left            =   5640
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   50
      Left            =   4440
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   49
      Left            =   3240
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   48
      Left            =   5640
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   47
      Left            =   4440
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   46
      Left            =   3240
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   45
      Left            =   5640
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   44
      Left            =   4440
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   43
      Left            =   3240
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   42
      Left            =   5640
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   41
      Left            =   4440
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   40
      Left            =   3240
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   39
      Left            =   5640
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   38
      Left            =   4440
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   37
      Left            =   3240
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   36
      Left            =   5640
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   35
      Left            =   4440
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   34
      Left            =   3240
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   33
      Left            =   5640
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   32
      Left            =   4440
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   31
      Left            =   3240
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   30
      Left            =   5640
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   29
      Left            =   4440
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   28
      Left            =   3240
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   27
      Left            =   360
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   2880
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   26
      Left            =   1440
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   2880
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   25
      Left            =   2520
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   2880
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   24
      Left            =   3600
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   2880
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   23
      Left            =   4680
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   2880
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   22
      Left            =   5760
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   2880
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   21
      Left            =   6840
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   2880
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   20
      Left            =   840
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   19
      Left            =   1920
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   18
      Left            =   3000
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   17
      Left            =   4080
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   16
      Left            =   5160
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   15
      Left            =   6240
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   14
      Left            =   1320
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   13
      Left            =   2400
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   12
      Left            =   3480
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   11
      Left            =   4560
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   10
      Left            =   5640
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   9
      Left            =   1800
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   8
      Left            =   2880
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   7
      Left            =   3960
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   6
      Left            =   5040
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   5
      Left            =   2280
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   960
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   4
      Left            =   3360
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   960
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   3
      Left            =   4440
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   960
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   2
      Left            =   3960
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   480
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   1
      Left            =   2880
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   480
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   0
      Left            =   3360
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Top             =   0
      Width           =   1065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dragged, draggedIndex, points, cardno&



Function Is13(cardone, cardtwo) As Boolean
If cardone = "" Then cardone = -1
Do
If cardone > 12 Then cardone = cardone - 13
If cardtwo > 12 Then cardtwo = cardtwo - 13
Loop Until cardone < 13 And cardtwo < 13

cardone = cardone + 1
cardtwo = cardtwo + 1
'MsgBox cardone
If cardone + cardtwo = 13 Then
Is13 = True

If cardtwo = 13 Then
points = points + 1
Else
points = points + 2
End If
Label3.Caption = points
If points = 52 Then MsgBox "YOU WIN!"
If points > 34 Then Image2.Picture = LoadResPicture(156, 0): Exit Function
If points > 17 Then Image2.Picture = LoadResPicture(155, 0): Exit Function
If points > 0 Then Image2.Picture = LoadResPicture(154, 0): Exit Function

End If

dragged = ""
End Function

Private Sub Command1_Click()
newgame
End Sub

Private Sub Form_Load()

newgame
End Sub
Public Sub newgame()
Randomize
For i = 28 To 51
Image1(i).Visible = False
List1.Clear
Next
For i = 0 To 27
Image1(i).Visible = True
Next

For i = 0 To 20
Image1(i).Enabled = False
Next
Do
bla = Int(Rnd * 52)
List1.Text = bla
If List1.Text <> bla Then List1.AddItem bla
Loop Until List1.ListCount = 52
For i = 0 To 51
Image1(i).Picture = LoadResPicture(101 + List1.List(i), 0)
Image1(i).Tag = List1.List(i)
Next
points = 0
Label3.Caption = "0"
cardno& = 27
Image2.Picture = LoadResPicture(153, 0)
Image3.Picture = LoadResPicture(156, 0)
Image3.Enabled = True
End Sub


Private Sub Image1_DblClick(Index As Integer)

If Is13("", Image1(Index).Tag) Then
Image1(Index).Visible = False
If Index < 28 Then EnableOpen (Index)
End If
End Sub

Private Sub Image1_OLECompleteDrag(Index As Integer, Effect As Long)
If Is13(dragged, Image1(Index).Tag) Then
Image1(Index).Visible = False
Image1(draggedIndex).Visible = False
EnableOpen draggedIndex
EnableOpen Index

End If
End Sub

Private Sub Image1_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
dragged = Image1(Index).Tag
draggedIndex = Index
End Sub
Sub EnableOpen(card)

Select Case card
Case 21
If Image1(22).Visible = False Then Image1(15).Enabled = True
Case 22
If Image1(21).Visible = False Then Image1(15).Enabled = True
If Image1(23).Visible = False Then Image1(16).Enabled = True
Case 23
If Image1(22).Visible = False Then Image1(16).Enabled = True
If Image1(24).Visible = flase Then Image1(17).Enabled = True
Case 24
If Image1(23).Visible = False Then Image1(17).Enabled = True
If Image1(25).Visible = False Then Image1(18).Enabled = True
Case 25
If Image1(24).Visible = False Then Image1(18).Enabled = True
If Image1(26).Visible = False Then Image1(19).Enabled = True
Case 26
If Image1(25).Visible = False Then Image1(19).Enabled = True
If Image1(27).Visible = False Then Image1(20).Enabled = True
Case 27
If Image1(26).Visible = False Then Image1(20).Enabled = True
Case 15
If Image1(16).Visible = False Then Image1(10).Enabled = True
Case 16
If Image1(15).Visible = False Then Image1(10).Enabled = True
If Image1(17).Visible = False Then Image1(11).Enabled = True
Case 17
If Image1(16).Visible = False Then Image1(11).Enabled = True
If Image1(18).Visible = False Then Image1(12).Enabled = True
Case 18
If Image1(17).Visible = False Then Image1(12).Enabled = True
If Image1(19).Visible = False Then Image1(13).Enabled = True
Case 19
If Image1(18).Visible = False Then Image1(13).Enabled = True
If Image1(20).Visible = False Then Image1(14).Enabled = True
Case 20
If Image1(19).Visible = False Then Image1(14).Enabled = True
Case 10
If Image1(11).Visible = False Then Image1(6).Enabled = True
Case 11
If Image1(10).Visible = False Then Image1(6).Enabled = True
If Image1(12).Visible = False Then Image1(7).Enabled = True
Case 12
If Image1(11).Visible = False Then Image1(7).Enabled = True
If Image1(13).Visible = False Then Image1(8).Enabled = True
Case 13
If Image1(12).Visible = False Then Image1(8).Enabled = True
If Image1(14).Visible = False Then Image1(9).Enabled = True
Case 14
If Image1(13).Visible = False Then Image1(9).Enabled = True
Case 6
If Image1(7).Visible = False Then Image1(3).Enabled = True
Case 7
If Image1(6).Visible = False Then Image1(3).Enabled = True
If Image1(8).Visible = False Then Image1(4).Enabled = True
Case 8
If Image1(7).Visible = False Then Image1(4).Enabled = True
If Image1(9).Visible = False Then Image1(5).Enabled = True
Case 9
If Image1(8).Visible = False Then Image1(5).Enabled = True
Case 3
If Image1(4).Visible = False Then Image1(2).Enabled = True
Case 4
If Image1(3).Visible = False Then Image1(2).Enabled = True
If Image1(5).Visible = False Then Image1(1).Enabled = True
Case 5
If Image1(4).Visible = False Then Image1(1).Enabled = True
Case 2
If Image1(1).Visible = False Then Image1(0).Enabled = True
Case 1
If Image1(2).Visible = False Then Image1(0).Enabled = True
End Select
End Sub

Private Sub Image3_DblClick()

For i = 1 To 3
cardno = cardno + 1
Image1(cardno).Visible = True
Next

If cardno = 51 Then Image3.Picture = LoadResPicture(153, 0): Image3.Enabled = False: Exit Sub
If cardno > 43 Then Image3.Picture = LoadResPicture(154, 0): Exit Sub
If cardno > 35 Then Image3.Picture = LoadResPicture(155, 0): Exit Sub
End Sub
