VERSION 5.00
Begin VB.Form List 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "        Word                        -----                         Description"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Edit2 
      Caption         =   "Edit"
      Height          =   375
      Index           =   1
      Left            =   5400
      TabIndex        =   7
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton Edit1 
      Caption         =   "Edit"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton Save 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Remove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   5160
      Width           =   975
   End
   Begin VB.ListBox List2 
      Height          =   4740
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Words(1000) As String
Dim Dess(1000) As String
Dim SCon As Boolean
Private Sub Add_Click()
Dim Word As String
Dim Des As String
SCon = True
Word = InputBox(" Type a Word", " Type a Word")
If Len(Word) Then
Des = InputBox(" Type Description", " Type Description")
If Len(Des) Then
List1.AddItem Word
List2.AddItem Des
End If
End If

End Sub

Private Sub Edit1_Click(Index As Integer)
Dim Con As String
Dim Word As String
SCon = True
If Len(List1.Text) Then
Con = MsgBox("Want to edit '" & List1.Text & "'", vbInformation + vbYesNo, "Edit ...")
If Con = vbYes Then
Word = InputBox("Type Word ", "Edit ...", List1.Text)
If Len(Word) Then List1.List(List1.ListIndex) = Word
End If
End If
End Sub

Private Sub Edit2_Click(Index As Integer)
Dim Con As String
Dim Word As String
SCon = True
If Len(List2.Text) Then
Con = MsgBox("Want to edit '" & List2.Text & "'", vbInformation + vbYesNo, "Edit ...")
If Con = vbYes Then
Word = InputBox("Type Word ", "Edit ...", List2.Text)
If Len(Word) Then List2.List(List2.ListIndex) = Word
End If
End If
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
LoadFile
SCon = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Con As String
If SCon = False Then Exit Sub
Con = MsgBox("Want to Save List", vbInformation + vbYesNoCancel, "Message !!")
If Con = vbYes Then Save_Click
If Con = vbCancel Then Cancel = 1
End Sub

Private Sub List1_Click()
On Error Resume Next
List2.Selected(List1.ListIndex) = True
End Sub

Private Sub List2_Click()
On Error Resume Next
List1.Selected(List2.ListIndex) = True
End Sub

Private Sub Remove_Click()
On Error Resume Next
Dim Con As String
Con = MsgBox("Want to Remove This : ' " & List1.List(List1.ListIndex) & " '", vbInformation + vbYesNo, "Removing ...")
If Con = vbYes Then
List1.RemoveItem List1.ListIndex
List2.RemoveItem List2.ListIndex
End If
SCon = True
End Sub

Private Sub Save_Click()
Dim WordFile As String
Dim DesFile As String
Dim Total As Integer
Dim L
Dim N
If List1.ListCount = 0 Then Exit Sub
If List2.ListCount = 0 Then Exit Sub
If List1.ListCount = List2.ListCount Then

WordFile = App.Path & "\WordFile.dat"
DesFile = App.Path & "\DesFile.dat"
If Dir(WordFile) <> "" Then Kill WordFile
If Dir(DesFile) <> "" Then Kill DesFile

Open WordFile For Binary As 1
Total = List1.ListCount
Put 1, , Total
For L = 0 To Total - 1
Words(L) = List1.List(L)
Next L
Put 1, , Words
Close 1

Open DesFile For Binary As 2
Total = List2.ListCount
Put 2, , Total
For N = 0 To Total - 1
Dess(N) = List2.List(N)
Next N
Put 2, , Dess
Close 2
SCon = False
End If
End Sub
Public Sub LoadFile()
Dim WordFile As String
Dim DesFile As String
Dim Total As Integer
Dim L
Dim N
WordFile = App.Path & "\WordFile.dat"
DesFile = App.Path & "\DesFile.dat"
Open WordFile For Binary As 1
Get 1, , Total
Get 1, , Words
Close 1
For L = 0 To Total - 1
List1.AddItem Words(L)
Next L
Open DesFile For Binary As 2
Get 2, , Total
Get 2, , Dess
Close 2
For L = 0 To Total - 1
List2.AddItem Dess(L)
Next L


End Sub
