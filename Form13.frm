VERSION 5.00
Begin VB.Form Form13 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13770
   BeginProperty Font 
      Name            =   "Tempus Sans ITC"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form13.frx":0000
   ScaleHeight     =   9105
   ScaleWidth      =   13770
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   13080
      Top             =   120
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4740
      ItemData        =   "Form13.frx":345DE
      Left            =   4440
      List            =   "Form13.frx":345E0
      TabIndex        =   3
      Top             =   2520
      Width           =   5055
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   10080
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   10080
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add Results"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   10080
      TabIndex        =   7
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Modify Slots"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   10080
      TabIndex        =   6
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add Participants"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      TabIndex        =   5
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search Participants"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      TabIndex        =   4
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   5040
      TabIndex        =   2
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome,"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmd1 As New ADODB.Command
Dim currdate As String
Dim time1 As String
Public feid As Integer


Private Sub Form_Load()
Label3.Caption = Form8.username
Label1.Caption = Form1.festname
Call livestatistics
End Sub
Private Sub Label4_Click()
Form12.Show
Unload Me
End Sub

Private Sub livestatistics()
Dim str As String
Dim str1 As String
Dim cnt As Integer
Dim rs1 As New ADODB.Recordset
Dim cmd2 As New ADODB.Command
Set Form3.rs = Nothing
cnt = 0
str = "select count(pname) as mycnt from participant where eventname in (select eventname from evtmain where festname ='" + Form1.festname + "');"
With cmd1
    .ActiveConnection = Form3.oconn
    .CommandText = str
    .CommandType = adCmdText
End With
With Form3.rs
    .CursorType = adOpenStatic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open cmd1
End With

'First thing in live statistics.
 List1.AddItem ("Total number of Participants is " + CStr(Form3.rs!mycnt))
 cnt = cnt + 1
Set Form3.rs = Nothing
Set cmd1 = Nothing
'Completed
List1.AddItem (vbCrLf)
cnt = cnt + 1
'Now displaying current events

currdate = Format(Now, "mmm-dd-yy")
currdate = Format(currdate, "DD-MMM-YY")
time1 = Time$

List1.AddItem (vbCrLf + "Events :                                Rounds :")
cnt = cnt + 1
str = "select eventname from evtmain where feid in (select feid from rounds where rdate='" + currdate + "'and stime <='" + time1 + "' and etime >=' " + time1 + "')"
str1 = "select roundno from rounds where rdate='" + currdate + "'and stime <='" + time1 + "' and etime >=' " + time1 + "'"
With cmd1
    .ActiveConnection = Form3.oconn
    .CommandText = str
    .CommandType = adCmdText
End With
With Form3.rs
    .CursorType = adOpenStatic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open cmd1
End With

With cmd2
    .ActiveConnection = Form3.oconn
    .CommandText = str1
    .CommandType = adCmdText
End With
With rs1
    .CursorType = adOpenStatic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open cmd2
End With

 Do Until Form3.rs.EOF = True
    List1.AddItem (CStr(Form3.rs!eventname) + "                             " + CStr(rs1!roundno) + vbCrLf)
    cnt = cnt + 1
    Form3.rs.MoveNext
    rs1.MoveNext
 Loop
 'done
 
Do Until cnt = 14
List1.AddItem (vbCrLf)
cnt = cnt + 1
Loop

End Sub

Private Sub Label5_Click()
Form11.Show
Unload Me
End Sub

Private Sub Label6_Click()
Form5.Show
Unload Me
End Sub

Private Sub Label7_Click()
Form10.Show
Unload Me
End Sub

Private Sub Timer1_Timer()
Dim noofenteries As Integer
Dim i As Integer
Dim temp As String
j = List1.ListCount - 2
temp = List1.List(List1.ListCount - 1)
While j >= 0
List1.List(j + 1) = List1.List(j)
j = j - 1
Wend
List1.RemoveItem (0)
List1.AddItem temp, 0
End Sub
