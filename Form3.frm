VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "College Fest Manager - Begin"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12360
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form3"
   MousePointer    =   4  'Icon
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   8250
   ScaleWidth      =   12360
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      ItemData        =   "Form3.frx":345DE
      Left            =   6480
      List            =   "Form3.frx":345E0
      TabIndex        =   5
      Top             =   2880
      Width           =   4935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "College Fest Manager"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   9255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   960
      Top             =   4440
      Width           =   3495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Fest"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   3
      Top             =   4560
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Set Up Fest"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   960
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a fest first!"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   6480
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Existing Fests"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   5880
      Width           =   3135
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cmd As New ADODB.Command
Public oconn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public strSQL As String
Dim selfest As String

Private Sub Command1_Click()
Form1.Show
Form3.Hide
End Sub

Private Sub Command2_Click()
selfest = List1.Text
If (selfest = "") Then
    Label3.Visible = True
    Exit Sub
ElseIf (selfest <> "") Then
    'Goto login form with selfest as current fest id
    Form1.festname = selfest
    Form8.Show
    Form3.Hide
End If
End Sub


Private Sub Form_Load()
Set oconn = New ADODB.Connection
oconn.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Suket;User Id=system;Password=tiger;"
Set rs = New ADODB.Recordset
With cmd
    .ActiveConnection = oconn
    .CommandText = "SELECT festname from fest;"
    .CommandType = adCmdText
End With
With rs
    .CursorType = adOpenStatic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open cmd
End With
If rs.EOF = False Then
    rs.MoveFirst
    Do
      List1.AddItem (CStr(rs!festname))
      rs.MoveNext
    Loop Until rs.EOF = True
End If

  Set cmd = Nothing
  Set rs = Nothing
End Sub


Public Function reload()
Call Form_Load
End Function

Private Sub Label1_Click()
Form1.Show
Form3.Hide
End Sub

Private Sub Label4_Click()
selfest = List1.Text
If (selfest = "") Then
    Label3.Visible = True
    Exit Sub
ElseIf (selfest <> "") Then
    'Goto login form with selfest as current fest id
    Form1.festname = selfest
    Form8.Show
    Form3.Hide
End If
End Sub

