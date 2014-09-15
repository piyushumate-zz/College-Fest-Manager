VERSION 5.00
Begin VB.Form Formtrial 
   Caption         =   "College Fest Manager - Begin"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12900
   LinkTopic       =   "Form3"
   Picture         =   "Formtrial.frx":0000
   ScaleHeight     =   7380
   ScaleWidth      =   12900
   StartUpPosition =   3  'Windows Default
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
      ItemData        =   "Formtrial.frx":26269
      Left            =   5280
      List            =   "Formtrial.frx":2626B
      TabIndex        =   2
      Top             =   2760
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Manage Fest"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Set Up Fest"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      UseMaskColor    =   -1  'True
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   6720
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   5520
      Width           =   3135
   End
End
Attribute VB_Name = "Formtrial"
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

Private Sub Image1_Click()

End Sub
