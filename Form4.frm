VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12195
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   8445
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   8
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   5
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   7680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2640
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "Form4.frx":345DE
      Left            =   600
      List            =   "Form4.frx":345EB
      TabIndex        =   1
      Text            =   "Select one of the following"
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Save Current"
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
      Height          =   735
      Left            =   4320
      TabIndex        =   9
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   4320
      Top             =   7440
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Submit All"
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
      Height          =   735
      Left            =   4320
      TabIndex        =   7
      Top             =   7560
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   4320
      Top             =   6240
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Names Of"
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
      Height          =   735
      Left            =   8040
      TabIndex        =   6
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Capacity of"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   4080
      TabIndex        =   3
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Number of"
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
      Height          =   855
      Left            =   600
      TabIndex        =   2
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Infrastructure Details"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   7455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmd1 As New ADODB.Command
Dim strText1() As String
Dim strText2() As String
Dim strText3() As String
Dim flag1, flag2, flag3 As Integer
Dim labcnt, labcap, audicnt, audicap, classcnt, classcap As Integer
Dim str As String


Private Sub Combo1_Click()
Dim count As Integer
cnt = 0

If Combo1.Text = "Classrooms" Then
Label2 = "Number of Classrooms"
Label3 = "Capacity of Classrooms"
Label4 = "Names of Classrooms"
Text1.Text = CStr(classcnt)
Text2.Text = CStr(classcap)
If (flag1 <> 0) Then
    Text3.Text = ""
    Do Until cnt >= classcnt
     Text3.Text = Text3.Text + strText1(cnt) + vbCrLf
     cnt = cnt + 1
    Loop
    GoTo exi
Else
     Text3.Text = ""
End If
ElseIf Combo1.Text = "Auditoriums" Then
Label2 = "Number of Auditoriums"
Label3 = "Capacity of Auditoriums"
Label4 = "Names of Auditoriums"
Text1.Text = CStr(audicnt)
Text2.Text = CStr(audicap)
If (flag3 <> 0) Then
    Text3.Text = ""
    Do Until cnt >= audicnt
        Text3.Text = Text3.Text + strText3(cnt) + vbCrLf
        cnt = cnt + 1
    Loop
    GoTo exi
Else
    Text3.Text = ""
    
End If
ElseIf Combo1.Text = "Laboratories" Then
Label2 = "Number of Laboratories"
Label3 = "Capacity of Laboratories"
Label4 = "Names of Laboratories"
Text1.Text = CStr(labcnt)
Text2.Text = CStr(labcap)
If (flag2 <> 0) Then
    Text3.Text = ""
    Do Until cnt >= labcnt
        Text3.Text = Text3.Text + strText2(cnt) + vbCrLf
        cnt = cnt + 1
    Loop
    GoTo exi
Else
    Text3.Text = ""
End If
End If
exi:
End Sub


Private Sub Form_Load()
audicnt = 0
audicap = 0
labcnt = 0
labcap = 0
classcnt = 0
classcap = 0
flag1 = 0
flag2 = 0
flag3 = 0
Text1.Text = CStr(0)
Text2.Text = CStr(0)
End Sub
Private Sub Label5_Click()
Dim str1, str2, str3 As String
Dim i
i = 0


'adding all classrooms
Do While i < classcnt
str = "INSERT INTO INFRA VALUES (" + "iid_infra.nextval" + "," + "'" + "Classrooms" + "'" + "," + "'" + strText1(i) + "'" + "," + "'" + CStr(classcap) + "'" + "," + "'" + Form1.festname + "'" + ");"
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
i = i + 1
Loop

'adding all labs
i = 0
Do While i < labcnt
str = "INSERT INTO INFRA VALUES (" + "iid_infra.nextval" + "," + "'" + "Laboratories" + "'" + "," + "'" + strText2(i) + "'" + "," + "'" + CStr(labcap) + "'" + "," + "'" + Form1.festname + "'" + ");"
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
i = i + 1
Loop


'adding all audis
i = 0
Do While i < audicnt
str = "INSERT INTO INFRA VALUES (" + "iid_infra.nextval" + "," + "'" + "Auditoriums" + "'" + "," + "'" + strText3(i) + "'" + "," + "'" + CStr(audicap) + "'" + "," + "'" + Form1.festname + "'" + ");"
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
i = i + 1
Loop

Form2.Show
Form4.Hide
End Sub

Private Sub Label6_Click()
If Combo1.Text = "Select one of the following" Then
 MsgBox "Choose first! "
 Exit Sub
End If
If Combo1.Text = "Classrooms" Then
classcnt = CInt(Text1.Text)
classcap = CInt(Text2.Text)
    If (Text3.Text = "") Then
        flag1 = 0
    Else
        strText1() = Split(Text3.Text, vbCrLf)
        flag1 = 1
    End If

ElseIf Combo1.Text = "Auditoriums" Then
audicnt = CInt(Text1.Text)
audicap = CInt(Text2.Text)
    If (Text3.Text = "") Then
        flag3 = 0
    Else
        strText3() = Split(Text3.Text, vbCrLf)
        flag3 = 1
    End If


ElseIf Combo1.Text = "Laboratories" Then
labcnt = CInt(Text1.Text)
labcap = CInt(Text2.Text)
    If (Text3.Text = "") Then
        flag2 = 0
    Else
        strText2() = Split(Text3.Text, vbCrLf)
        flag2 = 1
    End If
End If

End Sub


