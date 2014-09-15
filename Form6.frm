VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13365
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
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   8745
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7920
      TabIndex        =   11
      Text            =   "Select an Auditorium"
      Top             =   4920
      Width           =   5175
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7920
      TabIndex        =   10
      Text            =   "Select a Laboratory"
      Top             =   3480
      Width           =   5175
   End
   Begin VB.ComboBox Combo6 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7920
      TabIndex        =   9
      Text            =   "Select a Classroom"
      Top             =   2160
      Width           =   5175
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   2
      Text            =   "Select Event"
      Top             =   2160
      Width           =   4095
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "Form6.frx":345DE
      Left            =   360
      List            =   "Form6.frx":345F4
      TabIndex        =   1
      Text            =   "Select which Round"
      Top             =   3480
      Width           =   4095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   390
      Left            =   1680
      TabIndex        =   3
      Top             =   4920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   103809025
      CurrentDate     =   41542
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   6000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   103809026
      CurrentDate     =   41544
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   6000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   103809026
      CurrentDate     =   41544
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   5520
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   8280
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   5520
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   2880
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Finish"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5520
      TabIndex        =   16
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Book Auditorium"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8280
      TabIndex        =   15
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Book Laboratory"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5520
      TabIndex        =   14
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Book Classroom"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2880
      TabIndex        =   13
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1335
      Left            =   5760
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "                  Refresh"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   5760
      TabIndex        =   12
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   2280
      X2              =   11280
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Select End Time"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Start Time"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Date"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Set Rounds For Events"
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
      Height          =   855
      Left            =   2640
      TabIndex        =   0
      Top             =   720
      Width           =   7695
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmd1 As New ADODB.Command
Dim roundcnt, eventname, roundsel As String
Dim time1, time2 As String
Dim date1 As String
Dim str As String
Dim startdate As Date
Dim strevent As String
Dim currfeid, curriid As Integer

Private Sub MyRefresh()
Combo6.Clear
Combo5.Clear
Combo4.Clear
Call Addall
Combo6.Text = "Select a Classroom"
Combo5.Text = "Select a Laboratory"
Combo4.Text = "Select an Auditorium"
'clear all then add all and then remove those that don't match

strevent = Combo3.Text
date1 = CStr(DTPicker1.Value)
date1 = changemonth(date1)
time1 = Right(CStr(DTPicker2.Value), 8)
time2 = Right(CStr(DTPicker3.Value), 8)

'Step 1 - Get date,time of all bookings of fest classrooms

str = "select i.iid,r.rdate,r.stime,r.etime,i.name from rounds r,infra i,evtmain e where r.iid=i.iid and r.feid=e.feid  and i.type='Classrooms' and e.festname='" + Form1.festname + "';"
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

'Step 2 - Compare the data in the form with all the data from query. if satisfied remove name to combobox

If Form3.rs.EOF = False Then
    Form3.rs.MoveFirst
    Do
      Call checkslot(CStr(changemonth(Form3.rs!rDate)), CStr(Form3.rs!stime), CStr(Form3.rs!etime), 1)
      Form3.rs.MoveNext
    Loop Until Form3.rs.EOF = True
End If
Set Form3.rs = Nothing

'Now repeating steps for labs and audis
str = "select i.iid,r.rdate,r.stime,r.etime,i.name from rounds r,infra i,evtmain e where r.iid=i.iid and r.feid=e.feid  and i.type='Laboratories' and e.festname='" + Form1.festname + "';"
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

'Step 2 - Compare the data in the form with all the data from query. if satisfied add name to combobox

If Form3.rs.EOF = False Then
    Form3.rs.MoveFirst
    Do
      Call checkslot(CStr(changemonth(Form3.rs!rDate)), CStr(Form3.rs!stime), CStr(Form3.rs!etime), 2)
      Form3.rs.MoveNext
    Loop Until Form3.rs.EOF = True
End If
Set Form3.rs = Nothing

str = "select i.iid,r.rdate,r.stime,r.etime,i.name from rounds r,infra i,evtmain e where r.iid=i.iid and r.feid=e.feid  and i.type='Auditoriums' and e.festname='" + Form1.festname + "';"
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

'Step 2 - Compare the data in the form with all the data from query. if satisfied add name to combobox

If Form3.rs.EOF = False Then
    Form3.rs.MoveFirst
    Do
      Call checkslot(CStr(changemonth(Form3.rs!rDate)), CStr(Form3.rs!stime), CStr(Form3.rs!etime), 3)
      Form3.rs.MoveNext
    Loop Until Form3.rs.EOF = True
End If
Set Form3.rs = Nothing
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub



Private Sub Form_Load()
With cmd1
    .ActiveConnection = Form3.oconn
    .CommandText = "select eventname from evtmain where festname='" + Form1.festname + "';"
    .CommandType = adCmdText
End With
With Form3.rs
    .CursorType = adOpenStatic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open cmd1
End With
If Form3.rs.EOF = False Then
    Form3.rs.MoveFirst
    Do
      Combo3.AddItem (CStr(Form3.rs!eventname))
      Form3.rs.MoveNext
    Loop Until Form3.rs.EOF = True
End If

Set Form3.rs = Nothing

'now setting restriction on datetimepicker
DTPicker1.MinDate = Form2.evtstart
DTPicker1.MaxDate = Form2.evtend

End Sub


Function addtorounds(flag As Integer)

'First find corresponding feid and iid
str = "select feid from evtmain where eventname= " + "'" + strevent + "'" + "and festname = " + "'" + Form1.festname + "'" + ";"
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
currfeid = CInt(Form3.rs!feid)  'currfeid obtained
Set Form3.rs = Nothing

'now iid
If (flag = 1) Then
str = "select iid from infra where name = " + "'" + Combo6.Text + "'" + "and festname = " + "'" + Form1.festname + "'" + ";"
ElseIf (flag = 2) Then
str = "select iid from infra where name = " + "'" + Combo5.Text + "'" + "and festname = " + "'" + Form1.festname + "'" + ";"
Else
str = "select iid from infra where name = " + "'" + Combo4.Text + "'" + "and festname = " + "'" + Form1.festname + "'" + ";"
End If
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
curriid = CInt(Form3.rs!iid)   'curriid obtained
Set Form3.rs = Nothing

'now adding to round
str = "insert into rounds values(" + CStr(currfeid) + "," + "'" + CStr(changemonth(DTPicker1.Value)) + "'" + "," + "'" + Right(CStr(DTPicker2.Value), 8) + "'" + "," + "'" + Right(CStr(DTPicker3.Value), 8) + "'" + "," + CStr(CInt(Combo2.Text)) + "," + CStr(curriid) + ");"
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
Set Form3.rs = Nothing
End Function

Function changemonth(date1 As String) As String
Dim str, temp As String
str = Left(date1, 3)
date1 = Right(date1, Len(date1) - 3)
temp = Left(date1, 2)

If temp = "01" Then
str = str + "JAN-"
ElseIf temp = "02" Then
str = str + "FEB-"
ElseIf temp = "03" Then
str = str + "MAR-"
ElseIf temp = "04" Then
str = str + "APR-"
ElseIf temp = "05" Then
str = str + "MAY-"
ElseIf temp = "06" Then
str = str + "JUN-"
ElseIf temp = "07" Then
str = str + "JUL-"
ElseIf temp = "08" Then
str = str + "AUG-"
ElseIf temp = "09" Then
str = str + "SEP-"
ElseIf temp = "10" Then
str = str + "OCT-"
ElseIf temp = "11" Then
str = str + "NOV-"
ElseIf temp = "12" Then
str = str + "DEC-"
End If

str = str + Right(date1, 2)
changemonth = str
End Function

Function checkslot(datef As String, stimef As String, etimef As String, flag As Integer)
If (datef <> date1) Then

    ' If (flag = 1) Then
    '     Combo6.RemoveItem (FindCBIndex(Combo6, Form3.rs!Name))
    ' ElseIf (flag = 2) Then
    '     Combo5.RemoveItem (FindCBIndex(Combo5, Form3.rs!Name))
    ' Else
    '     Combo4.RemoveItem (FindCBIndex(Combo4, Form3.rs!Name))
    ' End If
Else
    If (time1 >= stimef And time2 <= etimef) Then   'completely inside
        If (flag = 1) Then
            Combo6.RemoveItem (FindCBIndex(Combo6, Form3.rs!Name))
        ElseIf (flag = 2) Then
            Combo5.RemoveItem (FindCBIndex(Combo5, Form3.rs!Name))
        Else
            Combo4.RemoveItem (FindCBIndex(Combo4, Form3.rs!Name))
        End If
         
    ElseIf (time1 >= stimef And time1 <= etimef) Then   'partially inside start
        If (flag = 1) Then
            Combo6.RemoveItem (FindCBIndex(Combo6, Form3.rs!Name))
        ElseIf (flag = 2) Then
            Combo5.RemoveItem (FindCBIndex(Combo5, Form3.rs!Name))
        Else
            Combo4.RemoveItem (FindCBIndex(Combo4, Form3.rs!Name))
        End If
        
    ElseIf (time2 >= stimef And time2 <= etimef) Then   'partially inside end
        If (flag = 1) Then
            Combo6.RemoveItem (FindCBIndex(Combo6, Form3.rs!Name))
        ElseIf (flag = 2) Then
            Combo5.RemoveItem (FindCBIndex(Combo5, Form3.rs!Name))
        Else
            Combo4.RemoveItem (FindCBIndex(Combo4, Form3.rs!Name))
        End If
    Else
        Exit Function    'not inside time slot
    End If
End If

End Function


Public Function FindCBIndex(ByRef cbComboBox As ComboBox, ByRef strSearchValue As String) As Integer
    Dim n As Integer
    For n = 0 To cbComboBox.ListCount - 1
        If cbComboBox.List(n) = strSearchValue Then
          ' // Return the found index
            FindCBIndex = n
          ' // and exit
            Exit Function
        End If
    Next
  ' // Set not found value
    FindCBIndex = -1
End Function


Function Addall()
str = "select name as Name,type as Type from infra where festname = " + "'" + Form1.festname + "'" + ";"
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
If Form3.rs.EOF = False Then
    Form3.rs.MoveFirst
    Do
      If (Form3.rs!Type = "Classrooms") Then
        Combo6.AddItem (Form3.rs!Name)
      ElseIf (Form3.rs!Type = "Laboratories") Then
        Combo5.AddItem (Form3.rs!Name)
      ElseIf (Form3.rs!Type = "Auditoriums") Then
        Combo4.AddItem (Form3.rs!Name)
      End If
      Form3.rs.MoveNext
    Loop Until Form3.rs.EOF = True
End If
Set Form3.rs = Nothing

End Function


Private Sub Label5_Click()
Call MyRefresh
End Sub

Private Sub Label6_Click()
If (Combo6.Text = "Select a Classroom") Then
    MsgBox "Select A classroom first "
Else
    Call addtorounds(1)
End If
Call MyRefresh
End Sub

Private Sub Label7_Click()
If (Combo5.Text = "Select a Laboratory") Then
    MsgBox "Select A laboratory first "
Else
    Call addtorounds(2)
End If
Call MyRefresh
End Sub

Private Sub Label8_Click()
If (Combo4.Text = "Select an Auditorium") Then
    MsgBox "Select an Auditorium first "
Else
    Call addtorounds(3)
End If
Call MyRefresh
End Sub

Private Sub Label9_Click()
Form6.Hide
Form7.Show

End Sub
