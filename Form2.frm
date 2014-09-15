VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form2"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   8175
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   12
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   3120
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4335
      Left            =   8760
      TabIndex        =   1
      Top             =   2160
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      DefColWidth     =   67
      HeadLines       =   1
      RowHeight       =   21
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   4560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
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
      Format          =   105644033
      CurrentDate     =   41541
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   3120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
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
      Format          =   105644033
      CurrentDate     =   41541
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Finish"
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
      Left            =   2160
      TabIndex        =   15
      Top             =   6960
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add Event"
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
      Left            =   2280
      TabIndex        =   14
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Create Event"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   2040
      TabIndex        =   13
      Top             =   600
      Width           =   10815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5520
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Event Head"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   10
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Event Name"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "    Current List Of Events Added in Fest"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   2
      Top             =   7200
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmd1 As New ADODB.Command
Public evtstart, evtend As Date
Dim currfeid
Public eventname, eventhead As String







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


Private Sub Form_Load()
'setting restrictions on date
DTPicker1.MinDate = Form1.startdate
DTPicker1.MaxDate = Form1.enddate
DTPicker2.MinDate = Form1.startdate
DTPicker2.MaxDate = Form1.enddate

End Sub

Private Sub Label5_Click()
Dim flag
flag = 0
eventname = Text2.Text
eventhead = Text3.Text
evtstart = DTPicker1.Value
evtend = DTPicker2.Value

If (evtstart <= evtendd) Then
 Label2.Visible = True
ElseIf (evtstart > evtendd) Then
 Label2.Visible = False
End If
 

If (eventname = "") Then
 Label4.Visible = True
Else
 Label4.Visible = False
End If
 
If (eventhead = "") Then
 Label8.Visible = True
Else
 Label8.Visible = False
End If


If (evtstart <= evtend And eventname <> "" And eventhead <> "") Then
'Validations completed so now add to event table
 Dim str As String
 Dim date1, date2 As String
 date1 = changemonth(CStr(evtstart))
 date2 = changemonth(CStr(evtend))
 
 'Adding to evtmain table and generating feid
str = "INSERT INTO evtmain VALUES (" + "'" + Form1.festname + "'" + "," + "'" + eventname + "'" + "," + "feid_evtmain.nextval" + ");"
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
'Reading the current value of feid_evtmain
str = "select feid_evtmain.currval as feid from dual;"
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
currfeid = CStr(Form3.rs!feid)
Set Form3.rs = Nothing

   
'Now adding to event


 str = "INSERT INTO EVENT VALUES (" + currfeid + "," + "'" + eventhead + "'" + "," + "'" + date1 + "'" + "," + "'" + date2 + "'" + ");"
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
  'reload the datagrid with newly added event
  
  str = "select ev.eventname As EventName,e.startdate As Startdate ,e.enddate As Enddate from evtmain ev, event e where e.feid=ev.feid and ev.festname='" + Form1.festname + "';"
  Form3.rs.Open str, Form3.oconn, , , adCmdText
  Set DataGrid1.DataSource = Form3.rs
  
  Set cmd1 = Nothing
  Set Form3.rs = Nothing
  

End If

End Sub

Private Sub Label9_Click()
Form6.Show
Form2.Hide
End Sub

