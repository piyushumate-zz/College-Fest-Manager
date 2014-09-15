VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14715
   BeginProperty Font 
      Name            =   "Tempus Sans ITC"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   8145
   ScaleWidth      =   14715
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   14295
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   2880
         TabIndex        =   2
         Top             =   1080
         Width           =   11295
         Begin VB.Frame Frame4 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5895
            Left            =   4320
            TabIndex        =   14
            Top             =   600
            Visible         =   0   'False
            Width           =   6855
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
               Left            =   480
               TabIndex        =   17
               Text            =   "Select Available Auditorium"
               Top             =   2280
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
               Left            =   480
               TabIndex        =   16
               Text            =   "Select Available Laboratory"
               Top             =   1560
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
               Left            =   480
               TabIndex        =   15
               Text            =   "Select Available Classroom"
               Top             =   840
               Width           =   5175
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Booking Finished"
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
               Height          =   615
               Left            =   3600
               TabIndex        =   32
               Top             =   4440
               Width           =   2655
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Book Auditorium"
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
               Height          =   615
               Left            =   480
               TabIndex        =   31
               Top             =   4440
               Width           =   2775
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Book Laboratory"
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
               Left            =   3600
               TabIndex        =   30
               Top             =   3600
               Width           =   2655
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Book Classroom"
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
               Left            =   480
               TabIndex        =   28
               Top             =   3600
               Width           =   2535
            End
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   5655
            Left            =   4320
            TabIndex        =   13
            Top             =   720
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   9975
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            AllowDelete     =   -1  'True
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   8
            Top             =   2520
            Width           =   2295
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
            Left            =   360
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "Select Event Here"
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00000000&
            Caption         =   "Required Slot Details"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   3015
            Left            =   120
            TabIndex        =   9
            Top             =   3240
            Width           =   3975
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   375
               Left            =   2040
               TabIndex        =   20
               Top             =   1920
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   105709570
               CurrentDate     =   41549
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   375
               Left            =   240
               TabIndex        =   19
               Top             =   1920
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   105709570
               CurrentDate     =   41549
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   375
               Left            =   840
               TabIndex        =   18
               Top             =   840
               Width           =   2055
               _ExtentX        =   3625
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
               Format          =   105709569
               CurrentDate     =   41549
            End
            Begin VB.Label Label15 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Go"
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
               Height          =   375
               Left            =   1440
               TabIndex        =   27
               Top             =   2520
               Width           =   975
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Finish Time"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   2280
               TabIndex        =   12
               Top             =   1560
               Width           =   1215
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Start time"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   360
               TabIndex        =   11
               Top             =   1560
               Width           =   1095
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   1560
               TabIndex        =   10
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   4560
            Width           =   1695
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Round Number"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   360
            TabIndex        =   5
            Top             =   1920
            Width           =   2535
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Event "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   360
            TabIndex        =   4
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
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
            Left            =   3360
            TabIndex        =   3
            Top             =   4920
            Width           =   5295
         End
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Back to Home Page"
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
         Left            =   360
         TabIndex        =   26
         Top             =   6120
         Width           =   2175
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Delete Slot"
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
         Height          =   615
         Left            =   480
         TabIndex        =   25
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add Slot"
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
         Left            =   120
         TabIndex        =   24
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Check Slots"
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
         Height          =   615
         Left            =   240
         TabIndex        =   22
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   5640
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Slots"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   5640
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Book Classroom"
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
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Check Slots"
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
      Height          =   615
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim eventname, str As String
Dim roundno As Integer
Dim cmd1 As ADODB.Command
Dim date1 As String
Dim time1, time2 As String



Private Sub Command1_Click()

End Sub


Private Sub Command5_Click()
Call addtorounds(1)
End Sub

Private Sub Command6_Click()
Call addtorounds(2)
End Sub

Private Sub Command7_Click()

End Sub




Private Sub Form_Load()
Set cmd1 = New ADODB.Command
Set Form3.rs = New ADODB.Recordset
With cmd1
    .ActiveConnection = Form3.oconn
    .CommandText = "select eventname from evtmain where festname = '" + Form1.festname + "';"
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
      Combo1.AddItem (CStr(Form3.rs!eventname))
      Form3.rs.MoveNext
    Loop Until Form3.rs.EOF = True
End If

Set Form3.rs = Nothing
Set cmd1 = Nothing
Frame3.Visible = False
End Sub


Function addtorounds(flag As Integer)
Dim currfeid, curriid As Integer
'First find corresponding feid and iid
str = "select feid from evtmain where eventname= " + "'" + eventname + "'" + "and festname = " + "'" + Form1.festname + "'" + ";"
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
str = "insert into rounds values(" + CStr(currfeid) + "," + "'" + CStr(changemonth(DTPicker1.Value)) + "'" + "," + "'" + Right(CStr(DTPicker2.Value), 8) + "'" + "," + "'" + Right(CStr(DTPicker3.Value), 8) + "'" + "," + Text1.Text + "," + CStr(curriid) + ");"
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
Set Form3.rs = Nothing
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
Private Sub MyRefresh()
Combo6.Clear
Combo5.Clear
Combo4.Clear
Call Addall
Combo6.Text = "Select a Classroom"
Combo5.Text = "Select a Laboratory"
Combo4.Text = "Select an Auditorium"
'clear all then add all and then remove those that don't match

strevent = eventname
date1 = CStr(DTPicker1.Value)
date1 = changemonth(date1)
time1 = Right(CStr(DTPicker2.Value), 8)
time2 = Right(CStr(DTPicker3.Value), 8)

'Step 1 - Get date,time of all bookings of fest classrooms
Set Form3.rs = Nothing
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

'Step 2 - Compare the data in the form with all the data from query. if satisfied add name to combobox

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


Private Sub Label10_Click()
eventname = Combo1.Text
roundno = CInt(Text1.Text)
Set cmd1 = New ADODB.Command
Set Form3.rs = New ADODB.Recordset
str = "select r.rdate,r.stime,r.etime,i.name from rounds r,infra i where r.iid=i.iid and feid =(select feid from evtmain where eventname='" + eventname + "' and festname = '" + Form1.festname + "') and roundno = " + CStr(roundno) + ";"
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
Set DataGrid1.DataSource = Form3.rs
End Sub

Private Sub Label12_Click()
Frame3.Visible = True
Frame4.Visible = True
End Sub

Private Sub Label13_Click()
Dialog1.Show
End Sub

Private Sub Label14_Click()
Form13.Show
Form5.Hide
End Sub

Private Sub Label15_Click()
date1 = changemonth(CStr(DTPicker1.Value))
stime1 = Right(CStr(DTPicker2.Value), 8)
etime1 = Right(CStr(DTPicker2.Value), 8)
eventname = Combo1.Text
roundno = CInt(Text1.Text)
Call MyRefresh
End Sub

Private Sub Label16_Click()
Call addtorounds(1)
End Sub

Private Sub Label18_Click()
Call addtorounds(2)
End Sub

Private Sub Label19_Click()
Call addtorounds(3)
End Sub

Private Sub Label20_Click()
Call Command1_Click
Frame3.Visible = False
Frame4.Visible = False
End Sub

Private Sub Label9_Click()
Dialog1.Show
End Sub
