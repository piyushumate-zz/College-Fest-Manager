VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   7140
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12555
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7140
   ScaleWidth      =   12555
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   4200
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   2640
      Width           =   3135
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   7560
      TabIndex        =   0
      Top             =   2640
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      _Version        =   393216
      Format          =   103743489
      CurrentDate     =   41541
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   7560
      TabIndex        =   6
      Top             =   4320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      _Version        =   393216
      Format          =   103743489
      CurrentDate     =   41541
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9000
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   4200
      Top             =   5640
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Create Fest"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4200
      TabIndex        =   9
      Top             =   5760
      Width           =   3495
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Setup A New Fest"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4080
      TabIndex        =   5
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "End date"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7560
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Start date"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "College name"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fest Name"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmd1 As New ADODB.Command
Public startdate, enddate As Date
Public festname, collegename As String

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

Private Sub Label1_Click()
festname = Text1.Text
collegename = Text2.Text
startdate = DTPicker1.Value
enddate = DTPicker2.Value
If (startdate > enddate) Then
 Label5.Visible = True
Else
 Label5.Visible = False
End If

If Text1.Text = "" Then
 Label6.Visible = True
Else
 Label6.Visible = False
End If

If Text2.Text = "" Then
 Label7.Visible = True
Else
  Label7.Visible = False
End If

If (startdate <= enddate And Text1.Text <> "" And Text2.Text <> "") Then
'validation checks end here
Dim str, date1, date2 As String
date1 = changemonth(CStr(startdate))
date2 = changemonth(CStr(enddate))
str = "INSERT INTO FEST VALUES (" + "'" + festname + "'" + "," + "'" + collegename + "'" + "," + "'" + date1 + "'" + "," + "'" + date2 + "'" + ");"

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
  Set cmd = Nothing
  Set rs = Nothing
 Form4.Show
 Form1.Hide
End If
End Sub

