VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form12 
   Caption         =   "Form8"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15030
   BeginProperty Font 
      Name            =   "Tempus Sans ITC"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form8"
   Picture         =   "form12.frx":0000
   ScaleHeight     =   9075
   ScaleWidth      =   15030
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo4 
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
      Left            =   3120
      TabIndex        =   9
      Top             =   4800
      Width           =   855
   End
   Begin VB.ComboBox Combo5 
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
      Left            =   240
      TabIndex        =   8
      Top             =   4800
      Width           =   2175
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   3120
      TabIndex        =   7
      Top             =   3000
      Width           =   2175
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
      Top             =   3000
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      Left            =   7440
      TabIndex        =   5
      Top             =   2400
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5106
      _Version        =   393216
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
         Size            =   12
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   11880
      TabIndex        =   11
      Top             =   7200
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
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
      Left            =   5280
      TabIndex        =   10
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Round"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Participant Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Event Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt ID"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5280
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CB_FINDSTRING = &H14C
Private Const CB_SHOWDROPDOWN = &H14F
Private Const LB_FINDSTRING = &H18F
Private Const CB_ERR = (-1)

Private Declare Function SendMessage Lib _
    "user32" Alias "SendMessageA" (ByVal _
    hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) _
    As Long
    'Dim Form3.oconn As New ADODB.Form3.oconnnection
    'Dim Form3.rs As New ADODB.Recordset
    Dim t1, t2, t3, t4, t5 As Integer
    Dim temp1, temp2, temp3 As String
    
    

Private Sub Combo1_Change()
t1 = 1
 'Form3.rs.CuForm3.rsorLocation = adUseClient
'If t3 = 3 And Combo1.Text <> "" And t4 <> 4 And t5 <> 5 And Combo3.Text <> "" Then
 
 ' Form3.rs.Open "select * from participant where receiptid=" & Combo1.Text & " and eventname='" & Combo3.Text & "'", Form3.oconn, adOpenDynamic, adLockOptimistic

'ElseIf t3 = 3 And Combo1.Text <> "" And t4 <> 4 And t5 <> 5 And Combo3.Text = "" Then
 ' Form3.rs.Open "select * from participant where receiptid=" & Combo1.Text & " ", Form3.oconn, adOpenDynamic, adLockOptimistic

'ElseIf t3 = 3 And Combo1.Text = "" And t4 <> 4 And t5 <> 5 And Combo3.Text = "" Then
 ' Form3.rs.Open "select * from participant ", Form3.oconn, adOpenDynamic, adLockOptimistic
  

End Sub

Private Sub Combo1_GotFocus()
SendMessage Combo1.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
Dim CB As Long
    Dim FindString As String
    
    If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub
    
    If Combo1.SelLength = 0 Then
        FindString = Combo1.Text & Chr$(KeyAscii)
    Else
        FindString = Left$(Combo1.Text, Combo1.SelStart) & Chr$(KeyAscii)
    End If
    
    SendMessage Combo1.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&

    CB = SendMessage(Combo1.hwnd, CB_FINDSTRING, -1, ByVal FindString)
    
    If CB <> CB_ERR Then
        Combo1.ListIndex = CB
        Combo1.SelStart = Len(FindString)
        Combo1.SelLength = Len(Combo1.Text) - Combo1.SelStart
    End If
    
    KeyAscii = 0
End Sub

Private Sub Combo3_Change()
t3 = 3

End Sub

Private Sub Combo3_GotFocus()
SendMessage Combo3.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
Dim CB As Long
    Dim FindString As String
    
    If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub
    
    If Combo3.SelLength = 0 Then
        FindString = Combo3.Text & Chr$(KeyAscii)
    Else
        FindString = Left$(Combo3.Text, Combo3.SelStart) & Chr$(KeyAscii)
    End If
    
    SendMessage Combo3.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&

    CB = SendMessage(Combo3.hwnd, CB_FINDSTRING, -1, ByVal FindString)
    
    If CB <> CB_ERR Then
        Combo3.ListIndex = CB
        Combo3.SelStart = Len(FindString)
        Combo3.SelLength = Len(Combo3.Text) - Combo3.SelStart
    End If
    
    KeyAscii = 0
End Sub

Private Sub Combo4_Change()
t4 = 4

End Sub

Private Sub Combo4_GotFocus()
SendMessage Combo4.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
Dim CB As Long
    Dim FindString As String
    
    If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub
    
    If Combo4.SelLength = 0 Then
        FindString = Combo4.Text & Chr$(KeyAscii)
    Else
        FindString = Left$(Combo4.Text, Combo4.SelStart) & Chr$(KeyAscii)
    End If
    
    SendMessage Combo4.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&

    CB = SendMessage(Combo4.hwnd, CB_FINDSTRING, -1, ByVal FindString)
    
    If CB <> CB_ERR Then
        Combo4.ListIndex = CB
        Combo4.SelStart = Len(FindString)
        Combo4.SelLength = Len(Combo4.Text) - Combo4.SelStart
    End If
    
    KeyAscii = 0
End Sub

Private Sub Combo5_Change()
t5 = 5
End Sub

Private Sub Combo5_GotFocus()
SendMessage Combo5.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
Dim CB As Long
    Dim FindString As String
    
    If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub
    
    If Combo5.SelLength = 0 Then
        FindString = Combo5.Text & Chr$(KeyAscii)
    Else
        FindString = Left$(Combo5.Text, Combo5.SelStart) & Chr$(KeyAscii)
    End If
    
    SendMessage Combo5.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&

    CB = SendMessage(Combo5.hwnd, CB_FINDSTRING, -1, ByVal FindString)
    
    If CB <> CB_ERR Then
        Combo5.ListIndex = CB
        Combo5.SelStart = Len(FindString)
        Combo5.SelLength = Len(Combo5.Text) - Combo1.SelStart
    End If
    
    KeyAscii = 0
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
Combo4.AddItem "ALL"
Combo4.AddItem "1"
Combo4.AddItem "2"
Combo4.AddItem "3"
Combo4.AddItem "4"
Combo4.ListIndex = 0

Form3.rs.Open "select distinct receiptid from participant;", Form3.oconn, adOpenDynamic, adLockOptimistic
While Not Form3.rs.EOF
Combo1.AddItem Form3.rs(0)
Form3.rs.MoveNext
Wend
Form3.rs.Close

Form3.rs.Open "select eventname from evtmain where festname='Credenz';", Form3.oconn, adOpenDynamic, adLockOptimistic           'select festname to be Form3.oconnsidered'
While Not Form3.rs.EOF
Combo3.AddItem Form3.rs(0)
Form3.rs.MoveNext
Wend
Form3.rs.Close

Form3.rs.Open "select pname from participant;", Form3.oconn, adOpenDynamic, adLockOptimistic
While Not Form3.rs.EOF
Combo5.AddItem Form3.rs(0)
Form3.rs.MoveNext
Wend

End Sub

Private Sub Label3_Click()
Form3.rs.Close

Form3.rs.CursorLocation = adUseClient

 If Combo1.Text = "" And Combo3.Text = "" And Combo4.Text = "ALL" And Combo5.Text = "" Then
   Form3.rs.Open "select * from participant ", Form3.oconn, adOpenDynamic, adLockOptimistic
   
 ElseIf Combo1.Text = "" And Combo5.Text = "" And Combo3.Text <> "" And Combo4.Text <> "ALL" Then
   Form3.rs.Open "select * from participant where eventname='" & Combo3.Text & "' AND cround = " & Combo4.Text & " ", Form3.oconn, adOpenDynamic, adLockOptimistic
   
   ElseIf Combo3.Text = "" And Combo1.Text = "" And Combo4.Text <> "ALL" And Combo5.Text <> "" Then
   Form3.rs.Open "select * from participant where pname='" & Combo5.Text & "' AND cround=" & Combo4.Text & "", Form3.oconn, adOpenDynamic, adLockOptimistic
   
   
   ElseIf Combo1.Text = "" And Combo4.Text = "ALL" And Combo3.Text <> "" And Combo5.Text <> "" Then
   Form3.rs.Open "select * from participant where eventname='" & Combo3.Text & "' AND pname='" & Combo5.Text & "'", Form3.oconn, adOpenDynamic, adLockOptimistic
   
   ElseIf Combo5.Text = "" And Combo4.Text = "ALL" And Combo1.Text <> "" And Combo3.Text <> "" Then
   Form3.rs.Open "select * from participant where receiptid=" & Combo1.Text & " AND eventname='" & Combo3.Text & "'", Form3.oconn, adOpenDynamic, adLockOptimistic
   
   ElseIf Combo5.Text = "" And Combo3.Text = "" And Combo1.Text <> "" And Combo4.Text <> "ALL" Then
   Form3.rs.Open "select * from participant where receiptid=" & Combo1.Text & " AND cround=" & Combo4.Text & "", Form3.oconn, adOpenDynamic, adLockOptimistic
   
   ElseIf Combo3.Text = "" And Combo4.Text = "ALL" And Combo1.Text <> "" And Combo5.Text <> "" Then
   Form3.rs.Open "select * from participant where receiptid=" & Combo1.Text & " AND pname='" & Combo5.Text & "'", Form3.oconn, adOpenDynamic, adLockOptimistic
   
   ElseIf Combo1.Text = "" And Combo5.Text <> "" And Combo3.Text <> "" And Combo4.Text <> "ALL" Then
   Form3.rs.Open "select * from participant where pname='" & Combo5.Text & "' and eventname='" & Combo3.Text & "' and cround=" & Combo4.Text & "", Form3.oconn, adOpenDynamic, adLockOptimistic
   
   ElseIf Combo3.Text = "" And Combo1.Text <> "" And Combo4.Text <> "ALL" And Combo5.Text <> "" Then
   Form3.rs.Open "select * from participant where  receiptid=" & Combo1.Text & "  and pname='" & Combo5.Text & "' and cround=" & Combo4.Text & "", Form3.oconn, adOpenDynamic, adLockOptimistic
   
   ElseIf Combo4.Text = "ALL" And Combo3.Text <> "" And Combo1.Text <> "" And Combo5.Text <> "" Then
   Form3.rs.Open "select * from participant where eventname='" & Combo3.Text & "' and receiptid=" & Combo1.Text & " and pname='" & Combo5.Text & "'", Form3.oconn, adOpenDynamic, adLockOptimistic
   
   ElseIf Combo5.Text = "" And Combo3.Text <> "" And Combo1.Text <> "" And Combo4.Text <> "ALL" Then
   Form3.rs.Open "select * from participant where eventname='" & Combo3.Text & "' and receiptid=" & Combo1.Text & " and cround=" & Combo4.Text & "", Form3.oconn, adOpenDynamic, adLockOptimistic
   
   ElseIf Combo1.Text = "" And Combo3.Text = "" And Combo4.Text = "ALL" And Combo5.Text <> "" Then
   Form3.rs.Open "select * from participant where pname='" & Combo5.Text & "'", Form3.oconn, adOpenDynamic, adLockOptimistic
   
   
   ElseIf Combo1.Text = "" And Combo3.Text = "" And Combo5.Text = "" And Combo4.Text <> "ALL" Then
   Form3.rs.Open "select * from participant where cround=" & Combo4.Text & "", Form3.oconn, adOpenDynamic, adLockOptimistic
   
   ElseIf Combo1.Text = "" And Combo4.Text = "ALL" And Combo5.Text = "" And Combo3.Text <> "" Then
   Form3.rs.Open "select * from participant where eventname='" & Combo3.Text & "'", Form3.oconn, adOpenDynamic, adLockOptimistic
   
   ElseIf Combo3.Text = "" And Combo4.Text = "ALL" And Combo5.Text = "" And Combo1.Text <> "" Then
   Form3.rs.Open "select * from participant where receiptid=" & Combo1.Text & "", Form3.oconn, adOpenDynamic, adLockOptimistic
   
   ElseIf Combo1.Text <> "" And Combo3.Text <> "" And Combo5.Text <> "" And Combo4.Text <> "ALL" Then
   Form3.rs.Open "select * from participant where receiptid=" & Combo1.Text & " AND eventname='" & Combo3.Text & "' AND cround=" & Combo4.Text & " and pname='" & Combo5.Text & "'", Form3.oconn, adOpenDynamic, adLockOptimistic
   
   Else
     MsgBox "MATCH NOT FOUND"
   
   
 End If
 
   Set DataGrid1.DataSource = Form3.rs
   
   
   'Form3.rs.Update
   
   
   
   
   

End Sub

Private Sub Label7_Click()
Form12.Hide
Form13.Show
End Sub
