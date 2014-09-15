VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form5"
   Picture         =   "Form8.frx":0000
   ScaleHeight     =   7590
   ScaleWidth      =   11835
   StartUpPosition =   3  'Windows Default
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
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4320
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1800
      TabIndex        =   0
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   4680
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
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
      Left            =   4680
      TabIndex        =   6
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOG IN"
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
      Height          =   1335
      Left            =   1800
      TabIndex        =   5
      Top             =   600
      Width           =   8655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Username And Password do not match."
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   6360
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   1800
      TabIndex        =   3
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
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
      Left            =   1800
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public username
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

Private Sub Form_Load()
Set Form3.rs = New ADODB.Recordset
Form3.rs.Open "select username from login where festname ='" + Form1.festname + "';", Form3.oconn, adLockOptimistic, adOpenDynamic

While Not Form3.rs.EOF
Combo1.AddItem Form3.rs(0)
Form3.rs.MoveNext
Wend

Set Form3.rs = Nothing
End Sub


Private Sub Label5_Click()
Set Form3.rs = New ADODB.Recordset
Form3.rs.Open "select password from login where username ='" + Combo1.Text + "';", Form3.oconn, adLockOptimistic, adOpenDynamic
If (Form3.rs!Password <> Text2.Text) Then
    Label3.Visible = True
Else
    username = Combo1.Text
    Form13.Show  'go to homepage
    Form8.Hide
End If
Set Form3.rs = Nothing

End Sub
