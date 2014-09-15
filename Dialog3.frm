VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Dialog3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   5235
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Dialog3.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7440
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
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
      Height          =   1575
      Left            =   360
      TabIndex        =   2
      Top             =   3360
      Width           =   5895
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
      Height          =   1455
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   6600
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Submit"
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
      Height          =   735
      Left            =   6600
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter body of email"
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
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter body of sms"
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
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "Dialog3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public string1 As String
Public string2 As String
Public m As String
Dim Attach(20) As String
Option Explicit

Private Sub SendEmail()
  Dim i As Integer
  On Error GoTo Errhandler
  
    Dim iMsg As CDO.Message
    Dim iConf As CDO.Configuration
    
    Set iMsg = New CDO.Message
    Set iConf = New Configuration
    Dim AttachmentCount As Integer
    i = 1
 '   AttachmentCount = 2     'sample value for sending two attachments
    
    'Apply settings to the configuration object
    With iConf.Fields
        ' Don't bother with the following 3 lines of code if the SMTP server you will be connecting to
        ' does not require authentication
    
        ' Specify the authentication mechanism to basic (clear-text) authentication.
        .Item(cdoSMTPAuthenticate) = cdoBasic
        ' The username for authenticating to an SMTP server
        .Item(cdoSendUserName) = "projectvsps@gmail.com"  'enter your actual username here
        ' The password used to authenticate to an SMTP server
        .Item(cdoSendPassword) = "asdfghjkl@"    'enter your password here
        
        .Item(cdoSMTPUseSSL) = True
        
        'Secure socket layer. Gmail needs this to ensure encryption of data which is sent.
        
        ' How to send the mail
        .Item(cdoSendUsingMethod) = cdoSendUsingPort
        
        'Specify mail server , smtp is simple mail transfer protocol
        .Item(cdoSMTPServer) = "smtp.gmail.com"
        
        'Specify the timeout in seconds
        .Item(cdoSMTPConnectionTimeout) = 10
        
        ' The port on which the SMTP service specified by the smtpserver field is listening for connections (typically 25)
        .Item(cdoSMTPServerPort) = 25
        
        ' Ensure configuration is up to date
        .Update
    End With

    Set iMsg.Configuration = iConf
    iMsg.To = m    'receiver
    iMsg.Subject = Form1.festname + "Notification" 'Subject here
    
    iMsg.TextBody = Text2.Text 'Body text here
      
    iMsg.From = "projectvsps@gmail.com"     'same username again
    'The next line will work for single attachment
    'iMsg.AddAttachment ("C:\Users\hp\Desktop\foryou.txt")
    
    
    'If you want to send multiple attachments then store the url string or filepaths in the array attach() and execute below code.
    
    'Attach(1) = "C:\Users\hp\Desktop\foryou.txt"
    'Attach(2) = "C:\Users\hp\Desktop\foryou1.txt"
    
    'If AttachmentCount Then 'Attachment's here
     'For i = 1 To AttachmentCount
      'If Attach(i) = "" Then Exit For 'make sure we got something here
      'iMsg.AddAttachment Attach(i)
     'DoEvents
     'Next
    'End If

    iMsg.Send
    DoEvents
    Set iMsg = Nothing
    
    Screen.MousePointer = 0
    
    DoEvents
    
    MsgBox "Email Sent!", , "E-Mail"
    
    
    Exit Sub
    
Errhandler:
  
  Screen.MousePointer = 0
  MsgBox Error$ + " - " + str(Err), vbOKOnly + vbCritical, "Email Error!"
  Exit Sub
End Sub

Private Sub SendSms()
Dim str As String
str = "http://login.smsgatewayhub.com/API/WebSMS/Http/v1.0a/index.php?username=suket22&password=385q40yszz&sender=" + Form1.festname + "to=" + m + "&message=" + "sendingthisdata" + "&reqid=1&format={json|text}&route_id=a:1:{i:22;s:11:%22Promotional%22;}&sendondate=30-07-2013T09:19:20"
Inet1.OpenURL (str)
'Inet1.OpenURL("http://login.smsgatewayhub.com/API/WebSMS/Http/v1.0a/index.php?username=suket22&password=385q40yszz&sender=Pict&to=9561640052&message=Hello+Test+Message+Vb&reqid=1&format={json|text}&route_id=a:1:{i:22;s:11:%22Promotional%22;}&sendondate=30-07-2013T09:19:20")
End Sub

Private Sub Label3_Click()
If (Text1.Text = "") Then
    MsgBox "Body Of SMS is empty!"
ElseIf (Text2.Text = "") Then
    MsgBox "Body Of Email is empty!"
Else
    string1 = Text1.Text
    string2 = Text2.Text
    Call changespaces
    Dialog3.Hide
    Form11.Show
    Call sendnotif
    
End If

End Sub


Private Sub sendnotif()
If Form11.email1 <> "" Then
     m = Form11.email1
     Call SendEmail
     End If
If Form11.email2 <> "" Then
     m = Form11.email2
     Call SendEmail
     End If
If Form11.email3 <> "" Then
     m = Form11.email3
     Call SendEmail
     End If
If Form11.email4 <> "" Then
     m = Form11.email4
     Call SendEmail
     End If
If Form11.sms1 <> "" Then
     m = Form11.sms1
     Call SendSms
     End If
If Form11.sms2 <> "" Then
     m = Form11.sms2
     Call SendSms
     End If
If Form11.sms3 <> "" Then
     m = Form11.sms3
     Call SendSms
     End If
If Form11.sms4 <> "" Then
     m = Form11.sms4
     Call SendSms
     End If


End Sub


Function changespaces()
Dim tempstr As String
Dim flag As Integer
flag = 0

Do Until flag = 1
    If (Left(string1, 1) = " ") Then
        tempstr = tempstr + "+"
    Else
        tempstr = tempstr + Left(string1, 1)
    End If
    string1 = Right(string1, Len(string1) - 1)
    If (Len(string1) = 0) Then
        flag = 1
    End If
Loop
string1 = tempstr
End Function
