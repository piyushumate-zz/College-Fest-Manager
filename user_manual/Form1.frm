VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6600
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4455
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   6495
      Begin VB.TextBox Text1 
         Height          =   3375
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   5895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Get"
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   3960
         Width           =   3735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = Inet1.OpenURL("http://login.smsgatewayhub.com/API/WebSMS/Http/v1.0a/index.php?username=suket22&password=385q40yszz&sender=Pict&to=9595798239&message=Hello+Test+Message+Vb&reqid=1&format={json|text}&route_id=a:1:{i:22;s:11:%22Promotional%22;}&sendondate=30-07-2013T09:19:20")
End Sub
