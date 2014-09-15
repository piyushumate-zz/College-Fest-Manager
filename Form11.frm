VERSION 5.00
Begin VB.Form Form11 
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14085
   LinkTopic       =   "Form2"
   Picture         =   "Form11.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   14085
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   11055
      Left            =   -120
      TabIndex        =   1
      Top             =   0
      Width           =   13455
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
         Left            =   3480
         TabIndex        =   47
         Top             =   1920
         Width           =   3255
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
         Left            =   3480
         TabIndex        =   46
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "Participant 1"
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
         Height          =   3135
         Left            =   240
         TabIndex        =   34
         Top             =   4560
         Width           =   4455
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1560
            TabIndex        =   39
            Top             =   360
            Width           =   2295
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   1560
            TabIndex        =   38
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   1560
            TabIndex        =   37
            Top             =   1560
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   3960
            TabIndex        =   36
            Top             =   960
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   3960
            TabIndex        =   35
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Tempus Sans ITC"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   45
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No:"
            BeginProperty Font 
               Name            =   "Tempus Sans ITC"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   44
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            BeginProperty Font 
               Name            =   "Tempus Sans ITC"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   43
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   2
            Left            =   720
            TabIndex        =   42
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   6
            Left            =   1080
            TabIndex        =   41
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   7
            Left            =   720
            TabIndex        =   40
            Top             =   1440
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Caption         =   "Participant 2"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   5160
         TabIndex        =   24
         Top             =   4560
         Width           =   4455
         Begin VB.TextBox Text8 
            Height          =   375
            Left            =   1560
            TabIndex        =   29
            Top             =   1560
            Width           =   2295
         End
         Begin VB.TextBox Text9 
            Height          =   375
            Left            =   1560
            TabIndex        =   28
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox Text10 
            Height          =   375
            Left            =   1560
            TabIndex        =   27
            Top             =   360
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   3960
            TabIndex        =   26
            Top             =   960
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00000000&
            Height          =   375
            Index           =   3
            Left            =   3960
            TabIndex        =   25
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            BeginProperty Font 
               Name            =   "Tempus Sans ITC"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No:"
            BeginProperty Font 
               Name            =   "Tempus Sans ITC"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   32
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Tempus Sans ITC"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   31
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   3
            Left            =   720
            TabIndex        =   30
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         Caption         =   "Participant 3"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   240
         TabIndex        =   14
         Top             =   7800
         Width           =   4455
         Begin VB.TextBox Text12 
            Height          =   375
            Left            =   1560
            TabIndex        =   19
            Top             =   1560
            Width           =   2295
         End
         Begin VB.TextBox Text13 
            Height          =   375
            Left            =   1560
            TabIndex        =   18
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox Text14 
            Height          =   375
            Left            =   1560
            TabIndex        =   17
            Top             =   360
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00000000&
            Height          =   375
            Index           =   4
            Left            =   3960
            TabIndex        =   16
            Top             =   960
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00000000&
            Height          =   375
            Index           =   5
            Left            =   3960
            TabIndex        =   15
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label14 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            BeginProperty Font 
               Name            =   "Tempus Sans ITC"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No:"
            BeginProperty Font 
               Name            =   "Tempus Sans ITC"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   22
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label16 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Tempus Sans ITC"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   4
            Left            =   720
            TabIndex        =   20
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Caption         =   "Participant 4"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   5160
         TabIndex        =   4
         Top             =   7800
         Width           =   4455
         Begin VB.TextBox Text16 
            Height          =   375
            Left            =   1560
            TabIndex        =   9
            Top             =   1560
            Width           =   2295
         End
         Begin VB.TextBox Text17 
            Height          =   375
            Left            =   1560
            TabIndex        =   8
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox Text18 
            Height          =   375
            Left            =   1560
            TabIndex        =   7
            Top             =   360
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00000000&
            Height          =   375
            Index           =   6
            Left            =   3960
            TabIndex        =   6
            Top             =   960
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00000000&
            Height          =   375
            Index           =   7
            Left            =   3960
            TabIndex        =   5
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label18 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            BeginProperty Font 
               Name            =   "Tempus Sans ITC"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label Label19 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No:"
            BeginProperty Font 
               Name            =   "Tempus Sans ITC"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label20 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Tempus Sans ITC"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   5
            Left            =   720
            TabIndex        =   10
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00000000&
         Caption         =   "Notify All"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10920
         TabIndex        =   2
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         Height          =   615
         Left            =   11640
         Top             =   8640
         Width           =   735
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   11640
         TabIndex        =   56
         Top             =   8640
         Width           =   735
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   495
         Left            =   9480
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Back"
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
         Left            =   9480
         TabIndex        =   55
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Participant Details"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   4440
         TabIndex        =   54
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Number:"
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
         Left            =   240
         TabIndex        =   53
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Event Name:"
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
         Left            =   240
         TabIndex        =   52
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Participants:"
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
         Left            =   240
         TabIndex        =   51
         Top             =   3720
         Width           =   3855
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   13560
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   0
         Left            =   2760
         TabIndex        =   50
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   49
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Submit"
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
         Left            =   9480
         TabIndex        =   48
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   495
         Left            =   9480
         Top             =   2280
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00404040&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   2415
   End
End
Attribute VB_Name = "Form11"
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
Public email1, email2, email3, email4 As String
Public sms1, sms2, sms3, sms4 As String
Dim no_participant As Integer
'Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rounder As Integer

Private Function ref()
Dim g As Integer

Text1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Combo2.ListIndex = 0
'Combo1.ListIndex = 0
Combo1.Text = ""

Text9.Text = ""
Text8.Text = ""
Text14.Text = ""
Text13.Text = ""
Text12.Text = ""
Text18.Text = ""
Text17.Text = ""
Text16.Text = ""
Text10.Text = ""

Check2.Value = 0

While g <> 8

Check1(g).Value = 0
g = g + 1

Wend
End Function




Private Sub Check2_Click()
Dim g As Integer
If Check2.Value = 1 Then

While g <> 8
Check1(g).Value = 1
g = g + 1
Wend

Else


While g <> 8
Check1(g).Value = 0
g = g + 1
Wend

End If
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

Private Sub Combo2_Click()
Dim i As Integer
i = 1
no_participant = Combo2.Text

If no_participant = 1 Then
Frame2.Visible = True
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
End If

If no_participant = 2 Then
Frame2.Visible = True
Frame3.Visible = True
Frame4.Visible = False
Frame5.Visible = False
End If


If no_participant = 3 Then
Frame2.Visible = True
Frame3.Visible = True
Frame4.Visible = True
Frame5.Visible = False
End If

If no_participant = 4 Then
Frame2.Visible = True
Frame3.Visible = True
Frame4.Visible = True
Frame5.Visible = True
End If



End Sub



Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
'Dim j As Integer
rounder = 1

Frame2.Visible = True
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False

no_participant = 1

Combo2.AddItem "1"
Combo2.AddItem "2"
Combo2.AddItem "3"
Combo2.AddItem "4"
Combo2.ListIndex = 0

Set Form3.rs = Nothing
Form3.rs.Open "select eventname from evtmain where festname='Credenz'", Form3.oconn, adOpenStatic, adLockOptimistic, adCmdText

While Not Form3.rs.EOF
Combo1.AddItem Form3.rs(0)
Form3.rs.MoveNext
Wend

Set Form3.rs = Nothing


End Sub

Private Sub Label13_Click()
Form11.Hide
Form13.Show

End Sub

Private Sub Label9_Click()
Dim z As Integer
Dim flag As Integer


         
If Text1.Text = "" Then
         flag = 1
         MsgBox "Please Enter the Receipt ID"

 'ElseIf Combo1.Text = "" Then
 
    '     flag = 1
        ' MsgBox "Please Enter the Event Name "
   ElseIf Text3.Text = "" Then
         flag = 1
         MsgBox "Enter the Participant Name"
ElseIf Text4.Text = "" And Text5.Text = "" Then
         flag = 1
         MsgBox "Please Enter a Contact Number or Email ID "
         End If
         

 
If flag = 0 Then

If no_participant >= 1 Then
   Form3.rs.Open "insert into participant values (" & Text1.Text & ",'" & Text3.Text & "','" & Text5.Text & "'," & Text4.Text & ",'" & Combo1.Text & "'," & rounder & ")", Form3.oconn, adOpenStatic, adLockOptimistic, adCmdText
   
   If Check1(0).Value = 1 And Text4.Text <> "" Then
   Form3.rs.Open "insert into mnotify values('" & Text3.Text & "','" & Combo1.Text & "'," & Text4.Text & ");", Form3.oconn, adOpenStatic, adLockOptimistic, adCmdText
    sms1 = Text4.Text
   ElseIf Check1(0).Value = 1 And Text4.Text = "" Then
   MsgBox "Enter Phone Number for Notification"
   End If
   

   If Check1(1).Value = 1 And Text5.Text <> "" Then
   Form3.rs.Open "insert into pnotify values('" & Text3.Text & "','" & Combo1.Text & "','" & Text5.Text & "');", Form3.oconn, adOpenStatic, adLockOptimistic, adCmdText
   email1 = Text5.Text
   ElseIf Check1(1).Value = 1 And Text5.Text = "" Then
   MsgBox "enter email id for notification"
   End If
   
      
   
   
   

If no_participant >= 2 Then
    Form3.rs.Open "insert into participant values ( " & Text1.Text & ",'" & Text10.Text & "','" & Text8.Text & "'," & Text9.Text & ",'" & Combo1.Text & "'," & rounder & ")", Form3.oconn, adOpenStatic, adLockOptimistic, adCmdText
    
    If Check1(2).Value = 1 And Text9.Text <> "" Then
   Form3.rs.Open "insert into mnotify values('" & Text10.Text & "','" & Combo1.Text & "'," & Text9.Text & ");", Form3.oconn, adOpenStatic, adLockOptimistic, adCmdText
   sms2 = Text9.Text
   ElseIf Check1(2).Value = 1 And Text9.Text = "" Then
   MsgBox "enter phone number for notification"
   End If
   
   If Check1(3).Value = 1 And Text8.Text <> "" Then
   Form3.rs.Open "insert into pnotify values('" & Text10.Text & "','" & Combo1.Text & "','" & Text8.Text & "');", Form3.oconn, adOpenStatic, adLockOptimistic, adCmdText
   email2 = Text8.Text
   
   ElseIf Check1(3).Value = 1 And Text9.Text = "" Then
   MsgBox "enter email id for notification"
   End If

If no_participant >= 3 Then
     Form3.rs.Open "insert into participant values (" & Text1.Text & ",'" & Text14.Text & "','" & Text12.Text & "'," & Text13.Text & ",'" & Combo1.Text & "'," & rounder & ")", Form3.oconn, adOpenStatic, adLockOptimistic, adCmdText
     
    If Check1(4).Value = 1 And Text13.Text <> "" Then
    Form3.rs.Open "insert into mnotify values('" & Text14.Text & "','" & Combo1.Text & "'," & Text13.Text & ");", Form3.oconn, adOpenStatic, adLockOptimistic, adCmdText
    sms3 = Text13.Text
    ElseIf Check1(4).Value = 1 And Text13.Text = "" Then
   
    MsgBox "enter phone number for notification"
    End If
    
    If Check1(5).Value = 1 And Text12.Text <> "" Then
    Form3.rs.Open "insert into pnotify values('" & Text14.Text & "','" & Combo1.Text & "','" & Text12.Text & "');", Form3.oconn, adOpenStatic, adLockOptimistic, adCmdText
    email3 = Text12.Text
    
    ElseIf Check1(5).Value = 1 And Text12.Text = "" Then
   
    MsgBox "enter email id for notification"
    End If
    

If no_participant = 4 Then
     Form3.rs.Open "insert into participant values (" & Text1.Text & ",'" & Text18.Text & "','" & Text16.Text & "'," & Text17.Text & ",'" & Combo1.Text & "'," & rounder & ")", Form3.oconn, adOpenStatic, adLockOptimistic, adCmdText
 If Check1(6).Value = 1 And Text17.Text <> "" Then
   Form3.rs.Open "insert into mnotify values('" & Text18.Text & "','" & Combo1.Text & "'," & Text17.Text & ");", Form3.oconn, adOpenStatic, adLockOptimistic, adCmdText
   sms4 = Text17.Text
  ElseIf Check1(6).Value = 1 And Text17.Text = "" Then
   MsgBox "enter phone number for notification"
    End If
    
   If Check1(7).Value = 1 And Text16.Text <> "" Then
   Form3.rs.Open "insert into pnotify values('" & Text18.Text & "','" & Combo1.Text & "','" & Text16.Text & "');", Form3.oconn, adOpenStatic, adLockOptimistic, adCmdText
    email4 = Text16.Text
    
    ElseIf Check1(7).Value = 1 And Text16.Text = "" Then
   MsgBox "enter email id for notification"
    End If

     
End If
End If
End If
End If
End If

z = ref()
MsgBox "values accepted"

Unload Me
Dialog3.Show

End Sub


