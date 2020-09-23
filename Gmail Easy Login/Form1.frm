VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Gmail Login"
   ClientHeight    =   6825
   ClientLeft      =   3270
   ClientTop       =   4140
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   9990
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   5895
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   9735
      ExtentX         =   17171
      ExtentY         =   10398
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      TabIndex        =   1
      Text            =   "Password"
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "GmailID"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   4440
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded By ViCiOuS'

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub form_load()
On Error Resume Next
Me.Height = 1335
Me.Width = 4455
URL = "https://gmail.google.com"
wb.Navigate URL
If URL = "https://gmail.google.com" Then
MsgBox "Please Login Now Thank You For Using GmailEasyLogin", vbInformation, "GmailLogin By ViCiOuS"
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
Sleep "4000"
SendKeys ("{tab}")
SendKeys (Text1.Text)
SendKeys ("{tab}")
SendKeys (Text1.Text)
SendKeys ("{tab}")
SendKeys (Text2.Text)
SendKeys ("{Enter}")
Me.Width = 10110
Me.Height = 7365
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End
End Sub

Private Sub form_unload(cancel As Integer)
On Error Resume Next
Unload Me
End
End Sub


