VERSION 5.00
Begin VB.Form frmsplash 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5430
   LinkTopic       =   "Form2"
   ScaleHeight     =   3165
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   360
      Picture         =   "frmabt.frx":0000
      ScaleHeight     =   735
      ScaleWidth      =   615
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   4560
      Top             =   1680
   End
   Begin VB.Label Label5 
      Caption         =   "www.onix.uni.cc"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1800
      MouseIcon       =   "frmabt.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Programming by"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Hemanth Kumar E.K"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Virtual Drv 1.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Label5_Click()
    Dim URL As String
    URL = "http://www.onix.uni.cc"
    Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub Timer1_Timer()
    Load frmmain
    frmmain.Show
    Unload Me
End Sub
