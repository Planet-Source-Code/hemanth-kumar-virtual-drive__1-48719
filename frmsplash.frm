VERSION 5.00
Begin VB.Form frmabt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmsplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   " Warning "
      Height          =   1335
      Left            =   480
      TabIndex        =   4
      Top             =   3000
      Width           =   4455
      Begin VB.Label Label4 
         Caption         =   $"frmsplash.frx":0742
         Height          =   735
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   4095
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   480
      Picture         =   "frmsplash.frx":07E1
      ScaleHeight     =   735
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   360
      Width           =   615
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
      Left            =   1920
      MouseIcon       =   "frmsplash.frx":0C23
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
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
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   3735
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
      Left            =   1800
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Programming by"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "frmabt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Label5_Click()
    Dim URL As String
    URL = "http://www.onix.uni.cc"
    Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Sub
