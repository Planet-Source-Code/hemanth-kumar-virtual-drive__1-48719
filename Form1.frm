VERSION 5.00
Begin VB.Form frmnewdrv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mount Drive"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   5160
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create &Drive"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   5520
      Width           =   4215
   End
   Begin VB.Frame Frame2 
      Caption         =   " Folder "
      Height          =   3855
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   4695
      Begin VB.DirListBox Dir1 
         Height          =   3015
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Drive "
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmnewdrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim drvl
Dim pth As String

Private Sub Command1_Click()
        
    If MsgBox("Are you sure you want to mount the drive ?", vbInformation + vbYesNo, "Vitual Drv 1.0") = vbYes Then
        a = lastdrvletter()
        a = a + 1
        MsgBox a
        MsgBox Chr(a) & ":"
        pth = Text1.Text
        If MountVirtualDrive(Chr(a) & ":", pth) = True Then
            MsgBox "Drive Sucessfully Mounted.", vbInformation, "Vitual Drv 1.0"
        End If
    End If
    
    Unload Me
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Dir1_Change()
Text1.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Text1.Text = Dir1.Path
End Sub


