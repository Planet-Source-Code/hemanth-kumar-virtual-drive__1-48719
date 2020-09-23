VERSION 5.00
Begin VB.Form frmddrv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unmount Drive"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   Icon            =   "frmddrv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3045
   Begin VB.CommandButton Command1 
      Caption         =   "&Unmount"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   " Drive "
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      Begin VB.ComboBox c 
         Height          =   315
         ItemData        =   "frmddrv.frx":0442
         Left            =   840
         List            =   "frmddrv.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmddrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    If c.Text <> "" Then
        UnMountVirtualDrive c.Text & ":"
    End If
    
    Unload Me
    
End Sub

Private Sub Drive1_Change()
Text1 = Drive1.Drive
End Sub

Private Sub Form_Load()
    
    For Each Drive In objfile.Drives
        Set drv = objfile.GetDrive(Drive)
        c.AddItem drv.DriveLetter
    Next

End Sub
