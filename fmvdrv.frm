VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmvdrv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View all Drives"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "fmvdrv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5741
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmvdrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim colh As ColumnHeader
Dim itm As ListItem


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Set colh = lv.ColumnHeaders.Add(, , "Drive Letter")
    Set colh = lv.ColumnHeaders.Add(, , "Type")

    lv.View = lvwReport
    
    For Each Drive In objfile.Drives
        
        Set drv = objfile.GetDrive(Drive)
        
        Select Case drv.DriveType
            Case 0: t = "Unknown"
            Case 1: t = "Removable"
            Case 2: t = "Fixed"
            Case 3: t = "Network"
            Case 4: t = "CD-ROM"
            Case 5: t = "RAM Disk"
        End Select

        If drv.IsReady = True Then
        
            If drv.VolumeName = "" Then
                Set itm = lv.ListItems.Add(, , "(" & drv.DriveLetter & ":)")
            Else
                Set itm = lv.ListItems.Add(, , drv.VolumeName & " (" & drv.DriveLetter & ":)")
            End If
            
            itm.SubItems(1) = t
        Else
            
            Set itm = lv.ListItems.Add(, , "(" & drv.DriveLetter & ":)")
            itm.SubItems(1) = t
            
        End If
        
    Next
    
End Sub
