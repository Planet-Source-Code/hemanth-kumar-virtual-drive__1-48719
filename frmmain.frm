VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmmain 
   BackColor       =   &H8000000C&
   Caption         =   " Virtual Drv 1.0"
   ClientHeight    =   7665
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8895
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList iml 
      Left            =   8040
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":113E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1592
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":19E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   741
      ButtonWidth     =   609
      Appearance      =   1
      ImageList       =   "iml"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "m"
            Object.ToolTipText     =   "Mount Drive"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "u"
            Object.ToolTipText     =   "Unmount Drive"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "v"
            Object.ToolTipText     =   "View all Drives"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stsbar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7290
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11695
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1147
            MinWidth        =   1147
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1147
            MinWidth        =   1147
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1147
            MinWidth        =   1147
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunewdrv 
         Caption         =   "&Mount Drive"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnudeldrv 
         Caption         =   "&Unmount Drive"
         Shortcut        =   ^U
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu muview 
      Caption         =   "&View"
      Begin VB.Menu mnuvalldrv 
         Caption         =   "&All Drives"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabt 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
    stsbar.Panels(1).Text = "Ready"
End Sub

Private Sub mnuabt_Click()
    frmabt.Show
End Sub

Private Sub mnudeldrv_Click()
    frmddrv.Show
End Sub

Private Sub mnunewdrv_Click()
    frmnewdrv.Show
End Sub

Private Sub mnuvalldrv_Click()
    frmmain.MousePointer = 11
    frmvdrv.Show
    frmmain.MousePointer = vbDefault
End Sub

Private Sub tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        Case "m"
            mnunewdrv_Click
        Case "u"
            mnudeldrv_Click
        Case "v"
            mnuvalldrv_Click
    End Select
End Sub
