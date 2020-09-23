Attribute VB_Name = "Module1"
Public objfile As New FileSystemObject
Public drv As Drive
Public fol As Folder


Public Function lastdrvletter() As Integer

Dim d As String

    For Each Drive In objfile.Drives
        Set drv = objfile.GetDrive(Drive)
        d = drv.DriveLetter
    Next
    
    lastdrvletter = Asc(d)

End Function


Public Function MountVirtualDrive(strVirtualDrive, strPhysicPath) As Boolean
On Error GoTo err1
Shell "subst.exe " & strVirtualDrive & Chr(32) & strPhysicPath, vbHide
MountVirtualDrive = True
err1:
MountVirtualDrive = False
End Function

Public Function UnMountVirtualDrive(strVirtualDrive) As Boolean
On Error GoTo err1
Shell "subst.exe " & strVirtualDrive & " /d", vbHide
UnMountVirtualDrive = True
err1:
UnMountVirtualDrive = False
End Function
