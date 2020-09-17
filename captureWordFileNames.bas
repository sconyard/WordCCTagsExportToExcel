Sub getfilenamesinexcel()
Dim fso As Scripting.FileSystemObject
Dim fsofolder As Scripting.Folder
Dim fsofile As Scripting.File

Set fso = CreateObject("scripting.filesystemobject")
Set fsofolder = fso.GetFolder("Word Document Repository")
ce = 2

For Each fsofile In fsofolder.Files
Range("A" & ce).Value = fsofile.Path
ce = ce + 1
Next fsofile
End Sub
