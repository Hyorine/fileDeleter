
'lista létrehozása
dim list
Set list = CreateObject("System.Collections.ArrayList")
' fájlkezelő létrehozása
Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = "C:\"
'változok
 counter = 0
 exite = false

list.add objStartFolder


Do While exite = false
	On Error Resume Next	'hiba kezelés (try)
		Set objFolder = objFSO.GetFolder(list.Item(counter))
		
			For Each Subfolder in objFolder.SubFolders
				list.add Subfolder.Path
				
			Next
			' mappában lévő fájlok törlése
			Set colFiles = objFolder.Files
			For Each objFile in colFiles
				'Wscript.Echo objFile.Name
				objFSO.DeleteFile(objFile.Path)
			Next
		
		'Wscript.Echo objFolder
		counter = counter + 1
	On Error Goto 0  'hiba kezelés vége(catch)
		If counter >= list.Count  Then
			exite = true
		End If
Loop

'Set objFolder = objFSO.GetFolder(objStartFolder)
'Set colFolder = objFolder.SubFolders
'Set colFiles = objFolder.Files
'For Each objFile in colFiles
'    Wscript.Echo objFile.Name
'Next