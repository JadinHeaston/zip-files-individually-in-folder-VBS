'Created by Jadin Heaston sometime around 2019-2020.

Option Explicit


'Declaring handle variable.
Dim objFSO
Dim objFile
Dim objFolder

Dim basename
Dim tempZip
Dim zippingFilePath

'Creating handle to browse folders.
dim objShell
Set objShell = CreateObject("Shell.Application")
'Checking if a zip file was found
dim zipFileFound
'Creating the file system handle needed to manipulate files.
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Finding the file path for files to be zipped.
Set zippingFilePath = objShell.BrowseForFolder(0, "What folder is data being compressed?", 1, 0)
If Not (zippingFilePath Is Nothing) Then
	'Updating zippingFilePath to be a file path.
	zippingFilePath = zippingFilePath.Self.path + "\"
	'Setting the collection of files to be within "zippingFilePath".
	Set objFolder = objFSO.GetFolder(zippingFilePath).Files

	For Each objFile in objFolder
		If objFSO.GetExtensionName(objFile) = "zip" Then
		'Do nothing
		Else
			'Create filename and append zip extension.
			baseName = objFSO.GetBaseName(objFile)
			tempZip = zippingFilePath&baseName&".zip"
			If objFSO.FileExists(tempZip) Then
				'A zip file under the same name exists. Add " - ZIP" to allow easy deletion/manipulation later.
				tempZip = zippingFilePath&baseName&" - ZIP.zip"
			Else
				'If the file does not exist, create it.
				NewZip(tempZip)
			End If
			
			'Copying file to new zip
			CopyToZip zippingFilePath&baseName&"."&objFSO.GetExtensionName(objFile), tempZip
			'Delete original file.
			objFSO.DeleteFile(objFile)
		End If
	Next
	
	wscript.echo("FINISHED")
	Wscript.quit()

Else
	Wscript.quit()
End If




















'''''SUBSCRIPTS''''''
Private Sub NewZip(pathToZipFile)
	Dim zipFSO
	Set zipFSO = CreateObject("Scripting.FileSystemObject")
	Dim zipFile
	Set zipFile = zipFSO.CreateTextFile(pathToZipFile)
 
	zipFile.Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)
 
	zipFile.Close
	Set zipFSO = Nothing
	Set zipFile = Nothing
 
	WScript.Sleep 500
End Sub



Private Sub CopyToZip(fileToCopy, fileDest)
	'Creating shell object
	Dim objShell
	Set objShell = CreateObject("shell.application")
	'Creating "File System Object" Object.
	Dim zipFSO
	Set zipFSO = CreateObject("Scripting.FileSystemObject")
	
	Dim counter

	Dim zipFolder
	Set zipFolder = objShell.NameSpace(fileDest)

	counter = zipFolder.Items.Count + 1
	zipFolder.CopyHere(fileToCopy)
	
	While zipFolder.Items.Count < counter
		WScript.Sleep 100
	Wend

End Sub