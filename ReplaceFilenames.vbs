folderPath = InputBox("enter directory path with your files:","directory path")
If folderPath = "" Then
  Wscript.Quit
End If
Set objFso = CreateObject("Scripting.FileSystemObject")
Set Folder = objFSO.GetFolder(folderPath+"\")

stringToBeReplaced = InputBox("enter string you want to be replaced:","existing string")
stringToReplace = InputBox("enter new string:","new string")
isCounter = InputBox("enter number of files (enter 0 if you do not want to remove indexes from filenames):","remove indexes")

For Each File In Folder.Files
  sNewFile = File.Name
  
  If (isCounter <> "0" AND isCounter <> 0) Then
    counter = 1
    Do While counter <= CInt(isCounter)
      sNewFile = Replace(sNewFile,counter,"")
      counter = counter + 1
    Loop
    sNewFile = Replace(sNewFile,"0","")
  End If

  sNewFile = Replace(sNewFile,stringToBeReplaced,stringToReplace)
  
  If(sNewFile<>File.Name) Then
    File.Move(File.ParentFolder+"\"+sNewFile)
  end if 
Next