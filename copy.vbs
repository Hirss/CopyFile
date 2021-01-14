Sub DoMyProc50Times()
Dim oFS,oFl 
   Dim x
   For x = 1 To 10
	Set oFSO = CreateObject("Scripting.FileSystemObject") 
	Set oFl = oFSO.GetFile("lol.txt") 
	oFl.Copy "lol.txt"&x,False
   Next
End Sub
DoMyProc50Times()