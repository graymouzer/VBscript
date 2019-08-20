'		Replication Checker
'		Jeffrey Kendrick
'		ACS Retail Solutions
'		Created 2/27/2002

'The purpose of this program is to aid the person doing replication
'tell where in the process he or she is and what log files have been generated.
'It assumes a network drive V mapped to \\NT3\D$ .


   Set wshShell = WScript.CreateObject("WScript.Shell")
 Set fsObj = CreateObject("Scripting.FileSystemObject")

 ' Location of Folder that contains files
 thisDir = "V:\program files\remacs\bccm\lab_nt3\logs"

 ' Find all files in the directory
 Set f = fsObj.GetFolder(thisDir)
 Set fc = f.Files
 Dim boo
 Dim goo
 Dim Headsup
 For Each fl in fc
	
 ' If a file was created today show it.
  If Date =< fl.DateCreated Then
	boo =  boo  &  VBCRLF & fl.name 
	goo = Mid(fl.name,1, 6)
		If goo = "GetCCM" Then
			GetM = "Y"
		End If

		If goo = "PutCCM" Then
		PutM = "Y"
		End If
		
	      	
  End If
 Next
 If GetM = "Y" Then Headsup = "GetMain was run today." End If
 If PutM = "Y" Then Headsup = Headsup & VBCRLF & "PutMain was run today." End If
 Msgbox Headsup & VBCRLF & "The following logs were generated today:" & VBCRLF & boo
	

