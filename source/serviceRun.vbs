'Written by KramWell.com - 12/JAN/18
'The purpose of this tool is to check a problematic service that may fail, log its time, and start the service when it does.

'put the service names in the bracket below followed by ; to seperate. service1;service2 etc.
fullService=Split("spooler;wuauserv",";")

'set log PATH
Dim LOGPATH
LOGPATH = "START.LOG"

'run loop
i = 0
Do While i = 0

	'here we split the services to check for multiple
	For each SERVICE in FullService
		'WScript.Echo(splitService)


		'###################################
		Set objShell = WScript.CreateObject("WScript.Shell")
		Set objExecObject = objShell.Exec("cmd /c SC QUERYEX " & SERVICE)
		strText = objExecObject.StdOut.ReadAll()
			
			If Instr(strText, "STOPPED") > 0 Then

				Call SaveToFile("STARTING '" & SERVICE & "' at: " & Now)
				'service stopped - need to start and log failure.
				strProgramPath = "cmd /c net start " & SERVICE & " >> " & LOGPATH
				
				SET objShell = CreateObject("Wscript.Shell")
				objShell.Run strProgramPath, ,True
				
				
				
			Else
				
				If NOT Instr(strText, "RUNNING") > 0 Then
					Call SaveToFile("ERROR '" & SERVICE & "' at: " & Now)
					Call SaveToFile(strText)				
				End If

			End If

		'###################################

	WScript.Sleep(5000) 'added here to give the script a chance to catch up.		
	Next 'next service	

WScript.Sleep(600000) 'Wait 10 minutes # (10000) = 10 seconds
Loop

	'save result to textfile
	Sub SaveToFile(ByVal x)
		Set fs = CreateObject("Scripting.FileSystemObject")
		'open file in append mode (8)
		set objLog = fs.OpenTextFile(LOGPATH, 8, true, 0)
		'write to file
		objLog.Write x & vbCrLf
		'close file
		objLog.close 
	End Sub