'Written By: Maxwell Kobe

OPTION EXPLICIT

'Subroutine that figures out if the script is running in cscript or not.  If not it calls itself again and runs in the cscript connotation.
'If not running in cscript, it will open a new window and hold the original window in memory until the script is complete.
SUB forceCScriptExecution
    DIM strArg, strTemp, strPath
	strPath = lCase(Left(WScript.FullName, Len(WScript.FullName) - 11))
    IF NOT LCase(Right(WScript.FullName, 11)) = "cscript.exe" OR NOT isAdmin() THEN
		IF NOT (strPath = lcase("C:\windows\system32")) THEN
			strPath = "C:\windows\system32\"
		END IF
        FOR EACH strArg IN WScript.Arguments
            IF InStr(strArg, " ") THEN strArg = """" & strArg & """"
            strTemp = strTemp & " " & strArg
        NEXT
        CreateObject("Shell.Application").ShellExecute """" & strPath & "cscript.exe""", """" & WScript.ScriptFullName & """ " & strTemp, "", "runas", 8
		WScript.Quit()
	END IF
END SUB

FUNCTION isAdmin()
	ON ERROR RESUME NEXT
	DIM objShell
	SET objShell = CreateObject("WScript.Shell")
	
	'This attempts to read the "NT Authority" user key in HKU.  If the script can successfully read
	'the key, then an return code of 0 is returned signifying the script has administrative rights
	objShell.RegRead("HKEY_USERS\s-1-5-19\")
	IF err.number <> 0 THEN
		isAdmin = FALSE
	ELSE
		isAdmin = TRUE
	END IF
END FUNCTION

FUNCTION getScriptPath()
	getScriptPath = CreateObject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder
END FUNCTION

FUNCTION getOUStructure()
	DIM arrOUs, objSysInfo, strOU, temp, i
	'get fully qualified name of current computer, split on commas
	SET objSysInfo = CreateObject("ADSystemInfo")
	arrOUs = Split(objSysInfo.ComputerName, ",")
	
	i = 0
	DO WHILE i < UBound(arrOUs)
		strOU = Split(arrOUs(i), "=")
		arrOUs(i) = strOU(1)
		i = i + 1
	LOOP
	
	getOUStructure = arrOUs
END FUNCTION

'write to file. use the VNC shortcut format
FUNCTION makeVNCShortcut(arrOUs)
	DIM outFile, objFSO, objFile
	
	outFile = "C:\" & arrOUs(2) & "-" & arrOUs(1) & ".vnc"
	SET objFSO = CreateObject("Scripting.FileSystemObject")
	SET objFile = objFSO.CreateTextFile(outFile, TRUE)
	objFile.Write "[Connection]" & vbCrLf
	objFile.Write "Host=" & arrOUs(0) & ".wvu-ad.wvu.edu:8173" & vbCrLf
	objFile.Write "[Options]" & vbCrLf
	objFile.Write "UseLocalCursor=1" & vbCrLf
	objFile.Write "UseDesktopResize=1" & vbCrLf
	objFile.Write "FullScreen=0" & vbCrLf
	objFile.Write "FullColour=1" & vbCrLf
	objFile.Write "LowColourLevel=1" & vbCrLf
	objFile.Write "PreferredEncoding=hextile" & vbCrLf
	objFile.Write "AutoSelect=1" & vbCrLf
	objFile.Write "Shared=0" & vbCrLf
	objFile.Write "SendPtrEvents=1" & vbCrLf
	objFile.Write "SendKeyEvents=1"& vbCrLf
	objFile.Write "SendCutText=1" & vbCrLf
	objFile.Write "AcceptCutText=1" & vbCrLf
	objFile.Write "DisableWinKeys=1" & vbCrLf
	objFile.Write "Emulate3=0" & vbCrLf
	objFile.Write "PointerEventInterval=0" & vbCrLf
	objFile.Write "Monitor=\\.\DISPLAY1" & vbCrLf
	objFile.Write "MenuKey=F8" & vbCrLf
	objFile.Write "AutoReconnect=1" & vbCrLf
	objFile.Close
	
	makeVNCShortcut = objFSO '.GetFileName(outFile)
END FUNCTION

FUNCTION copyShortCut(strSrcFile, strDestinationPath)
	DIM objFSO, objWNetwork, objShell
	
	SET objWNetwork = CreateObject("WScript.Network")
	SET objShell = CreateObject("WScript.Shell")
	SET objFSO = CreateObject("Scripting.FileSystemObject")
	
	'connect to network drive
	objWNetwork.MapNetworkDrive("A:", "\\wvu-ad.wvu.edu\data\Academic_Innovation\Secure\IDC", FALSE, "WVU-AD\mkobe", "password")
	'objShell.Run "NET USE a: \\wvu-ad.wvu.edu\data\Academic_Innovation\Secure\IDC"


END FUNCTION

FUNCTION main()
	forceCScriptExecution
	
	WScript.Echo makeVNCShortcut(getOUStructure)
	

END FUNCTION
main()