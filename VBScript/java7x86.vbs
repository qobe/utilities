OPTION EXPLICIT

CONST javaMSI = "jre1.7.0_25.msi"

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

'Function that returns a boolean based on if a process name fed in as a parameter currently exists
FUNCTION processExists(procName)
	DIM objWMI, colProc
	SET objWMI = GetObject("winmgmts:root\cimv2")
	SET colProc = objWMI.ExecQuery("select * from win32_process where name='" & procName & "'")
	IF (colProc.count <> 0) THEN
		processExists = TRUE
	ELSE
		processExists = FALSE
	END IF
END FUNCTION

'Function that kills all processes with the same name as the process name fed in as a paramter, waits .5 secs, and tests 
'to see if the process is still there.  It then requeries for the process name.  If one still exists, it sets colProcessList again
'and continues to loop until colProcessList.count is 0
FUNCTION killProcess(procName)
	DIM objWMI, objProcess, colProcessList
	SET objWMI = GetObject("winmgmts:root\cimv2")
	SET colProcessList = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & procName & "'")
	FOR EACH objProcess IN colProcessList
		objProcess.Terminate()
		WScript.Sleep 500
		SET colProcessList = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & procName & "'")
		IF colProcessList.Count = 0 THEN EXIT FOR
	NEXT
END FUNCTION

FUNCTION javaGUIDS()
	CONST HKLM = &H80000002
	DIM objReg, objShell, strComputer, strKeyPath, arrSubKeys
	strComputer = "."
	SET objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	
	SET objShell = CreateObject("WScript.Shell")
	strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"


	objReg.EnumKey HKLM, strKeyPath, arrSubKeys

	javaGUIDS = arrSubKeys
END FUNCTION

FUNCTION checkGUID(reg)
	DIM regEx
	SET regEx = NEW RegExp
	regEx.Pattern = "(^\{26A24AE4-039D-4CA4-87B4-2F8(32|64)17[0-9A-F]{5}\}$)|(^\{(32|64)48F0A8-6813-11D6-A77B-00B0D015[0-9A-F]{4}\}$)"
	regEx.IgnoreCase = TRUE
	regEx.Global = FALSE
	checkGUID = regEx.Test(reg)
END FUNCTION

FUNCTION installJava()
	DIM objShell, strInstallJava86, returnVal
	SET objShell = CreateObject("Wscript.Shell")
	
	strInstallJava86 = """msiexec"" /i""" & getScriptPath() & "\x86\" & javaMSI & """ /QUIET"

	IF processExists("iexplore.exe") THEN
		killProcess("iexplore.exe")
	END IF
	
	IF processExists("firefox.exe") THEN
		killProcess("firefox.exe")
	END IF
	returnVal = objShell.Run(strInstallJava86, 0, TRUE)
	installJava = returnVal
END FUNCTION

SUB copyDeploymentProperties()
	DIM objShell, objFSO, strDeploymentProperties, strDeploymentConfig, strWinDeploymentProperties, strWinDeploymentConfig, strDeploymentFolder
	SET objShell = CreateObject("WScript.Shell")
	SET objFSO = CreateObject("Scripting.FileSystemObject")
	
	strDeploymentProperties = "" & getScriptPath() & "\deployment.properties"
	strDeploymentConfig = "" & getScriptPath() & "\deployment.config"
	
	strWinDeploymentProperties = objShell.ExpandEnvironmentStrings("%WINDIR%\Sun\Java\Deployment\deployment.properties")
	strWinDeploymentConfig = objShell.ExpandEnvironmentStrings("%WINDIR%\Sun\Java\Deployment\deployment.config")
	
	strDeploymentFolder = objShell.ExpandEnvironmentStrings("%WINDIR%\Sun\Java\Deployment\")
	
	IF NOT objFSO.FileExists(strWinDeploymentProperties) THEN
		objFSO.CopyFile strDeploymentProperties, strDeploymentFolder
	END IF
	
	IF NOT objFSO.FileExists(strWinDeploymentConfig) THEN
		objFSO.CopyFile strDeploymentConfig, strDeploymentFolder
	END IF
END SUB

SUB main()
	forceCScriptExecution
	
	DIM objShell, regKey, strUninstall, returnVal
	SET objShell = CreateObject("WScript.Shell")

	FOR EACH regKey IN javaGUIDS
		IF (checkGUID(regKey) = TRUE) THEN
			strUninstall = """MsiExec.exe"" /x" & regKey & " /QUIET /NORESTART"
			'wscript.echo strUninstall
			objShell.Run strUninstall, 0, TRUE
		END IF
	NEXT
	
	returnVal = installJava()
	
	IF (returnVal = 0) THEN
		copyDeploymentProperties()
	END IF
	
	WScript.Quit(returnVal)
END SUB

main()