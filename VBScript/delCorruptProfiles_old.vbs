'Modified By: Maxwell Kobe
'Last Modified: November 12, 2013
'Copyright (c) 2013; West Virginia University

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

FUNCTION getProfileSIDs()
	CONST HKLM = &H80000002
	DIM objReg, objShell, strComputer, strKeyPath, arrSubKeys, temp
	strComputer = "."
	SET objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	
	SET objShell = CreateObject("WScript.Shell")
	strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"

	objReg.EnumKey HKLM, strKeyPath, arrSubKeys

	getProfileSIDs = arrSubKeys
END FUNCTION

FUNCTION checkSID(reg)
	DIM regEx
	SET regEx = NEW RegExp
	'regEx.Pattern = "(^\{26A24AE4-039D-4CA4-87B4-2F8(32|64)17[0-9A-F]{5}\}$)|(^\{(32|64)48F0A8-6813-11D6-A77B-00B0D015[0-9A-F]{4}\}$)"
	regEx.Pattern = "(^S-1-5-21-51*)"
	regEx.IgnoreCase = TRUE
	regEx.Global = FALSE
	checkSID = regEx.Test(reg)
END FUNCTION
	
SUB main()
	forceCScriptExecution
	
	DIM objShell, regKey, strDelKey, returnVal
	SET objShell = CreateObject("WScript.Shell")

	FOR EACH regKey IN getProfileSIDs
		IF (checkSID(regKey) = TRUE) THEN
			strDelKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" & regKey
			'wscript.echo strDelKey
			objShell.RegDelete strDelKey
		END IF
	NEXT
	
END SUB
main()