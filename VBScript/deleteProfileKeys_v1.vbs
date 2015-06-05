'Written By: Maxwell Kobe
'Date modified: 12/23/2013
'Script that deletes profile registry keys of local domain accounts

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

'Subroutine determinse which profiles are domain accounts by matching SIDs and deletes the account based on their SID
SUB deleteProfileKeys()
	CONST HKLM = &H80000002
	DIM strComputer, strKeyPath, regexKey, arrSubKeys, subKey, objReg
	strComputer = "."
	SET objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
	strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	
	SET regKey = NEW RegExp
	regexKey.Pattern = "(^S-1-5-21-5159*.bak)"
	regexKey.IgnoreCase = TRUE
	regexKey.Global = FALSE
	
	
	objReg.EnumKey HKLM, strKeyPath, arrSubKeys
	
	FOR EACH subKey IN arrSubKeys
		IF regexKey.Test(subKey) THEN
			Wscript.Echo strKeyPath & "\" & subKey
			'objReg.DeleteKey HKLM, strKeyPath & "\" & subKey
		END IF
	NEXT
	
END SUB

SUB main()
	forceCScriptExecution
	deleteProfileKeys

END SUB
main()