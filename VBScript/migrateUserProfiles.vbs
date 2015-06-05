'Written by: Maxwell Kobe
'this script is to "migrate" user profiles from the thawspace (z) to c:\users to prepare for a larger thawspace install
'Last Modified: 10/29/2013
OPTION EXPLICIT

SUB Main()
	Dim colUserAccounts, objAccount
	Dim objShell, objWMI
	SET objShell = CreateObject("WScript.Shell")
	SET objWMI = GetObject("winmgmts:root\cimv2")

	'Turn off Data Igloo profile redirection
	objShell.Run "IGC.exe /AutoRedirectUP /d",0,TRUE
	objShell.Run "IGC.exe /RedirectRegKeyLocation /d",0,TRUE
	
	'Migrate profiles to C drive
	SET colUserAccounts = objWMI.ExecQuery("SELECT Name FROM Win32_NetworkLoginProfile")
	FOR EACH objAccount in colUserAccounts
		IF InStr(objAccount.name, "WVU-AD\") THEN	
			objShell.Run "IGC.exe /RedirectUP "& objAccount.Name &" /loc:C:\Users\",0,TRUE
			'WScript.Echo "IGC.exe /RedirectUP "& objAccount.Name &" /loc:C:\Users\"
		END IF
	NEXT
END SUB

Main()
