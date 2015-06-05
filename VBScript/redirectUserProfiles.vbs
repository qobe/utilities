'Written by: Maxwell Kobe
'sets Faronics Data Igloo to redirect User Profiles to thawspace and copy's over any existing ones
'Last Modified: 10/29/2013
OPTION EXPLICIT

SUB Main()
	Dim colUserAccounts, objAccount
	Dim objShell, objWMI
	SET objShell = CreateObject("WScript.Shell")
	SET objWMI = GetObject("winmgmts:root\cimv2")

	'Turn on Data Igloo profile redirection
	objShell.Run "IGC.exe /RedirectRegKeyLocation /loc:z:\",0,TRUE
	objShell.Run "IGC.exe /AutoRedirectUP /loc:z:\",0,TRUE
	
	'Migrate profiles to C drive
	SET colUserAccounts = objWMI.ExecQuery("SELECT Name FROM Win32_NetworkLoginProfile")
	FOR EACH objAccount in colUserAccounts
		IF InStr(objAccount.name, "WVU-AD\") THEN	
			objShell.Run "IGC.exe /RedirectUP "& objAccount.Name &" /loc:z:\", 0, TRUE
			'WScript.Echo "IGC.exe /RedirectUP "& objAccount.Name &" /loc:z:\"
		END IF
	NEXT
END SUB

Main()