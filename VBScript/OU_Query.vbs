Set objSysInfo = CreateObject("ADSystemInfo")
strComputer = objSysInfo.ComputerName

Set objComputer = GetObject("LDAP://" & strComputer)

arrOUs = Split(objComputer.Parent, ",")
arrMainOU = Split(arrOUs(0), "=")

FOR EACH l IN arrMainOU
	Wscript.Echo l
	'Wscript.Echo objSysInfo.ComputerName
NEXT

arr2 = Split(objSysInfo.ComputerName, ",")
FOR EACH a IN arr2
	WScript.echo a
NEXT