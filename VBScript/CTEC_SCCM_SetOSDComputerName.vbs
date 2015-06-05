'************************************************************ 
Const constPrefix = "CTEC-"

set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

set colSMBIOS = objWMIService.ExecQuery("Select * from Win32_SystemEnclosure") 
For Each objSMBIOS in colSMBIOS
   If lcase(objSMBIOS.Manufacturer) = lcase("Apple Inc.") Then 
      strNewName = Right(objSMBIOS.SerialNumber, 14 - Len(constPrefix)) & "W"
   Else
      strNewName = Right(objSMBIOS.SerialNumber, 15 - Len(constPrefix))
   End If
next

strNewName = UCase(constPrefix & strNewName)

SET env = CreateObject("Microsoft.SMS.TSEnvironment") 
env("OSDCOMPUTERNAME") = strNewName
wscript.echo strNewName

'*********************************************************