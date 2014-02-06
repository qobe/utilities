# Add-type –AssemblyName System.Windows.Forms
#Set-Executionpolicy Unrestricted -force 

$dump = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
$output = [System.Windows.Forms.MessageBox]::Show("We are proceeding with next step.")

# $pass="FooBoo"|ConvertTo-SecureString -AsPlainText -Force
# $cred = New-Object System.Management.Automation.PsCredential("user@domain",$pass)
# New-PSDrive -Name "IDCShare" -PSProvider FileSystem -Root "\\wvu-ad\" -Credential $cred

# $objCS = Get-WmiObject -Class Win32_ComputerSystem
# $objCS.Domain
# $objCS.Name
$textToFile = "this" + " and " + "That"