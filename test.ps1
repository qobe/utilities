# Add-type –AssemblyName System.Windows.Forms
$dump = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
$output = [System.Windows.Forms.MessageBox]::Show("We are proceeding with next step.")