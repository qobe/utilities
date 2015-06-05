Function Set-Default-Playback-Device()
{
    $topkeypath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render'
    #$topkeypath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture'
    $topkey = Get-Item $topkeypath
    Foreach ($keyName in $topkey.GetSubkeyNames())
    {
        $key = $topkey.OpenSubKey($keyName, $True)
        $subkey = $key.OpenSubKey("Properties", $True)
        Foreach ($value in $subkey.GetValueNames())
        {
            If(([String]$subkey.GetValue($value)).Contains("Speaker"))
            {
                Write-Host $key.Name "`n" $value "`n" $key.GetValueNames()
                #$key.SetValue("DeviceState", $key.GetValue("DeviceState")+10000000)
            }
            
        }

    }
}

Function Disable-Audio-Devices()
{

}

Function main()
{
    Set-Default-Playback-Device
    #Write-Host "time: `t" ([timespan](Get-Date).ToShortTimeString()).TotalMilliseconds()
}

main