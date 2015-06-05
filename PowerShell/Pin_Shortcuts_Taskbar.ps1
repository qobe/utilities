# 4/20/2015
# Maxwell Kobe
# create short cut and pin it to taskbar


function Make-Shortcut($target)
{
    $wshell = New-Object -ComObject WScript.Shell
    $lnk = $wshell.CreateShortcut($target.Replace(".exe",".lnk"))
    $lnk.TargetPath = $target
    $lnk.Save()
}

function PinToTaskbar($target)
{
    $shell = New-Object -ComObject "Shell.Application"
    $folder = $shell.Namespace($target.Substring(0,$target.LastIndexOf("\")+1))
    $item = $target.split("\")
    $item = $folder.Parsename($item[$item.count - 1])
    $verb = $item.Verbs()| ? {$_.Name -eq "Pin to Task&bar"}
    if($verb)
    {
        $verb.DoIt()
    }
}

function PinToStartMenu($target)
{
    $shell = New-Object -ComObject "Shell.Application"
    $folder = $shell.Namespace($target.Substring(0,$target.LastIndexOf("\")+1))
    $item = $target.split("\")
    $item = $folder.Parsename($item[$item.count - 1])
    $verb = $item.Verbs()| ? {$_.Name -eq "Pin to Start Men&u"}
    if($verb)
    {
        $verb.DoIt()
    }
}

function main($text)
{    
    Make-Shortcut($text)
    #PinToStartMenu($text)
    PinToTaskbar($text)
    #Start-Process powershell -Verb runAs
}

# main("C:\Program Files (x86)\Orca\Orca.exe")