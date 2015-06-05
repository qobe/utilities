function Get-CMMACFromCollection ($SiteServer, $Site, $CollectionID) {

    $query = "SELECT sys.SystemOuName, sys.distinguishedname, sys.NetbiosName, sys.MacAddresses, sys.SystemOUName, sys.IPAddresses,  sys.IPSubnets FROM SMS_R_SYSTEM sys, SMS_FullCollectionMembership fcm WHERE sys.ResourceID = fcm.ResourceID AND fcm.CollectionID = '$collectionid'"
    Get-WmiObject -Query $query -Namespace "root\sms\site_$site" -ComputerName $siteserver | Select-Object SystemOuName, NetbiosName, MACAddresses, IPAddresses, IPSubnets
}


function main()
{
    $outfile = "PC_Network_Info"

    $collection = Get-CMMACFromCollection -SiteServer "wvu-adcm01-2.wvu-ad.wvu.edu" -Site "P01" -CollectionID "P010009C"
    $colObjects = @()

    Foreach ($pc in $collection)
    {
        #Split OU name WVU-AD.WVU.EDU/MAIN/CTEC/COMPUTERS/CLASSROOMS/ALH/610 in ALH-610
        $classroomName = $pc.SystemOUName[$pc.SystemOUName.length - 1].split('/')
        $classroomName = $classroomName[$classroomName.length - 2] + '-' + $classroomName[$classroomName.length - 1]
        #create new objects to hold relevent data for easier sorting
        $obj = New-Object System.Object
        $obj | Add-Member -type NoteProperty -name Classroom -value $classroomName
        $obj | Add-Member -type NoteProperty -name Computer -value $pc.NetbiosName
        $obj | Add-Member -type NoteProperty -name MACAddresses -Value ($pc.MACAddresses -join ",")
        $obj | Add-Member -type NoteProperty -name IPAddresses -value ($pc.IPAddresses -join ",")
        $obj | Add-Member -type NoteProperty -name IPSubnets -value ($pc.IPSubnets -join ",")

        $colObjects += $obj
        
    }

    $colObjects |where {$_.Classroom -inotlike "CLASSROOMS-*"} | Sort-Object Classroom | Export-CSV $outfile".csv" -NoTypeInformation
    $colObjects |where {$_.Classroom -inotlike "CLASSROOMS-*"} | Sort-Object Classroom > $outfile".txt"

}

main