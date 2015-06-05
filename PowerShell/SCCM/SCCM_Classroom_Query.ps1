$querySelect = "SELECT SMS_R_SYSTEM.DistinguishedName, SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client "

$queryFrom = "FROM SMS_R_System "

$queryWhere = " WHERE SMS_R_System.DistinguishedName LIKE '%CLASSROOM%' AND NOT (SMS_R_System.DistinguishedName LIKE '%LSB%' OR SMS_R_System.DistinguishedName LIKE '%OU=116,OU=WDB%')"
$query = $querySelect + $queryFrom + $queryWhere
get-wmiobject -query $query -namespace "root\sms\site_P01" -ComputerName "wvu-adcm01-2.wvu-ad.wvu.edu"
