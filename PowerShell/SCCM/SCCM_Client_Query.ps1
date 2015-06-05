$query = "select * from SMS_R_System where SMS_R_System.ClientVersion = '5.00.7958.1303'"


$collection = get-wmiobject -query $query -namespace "root\sms\site_P01" -ComputerName "wvu-adcm01-2.wvu-ad.wvu.edu"
