   $query = "SELECT * FROM SMS_G_SYSTEM_COMPUTER_SYSTEM"
    Get-WmiObject -Query $query -Namespace "root\sms\site_P01" -ComputerName "wvu-adcm01-2.wvu-ad.wvu.edu"