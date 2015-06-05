function replace($str1, $str2)
{

}


function main()
{

    Import-Module "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
    CD "P01:"

    ForEach($col in (Get-CMDeviceCollection | Where {$_.Name -Like "CTEC*"}))
    {
        #$rule = Get-CMDeviceCollectionQueryMembershipRule -CollectionID $col.CollectionID -RuleName $col.CollectionRules
        If($col.CollectionRules.QueryExpression -match "(m|M)(a|A)(i|I)(n|N)/(c|C)(t|T)(e|E)(c|C)")
        {
            $ruleName = $col.CollectionRules.RuleName
            $rule = $col.CollectionRules.QueryExpression
            $newRule = $rule -Replace "(m|M)(a|A)(i|I)(n|N)/(c|C)(t|T)(e|E)(c|C)", "MAIN/IDESIGN"
            #uncomment when using
            #Remove-CMDeviceCollectionQueryMembershipRule -CollectionID $col.CollectionID -RuleName $ruleName
            #Add-CMDeviceCollectionQueryMembershipRule -CollectionID $col.CollectionID -QueryExpression $newRule -RuleName $ruleName
        }
    }

}

main