$string = dsquery computer -inactive 52 -limit 2000 | findstr /v /i "Disabled" 
$string = $string.Replace('"',"")
$reports = $string | ForEach-Object{

    $dn = $_.Split(',')

    New-Object PSObject -Property @{
        Name = $dn[0] -replace '^cn='
        ParentContainer = $dn[$dn.length..1] -notlike 'dc=*' -join ','
    } 


}

$body = "<b>Computers Listed below have been inactive in Active Directory for 1 year and not in Disabled OU</b><br />"

##Assembles and sends completion email with DL information##
$emailFrom = "sam.kaufman@wcaa.us"
$emailTo = "sam.kaufman@wcaa.us"
$subject = "Wayne County IT Inactive AD Computers"
$smtpServer = "SEXCH01.wcaa.local"
foreach ($report in $reports) {$body = $body + "<body>Server Name:<b>" + $report.name + "</b>          ParentContainer:" + $report.ParentContainer + "<br /></body>"} 


Send-MailMessage -To $emailTo -From $emailFrom -Subject $subject -BodyAsHtml -Body $body -SmtpServer $smtpServer
