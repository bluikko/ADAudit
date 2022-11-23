$mailto = "your@email.address.here"
$mailfrom = "sender@email.address.here"
$mailserver = "mail.server.address.here"

$properties = @(
    @{n='Timestamp';e={$_.TimeCreated}},
    @{n='Message1';e={($_.Message -split '\n')[0].Trim()}},
    @{n='ID';e={$_.Id}},
    @{n='p0';e={$_.Properties[0].Value}},
    @{n='p1';e={$_.Properties[1].Value}},
    @{n='p2';e={$_.Properties[2].Value}},
    @{n='p3';e={$_.Properties[3].Value}},
    @{n='p4';e={$_.Properties[4].Value}},
    @{n='p5';e={$_.Properties[5].Value}},
    @{n='p6';e={$_.Properties[6].Value}},
    @{n='p7';e={$_.Properties[7].Value}},
    @{n='p8';e={$_.Properties[8].Value}},
    @{n='p9';e={$_.Properties[9].Value}},
    @{n='p10';e={$_.Properties[10].Value}},
    @{n='p11';e={$_.Properties[11].Value}},
    @{n='p12';e={$_.Properties[12].Value}}
)
$bodycss = @"
<style>
h1, h2, h5, h6, th { text-align: center; font-family: Segoe UI; }
table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
th { background: #464646; color: #fff; max-width: 400px; padding: 5px 5px; }
td { font-size: 11px; padding: 5px 5px; color: #000; }
td.del { background: #bf8484; }
td.add { background: #84bf84; }
td.pass { background: #8484bf; }
td.enable { background: #a2bf84; }
td.disable { background: #bfa284; }
tr { background: #efefef; }
tr:nth-child(even) { background: #f3f3f3; }
tr:nth-child(odd) { background: #e4e4e4; }
</style>
"@
$evquery = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[(EventID=4728 or EventID=4729 or EventID=4727 or EventID=4730 or
      EventID=4732 or EventID=4733 or EventID=4731 or EventID=4734 or
      EventID=4726 or EventID=4725 or EventID=4722 or EventID=4724 or
      EventID=4720 or EventID=4738 or EventID=4743 or EventID=4740 or EventID=4767 or EventID=4781) and
      TimeCreated[timediff(@SystemTime) &lt;= 86400000]]]
    </Select>
  </Query>
</QueryList>
"@

# 4728: A member was added to a security-enabled global group
# 4729: A member was removed from a security-enabled global group
# 4727: A security-enabled global group was created
# 4730: A security-enabled global group was deleted
# 4732: A member was added to a security-enabled local group
# 4733: A member was removed from a security-enabled local group
# 4731: A security-enabled local group was created
# 4734: A security-enabled local group was deleted
# 4726: A user account was deleted
# 4725: A user account was disabled
# 4722: A user account was enabled
# 4724: An attempt was made to reset an account's password
# 4720: A user account was created
# 4738: A user account was changed
# 4743: A computer account was deleted
# 4740: A user account was locked out
# 4767: A user account was unlocked
$outlist = @()
$starttime = Get-Date

$machineinfo = Get-ComputerInfo

Get-WinEvent -FilterXml $evquery | Select $properties | 
ForEach-Object {
    $in = $_
    $subject = ''
    $by = ''
    $target = ''
    if ($_.p0 -eq '-') {
        $subject = $in.p1
        $by = $in.p5
    }
    elseif ($_.p0 -like 'CN=*') {
        $subject = $in.p0
        $target = $in.p2
        $by = $in.p6
    }
    else {
        $subject = $in.p0
        $by = $in.p4
    }
    $out = [PSCustomObject]@{'Timestamp'=$in.Timestamp; 'ID'=$in.ID; 'Event'=$in.Message1; 'Changed by'=$by; 'Subject'=$subject; 'Target'=$target}
    $outlist += $out
}
$outfile = "C:\Program Files\Scripts\Data\DCAudit-$(Get-Date -Format yyyy-MM-dd).csv"
$outlist | Export-Csv -NoTypeInformation -Path $outfile

$bodyraw = (Get-Content $outfile) -replace ',OU=Users,OU=Company,DC=domain,DC=tld','...' -replace ',OU=Company,DC=domain,DC=tld','...'
$body = ConvertFrom-Csv $bodyraw | ConvertTo-Html -Head $bodycss -Body "<h1>DC Audit report</h1>" -PostContent "<h6>Generated on $(Get-Date) in `
 $(((Get-Date) - $starttime).TotalMilliseconds / 1000) seconds<br/>on $($machineinfo.WindowsProductName) $($machineinfo.WindowsBuildLabEx)<br/>with `
 $(($machineinfo.OsHotFixes | Measure-Object).Count) HotFixes installed and booted on $($machineinfo.OsLastBootUpTime)</h6>" | Out-String

$body = $body -replace '<td>(?=[^<]*? was deleted)','<td class="del">'
$body = $body -replace '<td>(?=[^<]*? was created)','<td class="add">'
$body = $body -replace '<td>(?=[^<]*? password)','<td class="pass">'
$body = $body -replace '<td>(?=[^<]*? was (enabled|unlocked))','<td class="enable">'
$body = $body -replace '<td>(?=[^<]*? was (disabled|locked))','<td class="disable">'

Send-MailMessage -To $mailto -From "DC audit $(HOSTNAME) <$($mailfrom)>" -Subject "DC Audit report for day $(Get-Date -Format FileDate) on $(HOSTNAME)" `
 -SmtpServer $mailserver -Body $body -BodyAsHtml -Priority Low -Attachments $outfile
